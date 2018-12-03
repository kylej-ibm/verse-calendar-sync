function saveOptions() {
  const calendarId = document.getElementById('calendar-id').value;
  const verseBaseUrl = document.getElementById('verse-base-url').value;
  let daysToSync = document.getElementById('days-to-sync').value;
  if (daysToSync < 1) daysToSync = 1;
  if (daysToSync > 90) daysToSync = 90;

  chrome.storage.sync.set(
    {
      calendarId,
      verseBaseUrl,
      daysToSync
    },
    function() {
      const status = document.getElementById('status');
      const oldStatus = status.textContent;
      status.textContent = 'Options saved.';
      setTimeout(() => {
        status.textContent = oldStatus;
      }, 750);
    }
  );
}

function makeLastSyncDate(date) {
  const lastSync = new Date(date);
  let minutes = String(lastSync.getMinutes());
  if (minutes.length === 1) minutes = '0' + minutes;
  let seconds = String(lastSync.getSeconds());
  if (seconds.length === 1) seconds = '0' + seconds;
  return lastSync.getHours() + ':' + minutes + ':' + seconds;
}

function restoreOptions() {
  chrome.storage.sync.get(
    {
      calendarId: defaults.VERSE_BASE_URL_DEFAULT,
      verseBaseUrl: defaults.VERSE_BASE_URL_DEFAULT,
      daysToSync: defaults.DAYS_TO_SYNC_DEFAULT,
      lastSync: null
    },
    function(items) {
      document.getElementById('calendar-id').value = items.calendarId;
      document.getElementById('verse-base-url').value = items.verseBaseUrl;
      document.getElementById('days-to-sync').value = items.daysToSync;
      document.getElementById('status').textContent =
        items.lastSync && 'Last sync: ' + makeLastSyncDate(items.lastSync);
      setDisabledForOptions(false);
    }
  );
}

function syncNow() {
  sync();
  document.getElementById('status').textContent = 'Syncing...';
}

function updateSyncResult() {
  chrome.storage.sync.get({ lastSync: null, lastError: null }, function(items) {
    document.getElementById('status').textContent = items.lastSync && 'Last sync: ' + makeLastSyncDate(items.lastSync);
    if (items.lastError) {
      document.getElementById('sync-error').textContent = 'Sync error: ' + items.lastError;
    } else {
      document.getElementById('sync-error').textContent = '';
    }
  });
}

async function authorize() {
  document.getElementById('auth-error').textContent = '';
  const btn = document.getElementById('grant-access');
  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Loading...';

  try {
    await utils.getAuthToken(true);
  } catch (error) {
    document.getElementById('auth-error').textContent = 'Auth error: ' + error.message;
  }

  setTimeout(() => {
    btn.disabled = false;
    btn.textContent = label;
    initialize();
  }, 250);
}

async function initialize() {
  setDisabledForOptions(true);
  const hasAccess = await checkGoogleCalendarAccess();
  if (hasAccess) {
    hideGrantAccessPanel();
    restoreOptions();
  } else {
    showGrantAccessPanel();
  }
}

async function checkGoogleCalendarAccess() {
  let hasAccess = false;
  try {
    const token = await utils.getAuthToken();
    hasAccess = true;
    chrome.identity.removeCachedAuthToken({ token });
  } catch (error) {
    console.debug(error);
  }
  return hasAccess;
}

function hideGrantAccessPanel() {
  document.getElementById('missing-permission').style.display = 'none';
  document.getElementById('options').style.display = 'block';
}

function showGrantAccessPanel() {
  document.getElementById('missing-permission').style.display = 'block';
  document.getElementById('options').style.display = 'none';
}

function setDisabledForOptions(disabled) {
  document.getElementById('calendar-id').disabled = disabled;
  document.getElementById('verse-base-url').disabled = disabled;
  document.getElementById('days-to-sync').disabled = disabled;
  document.getElementById('save').disabled = disabled;
}

const verseApi = new VerseApi();
const googleApi = new GoogleApi();
thing = googleApi.load();

  sync();
  setInterval(sync, 5 * 60 * 1000); // sync every 5 min

  async function sync() {
    chrome.storage.sync.set({ lastError: null });
    try {
      await executeSync();
    } catch (error) {
      console.warn(error.message);
      chrome.storage.sync.set({ lastError: error.message });
    } finally {
      chrome.storage.sync.set({ lastSync: Date.now() });
    }
  }

  async function executeSync() {
    const config = await getConfig();
    if (!config.calendarId) {
      console.warn('No calendar ID set in extension options. Select a calendar to enable sync.');
      return;
    }

    let authToken;
    try {
      authToken = await utils.getAuthToken();
    } catch (error) {
      console.warn('Access to Google Calendar not allowed. Open extension options to allow access.');
      return;
    }

    googleApi.authToken = authToken;
    verseApi.baseUrl = config.verseBaseUrl;

    console.debug('Start sync...');
    const start = makeDate(0); // now
    const until = makeDate(config.daysToSync);
    console.debug('Sync events from ' + start + ' to ' + until);

    let verseEntries;
    try {
      verseEntries = await verseApi.fetchCalendarEntries(start, until);
    } catch (error) {
      console.debug(error);
      throw new Error('Unable to fetch Verse entries. Ensure that you are logged into Verse.');
    }
    console.debug(verseEntries.length + ' events found in Verse');

    let googleEntires;
    try {
      googleEntires = await googleApi.fetchCalendarEntries(config.calendarId, start, until);
    } catch (error) {
      console.debug(error);
      chrome.identity.removeCachedAuthToken({ token: authToken });
      throw new Error('Unable to fetch Google Calendar entries. Renewing auth token for next sync.');
    }

    console.debug(googleEntires.length + ' events found in Google');

    const convertedEntries = verseEntries.map(convertToGoogleEntry);

    console.debug('Importing events to Google');
    await googleApi.createCalendarEntries(config.calendarId, convertedEntries);

    const verseEntryIds = verseEntries.map(e => e.syncId);
    const orphanedGoogleEntryIds = googleEntires.reduce((acc, entry) => {
      if (verseEntryIds.indexOf(entry.iCalUID) === -1 && entry.iCalUID.indexOf(verseApi.uidPrefix) === 0) {
        acc.push(entry.id);
      }
      return acc;
    }, []);

    if (orphanedGoogleEntryIds.length > 0) {
      console.debug('Remove ' + orphanedGoogleEntryIds.length + ' events from Google');
      await googleApi.deleteCalendarEntries(config.calendarId, orphanedGoogleEntryIds);
    }

    console.debug('Sync done.');
  }

  function makeDate(plusDays) {
    const date = new Date(new Date().getTime() + plusDays * 24 * 60 * 60 * 1000);
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function getConfig() {
    return new Promise(resolve =>
      chrome.storage.sync.get(
        { calendarId: defaults.CALENDAR_ID, verseBaseUrl: defaults.VERSE_BASE_URL_DEFAULT, daysToSync: defaults.DAYS_TO_SYNC_DEFAULT },
        resolve
      )
    );
  }

  function convertToGoogleEntry(entry) {
    const startDateTime = luxon.DateTime.fromFormat(entry.$StartDateTime, 'yyyyMMddTHHmmss,00ZZZ');
    const endDateTime = startDateTime.plus({ seconds: Number(entry.$Duration) });
    const start = {};
    const end = {};
    let description;

    if (entry.$AppointmentType === VerseApi.ENTRY_TYPE_ALL_DAY) {
      start.date = startDateTime.toISODate();
      end.date = endDateTime.toISODate();
    } else {
      start.dateTime = startDateTime.toISO();
      end.dateTime = endDateTime.toISO();
    }

    if (entry.$OnlineMeeting) {
      description = 'Online Meeting: ' + entry.$OnlineMeeting;
      if (entry.$OnlineMeetingCode) {
        description += '\nCode:' + entry.$OnlineMeetingCode;
      }
    }

    return {
      iCalUID: entry.syncId,
      summary: entry.$Subject,
      start,
      end,
      location: entry.$Location,
      description,
      sequence: Math.floor(Date.now() / 1000)
    };
  }

document.getElementById('save').addEventListener('click', saveOptions);
document.getElementById('sync-now').addEventListener('click', syncNow);
document.getElementById('grant-access').addEventListener('click', authorize);
chrome.storage.onChanged.addListener(updateSyncResult);

initialize();