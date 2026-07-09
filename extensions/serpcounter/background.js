chrome.alarms.get('myAlarm', function (alarm) {
    if (alarm == null) {
        chrome.alarms.create('myAlarm', { periodInMinutes: 1 });
    }
});
chrome.alarms.onAlarm.addListener(function (alarm) {
    chrome.storage.local.get('enabled', function(result) {
        if (result.enabled) {
            console.log('Show notification');
        }
    });
});