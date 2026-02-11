Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('enableBtn').onclick = enableAutoOpen;
        document.getElementById('disableBtn').onclick = disableAutoOpen;

        // Read and display current status
        updateStatus();
    }
});

function enableAutoOpen() {
    Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
    Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus();
        } else {
            alert('Failed to save setting: ' + result.error.message);
        }
    });
}

function disableAutoOpen() {
    Office.context.document.settings.remove('Office.AutoShowTaskpaneWithDocument');
    Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus();
        } else {
            alert('Failed to save setting: ' + result.error.message);
        }
    });
}

function updateStatus() {
    const isEnabled = Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument');
    const statusDiv = document.getElementById('status');

    if (isEnabled === true) {
        statusDiv.textContent = 'Status: Auto-open is ENABLED';
        statusDiv.className = 'status-enabled';
    } else {
        statusDiv.textContent = 'Status: Auto-open is DISABLED';
        statusDiv.className = 'status-disabled';
    }
}
