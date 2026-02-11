Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('birthdayBtn').onclick = writeBirthdayMessage;
        document.getElementById('enableBtn').onclick = enableAutoOpen;
        document.getElementById('disableBtn').onclick = disableAutoOpen;

        // Read and display current status
        updateStatus();

        // Automatically write birthday message when add-in opens
        writeBirthdayMessage();
    }
});

function writeBirthdayMessage() {
    // Use Office.context.document.setSelectedDataAsync for simpler API
    const birthdayText = `🎉🎂 HAPPY BIRTHDAY! 🎂🎉

Wishing you a day filled with happiness and a year filled with joy!

May all your dreams and wishes come true on this special day.

Have a wonderful birthday celebration!

With love and best wishes,
Your Office Add-in`;

    Office.context.document.setSelectedDataAsync(
        birthdayText,
        { coercionType: Office.CoercionType.Text },
        (result) => {
            const statusDiv = document.getElementById('status');

            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Show success message
                const originalClass = statusDiv.className;
                const originalText = statusDiv.textContent;

                statusDiv.textContent = '✅ Birthday message written successfully!';
                statusDiv.className = 'status-enabled';

                // Restore original status after 3 seconds
                setTimeout(() => {
                    statusDiv.textContent = originalText;
                    statusDiv.className = originalClass;
                }, 3000);
            } else {
                console.error('Error writing birthday message:', result.error.message);
                statusDiv.textContent = '❌ Failed to write birthday message';
                statusDiv.className = 'status-disabled';
            }
        }
    );
}

function enableAutoOpen() {
    Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
    Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus();
        } else {
            console.error('Failed to save setting:', result.error.message);
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = '❌ Failed to save auto-open setting';
            statusDiv.className = 'status-disabled';
        }
    });
}

function disableAutoOpen() {
    Office.context.document.settings.remove('Office.AutoShowTaskpaneWithDocument');
    Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            updateStatus();
        } else {
            console.error('Failed to save setting:', result.error.message);
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = '❌ Failed to save auto-open setting';
            statusDiv.className = 'status-disabled';
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
