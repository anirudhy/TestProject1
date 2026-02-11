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
    Word.run(async (context) => {
        // Get the current selection or end of document
        const selection = context.document.getSelection();

        // Create the birthday message
        const birthdayText = `
🎉🎂 HAPPY BIRTHDAY! 🎂🎉

Wishing you a day filled with happiness and a year filled with joy!

May all your dreams and wishes come true on this special day.

Have a wonderful birthday celebration!

With love and best wishes,
Your Office Add-in
`;

        // Insert the text at the selection
        selection.insertText(birthdayText, Word.InsertLocation.replace);

        // Format the text
        selection.font.size = 14;
        selection.font.color = '#0078d4';
        selection.font.bold = true;
        selection.paragraphs.first.alignment = Word.Alignment.centered;

        // Sync to apply changes
        await context.sync();

        // Show success message
        const statusDiv = document.getElementById('status');
        const originalClass = statusDiv.className;
        const originalText = statusDiv.textContent;

        statusDiv.textContent = '✅ Birthday message written successfully!';
        statusDiv.className = 'status-enabled';

        // Restore original status after 3 seconds
        setTimeout(() => {
            statusDiv.textContent = originalText;
            statusDiv.className = originalClass;
        }, 3000);

    }).catch((error) => {
        console.error('Error writing birthday message:', error);
        alert('Failed to write birthday message: ' + error.message);
    });
}

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
