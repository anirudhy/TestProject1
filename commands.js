// Headless add-in - runs automatically when FunctionFile loads
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        // Automatically write birthday message when commands.html loads
        writeBirthdayMessage();
    }
});

function writeBirthdayMessage() {
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
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('✅ Birthday message written successfully!');
            } else {
                console.error('❌ Error writing birthday message:', result.error);
            }
        }
    );
}

// Expose function for manifest to call
Office.actions = Office.actions || {};
Office.actions.writeBirthdayOnOpen = function (event) {
    writeBirthdayMessage();
    if (event && event.completed) {
        event.completed();
    }
};

