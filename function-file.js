// --- START OF REFINED SIMPLIFIED function-file.js FOR TESTING ---

console.log("TEST V2: function-file.js - Script execution started."); // Найперший лог

// Функція, що викликається кнопкою
function convertToPDF_Handler(event) {
    console.log("TEST V2: convertToPDF_Handler function EXECUTED!"); // Чи викликається функція?

    // Спроба показати сповіщення як індикатор успіху
    try {
        console.log("TEST V2: Attempting to add notification...");
        Office.context.mailbox.item.notificationMessages.addAsync("testLaunchV2", {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Тест V2: Функція успішно запущена!",
            icon: "icon16",
            persistent: false
        }, function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("TEST V2: Failed to add notification: " + asyncResult.error.message);
            } else {
                console.log("TEST V2: Notification added successfully.");
            }
            // Завершуємо подію в будь-якому випадку
            console.log("TEST V2: Completing event inside notification callback.");
            event.completed();
        });
    } catch (e) {
        console.error("TEST V2: Error during notification attempt: " + e.message);
        console.log("TEST V2: Completing event inside catch block.");
        event.completed(); // Завершуємо навіть при помилці сповіщення
    }
}

// Реєстрація функції при готовності Office.js
Office.onReady((info) => {
    console.log("TEST V2: Office.onReady callback executed. Host: " + info.host + ", Platform: " + info.platform);

    // Стандартний спосіб реєстрації функції для ExecuteFunction
    try {
        if (!Office.actions) {
            Office.actions = {};
        }
        Office.actions.convertToPDF = convertToPDF_Handler;
        console.log("TEST V2: Function 'convertToPDF' registered under Office.actions.");
    } catch (e) {
         console.error("TEST V2: Error registering function: " + e.message);
    }

});

console.log("TEST V2: function-file.js - Script execution finished."); // Останній лог у файлі

// --- END OF REFINED SIMPLIFIED function-file.js FOR TESTING ---
