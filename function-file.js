// --- START OF SIMPLIFIED function-file.js FOR TESTING ---

// Проста тестова функція
function convertToPDF(event) {
    console.log("TEST: convertToPDF function started!"); // Лог для налагодження

    // Спроба показати базове сповіщення
    try {
        Office.context.mailbox.item.notificationMessages.addAsync("testLaunch", {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Тест: Функція запущена!",
            icon: "icon16", // Використовуємо ID іконки з маніфесту
            persistent: false
        }, function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("TEST: Failed to add notification: " + asyncResult.error.message);
            }
            // Обов'язково завершуємо подію
            console.log("TEST: Completing event.");
            event.completed();
        });
    } catch (e) {
        // Якщо навіть notificationMessages не доступний, просто логуємо і завершуємо
        console.error("TEST: Error during notification attempt: " + e.message);
        event.completed();
    }
}

// Реєстрація функції при завантаженні скрипта
Office.onReady(() => {
    console.log("TEST: Office.js is ready. Registering function...");
    // Реєструємо функцію для обробки дії кнопки
    window.convertToPDF = convertToPDF; // Зробимо її глобальною для надійності
    console.log("TEST: convertToPDF function registered globally.");
});

console.log("TEST: Simplified function-file.js loaded."); // Лог для перевірки завантаження файлу

// --- END OF SIMPLIFIED function-file.js FOR TESTING ---
