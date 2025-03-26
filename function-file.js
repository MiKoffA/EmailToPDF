// Глобальна змінна для відстеження ідентифікатора поточного сповіщення
let notificationKey = null;

// Функція для показу або оновлення сповіщення
function showNotification(key, message, type = Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, persistent = false) {
    // Якщо вже є сповіщення, спочатку видалимо його
    if (notificationKey) {
        Office.context.mailbox.item.notificationMessages.removeAsync(notificationKey, (removeResult) => {
            if (removeResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to remove previous notification: " + removeResult.error.message);
            }
            // Після видалення (або якщо його не було) додаємо нове
            addNotification(key, message, type, persistent);
        });
    } else {
        addNotification(key, message, type, persistent);
    }
}

// Допоміжна функція для додавання сповіщення
function addNotification(key, message, type, persistent) {
    Office.context.mailbox.item.notificationMessages.addAsync(key, {
        type: type,
        message: message,
        icon: "icon16", // Використовуємо ID іконки з маніфесту
        persistent: persistent
    }, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            notificationKey = key; // Зберігаємо ключ поточного сповіщення
        } else {
            console.error(`Failed to add notification [${key}]: ${asyncResult.error.message}`);
            notificationKey = null; // Скидаємо ключ, якщо не вдалося додати
        }
    });
}

// Функція для видалення сповіщення
function removeNotification(callback) {
    if (notificationKey) {
        const keyToRemove = notificationKey;
        notificationKey = null; // Одразу скидаємо ключ
        Office.context.mailbox.item.notificationMessages.removeAsync(keyToRemove, (removeResult) => {
            if (removeResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to remove notification: " + removeResult.error.message);
            }
            if (callback) callback();
        });
    } else {
        if (callback) callback();
    }
}


Office.onReady(() => {
    // Цей код виконується при завантаженні function-file.js,
    // але для ExecuteFunction він не має особливого сенсу, крім логування
    console.log("Office.js is ready for function execution.");
});

// Основна функція, що викликається кнопкою
function convertToPDF(event) {
    console.log("Starting convertToPDF function");
    showNotification("progress", "Розпочато конвертацію в PDF...");

    // Отримуємо тему листа
    Office.context.mailbox.item.subject.getAsync((subjectResult) => {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get subject: " + subjectResult.error.message);
            showNotification("error", "Помилка: Не вдалося отримати тему листа. " + subjectResult.error.message, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
            event.completed(); // Завершуємо подію, бо без теми/тіла нічого не вийде
            return;
        }

        const subject = subjectResult.value || "Email";
        console.log("Subject retrieved: " + subject);
        showNotification("progress", `Тема "${subject}" отримана. Отримання тіла листа...`);

        // Отримуємо тіло листа у форматі HTML
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get body: " + bodyResult.error.message);
                showNotification("error", "Помилка: Не вдалося отримати тіло листа. " + bodyResult.error.message, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
                event.completed();
                return;
            }

            const htmlContent = bodyResult.value;
            console.log("Body retrieved, length: " + htmlContent.length);
            if (!htmlContent || htmlContent.trim() === "") {
                 console.error("Email body is empty.");
                 showNotification("error", "Помилка: Тіло листа порожнє.", Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
                 event.completed();
                 return;
            }
            
            showNotification("progress", "Тіло листа отримано. Надсилання запиту до PDF.co...");

            // !!! ПОПЕРЕДЖЕННЯ ПРО БЕЗПЕКУ !!!
            // Ваш API-ключ зараз знаходиться в клієнтському коді.
            // Це НЕБЕЗПЕЧНО для реального використання. Розгляньте можливість
            // використання серверного посередника для викликів PDF.co.
            const apiKey = "mykhailo.kovtun@streamtele.com_6L8GbpqyvNuzWM3loNDN9Qnf7VfOYFr95Rpd74qWx75784HoCPseW5thlw6wsYC0";

            // Відправляємо запит до PDF.co
            fetch("https://api.pdf.co/v1/pdf/convert/from/html", {
                method: "POST",
                headers: {
                    "x-api-key": apiKey,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    html: htmlContent,
                    name: `${subject}.pdf`,
                    async: false // Виконуємо синхронно для простоти, але для великих файлів краще async: true
                })
            })
            .then(response => {
                console.log("PDF.co response status: " + response.status);
                if (!response.ok) {
                    // Якщо статус відповіді не в діапазоні 200-299, спробуємо отримати текст помилки
                     return response.text().then(text => {
                         throw new Error(`PDF.co HTTP error! Status: ${response.status}, Body: ${text}`);
                     });
                }
                return response.json(); // Якщо відповідь ОК, парсимо JSON
            })
            .then(data => {
                if (data.error) {
                    console.error("PDF.co error: " + data.message);
                    // Показуємо помилку від PDF.co
                    showNotification("error", `Помилка PDF.co: ${data.message}`, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
                } else if (data.url) {
                    console.log("PDF.co success, URL: " + data.url);
                    // Показуємо сповіщення з посиланням на PDF
                    showNotification("success", `PDF успішно створено! Посилання: ${data.url}`, Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, true); // persistent: true - щоб посилання не зникло
                    // window.open(data.url); // Замінено на сповіщення вище
                } else {
                     console.error("PDF.co unexpected response:", data);
                     showNotification("error", "Помилка: Несподівана відповідь від сервісу PDF.co.", Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
                }
            })
            .catch(error => {
                // Обробка помилок мережі або помилок при обробці відповіді
                console.error("Fetch or processing error: " + error.message);
                showNotification("error", `Помилка мережі або обробки: ${error.message}`, Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, true);
            })
            .finally(() => {
                console.log("Completing event");
                // Важливо завжди викликати event.completed(), щоб повідомити Outlook, що функція завершила роботу.
                // Не видаляємо сповіщення тут, якщо воно показує результат (успіх або помилку)
                event.completed();
            });
        });
    });
}

// Потрібно зареєструвати функцію в глобальній області видимості,
// щоб Office міг її знайти за іменем з маніфесту
if (typeof Office !== 'undefined') {
    Office.actions = Office.actions || {};
    Office.actions.convertToPDF = convertToPDF;
}
