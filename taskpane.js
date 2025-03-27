'use strict';

// Обов'язкова ініціалізація Office Add-in
Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        // Запуск ініціалізації панелі завдань
        initializeTaskpane();
    }
});

// Ваш ключ API для pdf.co (ЗАМІНІТЬ НА ВАШ КЛЮЧ!)
// !!! ПОПЕРЕДЖЕННЯ ПРО БЕЗПЕКУ !!! Див. коментар вище.
const PDF_CO_API_KEY = "mykhailo.kovtun@streamtele.com_7GTwyNiN7t1y8qveIQs10I86ITvKhVTi1ROETcHq3WieNWnY9WCmyIKZ80YPLKjl";
const PDF_CO_BASE_URL = "https://api.pdf.co/v1";

// Отримання посилань на елементи DOM
let convertHtmlButton;
let convertEmlButton;
let statusDiv;
let resultDiv;

function initializeTaskpane() {
    convertHtmlButton = document.getElementById('convert-html');
    convertEmlButton = document.getElementById('convert-eml');
    statusDiv = document.getElementById('status');
    resultDiv = document.getElementById('result');

    // Додавання обробників подій до кнопок
    convertHtmlButton.onclick = () => handleConversion('html');
    convertEmlButton.onclick = () => handleConversion('eml');

    // Перевірка, чи вибрано лист
    if (!Office.context.mailbox.item) {
        showStatus("Будь ласка, відкрийте або виберіть лист для конвертації.", true);
        disableButtons();
    } else {
        enableButtons(); // Кнопки активні за замовчуванням, якщо лист вибрано
    }
}

// Функція для обробки запиту на конвертацію
async function handleConversion(method) {
    clearStatus();
    disableButtons();
    showStatus("Отримання даних листа...");

    try {
        let pdfUrl;
        if (method === 'html') {
            showStatus("Конвертація HTML тіла листа...");
            const htmlBody = await getEmailBodyAsHtml();
            if (htmlBody) {
                pdfUrl = await convertHtmlToPdf(htmlBody);
            } else {
                showStatus("Не вдалося отримати HTML тіло листа.", true);
            }
        } else if (method === 'eml') {
            showStatus("Конвертація всього листа (.eml)...");
            const emlData = await getEmailAsEmlBase64();
            if (emlData) {
                pdfUrl = await convertEmlToPdf(emlData);
            } else {
                 showStatus("Не вдалося отримати лист у форматі .eml. Можливо, ця функція не підтримується вашим клієнтом Outlook.", true);
            }
        }

        if (pdfUrl) {
            showStatus("Конвертація успішна!", false);
            showResultLink(pdfUrl);
        }
        // Якщо pdfUrl немає, помилка вже була показана у відповідній функції

    } catch (error) {
        console.error("Помилка конвертації:", error);
        showStatus(`Помилка: ${error.message || error}`, true);
    } finally {
        enableButtons(); // Завжди вмикаємо кнопки після завершення
    }
}

// Отримати тіло листа у форматі HTML
function getEmailBodyAsHtml() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, { asyncContext: "getHtmlBody" }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                console.error("Помилка getAsync(Html):", result.error);
                reject(new Error("Не вдалося отримати HTML тіло: " + result.error.message));
            }
        });
    });
}

// Отримати весь лист у форматі .eml (Base64)
async function getEmailAsEmlBase64() {
    // Перевірка, чи підтримується метод getAsFileAsync
    if (Office.context.mailbox.item.getAsFileAsync === undefined) {
         console.error("Функція getAsFileAsync не підтримується цим клієнтом Outlook.");
         return null; // Повертаємо null, щоб вказати на помилку
    }

    return new Promise(async (resolve, reject) => {
        Office.context.mailbox.item.getAsFileAsync({ asyncContext: "getEmlFile" }, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Файл отримано, тепер потрібно прочитати його вміст
                const fileContent = result.value;
                try {
                    const base64Content = await readFileContentAsBase64(fileContent);
                    resolve(base64Content);
                } catch (error) {
                    reject(error);
                } finally {
                   fileContent.closeAsync(); // Завжди закриваємо файл
                }
            } else {
                console.error("Помилка getAsFileAsync:", result.error);
                reject(new Error("Не вдалося отримати файл .eml: " + result.error.message));
            }
        });
    });
}

// Допоміжна функція для читання вмісту файлу (отриманого з getAsFileAsync) у Base64
function readFileContentAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (event) => {
            // Результат містить "data:application/octet-stream;base64,..."
            // Нам потрібна тільки частина після коми
            const base64String = event.target.result.split(',')[1];
            resolve(base64String);
        };

        reader.onerror = (event) => {
            console.error("Помилка читання файлу:", event.target.error);
            reject(new Error("Помилка читання вмісту файлу .eml"));
        };

        // Читаємо файл як Data URL, який містить Base64
        reader.readAsDataURL(file);
    });
}


// Виклик API pdf.co для конвертації HTML в PDF
async function convertHtmlToPdf(htmlContent) {
    const url = `${PDF_CO_BASE_URL}/pdf/convert/from/html`;
    const payload = {
        html: htmlContent,
        inline: false, // Отримати URL на PDF, а не вміст
        name: `Email_${Date.now()}.pdf` // Ім'я вихідного файлу
    };

    return await callPdfCoApi(url, payload);
}

// Виклик API pdf.co для конвертації EML (Base64) в PDF
async function convertEmlToPdf(base64EmlContent) {
    const url = `${PDF_CO_BASE_URL}/pdf/convert/from/email`;
    const payload = {
        body: base64EmlContent, // Передаємо Base64 контент тут
        inline: false,          // Отримати URL
        name: `Email_${Date.now()}.pdf`
        // Можна додати інші параметри pdf.co, якщо потрібно (наприклад, 'profiles')
    };

    return await callPdfCoApi(url, payload);
}

// Загальна функція для виклику API pdf.co
async function callPdfCoApi(apiUrl, payload) {
    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': PDF_CO_API_KEY
            },
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!response.ok || data.error) {
            throw new Error(data.message || `HTTP помилка ${response.status}`);
        }

        if (!data.url) {
            throw new Error("Відповідь від pdf.co не містить URL на PDF.");
        }

        console.log("PDF.co відповідь:", data);
        return data.url; // Повертаємо URL на згенерований PDF

    } catch (error) {
        console.error(`Помилка виклику API ${apiUrl}:`, error);
        showStatus(`Помилка pdf.co: ${error.message}`, true);
        return null; // Повертаємо null у разі помилки
    }
}


// Функції для управління інтерфейсом
function showStatus(message, isError = false) {
    statusDiv.innerText = message;
    statusDiv.style.color = isError ? 'red' : 'black';
    resultDiv.innerHTML = ''; // Очистити попередні результати
}

function clearStatus() {
    statusDiv.innerText = '';
    resultDiv.innerHTML = '';
}

function showResultLink(url) {
    resultDiv.innerHTML = `<p>PDF готовий: <a href="${url}" target="_blank">Завантажити PDF</a></p>`;
}

function disableButtons() {
    convertHtmlButton.disabled = true;
    convertEmlButton.disabled = true;
}

function enableButtons() {
    // Перевіряємо ще раз, чи є активний лист
    if (Office.context.mailbox.item) {
        convertHtmlButton.disabled = false;
         // Перевіряємо, чи підтримується getAsFileAsync перед тим, як вмикати кнопку
        if (Office.context.mailbox.item.getAsFileAsync !== undefined) {
             convertEmlButton.disabled = false;
        } else {
             convertEmlButton.disabled = true; // Залишаємо вимкненою, якщо метод не підтримується
             console.warn("Метод getAsFileAsync не підтримується, кнопка конвертації EML вимкнена.");
             // Можна додати повідомлення користувачеві про це
             if(!document.getElementById('eml-not-supported-msg')) {
                 const msg = document.createElement('p');
                 msg.id = 'eml-not-supported-msg';
                 msg.style.fontSize = '12px';
                 msg.style.color = '#666';
                 msg.style.marginTop = '-5px';
                 msg.innerText = '(Конвертація .eml може бути недоступна у вашій версії Outlook)';
                 convertEmlButton.parentNode.insertBefore(msg, convertEmlButton.nextSibling.nextSibling); // Вставити після опису
             }
        }
    } else {
        // Якщо листа немає, кнопки мають бути вимкнені
        disableButtons();
    }
}