Office.onReady(() => {
    // Реєструємо функцію для використання в надбудові
    Office.actions.associate("convertToPDF", convertToPDF);
    
    console.log("Office.js is ready");
    try {
        Office.context.ui.messageParent(JSON.stringify({ message: "Office.js is ready" }));
    } catch (e) {
        console.error("Error in Office.onReady: " + e.message);
    }
});

function convertToPDF(event) {
    console.log("Starting convertToPDF function");
    try {
        Office.context.ui.messageParent(JSON.stringify({ message: "Starting convertToPDF function" }));
    } catch (e) {
        console.error("Error in messageParent: " + e.message);
    }

    // Отримуємо тему листа
    Office.context.mailbox.item.subject.getAsync(function(subjectResult) {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get subject: " + subjectResult.error.message);
            try {
                Office.context.ui.messageParent(JSON.stringify({ message: "Failed to get subject: " + subjectResult.error.message }));
            } catch (e) {
                console.error("Error in messageParent: " + e.message);
            }
            event.completed();
            return;
        }

        const subject = subjectResult.value || "Email";
        console.log("Subject retrieved: " + subject);
        try {
            Office.context.ui.messageParent(JSON.stringify({ message: "Subject retrieved: " + subject }));
        } catch (e) {
            console.error("Error in messageParent: " + e.message);
        }

        // Отримуємо тіло листа у форматі HTML
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get body: " + bodyResult.error.message);
                try {
                    Office.context.ui.messageParent(JSON.stringify({ message: "Failed to get body: " + bodyResult.error.message }));
                } catch (e) {
                    console.error("Error in messageParent: " + e.message);
                }
                event.completed();
                return;
            }

            const htmlContent = bodyResult.value;
            console.log("Body retrieved, length: " + htmlContent.length);
            try {
                Office.context.ui.messageParent(JSON.stringify({ message: "Body retrieved, length: " + htmlContent.length }));
            } catch (e) {
                console.error("Error in messageParent: " + e.message);
            }

            // Відправляємо запит до PDF.co
            fetch("https://api.pdf.co/v1/pdf/convert/from/html", {
                method: "POST",
                headers: {
                    "x-api-key": "mykhailo.kovtun@streamtele.com_6L8GbpqyvNuzWM3loNDN9Qnf7VfOYFr95Rpd74qWx75784HoCPseW5thlw6wsYC0",
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    html: htmlContent,
                    name: `${subject}.pdf`
                })
            })
            .then(response => {
                console.log("PDF.co response status: " + response.status);
                try {
                    Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co response status: " + response.status }));
                } catch (e) {
                    console.error("Error in messageParent: " + e.message);
                }
                return response.json();
            })
            .then(data => {
                if (data.error) {
                    console.error("PDF.co error: " + data.message);
                    try {
                        Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co error: " + data.message }));
                    } catch (e) {
                        console.error("Error in messageParent: " + e.message);
                    }
                } else {
                    console.log("PDF.co success, URL: " + data.url);
                    try {
                        Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co success, URL: " + data.url }));
                    } catch (e) {
                        console.error("Error in messageParent: " + e.message);
                    }
                    window.open(data.url);
                }
            })
            .catch(error => {
                console.error("Fetch error: " + error.message);
                try {
                    Office.context.ui.messageParent(JSON.stringify({ message: "Fetch error: " + error.message }));
                } catch (e) {
                    console.error("Error in messageParent: " + e.message);
                }
            })
            .finally(() => {
                console.log("Completing event");
                try {
                    Office.context.ui.messageParent(JSON.stringify({ message: "Completing event" }));
                } catch (e) {
                    console.error("Error in messageParent: " + e.message);
                }
                event.completed();
            });
        });
    });
}
