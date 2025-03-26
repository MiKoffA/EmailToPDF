Office.onReady(() => {
    console.log("Office.js is ready");
});

function convertToPDF(event) {
    console.log("Starting convertToPDF function");
    Office.context.ui.messageParent(JSON.stringify({ message: "Starting convertToPDF function" }));

    // Отримуємо тему листа
    Office.context.mailbox.item.subject.getAsync(function(subjectResult) {
        if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to get subject: " + subjectResult.error.message);
            Office.context.ui.messageParent(JSON.stringify({ message: "Failed to get subject: " + subjectResult.error.message }));
            event.completed();
            return;
        }

        const subject = subjectResult.value || "Email";
        console.log("Subject retrieved: " + subject);
        Office.context.ui.messageParent(JSON.stringify({ message: "Subject retrieved: " + subject }));

        // Отримуємо тіло листа у форматі HTML
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
            if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to get body: " + bodyResult.error.message);
                Office.context.ui.messageParent(JSON.stringify({ message: "Failed to get body: " + bodyResult.error.message }));
                event.completed();
                return;
            }

            const htmlContent = bodyResult.value;
            console.log("Body retrieved, length: " + htmlContent.length);
            Office.context.ui.messageParent(JSON.stringify({ message: "Body retrieved, length: " + htmlContent.length }));

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
                Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co response status: " + response.status }));
                return response.json();
            })
            .then(data => {
                if (data.error) {
                    console.error("PDF.co error: " + data.message);
                    Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co error: " + data.message }));
                } else {
                    console.log("PDF.co success, URL: " + data.url);
                    Office.context.ui.messageParent(JSON.stringify({ message: "PDF.co success, URL: " + data.url }));
                    window.open(data.url);
                }
            })
            .catch(error => {
                console.error("Fetch error: " + error.message);
                Office.context.ui.messageParent(JSON.stringify({ message: "Fetch error: " + error.message }));
            })
            .finally(() => {
                console.log("Completing event");
                Office.context.ui.messageParent(JSON.stringify({ message: "Completing event" }));
                event.completed();
            });
        });
    });
}
