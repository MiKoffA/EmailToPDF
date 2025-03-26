Office.onReady(() => {});

function convertToPDF(event) {
    Office.context.mailbox.item.subject.getAsync(function(subjectResult) {
        const subject = subjectResult.value || "Email";
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
            const htmlContent = bodyResult.value;
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
            .then(response => response.json())
            .then(data => window.open(data.url))
            .finally(() => event.completed());
        });
    });
}	