Office.onReady(function(info) {
    // Office is ready
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("analyze-button").onclick = analyzeEmail;
    }
});

function analyzeEmail() {
    const analyzeAll = document.getElementById("analyze-all").checked;
    const autoAnalyze = document.getElementById("auto-analyze").checked;

    // Get the current item (email)
    const item = Office.context.mailbox.item;

    // Save user preferences
    Office.context.roamingSettings.set("autoAnalyze", autoAnalyze);
    Office.context.roamingSettings.set("analyzeAll", analyzeAll);
    Office.context.roamingSettings.saveAsync();

    // Display loading state
    const resultSection = document.getElementById("result-section");
    resultSection.style.display = "block";
    document.getElementById("analysis-results").innerHTML = "<p>Analyzing email...</p>";

    // Get email data
    item.body.getAsync(Office.CoercionType.Text, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;
            const emailSubject = item.subject;
            const sender = item.from.emailAddress;

            // Call your analysis service
            callAnalysisService(emailSubject, emailBody, sender, analyzeAll)
                .then(displayResults)
                .catch(handleError);
        } else {
            handleError(result.error);
        }
    });
}

function callAnalysisService(subject, body, sender, analyzeAll) {
    // Replace with your actual service endpoint
    const serviceUrl = "https://your-api.spotlight.ai/analyze";

    return fetch(serviceUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            subject: subject,
            body: body,
            sender: sender,
            analyzeAll: analyzeAll
        })
    })
        .then(response => {
            if (!response.ok) {
                throw new Error("Analysis service returned an error");
            }
            return response.json();
        });
}

function displayResults(results) {
    const resultsContainer = document.getElementById("analysis-results");

    // Create HTML for the results
    let html = "<div class='analysis-result'>";
    html += "<h3>Key Insights</h3>";
    html += "<ul>";

    // Display insights from your analysis
    for (const insight of results.insights) {
        html += `<li>${insight}</li>`;
    }

    html += "</ul>";

    if (results.actionItems && results.actionItems.length > 0) {
        html += "<h3>Action Items</h3>";
        html += "<ul>";
        for (const action of results.actionItems) {
            html += `<li>${action}</li>`;
        }
        html += "</ul>";
    }

    html += "</div>";

    resultsContainer.innerHTML = html;
}

function handleError(error) {
    document.getElementById("analysis-results").innerHTML =
        `<p class="error-message">Error analyzing email: ${error.message}</p>`;
}