let apiKey = '';

// Last updated: 2023-10-04 20:00:00 UTC
const lastUpdated = "2023-10-04 20:00:00 UTC";

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
        document.getElementById('copy-alternative').onclick = copyToClipboard;
        document.getElementById('last-updated').textContent = `Last updated: ${lastUpdated}`;
        setupCollapsibles();
    }
});

function setupCollapsibles() {
    var coll = document.getElementsByClassName("collapsible");
    for (var i = 0; i < coll.length; i++) {
        coll[i].addEventListener("click", function() {
            this.classList.toggle("active");
            var content = this.nextElementSibling;
            if (content.style.display === "block") {
                content.style.display = "none";
            } else {
                content.style.display = "block";
            }
        });
    }
}

function saveApiKey() {
    apiKey = document.getElementById('api-key').value;
    if (apiKey) {
        document.getElementById('api-key-input').classList.add('hidden');
        document.getElementById('review-section').classList.remove('hidden');
        setResult("<p><i class='fas fa-check-circle text-green-500 mr-2'></i>API Key saved. You can now use the review feature.</p>");
    } else {
        setResult("<p><i class='fas fa-exclamation-triangle text-yellow-500 mr-2'></i>Please enter a valid API Key.</p>");
    }
}

async function run() {
    if (!apiKey) {
        setResult("<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Please enter your API Key first.</p>");
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                setResult("<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>No text selected. Please select a paragraph to review.</p>");
                return;
            }
            
            // Prepare the result area
            setResult('');
            ensureResultElements();
            
            // Show loading indicator
            document.getElementById('loader').classList.remove('hidden');
            
            const reviewMode = document.getElementById('review-mode').value;
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            document.getElementById('loader').classList.add('hidden');
            
            // Display the review in the sidebar
            displayReview(review);
        });
    } catch (error) {
        setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
    }
}

function ensureResultElements() {
    const resultEl = document.getElementById('result');
    if (!resultEl) return;

    ['visualization', 'summary', 'suggested-changes', 'proposed-alternative'].forEach(id => {
        if (!document.getElementById(id)) {
            const div = document.createElement('div');
            div.id = id;
            div.className = 'mb-4';
            resultEl.appendChild(div);
        }
    });
}

function setResult(html) {
    const resultEl = document.getElementById('result');
    if (resultEl) {
        resultEl.innerHTML = html;
    }
}

async function reviewParagraph(text, mode) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    
    const prompt = `Review the following paragraph from a policy document against Ofsted's SCIFF framework for Outstanding:
    Paragraph: "${text}"
    
    Provide a review based on the mode "${mode}". Return your response in the following JSON format, using Markdown for formatting:
    
    {
      "visualization": {
        "ofstedOutstanding": {"score": "red|amber|green", "reason": "Brief reason"},
        "tristonePolicy": {"score": "red|amber|green", "reason": "Brief reason"},
        "readability": {"score": "red|amber|green", "reason": "Brief reason"}
      },
      "summary": "A brief summary of the review",
      "suggestedChanges": [
        "Change 1",
        "Change 2",
        "Change 3"
      ],
      "proposedAlternative": "A proposed alternative paragraph"
    }
    
    Ensure the JSON is not enclosed in any code blocks or quotation marks.`;
    
    try {
        const response = await axios.post(
            `${API_CONFIG.azureEndpoint}/openai/deployments/${API_CONFIG.deploymentName}/chat/completions?api-version=${API_CONFIG.apiVersion}`,
            {
                messages: [{ role: "user", content: prompt }],
                temperature: 0.5,
                max_tokens: 1000
            },
            {
                headers: {
                    'Content-Type': 'application/json',
                    'api-key': apiKey
                }
            }
        );
        return JSON.parse(response.data.choices[0].message.content);
    } catch (error) {
        console.error("Error calling OpenAI API:", error);
        return {
            visualization: {
                ofstedOutstanding: { score: "red", reason: "Error occurred" },
                tristonePolicy: { score: "red", reason: "Error occurred" },
                readability: { score: "red", reason: "Error occurred" }
            },
            summary: "An error occurred while reviewing the paragraph. Please check your API key and try again.",
            suggestedChanges: [],
            proposedAlternative: ""
        };
    }
}

function displayReview(review) {
    const visualizationEl = document.getElementById('visualization');
    const summaryEl = document.getElementById('summary');
    const changesEl = document.getElementById('suggested-changes');
    const alternativeEl = document.getElementById('proposed-alternative');
    const copyButton = document.getElementById('copy-alternative');

    if (visualizationEl) {
        visualizationEl.innerHTML = `
            <div class="flex justify-between">
                <div class="tooltip">
                    <span class="flag ${review.visualization.ofstedOutstanding.score}"></span> Ofsted Outstanding
                    <span class="tooltiptext">${review.visualization.ofstedOutstanding.reason}</span>
                </div>
                <div class="tooltip">
                    <span class="flag ${review.visualization.tristonePolicy.score}"></span> Tristone Policy
                    <span class="tooltiptext">${review.visualization.tristonePolicy.reason}</span>
                </div>
                <div class="tooltip">
                    <span class="flag ${review.visualization.readability.score}"></span> Readability
                    <span class="tooltiptext">${review.visualization.readability.reason}</span>
                </div>
            </div>
        `;
    }

    if (summaryEl) {
        summaryEl.innerHTML = marked.parse(review.summary);
    }
    
    if (changesEl) {
        changesEl.innerHTML = marked.parse(review.suggestedChanges.map(change => `- ${change}`).join('\n'));
    }
    
    if (alternativeEl) {
        alternativeEl.innerHTML = marked.parse(review.proposedAlternative);
    }
    
    if (copyButton) {
        copyButton.classList.remove('hidden');
    }
}

function copyToClipboard() {
    const alternativeEl = document.getElementById('proposed-alternative');
    if (alternativeEl) {
        const alternativeText = alternativeEl.textContent;
        navigator.clipboard.writeText(alternativeText).then(() => {
            const copyButton = document.getElementById('copy-alternative');
            if (copyButton) {
                copyButton.innerHTML = '<i class="fas fa-check mr-2"></i>Copied!';
                setTimeout(() => {
                    copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
                }, 2000);
            }
        });
    }
}
