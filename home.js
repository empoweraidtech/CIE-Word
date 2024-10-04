let apiKey = '';

// Last updated: 2023-10-04 15:30:00 UTC
const lastUpdated = "2023-10-04 15:30:00 UTC";

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
        document.getElementById('copy-alternative').onclick = copyToClipboard;
        document.getElementById('last-updated').textContent = `Last updated: ${lastUpdated}`;
    }
});

function saveApiKey() {
    apiKey = document.getElementById('api-key').value;
    if (apiKey) {
        document.getElementById('api-key-input').classList.add('hidden');
        document.getElementById('review-section').classList.remove('hidden');
        document.getElementById('result').innerHTML = "<p><i class='fas fa-check-circle text-green-500 mr-2'></i>API Key saved. You can now use the review feature.</p>";
    } else {
        document.getElementById('result').innerHTML = "<p><i class='fas fa-exclamation-triangle text-yellow-500 mr-2'></i>Please enter a valid API Key.</p>";
    }
}

async function run() {
    if (!apiKey) {
        document.getElementById('result').innerHTML = "<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Please enter your API Key first.</p>";
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                document.getElementById('result').innerHTML = "<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>No text selected. Please select a paragraph to review.</p>";
                return;
            }
            
            // Show loading indicator
            document.getElementById('loader').classList.remove('hidden');
            document.getElementById('result').innerHTML = '';
            
            const reviewMode = document.getElementById('review-mode').value;
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            document.getElementById('loader').classList.add('hidden');
            
            // Display the review in the sidebar
            displayReview(review);
        });
    } catch (error) {
        document.getElementById('result').innerHTML = `<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`;
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
            summary: "An error occurred while reviewing the paragraph. Please check your API key and try again.",
            suggestedChanges: [],
            proposedAlternative: ""
        };
    }
}

function displayReview(review) {
    const summaryEl = document.getElementById('summary');
    const changesEl = document.getElementById('suggested-changes');
    const alternativeEl = document.getElementById('proposed-alternative');
    const copyButton = document.getElementById('copy-alternative');

    summaryEl.innerHTML = marked.parse(`### <i class="fas fa-info-circle text-blue-500 mr-2"></i>Summary\n\n${review.summary}`);
    
    changesEl.innerHTML = marked.parse(`### <i class="fas fa-edit text-yellow-500 mr-2"></i>Suggested Changes\n\n${review.suggestedChanges.map(change => `- ${change}`).join('\n')}`);
    
    alternativeEl.innerHTML = marked.parse(`### <i class="fas fa-file-alt text-green-500 mr-2"></i>Proposed Alternative\n\n${review.proposedAlternative}`);
    
    if (review.proposedAlternative) {
        copyButton.classList.remove('hidden');
    } else {
        copyButton.classList.add('hidden');
    }
}

function copyToClipboard() {
    const alternativeText = document.getElementById('proposed-alternative').textContent;
    navigator.clipboard.writeText(alternativeText).then(() => {
        const copyButton = document.getElementById('copy-alternative');
        copyButton.innerHTML = '<i class="fas fa-check mr-2"></i>Copied!';
        setTimeout(() => {
            copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
        }, 2000);
    });
}
