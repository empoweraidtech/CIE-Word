let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
        document.getElementById('copy-alternative').onclick = copyAlternative;
    }
});

function saveApiKey() {
    apiKey = document.getElementById('api-key').value;
    if (apiKey) {
        document.getElementById('api-key-input').classList.add('hidden');
        document.getElementById('review-section').classList.remove('hidden');
        document.getElementById('result').innerHTML = "API Key saved. You can now use the review feature.";
    } else {
        document.getElementById('result').innerHTML = "Please enter a valid API Key.";
    }
}

async function run() {
    if (!apiKey) {
        document.getElementById('result').innerHTML = "Please enter your API Key first.";
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                document.getElementById('result').innerHTML = "No text selected. Please select a paragraph to review.";
                return;
            }
            
            // Show loading indicator
            document.getElementById('loader').classList.remove('hidden');
            document.getElementById('result').classList.add('hidden');
            
            const reviewMode = document.getElementById('review-mode').value;
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            document.getElementById('loader').classList.add('hidden');
            document.getElementById('result').classList.remove('hidden');
            
            // Display the review in the sidebar
            displayReview(review);
        });
    } catch (error) {
        document.getElementById('result').innerHTML = `Error: ${error.message}`;
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
    
    Provide a response in the following JSON format, without enclosing it in triple backticks:
    {
        "summary": "A brief summary of the review",
        "explanation": "An in-depth explanation of how the paragraph aligns with the SCIFF framework",
        "changes": "Suggested changes to improve the paragraph",
        "alternative": "A proposed alternative paragraph incorporating the suggested changes"
    }
    
    Focus on the ${mode} aspect in your review.`;
    
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
            summary: "An error occurred while reviewing the paragraph.",
            explanation: "Please check your API key and try again.",
            changes: "",
            alternative: ""
        };
    }
}

function displayReview(review) {
    document.getElementById('summary').innerHTML = `<p class="font-bold"><i class="fas fa-info-circle mr-2"></i>Summary:</p><p>${review.summary}</p>`;
    document.getElementById('explanation').innerHTML = review.explanation;
    document.getElementById('changes').innerHTML = review.changes;
    document.getElementById('alternative').innerHTML = review.alternative;
}

function copyAlternative() {
    const alternativeText = document.getElementById('alternative').innerText;
    navigator.clipboard.writeText(alternativeText).then(() => {
        const copyButton = document.getElementById('copy-alternative');
        copyButton.innerHTML = '<i class="fas fa-check mr-2"></i>Copied!';
        setTimeout(() => {
            copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
        }, 2000);
    });
}
