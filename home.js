let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const saveKeyButton = document.getElementById('save-key');
        const runButton = document.getElementById('run');
        const copyAlternativeButton = document.getElementById('copy-alternative');
        const refreshCodeButton = document.getElementById('refresh-code');

        if (saveKeyButton) saveKeyButton.onclick = saveApiKey;
        if (runButton) runButton.onclick = run;
        if (copyAlternativeButton) copyAlternativeButton.onclick = copyAlternative;
        if (refreshCodeButton) refreshCodeButton.onclick = refreshCode;
    }
});

function refreshCode() {
    const scriptElement = document.querySelector('script[src="home.js"]');
    if (scriptElement) {
        const newScriptElement = document.createElement('script');
        newScriptElement.src = `home.js?v=${new Date().getTime()}`;
        scriptElement.parentNode.replaceChild(newScriptElement, scriptElement);
        
        // Show a message to the user
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            resultDiv.innerHTML = "Code refreshed. Please wait a moment and try your operation again.";
        }
        
        // Reload the page after a short delay
        setTimeout(() => {
            location.reload();
        }, 2000);
    }
}

function saveApiKey() {
    const apiKeyInput = document.getElementById('api-key');
    const apiKeyInputSection = document.getElementById('api-key-input');
    const reviewSection = document.getElementById('review-section');
    const resultDiv = document.getElementById('result');

    if (apiKeyInput) apiKey = apiKeyInput.value;
    
    if (apiKey) {
        if (apiKeyInputSection) apiKeyInputSection.classList.add('hidden');
        if (reviewSection) reviewSection.classList.remove('hidden');
        if (resultDiv) resultDiv.innerHTML = "API Key saved. You can now use the review feature.";
    } else {
        if (resultDiv) resultDiv.innerHTML = "Please enter a valid API Key.";
    }
}

async function run() {
    const resultDiv = document.getElementById('result');
    const loaderDiv = document.getElementById('loader');
    const reviewModeSelect = document.getElementById('review-mode');

    if (!apiKey) {
        if (resultDiv) resultDiv.innerHTML = "Please enter your API Key first.";
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                if (resultDiv) resultDiv.innerHTML = "No text selected. Please select a paragraph to review.";
                return;
            }
            
            // Show loading indicator
            if (loaderDiv) loaderDiv.classList.remove('hidden');
            if (resultDiv) resultDiv.classList.add('hidden');
            
            const reviewMode = reviewModeSelect ? reviewModeSelect.value : 'general';
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            if (loaderDiv) loaderDiv.classList.add('hidden');
            if (resultDiv) resultDiv.classList.remove('hidden');
            
            // Display the review in the sidebar
            displayReview(review);
        });
    } catch (error) {
        if (resultDiv) resultDiv.innerHTML = `Error: ${error.message}`;
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
    const summaryDiv = document.getElementById('summary');
    const explanationDiv = document.getElementById('explanation');
    const changesDiv = document.getElementById('changes');
    const alternativeDiv = document.getElementById('alternative');

    if (summaryDiv) summaryDiv.innerHTML = `<p class="font-bold"><i class="fas fa-info-circle mr-2"></i>Summary:</p><p>${review.summary}</p>`;
    if (explanationDiv) explanationDiv.innerHTML = review.explanation;
    if (changesDiv) changesDiv.innerHTML = review.changes;
    if (alternativeDiv) alternativeDiv.innerHTML = review.alternative;
}

function copyAlternative() {
    const alternativeDiv = document.getElementById('alternative');
    const copyButton = document.getElementById('copy-alternative');

    if (alternativeDiv && copyButton) {
        const alternativeText = alternativeDiv.innerText;
        navigator.clipboard.writeText(alternativeText).then(() => {
            copyButton.innerHTML = '<i class="fas fa-check mr-2"></i>Copied!';
            setTimeout(() => {
                copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
            }, 2000);
        });
    }
}
