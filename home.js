let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office.js is ready");
        const saveKeyButton = document.getElementById('save-key');
        const runButton = document.getElementById('run');
        const refreshCodeButton = document.getElementById('refresh-code');

        if (saveKeyButton) {
            saveKeyButton.onclick = saveApiKey;
            console.log("Save key button initialized");
        } else {
            console.error("Save key button not found");
        }

        if (runButton) {
            runButton.onclick = run;
            console.log("Run button initialized");
        } else {
            console.error("Run button not found");
        }

        if (refreshCodeButton) {
            refreshCodeButton.onclick = refreshCode;
            console.log("Refresh code button initialized");
        } else {
            console.error("Refresh code button not found");
        }
    } else {
        console.error("This is not a Word document");
    }
});

function saveApiKey() {
    console.log("saveApiKey function called");
    const apiKeyInput = document.getElementById('api-key');
    const apiKeyInputSection = document.getElementById('api-key-input');
    const reviewSection = document.getElementById('review-section');
    const resultDiv = document.getElementById('result');

    if (apiKeyInput) {
        apiKey = apiKeyInput.value;
        console.log("API Key saved (length: " + apiKey.length + ")");
    } else {
        console.error("API Key input not found");
    }
    
    if (apiKey) {
        if (apiKeyInputSection) apiKeyInputSection.classList.add('hidden');
        if (reviewSection) reviewSection.classList.remove('hidden');
        if (resultDiv) resultDiv.innerHTML = "API Key saved. You can now use the review feature.";
    } else {
        if (resultDiv) resultDiv.innerHTML = "Please enter a valid API Key.";
    }
}

async function run() {
    console.log("Run function called");
    const resultDiv = document.getElementById('result');
    const loaderDiv = document.getElementById('loader');
    const reviewModeSelect = document.getElementById('review-mode');

    if (!apiKey) {
        console.error("No API Key found");
        if (resultDiv) resultDiv.innerHTML = "Please enter your API Key first.";
        return;
    }

    try {
        await Word.run(async (context) => {
            console.log("Word.run started");
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                console.error("No text selected");
                if (resultDiv) resultDiv.innerHTML = "No text selected. Please select a paragraph to review.";
                return;
            }
            console.log("Selected text: " + selectedText);
            
            // Show loading indicator
            if (loaderDiv) loaderDiv.classList.remove('hidden');
            if (resultDiv) {
                resultDiv.classList.remove('hidden');
                resultDiv.innerHTML = "Processing...";
            }
            
            const reviewMode = reviewModeSelect ? reviewModeSelect.value : 'general';
            console.log("Review mode: " + reviewMode);
            
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            if (loaderDiv) loaderDiv.classList.add('hidden');
            
            // Display the review in the sidebar
            displayReview(review);
        });
    } catch (error) {
        console.error("Error in run function:", error);
        if (resultDiv) resultDiv.innerHTML = `Error: ${error.message}`;
    }
}

async function reviewParagraph(text, mode) {
    console.log("reviewParagraph function called");
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
        console.log("Sending API request");
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
        console.log("API response received:", response.data);
        return JSON.parse(response.data.choices[0].message.content);
    } catch (error) {
        console.error("Error calling OpenAI API:", error);
        return {
            summary: "An error occurred while reviewing the paragraph.",
            explanation: "Error details: " + error.message,
            changes: "",
            alternative: ""
        };
    }
}

function displayReview(review) {
    console.log("displayReview function called");
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = `
            <h3>Summary:</h3>
            <p>${review.summary}</p>
            <h3>Explanation:</h3>
            <p>${review.explanation}</p>
            <h3>Suggested Changes:</h3>
            <p>${review.changes}</p>
            <h3>Alternative:</h3>
            <p>${review.alternative}</p>
        `;
    } else {
        console.error("Result div not found");
    }
}

function refreshCode() {
    console.log("refreshCode function called");
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
    } else {
        console.error("Script element not found");
    }
}
