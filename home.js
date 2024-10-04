let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
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
            document.getElementById('result').innerHTML = '';
            
            const reviewMode = document.getElementById('review-mode').value;
            const review = await reviewParagraph(selectedText, reviewMode);
            
            // Hide loading indicator
            document.getElementById('loader').classList.add('hidden');
            
            // Display the review in the sidebar
            document.getElementById('result').innerHTML = review;
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
    
    let prompt;
    switch (mode) {
        case 'general':
            prompt = `Review the following paragraph from a policy document against Ofsted's SCIFF framework for Outstanding:
            Paragraph: "${text}"
            Provide a general review and suggestions for improvement.`;
            break;
        case 'improvement':
            prompt = `Review the following paragraph from a policy document against Ofsted's SCIFF framework for Outstanding:
            Paragraph: "${text}"
            Focus on areas for improvement and provide specific suggestions.`;
            break;
        case 'alignment':
            prompt = `Review the following paragraph from a policy document against Ofsted's SCIFF framework for Outstanding:
            Paragraph: "${text}"
            Analyze how well this paragraph aligns with the SCIFF framework and suggest any necessary adjustments.`;
            break;
    }
    
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
        return response.data.choices[0].message.content;
    } catch (error) {
        console.error("Error calling OpenAI API:", error);
        return "An error occurred while reviewing the paragraph. Please check your API key and try again.";
    }
}
