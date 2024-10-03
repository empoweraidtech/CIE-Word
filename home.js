// home.js
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
        document.getElementById('result').innerText = "API Key saved. You can now use the review feature.";
    } else {
        document.getElementById('result').innerText = "Please enter a valid API Key.";
    }
}

async function run() {
    if (!apiKey) {
        document.getElementById('result').innerText = "Please enter your API Key first.";
        return;
    }
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();
            const selectedText = selection.text;
            if (!selectedText) {
                document.getElementById('result').innerText = "No text selected. Please select a paragraph to review.";
                return;
            }
            const review = await reviewParagraph(selectedText);
            
            // Display the formatted review in the sidebar
            displayFormattedReview(review);
        });
    } catch (error) {
        document.getElementById('result').innerText = `Error: ${error.message}`;
    }
}

function displayFormattedReview(review) {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = ''; // Clear previous content

    const sections = review.split(/\d+\.\s/).filter(Boolean);
    
    const overviewSection = document.createElement('div');
    overviewSection.innerHTML = `<h3>Overview</h3><p>${sections[0].trim()}</p>`;
    resultDiv.appendChild(overviewSection);

    if (sections.length > 1) {
        const improvementSection = document.createElement('div');
        improvementSection.innerHTML = `<h3>Areas for Improvement</h3><p>${sections[1].trim()}</p>`;
        resultDiv.appendChild(improvementSection);
    }

    if (sections.length > 2) {
        const suggestionsSection = document.createElement('div');
        suggestionsSection.innerHTML = `<h3>Suggestions</h3><p>${sections[2].trim()}</p>`;
        resultDiv.appendChild(suggestionsSection);
    }
}

async function reviewParagraph(text) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    const prompt = `
    Review the following paragraph from a policy document against Ofsted's SCIFF framework for Outstanding. Provide areas for improvement with explanations:
    Paragraph: "${text}"
    Please structure your response as follows:
    1. Brief overview of how the paragraph aligns with the SCIFF framework
    2. Areas for improvement (if any)
    3. Specific suggestions for enhancement
    `;
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
