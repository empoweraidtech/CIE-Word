// home.js
let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
        document.getElementById('review-document').onclick = reviewDocument; // Add this line
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
            
            // Insert the review below the selected paragraph
            selection.insertParagraph(review, Word.InsertLocation.after);
            await context.sync();

            document.getElementById('result').innerText = "Review inserted below the selected paragraph.";
        });
    } catch (error) {
        document.getElementById('result').innerText = `Error: ${error.message}`;
    }
}

async function reviewDocument() {
    if (!apiKey) {
        document.getElementById('result').innerText = "Please enter your API Key first.";
        return;
    }

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.load("text");
            await context.sync();

            const documentText = body.text;
            const suggestions = await getSuggestions(documentText);
            
            for (const suggestion of suggestions) {
                const range = body.getRange('Content').search(suggestion.original, { matchCase: true, matchWholeWord: false });
                range.load("text");
                await context.sync();

                if (range.items.length > 0) {
                    const firstRange = range.items[0];
                    firstRange.insertText(suggestion.suggested, Word.InsertLocation.replace);
                    firstRange.track();
                }
            }

            await context.sync();
            document.getElementById('result').innerText = "Document reviewed and suggestions added using track changes.";
        });
    } catch (error) {
        document.getElementById('result').innerText = `Error: ${error.message}`;
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

async function getSuggestions(text) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };

    const prompt = `
    Review the following document against Ofsted's SCIFF framework for Outstanding. Provide suggestions for improvement in JSON format:

    Document: "${text}"

    Please structure your response as a JSON array of objects, where each object represents a suggestion:
    [
        {
            "original": "Original text",
            "suggested": "Suggested improvement",
            "explanation": "Brief explanation for the change"
        },
        ...
    ]
    
    Only return the JSON no other text, do not enclose in '''
    `;

    try {
        const response = await axios.post(
            `${API_CONFIG.azureEndpoint}/openai/deployments/${API_CONFIG.deploymentName}/chat/completions?api-version=${API_CONFIG.apiVersion}`,
            {
                messages: [{ role: "user", content: prompt }],
                temperature: 0.5,
                max_tokens: 2000
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
        throw new Error("An error occurred while reviewing the document. Please check your API key and try again.");
    }
}
