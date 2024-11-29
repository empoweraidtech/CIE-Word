let apiKey = '';

// Last updated: 2023-10-04 22:30:00 UTC
const lastUpdated = "2023-10-18 09:30:00 UTC";

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('run').onclick = run;
        document.getElementById('full-page-review').onclick = fullPageReview;
        document.getElementById('last-updated').textContent = `Last updated: ${lastUpdated}`;
    }
});

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
            
            setResult('');
            document.getElementById('loader').classList.remove('hidden');
            
            const reviewMode = "standard";
            const review = await reviewParagraph(selectedText, reviewMode);
            
            document.getElementById('loader').classList.add('hidden');
            displayReview(review);
        });
    } catch (error) {
        setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
    }
}

async function fullPageReview() {
    if (!apiKey) {
        setResult("<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Please enter your API Key first.</p>");
        return;
    }
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("text");
            await context.sync();

            const documentParagraphs = paragraphs.items.map((p, index) => ({
                index: index,
                text: p.text
            }));

            document.getElementById('loader').classList.remove('hidden');
            const analysis = await analyzeFullDocument(documentParagraphs);
            document.getElementById('loader').classList.add('hidden');

            // Process each comment and add highlighted annotations
            for (const comment of analysis.comments) {
                const paragraphIndex = comment.paragraphIndex;
                if (paragraphIndex >= 0 && paragraphIndex < paragraphs.items.length) {
                    const paragraph = paragraphs.items[paragraphIndex];
                    
                    // Search for the specific text within the paragraph
                    const range = paragraph.getRange();
                    const searchResults = range.search(comment.targetText);
                    context.load(searchResults);
                    await context.sync();

                    // If the specific text is found, add comment to just that portion
                    if (searchResults.items.length > 0) {
                        searchResults.items[0].insertComment(comment.text);
                        // Optional: Add highlighting to the specific text
                        searchResults.items[0].font.highlightColor = 'yellow';
                    } else {
                        // Fallback to paragraph if exact text not found
                        range.insertComment(comment.text);
                    }
                }
            }

            await context.sync();
            setResult(`<p><i class='fas fa-check-circle text-green-500 mr-2'></i>Review completed. ${analysis.comments.length} issues identified.</p>`);
        });
    } catch (error) {
        setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
    }
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

async function analyzeFullDocument(documentParagraphs) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    
    const prompt = `Analyze the following document paragraphs for inconsistencies, errors, and quality issues. For each issue, identify the specific text that needs attention.

For each issue found, examine:

1. Timeline and Chronological Issues
- Events occurring in impossible orders
- Inconsistent date formats
- Future dates in historical sections
- Illogical sequence of interventions

2. Personal Information
- Name spelling variations
- Address inconsistencies
- Living arrangement contradictions
- Pronoun inconsistencies
- Contradictory family relationships

3. Clinical Details
- Medical condition spelling errors
- Treatment timeline inconsistencies
- Facility name inconsistencies
- Healthcare provider reference errors
- Medication inconsistencies

4. Support and Care Information
- Contradictory support frequencies
- Inconsistent caregiver arrangements
- Contradictory independence levels
- Care package inconsistencies

5. Quality Standards
- Readability issues
- Structural problems

Document paragraphs:
${JSON.stringify(documentParagraphs)}

Provide your response in EXACTLY this JSON format:
{
  "comments": [
    {
      "paragraphIndex": 0,
      "targetText": "specific text that has the issue",
      "text": "[Category]: Issue found - Explanation - Suggested correction"
    }
  ]
}

Focus on finding specific phrases, words, or sentences that contain issues rather than flagging entire paragraphs. For each issue:
- Include the exact problematic text in targetText
- Provide a clear explanation and suggestion in the text field
- Format the text field as [Category]: Issue - Explanation - Correction
Empty paragraphs don't need comments!`;
    
    try {
        const response = await axios.post(
            `${API_CONFIG.azureEndpoint}/openai/deployments/${API_CONFIG.deploymentName}/chat/completions?api-version=${API_CONFIG.apiVersion}`,
            {
                messages: [{ role: "user", content: prompt }],
                temperature: 0.7,
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
        throw new Error("Failed to analyze the document. Please try again.");
    }
}

function displayReview(review) {
    const resultEl = document.getElementById('result');
    if (!resultEl) return;

    resultEl.innerHTML = '';

    const visualizationEl = document.createElement('div');
    visualizationEl.className = 'mb-4 p-4 border rounded';
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
    resultEl.appendChild(visualizationEl);

    const summarySection = createCollapsibleSection('Summary', review.summary);
    resultEl.appendChild(summarySection);

    const changesContent = review.suggestedChanges.map(change => `<li>${change}</li>`).join('');
    const changesSection = createCollapsibleSection('Suggested Changes', `<ul>${changesContent}</ul>`);
    resultEl.appendChild(changesSection);

    const alternativeSection = createCollapsibleSection('Proposed Alternative', review.proposedAlternative);
    resultEl.appendChild(alternativeSection);

    const copyButton = document.createElement('button');
    copyButton.id = 'copy-alternative';
    copyButton.className = 'bg-blue-500 text-white p-2 rounded hover:bg-blue-600 mt-2';
    copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
    copyButton.onclick = copyToClipboard;
    alternativeSection.appendChild(copyButton);

    setupCollapsibles();
}

function createCollapsibleSection(title, content) {
    const section = document.createElement('div');
    section.className = 'mb-4';
    section.innerHTML = `
        <button class="collapsible">${title}</button>
        <div class="content">
            <div class="p-4">${marked.parse(content)}</div>
        </div>
    `;
    return section;
}

function setupCollapsibles() {
    var coll = document.getElementsByClassName("collapsible");
    for (var i = 0; i < coll.length; i++) {
        coll[i].addEventListener("click", function() {
            this.classList.toggle("active");
            var content = this.nextElementSibling;
            content.classList.toggle("show");
        });
    }
}

function copyToClipboard() {
    const alternativeContent = document.querySelector('#result .content:nth-of-type(3) .p-4');
    if (alternativeContent) {
        const alternativeText = alternativeContent.textContent;
        navigator.clipboard.writeText(alternativeText)
            .then(() => {
                const copyButton = document.getElementById('copy-alternative');
                if (copyButton) {
                    copyButton.innerHTML = '<i class="fas fa-check mr-2"></i>Copied!';
                    setTimeout(() => {
                        copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
                    }, 2000);
                }
            })
            .catch(err => {
                console.error('Failed to copy text: ', err);
                alert('Failed to copy text. Please try again.');
            });
    } else {
        console.error('Alternative content not found');
        alert('Content not found. Please try again.');
    }
}
