let apiKey = '';

// Last updated: 2023-10-04 22:30:00 UTC
const lastUpdated = "2023-10-04 22:30:00 UTC";

// Add this near the top of your file, after the apiKey declaration

const tristonePolicies = {
    "TP001": {
        name: "Child-Centered Approach",
        summary: "Ensure all practices prioritize the child's needs and well-being.",
        fullPolicyLink: "https://tristone.org/policies/TP001"
    },
    "TP002": {
        name: "Staff Training and Development",
        summary: "Regular training programs for all staff members to maintain high standards of care.",
        fullPolicyLink: "https://tristone.org/policies/TP002"
    },
    "TP003": {
        name: "Safety and Risk Assessment",
        summary: "Comprehensive risk assessments for all activities and environments.",
        fullPolicyLink: "https://tristone.org/policies/TP003"
    },
    "TP004": {
        name: "Inclusive Practice",
        summary: "Ensure equal opportunities and support for all children regardless of background or abilities.",
        fullPolicyLink: "https://tristone.org/policies/TP004"
    },
    "TP005": {
        name: "Safeguarding Procedures",
        summary: "Robust procedures to protect children from harm and respond to concerns.",
        fullPolicyLink: "https://tristone.org/policies/TP005"
    }
};

const ofstedStandards = {
    "OS001": {
        name: "The overall experiences and progress of children and young people",
        summary: "Focus on the quality of care and support provided to children.",
        fullStandardLink: "https://www.gov.uk/government/publications/introduction-to-childrens-homes"
    },
    "OS002": {
        name: "How well children and young people are helped and protected",
        summary: "Emphasis on safeguarding and risk management practices.",
        fullStandardLink: "https://www.gov.uk/government/publications/introduction-to-childrens-homes"
    },
    "OS003": {
        name: "The effectiveness of leaders and managers",
        summary: "Evaluation of leadership and management in improving outcomes for children.",
        fullStandardLink: "https://www.gov.uk/government/publications/introduction-to-childrens-homes"
    },
    "OS004": {
        name: "Quality of education",
        summary: "Assessment of the educational provisions and support for children.",
        fullStandardLink: "https://www.gov.uk/government/publications/education-inspection-framework"
    },
    "OS005": {
        name: "Behaviour and attitudes",
        summary: "Focus on promoting positive behaviour and attitudes among children.",
        fullStandardLink: "https://www.gov.uk/government/publications/education-inspection-framework"
    },
    "OS006": {
        name: "Personal development",
        summary: "Evaluation of support for children's personal growth and development.",
        fullStandardLink: "https://www.gov.uk/government/publications/education-inspection-framework"
    },
    "OS007": {
        name: "Leadership and management",
        summary: "Assessment of the effectiveness of leadership in driving improvements.",
        fullStandardLink: "https://www.gov.uk/government/publications/education-inspection-framework"
    },
    "OS008": {
        name: "Overall effectiveness",
        summary: "Overall evaluation of the quality of care and support provided.",
        fullStandardLink: "https://www.gov.uk/government/publications/education-inspection-framework"
    },
    "OS009": {
        name: "The experiences and progress of children and young people",
        summary: "Focus on the outcomes and experiences of children in care.",
        fullStandardLink: "https://www.gov.uk/government/publications/inspecting-childrens-homes-framework"
    }
};

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
            
            // Prepare the result area
            setResult('');
            
            // Show loading indicator
            document.getElementById('loader').classList.remove('hidden');
            
            // Use a default review mode
            const reviewMode = "standard";
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

            for (const comment of analysis.comments) {
                const paragraphIndex = comment.paragraphIndex;
                if (paragraphIndex >= 0 && paragraphIndex < paragraphs.items.length) {
                    const paragraph = paragraphs.items[paragraphIndex];
                    const commentRange = paragraph.getRange();
                    const commentText = addPolicyLinks(comment.text, comment.policyReferences);
                    commentRange.insertComment(commentText);
                }
            }

            await context.sync();

            setResult(`<p><i class='fas fa-check-circle text-green-500 mr-2'></i>Full page review completed. ${analysis.comments.length} comments added.</p>`);
        });
    } catch (error) {
        setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
    }
}

function addPolicyLinks(commentText, policyReferences) {
    policyReferences.forEach(ref => {
        const policy = tristonePolicies[ref] || ofstedStandards[ref];
        if (policy) {
            const link = `<a href="#" onclick="showPolicyDetails('${ref}'); return false;">${policy.name}</a>`;
            commentText = commentText.replace(new RegExp(policy.name, 'g'), link);
        }
    });
    return commentText;
}

function showPolicyDetails(policyRef) {
    const policy = tristonePolicies[policyRef] || ofstedStandards[policyRef];
    if (policy) {
        const policyDetails = `
            <h3>${policy.name}</h3>
            <p>${policy.summary}</p>
            <p><a href="${policy.fullPolicyLink}" target="_blank">View full policy</a></p>
        `;
        document.getElementById('policy-details').innerHTML = policyDetails;
        document.getElementById('policy-details').classList.remove('hidden');
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

function displayReview(review) {
    const resultEl = document.getElementById('result');
    if (!resultEl) return;

    // Clear previous content
    resultEl.innerHTML = '';

    // Visualization
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

    // Summary
    const summarySection = createCollapsibleSection('Summary', review.summary);
    resultEl.appendChild(summarySection);

    // Suggested Changes
    const changesContent = review.suggestedChanges.map(change => `<li>${change}</li>`).join('');
    const changesSection = createCollapsibleSection('Suggested Changes', `<ul>${changesContent}</ul>`);
    resultEl.appendChild(changesSection);

    // Proposed Alternative
    const alternativeSection = createCollapsibleSection('Proposed Alternative', review.proposedAlternative);
    resultEl.appendChild(alternativeSection);

    // Copy to Clipboard button
    const copyButton = document.createElement('button');
    copyButton.id = 'copy-alternative';
    copyButton.className = 'bg-blue-500 text-white p-2 rounded hover:bg-blue-600 mt-2';
    copyButton.innerHTML = '<i class="fas fa-copy mr-2"></i>Copy to Clipboard';
    copyButton.onclick = copyToClipboard;
    alternativeSection.appendChild(copyButton);

    // Setup collapsible functionality
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

async function analyzeFullDocument(documentParagraphs) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    
    const prompt = `Analyze the following document paragraphs and identify those that need specific attention, focusing on Tristone policy, Ofsted standards, and readability. Focus on body paragraphs, ignore formatting / title issues / blank spaces. For each paragraph that needs attention, provide:
    1. The index of the paragraph (as provided in the input)
    2. A comment explaining what needs to change and why, considering Tristone policy, Ofsted standards, and readability.
    3. References to specific Tristone policies (TP001-TP005) or Ofsted standards (OS001-OS009) that are relevant.

    Tristone Policies:
    ${JSON.stringify(tristonePolicies)}

    Ofsted Standards:
    ${JSON.stringify(ofstedStandards)}

    Document paragraphs:
    ${JSON.stringify(documentParagraphs)}

    Provide your response in the following JSON format:
    {
      "comments": [
        {
          "paragraphIndex": 0,
          "text": "Comment text explaining what needs to change and why",
          "policyReferences": ["TP001", "OS002"]
        },
        ...
      ]
    }

    Ensure the JSON is not enclosed in any code blocks or quotation marks.`;
    
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
