let apiKey = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('save-key').onclick = saveApiKey;
        document.getElementById('upload-pdf').onclick = uploadPdfAndFillDocument;
    }
});

function saveApiKey() {
    apiKey = document.getElementById('api-key').value;
    if (apiKey) {
        document.getElementById('api-key-input').classList.add('hidden');
        document.getElementById('filler-section').classList.remove('hidden');
        setResult("<p><i class='fas fa-check-circle text-green-500 mr-2'></i>API Key saved. You can now use the document filler feature.</p>");
    } else {
        setResult("<p><i class='fas fa-exclamation-triangle text-yellow-500 mr-2'></i>Please enter a valid API Key.</p>");
    }
}

async function uploadPdfAndFillDocument() {
    if (!apiKey) {
        setResult("<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Please enter your API Key first.</p>");
        return;
    }

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.pdf';
    fileInput.onchange = async (event) => {
        const file = event.target.files[0];
        if (file) {
            try {
                document.getElementById('loader').classList.remove('hidden');
                const pdfText = await parsePdf(file);
                await fillDocument(pdfText);
                document.getElementById('loader').classList.add('hidden');
            } catch (error) {
                document.getElementById('loader').classList.add('hidden');
                setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
            }
        }
    };
    fileInput.click();
}

async function parsePdf(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async function(event) {
            const typedarray = new Uint8Array(event.target.result);
            try {
                const pdf = await pdfjsLib.getDocument(typedarray).promise;
                let fullText = '';
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }
                resolve(fullText);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

async function fillDocument(pdfText) {
    try {
        await Word.run(async (context) => {
            const document = context.document;
            document.load("sections");
            await context.sync();

            if (!document.sections) {
                throw new Error("Unable to access document sections");
            }

            const sections = document.sections;
            sections.load("items");
            await context.sync();

            if (!sections.items || sections.items.length === 0) {
                throw new Error("No sections found in the document");
            }

            const documentStructure = [];
            for (let i = 0; i < sections.items.length; i++) {
                const sectionBody = sections.items[i].body;
                sectionBody.load("text");
                await context.sync();
                documentStructure.push({
                    index: i,
                    text: sectionBody.text
                });
            }

            console.log("Document structure:", JSON.stringify(documentStructure, null, 2));

            const filledContent = await analyzeAndFillDocument(documentStructure, pdfText);

            for (const section of filledContent) {
                if (sections.items[section.index]) {
                    const range = sections.items[section.index].body.getRange();
                    range.insertText(section.filledText, Word.InsertLocation.replace);
                } else {
                    console.warn(`Section with index ${section.index} not found`);
                }
            }

            await context.sync();
            setResult("<p><i class='fas fa-check-circle text-green-500 mr-2'></i>Document filled successfully.</p>");
        });
    } catch (error) {
        console.error("Error in fillDocument:", error);
        setResult(`<p><i class='fas fa-exclamation-circle text-red-500 mr-2'></i>Error: ${error.message}</p>`);
    }
}

async function analyzeAndFillDocument(documentStructure, pdfText) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    
    const prompt = `Analyze the following document structure and PDF content. Fill each section of the document with relevant information from the PDF. If a section doesn't have relevant information, leave it as is.

    Document structure:
    ${JSON.stringify(documentStructure)}

    PDF content:
    ${pdfText}

    Provide your response in the following JSON format:
    [
      {
        "index": 0,
        "filledText": "Filled content for section 0"
      },
      {
        "index": 1,
        "filledText": "Filled content for section 1"
      },
      ...
    ]

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
        throw new Error("Failed to analyze and fill the document. Please try again.");
    }
}

function setResult(html) {
    const resultEl = document.getElementById('result');
    if (resultEl) {
        resultEl.innerHTML = html;
    }
}
