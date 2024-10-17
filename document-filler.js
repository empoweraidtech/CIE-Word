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
            const body = document.body;
            body.load("paragraphs,tables,font");
            await context.sync();

            const documentStructure = await analyzeDocumentStructure(body);
            console.log("Document structure:", JSON.stringify(documentStructure, null, 2));

            const filledContent = await analyzeAndFillDocument(documentStructure, pdfText);

            for (const item of filledContent) {
                if (item.type === 'paragraph') {
                    const paragraphs = body.paragraphs.items;
                    if (item.index < paragraphs.length) {
                        const paragraph = paragraphs[item.index];
                        const newParagraph = paragraph.insertParagraph(item.filledText, Word.InsertLocation.after);
                        newParagraph.font.set(body.font);
                    }
                } else if (item.type === 'table') {
                    const tables = body.tables.items;
                    if (item.tableIndex < tables.length) {
                        const table = tables[item.tableIndex];
                        const cell = table.getCell(item.rowIndex, item.columnIndex);
                        const newRow = table.insertRow(item.rowIndex + 1, Word.InsertLocation.after);
                        const newCell = newRow.getCell(item.columnIndex);
                        newCell.body.insertParagraph(item.filledText, Word.InsertLocation.replace);
                        newCell.body.font.set(body.font);
                    }
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

async function analyzeDocumentStructure(body) {
    const structure = [];
    let paragraphIndex = 0;
    let tableIndex = 0;

    body.load("paragraphs,tables");
    await body.context.sync();

    for (let i = 0; i < body.paragraphs.items.length; i++) {
        const paragraph = body.paragraphs.items[i];
        paragraph.load("text,style");
        await paragraph.context.sync();

        structure.push({
            type: 'paragraph',
            index: paragraphIndex,
            text: paragraph.text,
            style: paragraph.style
        });
        paragraphIndex++;
    }

    for (let i = 0; i < body.tables.items.length; i++) {
        const table = body.tables.items[i];
        table.load("rowCount,columnCount");
        await table.context.sync();

        for (let row = 0; row < table.rowCount; row++) {
            for (let col = 0; col < table.columnCount; col++) {
                const cell = table.getCell(row, col);
                cell.load("body");
                await cell.context.sync();

                structure.push({
                    type: 'table',
                    tableIndex: tableIndex,
                    rowIndex: row,
                    columnIndex: col,
                    text: cell.body.text
                });
            }
        }
        tableIndex++;
    }

    return structure;
}

async function analyzeAndFillDocument(documentStructure, pdfText) {
    const API_CONFIG = {
        model: 'gpt-4o',
        apiVersion: '2023-12-01-preview',
        deploymentName: 'gpt4o',
        azureEndpoint: 'https://cieuk1.openai.azure.com',
    };
    
    const prompt = `Analyze the following document structure and PDF content. Fill each section of the document with relevant information from the PDF. If a section doesn't have relevant information, use the phrase 'No relevant information found in the document'.

    Document structure:
    ${JSON.stringify(documentStructure)}

    PDF content:
    ${pdfText}

    Provide your response in the following JSON format:
    [
      {
        "type": "paragraph",
        "index": 0,
        "filledText": "Filled content for paragraph 0"
      },
      {
        "type": "table",
        "tableIndex": 0,
        "rowIndex": 0,
        "columnIndex": 0,
        "filledText": "Filled content for table 0, cell (0,0)"
      },
      ...
    ]

    Rules:
    1. Do not modify existing text. Only add new content.
    2. For paragraphs, insert the new content after the existing paragraph.
    3. For table cells, provide content to be inserted in a new row below the current cell.
    4. If no relevant information is found for a section, set "filledText" to "No relevant information found in the document".
    5. Use actual line breaks instead of \\n for new lines.
    6. Ensure the JSON is not enclosed in any code blocks or quotation marks.`;
    
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
