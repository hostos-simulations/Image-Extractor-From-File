// Initialize PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const log = document.getElementById('log');

// Event Listeners
dropZone.onclick = () => fileInput.click();
fileInput.onchange = (e) => handleFile(e.target.files[0]);

dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('active'); };
dropZone.ondragleave = () => dropZone.classList.remove('active');
dropZone.ondrop = (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    handleFile(e.dataTransfer.files[0]);
};

async function handleFile(file) {
    if (!file) return;
    
    const ext = file.name.split('.').pop().toLowerCase();
    updateStatus(`Processing ${file.name}...`, "info");
    log.innerHTML = "";
    log.style.display = "block";

    try {
        if (ext === 'pdf') {
            await extractFromPDF(file);
        } else if (ext === 'docx' || ext === 'pptx') {
            await extractFromOffice(file, ext);
        } else {
            updateStatus("Unsupported file format.", "error");
        }
    } catch (err) {
        console.error(err);
        updateStatus("Error during extraction.", "error");
    }
}

async function extractFromOffice(file, ext) {
    const zip = await JSZip.loadAsync(file);
    const outputZip = new JSZip();
    let count = 0;

    // Word uses 'word/media/', PPT uses 'ppt/media/'
    const mediaPath = ext === 'docx' ? 'word/media/' : 'ppt/media/';
    
    for (const [path, zipEntry] of Object.entries(zip.files)) {
        if (path.startsWith(mediaPath)) {
            const blob = await zipEntry.async("blob");
            const fileName = path.split('/').pop();
            outputZip.file(`extracted_${fileName}`, blob);
            addLog(`Found image: ${fileName}`);
            count++;
        }
    }
    finalize(outputZip, count, file.name);
}

async function extractFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
    const outputZip = new JSZip();
    let count = 0;

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const ops = await page.getOperatorList();
        
        for (let j = 0; j < ops.fnArray.length; j++) {
            // Check if operator is an image
            if (ops.fnArray[j] === pdfjsLib.OPS.paintImageXObject || ops.fnArray[j] === pdfjsLib.OPS.paintJpegXObject) {
                const imgName = ops.argsArray[j][0];
                try {
                    const img = await page.objs.get(imgName);
                    
                    // Render PDF image data to canvas to convert to PNG
                    const canvas = document.createElement('canvas');
                    canvas.width = img.width;
                    canvas.height = img.height;
                    const ctx = canvas.getContext('2d');
                    const imgData = ctx.createImageData(img.width, img.height);
                    imgData.data.set(img.data);
                    ctx.putImageData(imgData, 0, 0);
                    
                    const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
                    outputZip.file(`page${i}_img${j}.png`, blob);
                    addLog(`Extracted image from page ${i}`);
                    count++;
                } catch (e) {
                    addLog(`Skipped specialized image on page ${i}`);
                }
            }
        }
    }
    finalize(outputZip, count, file.name);
}

function updateStatus(msg) { status.innerText = msg; }

function addLog(msg) {
    const div = document.createElement('div');
    div.className = 'log-entry';
    div.innerText = msg;
    log.appendChild(div);
}

async function finalize(zip, count, originalName) {
    if (count === 0) {
        updateStatus("No images found.");
        return;
    }

    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = `images_from_${originalName.split('.')[0]}.zip`;
    link.click();
    updateStatus(`Successfully extracted ${count} images!`);
}