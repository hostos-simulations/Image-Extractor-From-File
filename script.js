// Initialize PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const log = document.getElementById('log');

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
    updateStatus(`Analyzing ${file.name}...`);
    log.innerHTML = "";
    log.style.display = "block";

    try {
        if (ext === 'pdf') {
            await extractFromPDF(file);
        } else if (['docx', 'pptx', 'xlsx'].includes(ext)) {
            await extractFromOffice(file);
        } else {
            updateStatus("Unsupported file format.");
        }
    } catch (err) {
        console.error(err);
        updateStatus("Processing failed. The file might be encrypted or corrupted.");
    }
}

// Robust Office Extraction: Scans all internal files for image extensions
async function extractFromOffice(file) {
    const zip = await JSZip.loadAsync(file);
    const outputZip = new JSZip();
    let count = 0;
    const imageExtensions = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'emf', 'wmf', 'svg', 'tiff'];

    for (const [path, zipEntry] of Object.entries(zip.files)) {
        const fileExt = path.split('.').pop().toLowerCase();
        if (imageExtensions.includes(fileExt)) {
            const blob = await zipEntry.async("blob");
            const fileName = path.split('/').pop();
            outputZip.file(`img_${count}_${fileName}`, blob);
            addLog(`Extracted: ${fileName}`);
            count++;
        }
    }
    finalize(outputZip, count, file.name);
}

// Deep Scan PDF Extraction
async function extractFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
    const outputZip = new JSZip();
    let count = 0;

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const ops = await page.getOperatorList();
        
        // Find all Image Object keys
        const imageKeys = [];
        for (let j = 0; j < ops.fnArray.length; j++) {
            const fn = ops.fnArray[j];
            if (fn === pdfjsLib.OPS.paintImageXObject || 
                fn === pdfjsLib.OPS.paintJpegXObject || 
                fn === pdfjsLib.OPS.paintInlineImageXObject) {
                imageKeys.push(ops.argsArray[j][0]);
            }
        }

        // Process unique keys found on this page
        const uniqueKeys = [...new Set(imageKeys)];
        for (const key of uniqueKeys) {
            try {
                // PDF.js uses a callback for object resolution
                const img = await new Promise((resolve) => {
                    page.objs.get(key, (obj) => resolve(obj));
                });

                if (img && (img.data || img.bitmap)) {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.width;
                    canvas.height = img.height;
                    const ctx = canvas.getContext('2d');

                    if (img.bitmap) {
                        // Newer PDF.js versions might return a ImageBitmap
                        ctx.drawImage(img.bitmap, 0, 0);
                    } else {
                        const imgData = ctx.createImageData(img.width, img.height);
                        // Convert RGB to RGBA if necessary
                        if (img.data.length === img.width * img.height * 3) {
                            const rgba = new Uint8ClampedArray(img.width * img.height * 4);
                            for (let k = 0, l = 0; k < img.data.length; k += 3, l += 4) {
                                rgba[l] = img.data[k]; rgba[l+1] = img.data[k+1];
                                rgba[l+2] = img.data[k+2]; rgba[l+3] = 255;
                            }
                            imgData.data.set(rgba);
                        } else {
                            imgData.data.set(img.data);
                        }
                        ctx.putImageData(imgData, 0, 0);
                    }

                    const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
                    outputZip.file(`page${i}_img${count}.png`, blob);
                    addLog(`Page ${i}: Extracted image ${count}`);
                    count++;
                }
            } catch (e) {
                addLog(`Page ${i}: Skipped an unreadable image format.`);
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
        updateStatus("No extractable images found in this file.");
        return;
    }
    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = `extracted_images_${originalName.split('.')[0]}.zip`;
    link.click();
    updateStatus(`Success! Downloaded ${count} images.`);
}
