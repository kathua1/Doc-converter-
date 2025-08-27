// script.js - client-side conversion logic
const { jsPDF } = window.jspdf;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.8.162/pdf.worker.min.js';

// Image -> PDF
document.getElementById('imgToPdfBtn').addEventListener('click', async () => {
  const input = document.getElementById('imgToPdfInput');
  if (!input.files.length) return alert('Choose one or more images.');
  const doc = new jsPDF();
  for (let i = 0; i < input.files.length; i++) {
    const file = input.files[i];
    const imgData = await fileToDataURL(file);
    const img = new Image();
    await new Promise(r => { img.onload = r; img.src = imgData; });
    const w = doc.internal.pageSize.getWidth();
    const h = (img.height * w) / img.width;
    if (i > 0) doc.addPage();
    doc.addImage(img, 'JPEG', 0, 0, w, h);
  }
  const pdfBlob = doc.output('blob');
  const link = document.getElementById('imgToPdfDownload');
  link.href = URL.createObjectURL(pdfBlob);
  link.classList.remove('hidden');
});

// PDF -> Images (renders pages to canvas and allows download)
document.getElementById('pdfToImgBtn').addEventListener('click', async () => {
  const input = document.getElementById('pdfToImgInput');
  if (!input.files.length) return alert('Choose a PDF file.');
  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({data: arrayBuffer});
  const pdf = await loadingTask.promise;
  const container = document.getElementById('pdfImgResults');
  container.innerHTML = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({scale: 2});
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({canvasContext: ctx, viewport}).promise;
    container.appendChild(canvas);

    // create download link for each page
    canvas.toBlob(blob => {
      const a = document.createElement('a');
      a.textContent = `Download page ${p} as PNG`;
      a.href = URL.createObjectURL(blob);
      a.download = `page-${p}.png`;
      a.style.display = 'inline-block';
      a.style.marginRight = '10px';
      container.appendChild(a);
    }, 'image/png');
  }
});

// PDF -> DOCX (creates docx with page images)
document.getElementById('pdfToDocxBtn').addEventListener('click', async () => {
  const input = document.getElementById('pdfToDocxInput');
  if (!input.files.length) return alert('Choose a PDF file.');
  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;

  const images = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({scale: 2});
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({canvasContext: canvas.getContext('2d'), viewport}).promise;
    const dataUrl = canvas.toDataURL('image/png');
    images.push(dataUrl);
  }

  // build docx
  const { Document, Packer, Paragraph, ImageRun } = docx;
  const doc = new Document();
  images.forEach(dataUrl => {
    const imgBase64 = dataUrl.split(',')[1];
    const img = new ImageRun({
      data: Uint8Array.from(atob(imgBase64), c => c.charCodeAt(0)),
      transformation: { width: 600, height: 800 },
    });
    doc.addSection({ children: [ new Paragraph({ children:[img] }) ] });
  });

  const blob = await Packer.toBlob(doc);
  const link = document.getElementById('pdfToDocxDownload');
  link.href = URL.createObjectURL(blob);
  link.classList.remove('hidden');
});

// PDF -> XLSX (extract text per page into rows)
document.getElementById('pdfToXlsxBtn').addEventListener('click', async () => {
  const input = document.getElementById('pdfToXlsxInput');
  if (!input.files.length) return alert('Choose a PDF file.');
  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;

  const rows = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const txt = await page.getTextContent();
    const strings = txt.items.map(i => i.str).join(' ');
    rows.push([`Page ${p}`, strings]);
  }

  const ws = XLSX.utils.aoa_to_sheet([['Page','Text'], ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Extracted');
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const blob = new Blob([wbout], {type:'application/octet-stream'});
  const link = document.getElementById('pdfToXlsxDownload');
  link.href = URL.createObjectURL(blob);
  link.classList.remove('hidden');
});

// DOCX -> PDF (mammoth -> html -> jsPDF)
document.getElementById('docxToPdfBtn').addEventListener('click', async () => {
  const input = document.getElementById('docxToPdfInput');
  if (!input.files.length) return alert('Choose a DOCX file.');
  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({arrayBuffer});
  const html = result.value;
  // create a hidden iframe to render HTML then convert to PDF via jsPDF html
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.left = '-9999px';
  document.body.appendChild(iframe);
  iframe.contentDocument.open();
  iframe.contentDocument.write(html);
  iframe.contentDocument.close();
  await new Promise(r => setTimeout(r, 500));
  const doc = new jsPDF('p','pt','a4');
  await doc.html(iframe.contentDocument.body, {callback: () => {
    const blob = doc.output('blob');
    const link = document.getElementById('docxToPdfDownload');
    link.href = URL.createObjectURL(blob);
    link.classList.remove('hidden');
    document.body.removeChild(iframe);
  }, x: 10, y: 10, width: 560});
});

// Images -> PPTX (each image becomes a slide)
document.getElementById('imgToPptxBtn').addEventListener('click', async () => {
  const input = document.getElementById('imgToPptxInput');
  if (!input.files.length) return alert('Choose one or more images.');
  const pptx = new PptxGenJS();
  for (let i = 0; i < input.files.length; i++) {
    const file = input.files[i];
    const dataUrl = await fileToDataURL(file);
    const slide = pptx.addSlide();
    slide.addImage({ data: dataUrl, x:0.5, y:0.5, w:9, h:5 });
  }
  const outName = 'presentation.pptx';
  pptx.writeFile({ fileName: outName }).then(() => {
    const link = document.getElementById('imgToPptxDownload');
    // PptxGenJS already triggers download; show link for consistency
    link.classList.remove('hidden');
  });
});

// utility
function fileToDataURL(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(fr.result);
    fr.onerror = rej;
    fr.readAsDataURL(file);
  });
}
