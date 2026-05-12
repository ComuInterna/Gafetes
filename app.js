const excelInput  = document.getElementById('excelFile');
const photoInput  = document.getElementById('photoFiles');
const generateBtn = document.getElementById('generateBtn');
const previewArea = document.getElementById('previewArea');

// UI refs
const excelCard    = document.getElementById('excelCard');
const photoCard    = document.getElementById('photoCard');
const excelHint    = document.getElementById('excelHint');
const photoHint    = document.getElementById('photoHint');
const statusChip   = document.getElementById('statusChip');
const statusDot    = statusChip.querySelector('.status-dot');
const statusText   = statusChip.querySelector('.status-text');
const previewEmpty = document.getElementById('previewEmpty');
const previewCount = document.getElementById('previewCount');

let employees  = [];
let photoMap   = {};
let canvasCache = []; // { canvas, filename, nombre }

// ── Observer: contador de previews ─────────────────────
const observer = new MutationObserver(() => {
  const n = previewArea.children.length;
  if (n > 0) {
    previewEmpty.style.display = 'none';
    previewCount.textContent = n + ' gafete' + (n !== 1 ? 's' : '');
  } else {
    previewEmpty.style.display = 'flex';
    previewCount.textContent = '';
  }
});
observer.observe(previewArea, { childList: true });

// ── UI helpers ──────────────────────────────────────────
function checkReady() {
  const r = excelCard.classList.contains('loaded') && photoCard.classList.contains('loaded');
  statusDot.className = 'status-dot ' + (r ? 'ready' : 'idle');
  statusText.textContent = r ? 'Listo para generar' : 'En espera';
}

// ── Listeners de archivos ───────────────────────────────
photoInput.addEventListener('change', (e) => {
  const files = e.target.files;
  for (const file of files) {
    const key = file.name.split('.')[0].toLowerCase();
    photoMap[key] = URL.createObjectURL(file);
  }
  const n = files.length;
  photoHint.textContent = n + ' foto' + (n !== 1 ? 's' : '') + ' seleccionada' + (n !== 1 ? 's' : '');
  photoCard.classList.add('loaded');
  checkReady();
});

excelInput.addEventListener('change', handleExcel);
generateBtn.addEventListener('click', generateBadges);

function handleExcel(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    employees = XLSX.utils.sheet_to_json(sheet);
    excelHint.textContent = file.name + ' · ' + employees.length + ' registros';
    excelCard.classList.add('loaded');
    checkReady();
  };
  reader.readAsArrayBuffer(file);
}

// ── Generación ──────────────────────────────────────────
async function generateBadges() {
  if (!employees.length) {
    alert('Carga primero el Excel');
    return;
  }

  previewArea.innerHTML = '';
  canvasCache = [];

  // Ocultar botones de descarga anteriores
  const oldBtns = document.getElementById('downloadBtns');
  if (oldBtns) oldBtns.remove();

  // Mostrar progress
  const progressWrap = document.getElementById('progressWrap');
  const progressFill = document.getElementById('progressFill');
  const progressLabel = document.getElementById('progressLabel');
  progressWrap.style.display = 'block';

  const zip = new JSZip();
  const total = employees.length;

  for (let i = 0; i < total; i++) {
    const employee = employees[i];

    // Progress
    progressFill.style.width = Math.round(((i) / total) * 100) + '%';
    progressLabel.textContent = 'Procesando ' + (i + 1) + ' de ' + total + '…';

    const badge = createBadge(employee);
    badge.style.position = "absolute";
    badge.style.left = "-99999px";
    document.body.appendChild(badge);

    // Clon para preview
    const previewBadge = badge.cloneNode(true);
    previewBadge.style.transform = "scale(.45)";
    previewBadge.style.transformOrigin = "top left";
    previewBadge.style.position = "absolute";
    previewBadge.style.top = "0";
    previewBadge.style.left = "0";

    const previewContainer = document.createElement("div");
    previewContainer.style.width = "287px";
    previewContainer.style.height = "457px";
    previewContainer.style.position = "relative";
    previewContainer.style.overflow = "hidden";
    previewContainer.style.background = "#fff";
    previewContainer.appendChild(previewBadge);
    previewArea.appendChild(previewContainer);

    await waitImages();

    const canvas = await html2canvas(badge, {
      scale: 2,
      useCORS: true,
      width: 638,
      height: 1016,
    });

    document.body.removeChild(badge);

    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));

    const cleanName = (employee.nombre || 'gafete')
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .split(" ");

    const firstName     = cleanName[0] || "";
    const firstLastName = cleanName[1] || "";
    const filename      = firstName + '-' + firstLastName + '-' + employee.id + '.png';

    zip.file(filename, blob);
    canvasCache.push({ canvas, filename, nombre: employee.nombre || filename });

    // Overlay botones por tarjeta
    const blobUrl = URL.createObjectURL(blob);

    const overlay = document.createElement("div");
    overlay.className = "preview-overlay";

    const nameLabel = document.createElement("span");
    nameLabel.className = "preview-name-label";
    nameLabel.textContent = employee.nombre || filename;

    const btnDownload = document.createElement("a");
    btnDownload.className = "preview-action-btn download";
    btnDownload.href = blobUrl;
    btnDownload.download = filename;
    btnDownload.title = "Descargar PNG";
    btnDownload.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';

    const btnDelete = document.createElement("button");
    btnDelete.className = "preview-action-btn delete";
    btnDelete.title = "Eliminar";
    btnDelete.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>';
    btnDelete.addEventListener("click", () => {
      previewContainer.classList.add("removing");
      setTimeout(() => previewContainer.remove(), 280);
    });

    overlay.appendChild(nameLabel);
    overlay.appendChild(btnDownload);
    overlay.appendChild(btnDelete);
    previewContainer.appendChild(overlay);
  }

  progressFill.style.width = '100%';
  progressLabel.textContent = '¡Listo! ' + total + ' gafetes generados.';

  // Mostrar botones de descarga global
  showDownloadButtons(zip);
}

// ── Botones de descarga global ──────────────────────────
function showDownloadButtons(zip) {
  const container = document.createElement('div');
  container.id = 'downloadBtns';
  container.className = 'download-btns-row';

  const btnZip = document.createElement('button');
  btnZip.className = 'btn-dl btn-dl-zip';
  btnZip.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Descargar ZIP';
  btnZip.addEventListener('click', () => {
    zip.generateAsync({ type: 'blob' }).then(content => saveAs(content, 'gafetes.zip'));
  });

  const btnPdf = document.createElement('button');
  btnPdf.className = 'btn-dl btn-dl-pdf';
  btnPdf.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="15" y2="17"/></svg> Descargar PDF';
  btnPdf.addEventListener('click', generatePDF);

  container.appendChild(btnZip);
  container.appendChild(btnPdf);

  // Insertar antes de la preview section
  const previewSection = document.querySelector('.preview-section');
  previewSection.parentNode.insertBefore(container, previewSection);
}

// ── Generación de PDF ───────────────────────────────────
async function generatePDF() {
  if (!canvasCache.length) return;

  const { jsPDF } = window.jspdf;

  // Dimensiones del gafete en mm (proporcional a 638x1016px a 96dpi)
  // Usamos A6 portrait (105 x 148 mm) que es tamaño credencial estándar impresa
  const pageW = 105;
  const pageH = 148;

  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: [pageW, pageH],
    compress: true,
  });

  // Cargar Montserrat como fuente embebida para texto vectorial
  // (jsPDF incluye Helvetica nativa — usamos addFont si se quiere custom,
  //  pero la imagen del canvas ya lleva la foto; el texto del gafete
  //  proviene del canvas renderizado con html2canvas, así que ya está
  //  rasterizado en el canvas. Lo que agregamos aquí es la imagen completa.)

  for (let i = 0; i < canvasCache.length; i++) {
    const { canvas, nombre } = canvasCache[i];

    if (i > 0) pdf.addPage([pageW, pageH], 'portrait');

    const imgData = canvas.toDataURL('image/jpeg', 0.95);

    // Ajustar imagen al 100% de la página manteniendo proporción
    const canvasRatio = canvas.height / canvas.width;
    const pageRatio   = pageH / pageW;

    let drawW, drawH, offsetX = 0, offsetY = 0;

    if (canvasRatio > pageRatio) {
      // Alto mayor → ajustar por alto
      drawH = pageH;
      drawW = pageH / canvasRatio;
      offsetX = (pageW - drawW) / 2;
    } else {
      // Ancho mayor → ajustar por ancho
      drawW = pageW;
      drawH = pageW * canvasRatio;
      offsetY = (pageH - drawH) / 2;
    }

    pdf.addImage(imgData, 'JPEG', offsetX, offsetY, drawW, drawH, undefined, 'FAST');
  }

  pdf.save('gafetes.pdf');
}

// ── Badge HTML ──────────────────────────────────────────
function createBadge(employee) {
  const badge = document.createElement('div');
  badge.className = 'badge';

  const photo = document.createElement('img');
  photo.className = 'photo';
  const photoKey = employee.foto?.toLowerCase();
  photo.src = photoMap[photoKey] || 'https://via.placeholder.com/300x300';

  const name = document.createElement('div');
  name.className = 'name';
  name.textContent = employee.nombre || '';

  const position = document.createElement('div');
  position.className = 'position';
  position.textContent = employee.puesto || '';

  const business = document.createElement('div');
  business.className = 'business';
  business.textContent = employee.unidad || '';

  const dateTitle = document.createElement('div');
  dateTitle.className = 'dateTitle';
  dateTitle.textContent = 'Fecha de ingreso:';

  const date = document.createElement('div');
  date.className = 'date';
  date.textContent = formatExcelDate(employee.fecha);

  const employeeId = document.createElement('div');
  employeeId.className = 'employeeId';
  employeeId.textContent = 'ID: ' + (employee.id || '');

  badge.appendChild(photo);
  badge.appendChild(name);
  badge.appendChild(position);
  badge.appendChild(business);
  badge.appendChild(dateTitle);
  badge.appendChild(date);
  badge.appendChild(employeeId);

  return badge;
}

function waitImages() {
  return Promise.all(
    Array.from(document.images)
      .filter(img => !img.complete)
      .map(img => new Promise(resolve => {
        img.onload = img.onerror = resolve;
      }))
  );
}

function formatExcelDate(excelDate) {

  if (!excelDate) return '';

  // Convertir serial Excel → UTC real
  const utc_days = Math.floor(excelDate - 25569);
  const utc_value = utc_days * 86400;

  const date_info = new Date(utc_value * 1000);

  const day = String(date_info.getUTCDate()).padStart(2, '0');
  const month = String(date_info.getUTCMonth() + 1).padStart(2, '0');
  const year = date_info.getUTCFullYear();

  return `${day}/${month}/${year}`;
}
