/**
 * ============================================================================
 * LÓGICA DE LA APLICACIÓN DE PRUEBA / DEMO (app.js)
 * ============================================================================
 * Esta versión gratuita de demostración no requiere clave de activación,
 * pero restringe la carga a un máximo de 15 invitados por lista Excel.
 * 
 * Si se intenta procesar una lista mayor, muestra un modal avisando al usuario
 * que debe adquirir la Licencia PRO en la página de ventas.
 */

// Límite de invitados máximo permitido en esta versión de demostración
const MAX_TRIAL_GUESTS = 15;

// Variables de estado global
let guestData = [];
let originalWorkbook = null;
let activeSheetName = "";
let originalColumnsOrder = [];
let isScanning = false;
let videoStream = null;
let scanInterval = null;
let lastScannedCode = "";
let lastScannedTime = 0;

// DOM Elements
const dropZone = document.getElementById("drop-zone");
const excelFileInput = document.getElementById("excel-file-input");
const fileInfo = document.getElementById("file-info");
const fileNameDisplay = document.getElementById("file-name-display");
const resetFileBtn = document.getElementById("reset-file-btn");

const statTotal = document.getElementById("stat-total");
const statPresent = document.getElementById("stat-present");
const statPending = document.getElementById("stat-pending");

const toggleCameraBtn = document.getElementById("toggle-camera-btn");
const scannerContainer = document.getElementById("scanner-container");
const videoPreview = document.getElementById("video-preview");

const searchInput = document.getElementById("search-input");
const guestsTableBody = document.getElementById("guests-table-body");
const downloadExcelBtn = document.getElementById("download-excel-btn");

const upgradeModal = document.getElementById("upgrade-modal");
const upgradeModalMsg = document.getElementById("upgrade-modal-msg");
const closeUpgradeModalBtn = document.getElementById("close-upgrade-modal-btn");

const EVENT_PALETTES = {
    terracotta: { name: "Terracota Gala", accent: "#DE6443", secondary: "#BA6E5B", hover: "#C94E2D", bgDark: "#1A1412", cardBg: "#271E1C" },
    gold:       { name: "Oro Champagne", accent: "#C5A059", secondary: "#A47B2C", hover: "#9A7432", bgDark: "#141311", cardBg: "#201E1A" },
    emerald:    { name: "Esmeralda Hacienda", accent: "#2E9E6B", secondary: "#52B788", hover: "#217A51", bgDark: "#0E1A15", cardBg: "#162821" },
    rose:       { name: "Rosa Cuarzo", accent: "#D96B82", secondary: "#B3546A", hover: "#B85066", bgDark: "#1B1316", cardBg: "#291D22" },
    royal:      { name: "Azul Medianoche", accent: "#4A80E8", secondary: "#60A5FA", hover: "#3366CC", bgDark: "#0F1624", cardBg: "#182338" },
    lavender:   { name: "Lavanda Imperial", accent: "#9B6B9E", secondary: "#B88BB8", hover: "#7E5181", bgDark: "#17131A", cardBg: "#241E29" }
};

let currentTrialAccent = "#DE6443";
let isTrialLightTheme = localStorage.getItem("trial-theme") === "light";

const themeColorMeta = document.getElementById("theme-color-meta");
const trialPaletteChips = document.getElementById("trial-palette-chips");
const trialColorPicker = document.getElementById("trial-color-picker");
const trialThemeToggle = document.getElementById("trial-theme-toggle");
const trialThemeIcon = document.getElementById("trial-theme-icon");
const trialThemeText = document.getElementById("trial-theme-text");

document.addEventListener("DOMContentLoaded", () => {
    setupPwa();
    setupTrialPaletteControls();
    setupFileLoaders();
    setupScanner();
    setupSearch();
    setupModalEvents();
    setupGuideModal();

    const downloadTemplateBtn = document.getElementById("download-template-btn");
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener("click", downloadSampleTemplate);
    }
});

function normalizeHex(hex) {
    if (!hex || typeof hex !== "string") return "#DE6443";
    let clean = hex.trim().replace(/^#/, "");
    if (clean.length === 3) clean = clean.split("").map(c => c + c).join("");
    if (!/^[0-9a-fA-F]{6}$/.test(clean)) return "#DE6443";
    return "#" + clean.toUpperCase();
}

function hexToRgb(hex) {
    const norm = normalizeHex(hex).slice(1);
    return {
        r: parseInt(norm.slice(0, 2), 16),
        g: parseInt(norm.slice(2, 4), 16),
        b: parseInt(norm.slice(4, 6), 16)
    };
}

function adjustHexBrightness(hex, percent) {
    const { r, g, b } = hexToRgb(hex);
    const factor = percent / 100;
    const adj = (ch) => {
        const val = factor >= 0 ? ch + (255 - ch) * factor : ch * (1 + factor);
        return Math.max(0, Math.min(255, Math.round(val)));
    };
    const toHex = (n) => n.toString(16).padStart(2, "0").toUpperCase();
    return `#${toHex(adj(r))}${toHex(adj(g))}${toHex(adj(b))}`;
}

function hexToRgba(hex, alpha) {
    const { r, g, b } = hexToRgb(hex);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function getContrastForeground(hex) {
    const { r, g, b } = hexToRgb(hex);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.62 ? "#1A1412" : "#FFFFFF";
}

function applyTrialPalette(paletteKey, customHex = null) {
    const palette = (paletteKey && EVENT_PALETTES[paletteKey]) ? EVENT_PALETTES[paletteKey] : null;
    const accent = normalizeHex(customHex || (palette ? palette.accent : "#DE6443"));
    const secondary = palette ? palette.secondary : adjustHexBrightness(accent, 18);
    const hover = palette ? palette.hover : adjustHexBrightness(accent, -18);
    const bgDark = palette ? palette.bgDark : "#1A1412";
    const cardBg = palette ? palette.cardBg : "#271E1C";
    const fg = getContrastForeground(accent);

    currentTrialAccent = accent;

    const root = document.documentElement.style;
    root.setProperty("--accent", accent);
    root.setProperty("--accent-secondary", secondary);
    root.setProperty("--accent-hover", hover);
    root.setProperty("--accent-fg", fg);
    root.setProperty("--accent-gradient", `linear-gradient(90deg, ${hover}, ${accent}, ${secondary})`);
    root.setProperty("--accent-glow", hexToRgba(accent, 0.16));
    root.setProperty("--border-color", hexToRgba(accent, 0.25));
    root.setProperty("--bg-dark", bgDark);
    root.setProperty("--card-bg", cardBg);

    if (themeColorMeta) themeColorMeta.setAttribute("content", accent);
    if (trialColorPicker) trialColorPicker.value = accent;

    document.querySelectorAll(".palette-chip").forEach(btn => {
        btn.classList.toggle("active", Boolean(palette && btn.dataset.palette === paletteKey));
    });
}

function applyTrialTheme(light) {
    isTrialLightTheme = light;
    document.body.classList.toggle("light-theme", light);
    if (trialThemeIcon) trialThemeIcon.className = light ? "ti ti-moon" : "ti ti-sun";
    if (trialThemeText) trialThemeText.textContent = light ? "Modo Noche" : "Modo Día";
    localStorage.setItem("trial-theme", light ? "light" : "dark");
}

function setupTrialPaletteControls() {
    if (trialPaletteChips) {
        trialPaletteChips.innerHTML = Object.entries(EVENT_PALETTES).map(([key, pal]) => `
            <button type="button" class="palette-chip" data-palette="${key}" title="${pal.name}">
                <span class="palette-chip-dot" style="background: linear-gradient(135deg, ${pal.accent}, ${pal.secondary});"></span>
                <span>${pal.name}</span>
            </button>
        `).join("");

        trialPaletteChips.querySelectorAll(".palette-chip").forEach(chip => {
            chip.addEventListener("click", () => {
                const key = chip.dataset.palette;
                const pal = EVENT_PALETTES[key];
                if (!pal) return;
                localStorage.setItem("trial-event-palette", key);
                localStorage.setItem("trial-accent-color", pal.accent);
                applyTrialPalette(key);
            });
        });
    }

    if (trialColorPicker) {
        trialColorPicker.addEventListener("input", (e) => {
            const hex = normalizeHex(e.target.value);
            localStorage.setItem("trial-event-palette", "custom");
            localStorage.setItem("trial-accent-color", hex);
            applyTrialPalette("custom", hex);
        });
    }

    if (trialThemeToggle) {
        trialThemeToggle.addEventListener("click", () => applyTrialTheme(!isTrialLightTheme));
    }

    const savedPalette = localStorage.getItem("trial-event-palette") || "terracotta";
    const savedColor = localStorage.getItem("trial-accent-color");
    applyTrialPalette(savedPalette, savedPalette === "custom" ? savedColor : null);
    applyTrialTheme(isTrialLightTheme);
}

function setupGuideModal() {
    const openGuideBtn = document.getElementById("open-guide-btn");
    const guideModal = document.getElementById("guide-modal");
    const closeGuideModalBtn = document.getElementById("close-guide-modal-btn");
    const guideModalOkBtn = document.getElementById("guide-modal-ok-btn");

    if (openGuideBtn && guideModal) {
        openGuideBtn.addEventListener("click", () => guideModal.classList.remove("hidden"));
        if (closeGuideModalBtn) closeGuideModalBtn.addEventListener("click", () => guideModal.classList.add("hidden"));
        if (guideModalOkBtn) guideModalOkBtn.addEventListener("click", () => guideModal.classList.add("hidden"));
    }
}

function downloadSampleTemplate() {
    const sampleData = [
        { "ID": "101", "Invitado": "Juan Pérez", "Cantidad": 2, "Mesa": "Mesa 1", "VIP": "VIP", "Teléfono": "5512345678", "Notas": "Vegetariano", "QR": "INV-101", "Asistencia": "" },
        { "ID": "102", "Invitado": "María Rodríguez", "Cantidad": 1, "Mesa": "Mesa 1", "VIP": "VIP", "Teléfono": "5587654321", "Notas": "", "QR": "INV-102", "Asistencia": "" },
        { "ID": "103", "Invitado": "Carlos López", "Cantidad": 3, "Mesa": "Mesa 2", "VIP": "No", "Teléfono": "5533221100", "Notas": "", "QR": "INV-103", "Asistencia": "" },
        { "ID": "104", "Invitado": "Ana Martínez", "Cantidad": 2, "Mesa": "Mesa 2", "VIP": "No", "Teléfono": "5544556677", "Notas": "Alergia mariscos", "QR": "INV-104", "Asistencia": "" },
        { "ID": "105", "Invitado": "Luis García", "Cantidad": 1, "Mesa": "Mesa 3", "VIP": "No", "Teléfono": "5599887766", "Notas": "", "QR": "INV-105", "Asistencia": "" }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData, { header: ["ID", "Invitado", "Cantidad", "Mesa", "VIP", "Teléfono", "Notas", "QR", "Asistencia"] });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invitados");
    XLSX.writeFile(workbook, "Plantilla_Invitados_AccesoQR.xlsx");
}

function setupPwa() {
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js?v=8.0').catch(err => console.log(err));
    }
}

function setupModalEvents() {
    closeUpgradeModalBtn.addEventListener("click", () => {
        upgradeModal.classList.add("hidden");
    });
}

function showUpgradeModal(msg) {
    if (msg) upgradeModalMsg.textContent = msg;
    upgradeModal.classList.remove("hidden");
}

function setupFileLoaders() {
    dropZone.addEventListener("click", () => excelFileInput.click());
    
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    
    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });
    
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length) {
            handleExcelFile(e.dataTransfer.files[0]);
        }
    });
    
    excelFileInput.addEventListener("change", (e) => {
        if (e.target.files.length) {
            handleExcelFile(e.target.files[0]);
        }
    });

    resetFileBtn.addEventListener("click", () => {
        resetState();
    });

    downloadExcelBtn.addEventListener("click", () => {
        exportUpdatedExcel();
    });
}

function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            originalWorkbook = workbook;
            activeSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[activeSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            if (jsonData.length === 0) {
                alert("El archivo Excel está vacío.");
                return;
            }

            const firstRowKeys = Object.keys(jsonData[0]);
            originalColumnsOrder = firstRowKeys;

            let finalRows = jsonData;
            let limitReached = false;

            if (jsonData.length > MAX_TRIAL_GUESTS) {
                finalRows = jsonData.slice(0, MAX_TRIAL_GUESTS);
                limitReached = true;
            }

            guestData = finalRows.map((row, index) => {
                return {
                    id: String(row["ID"] || row["id"] || index + 1).trim(),
                    name: String(row["Invitado"] || row["invitado"] || row["Nombre"] || row["nombre"] || "Sin Nombre").trim(),
                    quantity: parseInt(row["Cantidad"] || row["cantidad"] || row["Pases"] || row["pases"] || 1),
                    qrValue: String(row["QR"] || row["qr"] || row["Codigo"] || "").trim(),
                    attendance: String(row["Asistencia"] || row["asistencia"] || "").trim(),
                    rawRow: row
                };
            });

            fileNameDisplay.textContent = file.name;
            dropZone.classList.add("hidden");
            fileInfo.classList.remove("hidden");
            
            searchInput.removeAttribute("disabled");
            toggleCameraBtn.removeAttribute("disabled");
            downloadExcelBtn.removeAttribute("disabled");

            renderTable(guestData);
            updateStats();

            if (limitReached) {
                showUpgradeModal(`⚠️ Tu archivo original contiene ${jsonData.length} invitados. En esta versión de demostración únicamente se cargaron los primeros 15 invitados. Para procesar tu lista completa, adquiere una Licencia PRO.`);
            }

        } catch (err) {
            console.error(err);
            alert("Error al leer el archivo Excel. Verifica que sea un formato válido (.xlsx).");
        }
    };
    reader.readAsArrayBuffer(file);
}

function resetState() {
    guestData = [];
    originalWorkbook = null;
    activeSheetName = "";
    originalColumnsOrder = [];
    stopScanner();

    dropZone.classList.remove("hidden");
    fileInfo.classList.add("hidden");
    excelFileInput.value = "";

    searchInput.setAttribute("disabled", "true");
    toggleCameraBtn.setAttribute("disabled", "true");
    downloadExcelBtn.setAttribute("disabled", "true");

    guestsTableBody.innerHTML = `
        <tr>
            <td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 2rem;">
                Carga un archivo Excel para ver la lista de prueba.
            </td>
        </tr>
    `;
    updateStats();
}

function updateStats() {
    const total = guestData.length;
    const presentCount = guestData.filter(g => g.attendance.toLowerCase().includes("asistio") || g.attendance.toLowerCase().includes("si") || g.attendance.toLowerCase().includes("presente")).length;
    const pendingCount = total - presentCount;

    statTotal.textContent = `${total}/${MAX_TRIAL_GUESTS}`;
    statPresent.textContent = presentCount;
    statPending.textContent = pendingCount;
}

function renderTable(dataToRender) {
    if (dataToRender.length === 0) {
        guestsTableBody.innerHTML = `
            <tr>
                <td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 1.5rem;">
                    No se encontraron coincidencias.
                </td>
            </tr>
        `;
        return;
    }

    guestsTableBody.innerHTML = dataToRender.map((guest, idx) => {
        const isPresent = guest.attendance.toLowerCase().includes("asistio") || guest.attendance.toLowerCase().includes("si") || guest.attendance.toLowerCase().includes("presente");
        const statusBadge = isPresent 
            ? `<span class="status-badge present">Asistió</span>`
            : `<span class="status-badge pending">Pendiente</span>`;

        const actionBtn = isPresent
            ? `<button class="btn btn-warning" onclick="toggleCheckIn('${guest.id}', false)" style="padding: 4px 8px; font-size: 0.75rem;">Desmarcar</button>`
            : `<button class="btn btn-primary" onclick="toggleCheckIn('${guest.id}', true)" style="padding: 4px 8px; font-size: 0.75rem;">Marcar Entrada</button>`;

        return `
            <tr>
                <td>${idx + 1}</td>
                <td><strong>${guest.name}</strong></td>
                <td>${guest.quantity}</td>
                <td>${statusBadge}</td>
                <td>${actionBtn}</td>
            </tr>
        `;
    }).join("");
}

function toggleCheckIn(guestId, markPresent) {
    const guest = guestData.find(g => String(g.id) === String(guestId));
    if (!guest) return;

    if (markPresent) {
        const now = new Date();
        const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        guest.attendance = `Asistió (${timeStr})`;
        if (typeof confetti === 'function') {
            confetti({
                particleCount: 55,
                spread: 60,
                origin: { y: 0.7 },
                colors: [currentTrialAccent, '#F5EDE6', '#34D399']
            });
        }
    } else {
        guest.attendance = "";
    }

    renderTable(filterData(searchInput.value));
    updateStats();
}

function setupSearch() {
    searchInput.addEventListener("input", (e) => {
        const query = e.target.value;
        renderTable(filterData(query));
    });
}

function filterData(query) {
    if (!query) return guestData;
    const q = query.toLowerCase().trim();
    return guestData.filter(g => 
        g.name.toLowerCase().includes(q) || 
        String(g.id).toLowerCase().includes(q) ||
        g.qrValue.toLowerCase().includes(q)
    );
}

// Camera Scanner Setup
function setupScanner() {
    toggleCameraBtn.addEventListener("click", () => {
        if (isScanning) {
            stopScanner();
        } else {
            startScanner();
        }
    });
}

function startScanner() {
    navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
        .then(stream => {
            videoStream = stream;
            videoPreview.srcObject = stream;
            videoPreview.setAttribute("playsinline", true);
            videoPreview.play();

            isScanning = true;
            scannerContainer.classList.remove("hidden");
            toggleCameraBtn.innerHTML = '<i class="ti ti-camera-off"></i> Detener Cámara';
            toggleCameraBtn.classList.replace("btn-primary", "btn-warning");

            scanInterval = setInterval(scanFrame, 250);
        })
        .catch(err => {
            console.error(err);
            alert("No se pudo acceder a la cámara. Asegúrate de dar los permisos correspondientes.");
        });
}

function stopScanner() {
    if (scanInterval) clearInterval(scanInterval);
    if (videoStream) {
        videoStream.getTracks().forEach(track => track.stop());
        videoStream = null;
    }
    isScanning = false;
    scannerContainer.classList.add("hidden");
    toggleCameraBtn.innerHTML = '<i class="ti ti-camera"></i> Activar Cámara';
    toggleCameraBtn.classList.replace("btn-warning", "btn-primary");
}

function scanFrame() {
    if (!isScanning || videoPreview.readyState !== videoPreview.HAVE_ENOUGH_DATA) return;

    const canvas = document.createElement("canvas");
    canvas.width = videoPreview.videoWidth;
    canvas.height = videoPreview.videoHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(videoPreview, 0, 0, canvas.width, canvas.height);
    
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const code = jsQR(imageData.data, imageData.width, imageData.height, {
        inversionAttempts: "dontInvert",
    });

    if (code && code.data) {
        processQrCode(code.data);
    }
}

function processQrCode(qrString) {
    const val = qrString.trim();
    const now = Date.now();
    if (val === lastScannedCode && (now - lastScannedTime) < 2800) {
        return;
    }
    lastScannedCode = val;
    lastScannedTime = now;

    const guest = guestData.find(g => 
        g.qrValue === val || 
        String(g.id) === val || 
        g.name.toLowerCase() === val.toLowerCase()
    );

    if (guest) {
        if (!guest.attendance) {
            toggleCheckIn(guest.id, true);
            alert(`✅ ASISTENCIA REGISTRADA:\n\nInvitado: ${guest.name}\nPases: ${guest.quantity}`);
        } else {
            alert(`⚠️ ASISTENCIA PREVIA:\n\n${guest.name} ya fue registrado anteriormente (${guest.attendance}).`);
        }
    } else {
        alert(`❌ CÓDIGO NO ENCONTRADO:\n\nEl código "${val}" no pertenece a la lista de prueba.`);
    }
}

function exportUpdatedExcel() {
    if (!guestData.length || !originalWorkbook) return;

    const updatedRows = guestData.map(g => {
        const row = { ...g.rawRow };
        // Buscar o crear la columna de Asistencia
        let attendanceColKey = originalColumnsOrder.find(k => k.toLowerCase() === "asistencia") || "Asistencia";
        row[attendanceColKey] = g.attendance;
        return row;
    });

    const newWorksheet = XLSX.utils.json_to_sheet(updatedRows, { header: originalColumnsOrder });
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, activeSheetName);
    
    XLSX.writeFile(newWorkbook, `Reporte_Asistencia_DEMO_${activeSheetName}.xlsx`);
}
