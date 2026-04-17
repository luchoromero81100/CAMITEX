document.addEventListener('DOMContentLoaded', () => {

    // =========================================================
    // REFERENCIAS AL DOM
    // =========================================================
    // Pantallas (Ventanas)
    const screenHome   = document.getElementById('screen-home');
    const screenMasive = document.getElementById('screen-masive');
    const screenSingle = document.getElementById('screen-single');

    // Navegación HOME
    const btnGoMasive  = document.getElementById('btnGoMasive');
    const btnGoSingle  = document.getElementById('btnGoSingle');
    const btnBackMasive = document.getElementById('btnBackMasive');
    const btnBackSingle = document.getElementById('btnBackSingle');

    // Controles pantalla MASIVO
    const fileInput  = document.getElementById('excelFile');
    const processBtn = document.getElementById('processBtn');
    const printBtn   = document.getElementById('printBtn');
    const resetBtn   = document.getElementById('resetBtn');
    const statusDiv  = document.getElementById('status');
    const statsDiv   = document.getElementById('stats');
    const masiveStatusBar = document.getElementById('masiveStatusBar');

    // Controles pantalla ÚNICO
    const singleNombre      = document.getElementById('singleNombre');
    const singleSKU         = document.getElementById('singleSKU');
    const singlePrecio      = document.getElementById('singlePrecio');
    const singleGenerateBtn = document.getElementById('singleGenerateBtn');
    const singlePrintBtn    = document.getElementById('singlePrintBtn');
    const singleResetBtn    = document.getElementById('singleResetBtn');
    const singleStatus      = document.getElementById('singleStatus');

    // Previsualización Desktop
    const previewPanel = document.getElementById('previewPanel');
    const printArea = document.getElementById('printArea');
    const taskbarClock = document.getElementById('taskbarClock');

    // =========================================================
    // CLOCK TASKBAR
    // =========================================================
    function updateClock() {
        const now = new Date();
        const hrs = String(now.getHours()).padStart(2, '0');
        const min = String(now.getMinutes()).padStart(2, '0');
        taskbarClock.textContent = `${hrs}:${min}`;
    }
    setInterval(updateClock, 1000);
    updateClock();

    // =========================================================
    // NAVEGACIÓN ENTRE VENTANAS
    // =========================================================
    function showScreen(screen) {
        [screenHome, screenMasive, screenSingle].forEach(s => s.classList.remove('active'));
        screen.classList.add('active');
        
        // Ocultar previsualización al volver al inicio
        if (screen === screenHome) {
            previewPanel.style.display = 'none';
        }
    }

    btnGoMasive.addEventListener('click', () => {
        clearPrintArea();
        showScreen(screenMasive);
    });

    btnGoSingle.addEventListener('click', () => {
        clearPrintArea();
        showScreen(screenSingle);
    });

    btnBackMasive.addEventListener('click', () => {
        resetMasive();
        showScreen(screenHome);
    });

    btnBackSingle.addEventListener('click', () => {
        resetSingle();
        showScreen(screenHome);
    });

    // =========================================================
    // UTILIDADES COMPARTIDAS
    // =========================================================

    function formatPrice(value) {
        if (value === undefined || value === null || value === "") return "$ 0";
        let num = parseInt(value, 10);
        if (isNaN(num)) num = 0;
        let formattedStr = num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        return `$ ${formattedStr}`;
    }

    function clearPrintArea() {
        printArea.innerHTML = "";
        previewPanel.style.display = 'none';
    }

    // =========================================================
    // GENERADOR DE ETIQUETAS
    // =========================================================
    function renderLabels(expandedLabels) {
        printArea.innerHTML = "";
        previewPanel.style.display = 'block';
        let svgIdCounter = 0;

        const pairs = [];
        for (let i = 0; i < expandedLabels.length; i += 2) {
            pairs.push({
                left: expandedLabels[i],
                right: (i + 1 < expandedLabels.length) ? expandedLabels[i + 1] : null
            });
        }

        pairs.forEach((pair) => {
            const rowDiv = document.createElement('div');
            rowDiv.className = 'label-row';

            const leftLabel = createLabelElement(pair.left, svgIdCounter++);
            rowDiv.appendChild(leftLabel);

            const gapDiv = document.createElement('div');
            gapDiv.className = 'label-gap';
            rowDiv.appendChild(gapDiv);

            if (pair.right) {
                const rightLabel = createLabelElement(pair.right, svgIdCounter++);
                rowDiv.appendChild(rightLabel);
            } else {
                const placeholder = document.createElement('div');
                placeholder.className = 'label-placeholder';
                rowDiv.appendChild(placeholder);
            }

            printArea.appendChild(rowDiv);
        });

        // Generar barcodes
        expandedLabels.forEach((label, i) => {
            if (label.sku) {
                try {
                    const svgNode = document.querySelector(`#barcode-${i}`);
                    JsBarcode(svgNode, label.sku, {
                        format: "CODE128",
                        displayValue: false,
                        margin: 0,
                        width: 1.5,
                        height: 35
                    });
                    svgNode.setAttribute('preserveAspectRatio', 'none');
                    svgNode.setAttribute('shape-rendering', 'crispEdges');
                } catch (e) {
                    console.error("Error al generar barcode para SKU:", label.sku, e);
                }
            }
        });

        return pairs.length;
    }

    function createLabelElement(data, idIndex) {
        const wrapper = document.createElement('div');
        wrapper.className = 'label';

        const title = document.createElement('p');
        title.className = 'label-title';
        title.textContent = data.texto;

        const barcodeContainer = document.createElement('div');
        barcodeContainer.className = 'label-barcode-container';

        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.id = `barcode-${idIndex}`;
        barcodeContainer.appendChild(svg);

        const skuText = document.createElement('p');
        skuText.className = 'label-sku';
        skuText.textContent = data.sku;

        const priceText = document.createElement('p');
        priceText.className = 'label-price';
        priceText.textContent = data.precio;

        wrapper.appendChild(title);
        wrapper.appendChild(barcodeContainer);
        wrapper.appendChild(skuText);
        wrapper.appendChild(priceText);

        return wrapper;
    }

    // =========================================================
    // LÓGICA MASIVO
    // =========================================================
    processBtn.addEventListener('click', processFile);
    printBtn.addEventListener('click', () => window.print());
    resetBtn.addEventListener('click', resetMasive);

    function showError(msg) {
        statusDiv.textContent = msg;
        statsDiv.textContent  = "";
        if (masiveStatusBar) masiveStatusBar.textContent = "Error al procesar.";
    }

    function showStats(msg) {
        statusDiv.textContent = "";
        statsDiv.textContent  = msg;
        if (masiveStatusBar) masiveStatusBar.textContent = "Proceso completado.";
    }

    function resetMasive() {
        fileInput.value       = "";
        statusDiv.textContent = "";
        statsDiv.textContent  = "";
        if (masiveStatusBar) masiveStatusBar.textContent = "Esperando archivo…";
        clearPrintArea();
        printBtn.disabled = true;
    }

    function processFile() {
        const file = fileInput.files[0];
        if (!file) {
            showError("Por favor, selecciona un archivo Excel primero.");
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                if (jsonData.length === 0) {
                    showError("El archivo está vacío.");
                    return;
                }

                validateAndGenerateLabels(jsonData);
            } catch (err) {
                console.error(err);
                showError("Error al procesar archivo.");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function validateAndGenerateLabels(data) {
        const firstRow = data[0];
        const requiredCols = ['TEXTO', 'SKU', 'PRECIO', 'CANTIDAD'];
        const missingCols = requiredCols.filter(col => !(col in firstRow));

        if (missingCols.length > 0) {
            showError(`Faltan columnas: ${missingCols.join(', ')}`);
            return;
        }

        const expandedLabels = [];
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const texto  = String(row["TEXTO"]  || "").trim();
            const sku    = String(row["SKU"]    || "").trim();
            const precio = row["PRECIO"];
            let cantidad = parseInt(row["CANTIDAD"], 10);

            if (isNaN(cantidad) || cantidad <= 0) continue;
            if (!sku && !texto) continue;

            for (let j = 0; j < cantidad; j++) {
                expandedLabels.push({ texto, sku, precio: formatPrice(precio) });
            }
        }

        if (expandedLabels.length === 0) {
            showError("Sin etiquetas para generar.");
            return;
        }

        const totalPairs = renderLabels(expandedLabels);
        showStats(`Generadas ${expandedLabels.length} etiquetas.`);
        printBtn.disabled = false;
    }

    // =========================================================
    // LÓGICA ÚNICO
    // =========================================================
    singleGenerateBtn.addEventListener('click', generateSingleLabel);
    singlePrintBtn.addEventListener('click', () => window.print());
    singleResetBtn.addEventListener('click', resetSingle);

    function generateSingleLabel() {
        const nombre = singleNombre.value.trim();
        const sku    = singleSKU.value.trim();
        const precio = singlePrecio.value.trim();

        if (!nombre && !sku) {
            singleStatus.textContent = "Completá Nombre o SKU.";
            return;
        }

        singleStatus.textContent = "";
        const labelData = [{
            texto: nombre,
            sku:   sku,
            precio: formatPrice(precio)
        }];

        renderLabels(labelData);
        singlePrintBtn.disabled = false;
    }

    function resetSingle() {
        singleNombre.value       = "";
        singleSKU.value          = "";
        singlePrecio.value       = "";
        singleStatus.textContent = "";
        clearPrintArea();
        singlePrintBtn.disabled = true;
    }
});
