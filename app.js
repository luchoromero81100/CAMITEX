document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFile');
    const processBtn = document.getElementById('processBtn');
    const printBtn = document.getElementById('printBtn');
    const resetBtn = document.getElementById('resetBtn');
    const statusDiv = document.getElementById('status');
    const statsDiv = document.getElementById('stats');
    const printArea = document.getElementById('printArea');

    processBtn.addEventListener('click', processFile);
    printBtn.addEventListener('click', () => window.print());
    resetBtn.addEventListener('click', resetApp);

    /**
     * Da formato al precio recibido.
     * Ejemplo: 1990 -> "$ 1.990", 1234567 -> "$ 1.234.567"
     * Se ignora o redondea cualquier decimal, se utiliza un punto como separador de miles.
     */
    function formatPrice(value) {
        if (value === undefined || value === null || value === "") return "$ 0";
        
        let num = parseInt(value, 10);
        if (isNaN(num)) num = 0;
        
        // Convertir a string e insertar puntos cada 3 dígitos (separador de miles)
        let formattedStr = num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        return `$ ${formattedStr}`;
    }

    function showError(msg) {
        statusDiv.textContent = msg;
        statsDiv.textContent = "";
    }

    function showStats(msg) {
        statusDiv.textContent = "";
        statsDiv.textContent = msg;
    }

    function resetApp() {
        fileInput.value = "";
        statusDiv.textContent = "";
        statsDiv.textContent = "";
        printArea.innerHTML = "";
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
                // Utilizando SheetJS para interpretar el .xlsx
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Usamos siempre la primera hoja (Sheet) del libro
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convertir la hoja a un array de objetos (JSON)
                // defval: "" asegura que si una celda está vacía se devuelva como cadena vacía
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                
                if (jsonData.length === 0) {
                    showError("El archivo está vacío o no contiene datos tabulares válidos.");
                    return;
                }

                validateAndGenerateLabels(jsonData);
            } catch (err) {
                console.error(err);
                showError("Hubo un error al procesar el archivo. Asegúrate de que sea un archivo Excel (.xlsx) válido.");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function validateAndGenerateLabels(data) {
        // Validación de columnas obligatorias en base al primer objeto del JSON (que proviene del encabezado)
        const firstRow = data[0];
        const requiredCols = ['TEXTO', 'SKU', 'PRECIO', 'CANTIDAD'];
        const missingCols = requiredCols.filter(col => !(col in firstRow));
        
        if (missingCols.length > 0) {
            showError(`El Excel no tiene el formato correcto. Faltan las siguientes columnas: ${missingCols.join(', ')}`);
            return;
        }

        // 1. Expansión plana de las etiquetas de acuerdo a su 'CANTIDAD'
        const expandedLabels = [];

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            
            // Extracción de datos
            const texto = String(row["TEXTO"] || "").trim();
            const sku = String(row["SKU"] || "").trim();
            const precio = row["PRECIO"];
            let cantidad = parseInt(row["CANTIDAD"], 10);

            // Casos límite: ignorar filas vacías de SKU/Texto y las que tienen cantidad inválida (<= 0 o NaN)
            if (isNaN(cantidad) || cantidad <= 0) continue;
            if (!sku && !texto) continue;

            // Expandir tantas veces como indique la columna 'CANTIDAD'
            for (let j = 0; j < cantidad; j++) {
                expandedLabels.push({
                    texto: texto,
                    sku: sku,
                    precio: formatPrice(precio)
                });
            }
        }

        if (expandedLabels.length === 0) {
            showError("Ningún producto cumple con los criterios. Verifica que CANTIDAD sea mayor a 0 y que haya SKUs/Textos.");
            return;
        }

        // 2. Agrupamiento en pares de izquierda (left) y derecha (right)
        const pairs = [];
        for (let i = 0; i < expandedLabels.length; i += 2) {
            pairs.push({
                left: expandedLabels[i],
                right: (i + 1 < expandedLabels.length) ? expandedLabels[i + 1] : null // null cuando tenemos un final impar
            });
        }

        // 3. Renderizado del HTML en el DOM
        printArea.innerHTML = "";
        let svgIdCounter = 0; // Utilizado para inyectar IDs únicos para los barcodes

        pairs.forEach((pair) => {
            // Contenedor principal de la fila (ancho total: 103mm)
            const rowDiv = document.createElement('div');
            rowDiv.className = 'label-row';

            // Etiqueta Izquierda
            const leftLabel = createLabelElement(pair.left, svgIdCounter++);
            rowDiv.appendChild(leftLabel);

            // Brecha central obligatoria (3mm)
            const gapDiv = document.createElement('div');
            gapDiv.className = 'label-gap';
            rowDiv.appendChild(gapDiv);

            // Etiqueta Derecha o Espacio vacío (si es impar)
            if (pair.right) {
                const rightLabel = createLabelElement(pair.right, svgIdCounter++);
                rowDiv.appendChild(rightLabel);
            } else {
                // EXTREMADAMENTE IMPORTANTE: Placeholder vacío
                // No se centra automáticamente la izquierda. El lado derecho se mantiene 100% en blanco.
                const placeholder = document.createElement('div');
                placeholder.className = 'label-placeholder';
                rowDiv.appendChild(placeholder);
            }

            printArea.appendChild(rowDiv);
        });

        // 4. Invocación de JsBarcode sobre todos los SVG inyectados dinámicamente
        // Se hace en este paso porque los perfiles SVG recién ahora existen en el DOM
        expandedLabels.forEach((label, i) => {
            if (label.sku) {
                try {
                    const svgNode = document.querySelector(`#barcode-${i}`);
                    // Genera código CODE128
                    JsBarcode(svgNode, label.sku, {
                        format: "CODE128", 
                        displayValue: false, // Oculamos el valor de JsBarcode, usamos el texto renderizado abajo
                        margin: 0,
                        width: 1.5, // Ancho de barra que permite acomodar strings de largo moderado dentro de 50mm
                        height: 35  // Altura del código de barras
                    });
                    // IMPORTANTÍSIMO: Forza que al encoger horizontalmente, NO se encoja verticalmente.
                    // Manteniendo siempre los 12mm de altura estipulados en el CSS, sin importar lo largo del SKU.
                    svgNode.setAttribute('preserveAspectRatio', 'none');
                    // Fuerza renderizado de píxeles puros para evitar borrosidad (anti-aliasing)
                    svgNode.setAttribute('shape-rendering', 'crispEdges');
                } catch (e) {
                    console.error("Error al generar barcode para el SKU:", label.sku, e);
                    // Si falla silenciosamente generará una etiqueta en blanco, lo cual es manejable por el usuario.
                }
            }
        });

        // Listo, habilitar la impresión
        showStats(`Éxito: Se han generado ${expandedLabels.length} etiquetas distribuidas en ${pairs.length} fila/s preparadas para impresión.`);
        printBtn.disabled = false;
    }

    /**
     * Construye individualmente una etiqueta con todos sus elementos (50x24.5mm)
     * @param {Object} data - Datos {texto, sku, precio}
     * @param {Number} idIndex - Índice numérico para el ID único del SVG
     * @returns {HTMLElement}
     */
    function createLabelElement(data, idIndex) {
        const wrapper = document.createElement('div');
        wrapper.className = 'label';

        const title = document.createElement('p');
        title.className = 'label-title';
        title.textContent = data.texto;

        const barcodeContainer = document.createElement('div');
        barcodeContainer.className = 'label-barcode-container';
        
        // Elemento SVG nativo para el código de barras
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.id = `barcode-${idIndex}`;
        barcodeContainer.appendChild(svg);

        const skuText = document.createElement('p');
        skuText.className = 'label-sku';
        skuText.textContent = data.sku; // SKUs legibles en monospace sin el JsBarcode

        const priceText = document.createElement('p');
        priceText.className = 'label-price';
        priceText.textContent = data.precio; // Muestra algo tipo "$ 1.990"

        // Ensamblado según el orden definido en la etiqueta (Top a Bottom)
        wrapper.appendChild(title);
        wrapper.appendChild(barcodeContainer);
        wrapper.appendChild(skuText);
        wrapper.appendChild(priceText);

        return wrapper;
    }
});
