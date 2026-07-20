// ==========================================
// 1. VARIABLES GLOBALES Y LECTURA DE EXCEL
// ==========================================
let valoresX = [];
let valoresY = [];

document.getElementById('excel-file').addEventListener('change', function(evento) {
    const archivo = evento.target.files[0];
    if (!archivo) return;
    const lector = new FileReader();
    lector.onload = function(e) {
        const datos = new Uint8Array(e.target.result);
        const libro = XLSX.read(datos, {type: 'array'});
        const hoja = libro.Sheets[libro.SheetNames[0]];
        
        // CORRECCIÓN: Matriz pura
        const datosJSON = XLSX.utils.sheet_to_json(hoja, {header: 1});
        procesarYMostrarDatos(datosJSON);
    };
    lector.readAsArrayBuffer(archivo);
});

function procesarYMostrarDatos(datos) {
    valoresX = [];
    valoresY = [];
    let tablaHTML = '<h3>Datos Cargados:</h3><table><tr><th>X (Columna 1)</th><th>Y (Columna 2)</th></tr>';
    
    datos.forEach(fila => {
        if (fila.length >= 2) {
            let x = parseFloat(fila[0]);
            let y = parseFloat(fila[1]);
            
            if (!isNaN(x) && !isNaN(y)) {
                valoresX.push(x);
                valoresY.push(y);
                tablaHTML += `<tr><td>${x}</td><td>${y}</td></tr>`;
            }
        }
    });
    tablaHTML += '</table>';
    document.getElementById('data-preview').innerHTML = tablaHTML;
    document.getElementById('resultado-matematico').innerHTML = ''; 
}

// ==========================================
// 2. FUNCIÓN GENERAL DE GRÁFICOS (PLOTLY)
// ==========================================
function dibujarGrafico(titulo, xOriginal, yOriginal, xCurva, yCurva) {
    let trazoOriginal = {
        x: xOriginal, y: yOriginal, mode: 'markers', type: 'scatter', name: 'Datos Originales',
        marker: { size: 8, color: 'rgba(225, 29, 72, 0.6)', line: { color: '#e11d48', width: 1 } }
    };
    let trazoCurva = {
        x: xCurva, y: yCurva, mode: 'lines', type: 'scatter', name: 'Curva Aproximada',
        line: { color: '#2563eb', width: 3 }
    };
    let diseño = {
        title: titulo, xaxis: { title: 'Eje X' }, yaxis: { title: 'Eje Y' },
        plot_bgcolor: '#f9fafb', paper_bgcolor: '#ffffff', scattermode: 'group'
    };
    Plotly.newPlot('plot-container', [trazoOriginal, trazoCurva], diseño);
}

// ==========================================
// 3. INTERPOLACIÓN DE LAGRANGE
// ==========================================
function evaluarLagrange(x, arrX, arrY) {
    let resultado = 0;
    for (let i = 0; i < arrX.length; i++) {
        let termino = arrY[i];
        for (let j = 0; j < arrX.length; j++) {
            if (i !== j) termino = termino * (x - arrX[j]) / (arrX[i] - arrX[j]);
        }
        resultado += termino;
    }
    return resultado;
}

function obtenerPolinomioLagrangeTexto(arrX, arrY) {
    let terminos = [];
    for (let i = 0; i < arrX.length; i++) {
        let numerador = "";
        let denominadorValor = 1;
        for (let j = 0; j < arrX.length; j++) {
            if (i !== j) {
                let valorXj = arrX[j];
                numerador += valorXj < 0 ? `(x + ${Math.abs(valorXj)})` : (valorXj > 0 ? `(x - ${valorXj})` : `x`);
                denominadorValor *= (arrX[i] - arrX[j]);
            }
        }
        let coeficiente = (arrY[i] / denominadorValor).toFixed(4);
        if (parseFloat(coeficiente) !== 0) {
            let signo = terminos.length > 0 ? (parseFloat(coeficiente) >= 0 ? " + " : " - ") : (parseFloat(coeficiente) < 0 ? "-" : "");
            terminos.push(`${signo}${Math.abs(parseFloat(coeficiente))} · ${numerador}`);
        }
    }
    return "P(x) = " + terminos.join("");
}
document.getElementById('btn-lagrange').addEventListener('click', function() {
    if (valoresX.length === 0) return alert("Sube un archivo Excel primero.");

    // --- NUEVA ALERTA DE PREVENCIÓN (FENÓMENO DE RUNGE) ---
    if (valoresX.length >= 15) {
        alert("⚠️ Advertencia Académica: Has cargado 15 o más datos. Forzar un único polinomio de Lagrange de grado tan alto generará oscilaciones extremas en los extremos (Fenómeno de Runge).\n\n💡 Recomendación: Utiliza el método de 'Splines Cúbicos' para garantizar la estabilidad de la curva.");
    }
    // ------------------------------------------------------

    let xMin = Math.min(...valoresX), xMax = Math.max(...valoresX);
    let ptosX = [], ptosY = [], paso = (xMax - xMin) / 100;
    
    for (let i = 0; i <= 100; i++) {
        let xAct = xMin + (i * paso);
        ptosX.push(xAct); ptosY.push(evaluarLagrange(xAct, valoresX, valoresY));
    }
    
    dibujarGrafico('Interpolación de Lagrange', valoresX, valoresY, ptosX, ptosY);
    
    document.getElementById('resultado-matematico').innerHTML = `
        <h3 style="text-align: center; color: #1f2937;">Polinomio de Lagrange:</h3>
        <p style="background: white; padding: 15px; border: 1px solid #e2e8f0; overflow-x: auto; text-align: center;">
            <strong>${obtenerPolinomioLagrangeTexto(valoresX, valoresY)}</strong>
        </p>`;
});

// ==========================================
// 4. DIFERENCIAS CENTRADAS
// ==========================================
document.getElementById('btn-diferencias').addEventListener('click', function() {
    if (valoresX.length < 2) return alert("Se requieren al menos 2 puntos.");
    let derivadas = [], n = valoresX.length;
    for (let i = 0; i < n; i++) {
        if (i === 0) derivadas.push((valoresY[1] - valoresY[0]) / (valoresX[1] - valoresX[0]));
        else if (i === n - 1) derivadas.push((valoresY[n - 1] - valoresY[n - 2]) / (valoresX[n - 1] - valoresX[n - 2]));
        else derivadas.push((valoresY[i + 1] - valoresY[i - 1]) / (valoresX[i + 1] - valoresX[i - 1]));
    }
    let trazoOrg = { x: valoresX, y: valoresY, mode: 'lines+markers', name: 'Datos Originales', line: { color: '#e11d48' } };
    let trazoDer = { x: valoresX, y: derivadas, mode: 'lines+markers', name: "Derivada f'(x)", line: { color: '#059669', dash: 'dash' } };
    Plotly.newPlot('plot-container', [trazoOrg, trazoDer], { title: 'Diferencias Centradas' });
    
    let html = `<h3>Derivadas Aproximadas:</h3><table style="width:100%; text-align:center; border:1px solid #ccc;"><tr><th>x</th><th>f'(x)</th></tr>`;
    for(let i=0; i<n; i++) html += `<tr><td>${valoresX[i]}</td><td style="color:#059669; font-weight:bold;">${derivadas[i].toFixed(4)}</td></tr>`;
    document.getElementById('resultado-matematico').innerHTML = html + `</table>`;
});

// ==========================================
// 5. INTEGRACIÓN (TRAPECIO Y SIMPSON)
// ==========================================
function dibujarArea(titulo, xD, yD, area, error = null) {
    if (error) {
        Plotly.purge('plot-container');
        document.getElementById('resultado-matematico').innerHTML = `<div style="color:red; padding:15px; border:1px solid red;">⚠️ ${error}</div>`;
        return;
    }
    let trazo = { x: xD, y: yD, fill: 'tozeroy', type: 'scatter', mode: 'lines+markers', line: { color: '#8b5cf6' }, fillcolor: 'rgba(139, 92, 246, 0.3)' };
    Plotly.newPlot('plot-container', [trazo], { title: titulo });
    document.getElementById('resultado-matematico').innerHTML = `<h3 style="text-align:center;">Área Aproximada: <strong>${area.toFixed(6)} u²</strong></h3>`;
}

document.getElementById('btn-trapecio').addEventListener('click', function() {
    let n = valoresX.length - 1, h = (valoresX[n] - valoresX[0]) / n, suma = valoresY[0] + valoresY[n];
    for (let i = 1; i < n; i++) suma += 2 * valoresY[i];
    dibujarArea('Regla del Trapecio', valoresX, valoresY, (h / 2) * suma);
});

document.getElementById('btn-simpson').addEventListener('click', function() {
    let n = valoresX.length - 1;
    if (n % 2 !== 0) return dibujarArea('', [], [], 0, "Simpson 1/3 requiere una cantidad IMPAR de datos (intervalos pares).");
    let h = (valoresX[n] - valoresX[0]) / n, suma = valoresY[0] + valoresY[n];
    for (let i = 1; i < n; i++) suma += (i % 2 === 0 ? 2 : 4) * valoresY[i];
    dibujarArea('Simpson 1/3', valoresX, valoresY, (h / 3) * suma);
});

// ==========================================
// 6. SPLINES CÚBICOS NATURALES
// ==========================================
function calcularSplines(arrX, arrY) {
    let n = arrX.length - 1, a = [...arrY], h = new Array(n);
    for (let i = 0; i < n; i++) h[i] = arrX[i + 1] - arrX[i];
    let alpha = new Array(n).fill(0);
    for (let i = 1; i < n; i++) alpha[i] = (3 / h[i]) * (a[i + 1] - a[i]) - (3 / h[i - 1]) * (a[i] - a[i - 1]);
    let l = new Array(n + 1).fill(1), mu = new Array(n + 1).fill(0), z = new Array(n + 1).fill(0);
    for (let i = 1; i < n; i++) {
        l[i] = 2 * (arrX[i + 1] - arrX[i - 1]) - h[i - 1] * mu[i - 1];
        mu[i] = h[i] / l[i];
        z[i] = (alpha[i] - h[i - 1] * z[i - 1]) / l[i];
    }
    let c = new Array(n + 1).fill(0), b = new Array(n).fill(0), d = new Array(n).fill(0);
    l[n] = 1; z[n] = 0; c[n] = 0;
    for (let j = n - 1; j >= 0; j--) {
        c[j] = z[j] - mu[j] * c[j + 1];
        b[j] = (a[j + 1] - a[j]) / h[j] - h[j] * (c[j + 1] + 2 * c[j]) / 3;
        d[j] = (c[j + 1] - c[j]) / (3 * h[j]);
    }
    let coef = [];
    for (let i = 0; i < n; i++) coef.push({ a: a[i], b: b[i], c: c[i], d: d[i], x_i: arrX[i] });
    return coef;
}

function evaluarS(x, coef, arrX) {
    let n = coef.length, i = 0;
    for (i = 0; i < n - 1; i++) if (x >= arrX[i] && x <= arrX[i + 1]) break;
    if (x > arrX[n]) i = n - 1;
    let dx = x - coef[i].x_i;
    return coef[i].a + coef[i].b * dx + coef[i].c * Math.pow(dx, 2) + coef[i].d * Math.pow(dx, 3);
}

document.getElementById('btn-spline').addEventListener('click', function() {
    if (valoresX.length < 3) return alert("Se requieren al menos 3 puntos.");
    
    // Ordenamos los datos para evitar fallos matemáticos
    let pts = valoresX.map((x, idx) => ({ x: x, y: valoresY[idx] })).sort((p1, p2) => p1.x - p2.x);
    let xOrd = pts.map(p => p.x), yOrd = pts.map(p => p.y);

    let coefs = calcularSplines(xOrd, yOrd);
    let xMin = xOrd[0], xMax = xOrd[xOrd.length - 1];
    let ptosX = [], ptosY = [], paso = (xMax - xMin) / 150; 

    for (let i = 0; i <= 150; i++) {
        let xAct = i === 150 ? xMax : xMin + (i * paso);
        ptosX.push(xAct); ptosY.push(evaluarS(xAct, coefs, xOrd));
    }

    dibujarGrafico('Splines Cúbicos Naturales', xOrd, yOrd, ptosX, ptosY);

    let tablaE = `<h3 style="text-align:center;">Polinomios Fragmentados:</h3><table style="width:100%; border-collapse:collapse; border:1px solid #ccc; text-align:center;"><tr><th style="border:1px solid #ccc;">Intervalo</th><th style="border:1px solid #ccc;">Polinomio Cúbico S_i(x)</th></tr>`;
    for (let i = 0; i < coefs.length; i++) {
        let c = coefs[i], dx = c.x_i === 0 ? "x" : `(x - ${c.x_i})`;
        let eq = `${c.a.toFixed(4)} ${c.b>=0?'+':'-'} ${Math.abs(c.b).toFixed(4)}${dx} ${c.c>=0?'+':'-'} ${Math.abs(c.c).toFixed(4)}${dx}² ${c.d>=0?'+':'-'} ${Math.abs(c.d).toFixed(4)}${dx}³`;
        tablaE += `<tr><td style="border:1px solid #ccc; padding:5px;">[${xOrd[i]}, ${xOrd[i+1]}]</td><td style="border:1px solid #ccc; font-family:monospace; font-size:12px;">S_${i}(x) = ${eq}</td></tr>`;
    }
    document.getElementById('resultado-matematico').innerHTML = tablaE + `</table>`;
});
