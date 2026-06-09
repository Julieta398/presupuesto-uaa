// =============================================
// CONFIGURACION GRAPH API
// =============================================
const CONFIG = {
    clientId: "8e219d60-b3b3-4bee-940f-cf1e303b44e2",
    tenantId: "cf6d7bb3-e5d3-4a5b-ad2a-c20744ff225a",
    driveId: "b!6e8IVbC020eZ80xzh-guvFGEGZPPitlAtzLHyskTtba5ocGAk2jbTJqLGUIDyi3c",
    fileId: "013GIXEMTDO7AQ3XPDANGZHNF77YA4RJPO",
    sheetName: "PRODUCTOS_AGRÍCOLAS"
};

const COLUMNAS = {
    GERENCIA: 0,
    SECCION: 1,
    SUCURSAL: 2,
    GRUPO: 3,
    SUBGRUPO: 4,
    MES: 5,
    CUENTA: 6,
    CUENTA_CONTABLE: 7,
    CUENTA_GENERAL: 8,
    DESCRIPCION: 9,
    SALDO_2223: 10,
    SALDO_2324: 11,
    SALDO_2425: 12,
    PROM_NOMINAL: 13,
    PROM_VENTAS: 14,
    PVTAS_2223: 15,
    PVTAS_2324: 16,
    PVTAS_2425: 17,
    PROM_PVTAS: 18,
    PROYECCION: 19,
    PROM_PVTAS_RECONFIG: 20,
    PROYECCION_RECONFIG: 21,
    PROM_PVTAS_RESULT: 22,
    PROYECCION_RESULT: 23,
    FILA_EXCEL: 24
};

// =============================================
// VARIABLES GLOBALES
// =============================================
let datosBrutos = [];
let datosBrutosResultados = [];
let datosFiltrados = [];
let msalInstance = null;
let accessToken = null;
let modoResultados = false;

// =============================================
// MSAL - AUTENTICACION
// =============================================
const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: "https://julieta398.github.io/presupuesto-uaa/"
    }
};

const loginRequest = {
    scopes: ["Files.ReadWrite", "User.Read"]
};

// =============================================
// INICIALIZACION
// =============================================
document.addEventListener("DOMContentLoaded", async () => {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    await msalInstance.initialize();
    await msalInstance.handleRedirectPromise();

    document.getElementById("btn-login").addEventListener("click", iniciarSesion);
    document.getElementById("btn-filtrar").addEventListener("click", aplicarFiltros);
    document.getElementById("btn-limpiar").addEventListener("click", limpiarFiltros);
    document.getElementById("btn-guardar").addEventListener("click", guardarCambios);
    document.getElementById("btn-modo").addEventListener("click", toggleModo);

    document.getElementById("filtro-gerencia").addEventListener("change", () => actualizarCascada("gerencia"));
    document.getElementById("filtro-seccion").addEventListener("change", () => actualizarCascada("seccion"));
    document.getElementById("filtro-sucursal").addEventListener("change", () => actualizarCascada("sucursal"));
    document.getElementById("filtro-mes").addEventListener("change", () => actualizarCascada("mes"));
    document.getElementById("filtro-grupo").addEventListener("change", () => actualizarCascada("grupo"));
    document.getElementById("filtro-subgrupo").addEventListener("change", () => actualizarCascada("subgrupo"));
    document.getElementById("filtro-cuenta").addEventListener("change", () => actualizarCascada("cuenta"));

    const cuentas = msalInstance.getAllAccounts();
    if (cuentas.length > 0) {
        await obtenerToken();
        cargarDatosDesdeGraph();
    }
});

// =============================================
// LOGIN
// =============================================
async function iniciarSesion() {
    try {
        const btn = document.getElementById("btn-login");
        btn.disabled = true;
        btn.textContent = "Iniciando sesión...";
        document.getElementById("mensaje-inicial").textContent = "Redirigiendo a Microsoft para iniciar sesión...";
        await new Promise(resolve => setTimeout(resolve, 800));
        await msalInstance.loginRedirect(loginRequest);
    } catch (err) {
        console.error("Error al iniciar sesión:", err);
        const btn = document.getElementById("btn-login");
        btn.disabled = false;
        btn.textContent = "Iniciar sesión con Microsoft";
        alert("No se pudo iniciar sesión.");
    }
}

async function obtenerToken() {
    const cuentas = msalInstance.getAllAccounts();
    if (cuentas.length === 0) return;

    try {
        const resultado = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: cuentas[0]
        });
        accessToken = resultado.accessToken;
        document.getElementById("btn-login").style.display = "none";
        document.getElementById("usuario").textContent = cuentas[0].username;
        document.getElementById("info-usuario").style.display = "block";
    } catch (err) {
        await msalInstance.acquireTokenRedirect(loginRequest);
    }
}

// =============================================
// LEER DATOS DESDE GRAPH API
// =============================================
async function cargarDatosDesdeGraph(paraResultados = false) {
    if (!paraResultados) {
        document.getElementById("mensaje-inicial").textContent = "Cargando datos...";
    }

    try {
        const url = `https://graph.microsoft.com/v1.0/drives/${CONFIG.driveId}/items/${CONFIG.fileId}/workbook/worksheets('${CONFIG.sheetName}')/usedRange`;

        const response = await fetch(url, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        if (!response.ok) throw new Error("Error al leer el archivo Excel.");

        const data = await response.json();
        const filas = data.values;

        const datos = filas.slice(1).map((fila, index) => {
            const filaCopia = [...fila];
            filaCopia[COLUMNAS.FILA_EXCEL] = index + 2;
            return filaCopia;
        });

        if (paraResultados) {
            datosBrutosResultados = datos;
        } else {
            datosBrutos = datos;
            if (datosBrutos.length === 0 || document.getElementById("filtro-gerencia").options.length <= 1) {
                inicializarFiltros();
            }
            document.getElementById("mensaje-inicial").textContent =
                "Aplicá los filtros y hacé clic en Filtrar para ver los datos.";
        }

    } catch (err) {
        console.error(err);
        document.getElementById("mensaje-inicial").textContent =
            "Error al cargar los datos: " + err.message;
    }
}

// =============================================
// TOGGLE MODO
// =============================================
async function toggleModo() {
    const btn = document.getElementById("btn-modo");

    if (!modoResultados) {
        // Cambiar a modo resultados
        btn.disabled = true;
        btn.textContent = "Cargando resultados...";

        await cargarDatosDesdeGraph(true);

        modoResultados = true;
        btn.textContent = "Volver a edición";
        btn.classList.add("modo-resultados");
        document.getElementById("btn-guardar").style.display = "none";

        // Actualizar headers
        document.getElementById("th-nuevo-pvtas").style.display = "none";
        document.getElementById("th-nuevo-usd").style.display = "none";
        document.getElementById("th-result-pvtas").style.display = "";
        document.getElementById("th-result-usd").style.display = "";

        btn.disabled = false;
    } else {
        // Volver a modo edición
        modoResultados = false;
        btn.textContent = "Ver datos modificados";
        btn.classList.remove("modo-resultados");
        document.getElementById("btn-guardar").style.display = "";

        // Actualizar headers
        document.getElementById("th-nuevo-pvtas").style.display = "";
        document.getElementById("th-nuevo-usd").style.display = "";
        document.getElementById("th-result-pvtas").style.display = "none";
        document.getElementById("th-result-usd").style.display = "none";
    }

    aplicarFiltros();
}

// =============================================
// GUARDAR CAMBIOS VIA GRAPH API
// =============================================
async function guardarCambios() {
    const cambios = [];

    document.querySelectorAll(".input-pvtas, .input-usd").forEach(input => {
        if (input.value !== "") {
            const index = parseInt(input.dataset.index);
            const fila = datosFiltrados[index];
            const esPvtas = input.classList.contains("input-pvtas");

            cambios.push({
                filaExcel: fila[COLUMNAS.FILA_EXCEL],
                columna: esPvtas ? 20 : 21,
                valor: parsearNumero(input.value)
            });
        }
    });

    if (cambios.length === 0) {
        alert("No hay cambios para guardar.");
        return;
    }

    const btnGuardar = document.getElementById("btn-guardar");
    btnGuardar.disabled = true;
    btnGuardar.textContent = "Guardando...";

    try {
        for (const cambio of cambios) {
            const celda = columnToA1(cambio.columna, cambio.filaExcel);
            const url = `https://graph.microsoft.com/v1.0/drives/${CONFIG.driveId}/items/${CONFIG.fileId}/workbook/worksheets('${CONFIG.sheetName}')/range(address='${celda}')`;

            await fetch(url, {
                method: "PATCH",
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ values: [[cambio.columna === 20 ? cambio.valor / 100 : cambio.valor]] })
            });
        }

        btnGuardar.disabled = false;
        btnGuardar.textContent = "Guardar cambios";
        alert(`${cambios.length} cambio(s) guardado(s). Hacé clic en "Ver datos modificados" para ver el impacto en los resultados.`);

    } catch (err) {
        console.error(err);
        btnGuardar.disabled = false;
        btnGuardar.textContent = "Guardar cambios";
        alert("Error al guardar cambios: " + err.message);
    }
}

function columnToA1(colIndex, rowIndex) {
    let col = "";
    let n = colIndex + 1;
    while (n > 0) {
        col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
        n = Math.floor((n - 1) / 26);
    }
    return `${col}${rowIndex}`;
}

// =============================================
// INICIALIZAR FILTROS
// =============================================
function inicializarFiltros() {
    poblarSelect("filtro-gerencia", COLUMNAS.GERENCIA, datosBrutos);
    ["filtro-seccion", "filtro-sucursal", "filtro-mes", "filtro-grupo", "filtro-subgrupo", "filtro-cuenta", "filtro-cuentageneral"].forEach(id => {
        document.getElementById(id).disabled = true;
    });
}

// =============================================
// CASCADA DE FILTROS
// =============================================
function actualizarCascada(nivel) {
    const gerencia = document.getElementById("filtro-gerencia").value;
    const seccion = document.getElementById("filtro-seccion").value;
    const sucursal = document.getElementById("filtro-sucursal").value;
    const mes = document.getElementById("filtro-mes").value;
    const grupo = document.getElementById("filtro-grupo").value;
    const subgrupo = document.getElementById("filtro-subgrupo").value;
    const cuenta = document.getElementById("filtro-cuenta").value;

    const datos = modoResultados ? datosBrutosResultados : datosBrutos;

    const filtrarDatos = (data) => data.filter(f =>
        (gerencia === "" || f[COLUMNAS.GERENCIA] === gerencia) &&
        (seccion === "" || f[COLUMNAS.SECCION] === seccion) &&
        (sucursal === "" || f[COLUMNAS.SUCURSAL] === sucursal) &&
        (mes === "" || f[COLUMNAS.MES] === mes) &&
        (grupo === "" || f[COLUMNAS.GRUPO] === grupo) &&
        (subgrupo === "" || f[COLUMNAS.SUBGRUPO] === subgrupo) &&
        (cuenta === "" || f[COLUMNAS.CUENTA] === cuenta)
    );

    const orden = ["gerencia", "seccion", "sucursal", "mes", "grupo", "subgrupo", "cuenta", "cuentageneral"];
    const posicion = orden.indexOf(nivel);

    orden.slice(posicion + 1).forEach(n => resetearSelect(`filtro-${n}`));

    const datosFiltro = filtrarDatos(datos);

    if (posicion <= 0) { poblarSelect("filtro-seccion", COLUMNAS.SECCION, datosFiltro); document.getElementById("filtro-seccion").disabled = false; }
    if (posicion <= 1) { poblarSelect("filtro-sucursal", COLUMNAS.SUCURSAL, datosFiltro); document.getElementById("filtro-sucursal").disabled = false; }
    if (posicion <= 2) { poblarSelect("filtro-mes", COLUMNAS.MES, datosFiltro); document.getElementById("filtro-mes").disabled = false; }
    if (posicion <= 3) { poblarSelect("filtro-grupo", COLUMNAS.GRUPO, datosFiltro); document.getElementById("filtro-grupo").disabled = false; }
    if (posicion <= 4) { poblarSelect("filtro-subgrupo", COLUMNAS.SUBGRUPO, datosFiltro); document.getElementById("filtro-subgrupo").disabled = false; }
    if (posicion <= 5) { poblarSelect("filtro-cuenta", COLUMNAS.CUENTA, datosFiltro); document.getElementById("filtro-cuenta").disabled = false; }
    if (posicion <= 6) { poblarSelect("filtro-cuentageneral", COLUMNAS.CUENTA_GENERAL, datosFiltro); document.getElementById("filtro-cuentageneral").disabled = false; }
}

// =============================================
// POBLAR / RESETEAR SELECT
// =============================================
function poblarSelect(idSelect, indiceColumna, datos) {
    const select = document.getElementById(idSelect);
    const valorActual = select.value;
    select.innerHTML = '<option value="">-- Todos --</option>';

    const valores = [...new Set(datos
        .map(fila => fila[indiceColumna])
        .filter(v => v !== null && v !== undefined && v !== "")
    )].sort();

    valores.forEach(valor => {
        const option = document.createElement("option");
        option.value = valor;
        option.textContent = valor;
        select.appendChild(option);
    });

    if (valorActual) select.value = valorActual;
}

function resetearSelect(idSelect) {
    const select = document.getElementById(idSelect);
    select.innerHTML = '<option value="">-- Todos --</option>';
    select.disabled = true;
}

// =============================================
// APLICAR / LIMPIAR FILTROS
// =============================================
function aplicarFiltros() {
    const gerencia = document.getElementById("filtro-gerencia").value;
    const seccion = document.getElementById("filtro-seccion").value;
    const sucursal = document.getElementById("filtro-sucursal").value;
    const mes = document.getElementById("filtro-mes").value;
    const grupo = document.getElementById("filtro-grupo").value;
    const subgrupo = document.getElementById("filtro-subgrupo").value;
    const cuenta = document.getElementById("filtro-cuenta").value;
    const cuentageneral = document.getElementById("filtro-cuentageneral").value;

    const datos = modoResultados ? datosBrutosResultados : datosBrutos;

    datosFiltrados = datos.filter(fila => {
        return (
            (gerencia === "" || fila[COLUMNAS.GERENCIA] === gerencia) &&
            (seccion === "" || fila[COLUMNAS.SECCION] === seccion) &&
            (sucursal === "" || fila[COLUMNAS.SUCURSAL] === sucursal) &&
            (mes === "" || fila[COLUMNAS.MES] === mes) &&
            (grupo === "" || fila[COLUMNAS.GRUPO] === grupo) &&
            (subgrupo === "" || fila[COLUMNAS.SUBGRUPO] === subgrupo) &&
            (cuenta === "" || fila[COLUMNAS.CUENTA] === cuenta) &&
            (cuentageneral === "" || fila[COLUMNAS.CUENTA_GENERAL] === cuentageneral)
        );
    });

    renderizarTabla();
}

function limpiarFiltros() {
    document.getElementById("filtro-gerencia").value = "";
    ["filtro-seccion", "filtro-sucursal", "filtro-mes", "filtro-grupo", "filtro-subgrupo", "filtro-cuenta", "filtro-cuentageneral"].forEach(id => {
        resetearSelect(id);
    });
    datosFiltrados = [];
    document.getElementById("tabla-datos").style.display = "none";
    document.getElementById("mensaje-inicial").style.display = "block";
    document.getElementById("mensaje-inicial").textContent =
        "Aplicá los filtros y hacé clic en Filtrar para ver los datos.";
}

// =============================================
// RENDERIZAR TABLA
// =============================================
function renderizarTabla() {
    const tbody = document.getElementById("tabla-body");
    const tabla = document.getElementById("tabla-datos");
    const mensaje = document.getElementById("mensaje-inicial");

    tbody.innerHTML = "";

    if (datosFiltrados.length === 0) {
        tabla.style.display = "none";
        mensaje.style.display = "block";
        mensaje.textContent = "No se encontraron datos con los filtros seleccionados.";
        return;
    }

    mensaje.style.display = "none";
    tabla.style.display = "table";

    datosFiltrados.forEach((fila, index) => {
        const tr = document.createElement("tr");

        if (modoResultados) {
            tr.innerHTML = `
                <td>${fila[COLUMNAS.GERENCIA] || ""}</td>
                <td>${fila[COLUMNAS.SECCION] || ""}</td>
                <td>${fila[COLUMNAS.SUCURSAL] || ""}</td>
                <td>${fila[COLUMNAS.MES] || ""}</td>
                <td>${fila[COLUMNAS.GRUPO] || ""}</td>
                <td>${fila[COLUMNAS.SUBGRUPO] || ""}</td>
                <td>${fila[COLUMNAS.CUENTA] || ""}</td>
                <td>${fila[COLUMNAS.CUENTA_GENERAL] || ""}</td>
                <td>${formatearPorcentaje(fila[COLUMNAS.PROM_PVTAS])}</td>
                <td>${formatearUSD(fila[COLUMNAS.PROYECCION])}</td>
                <td>${formatearPorcentaje(fila[COLUMNAS.PROM_PVTAS_RESULT])}</td>
                <td>${formatearUSD(fila[COLUMNAS.PROYECCION_RESULT])}</td>
            `;
        } else {
            tr.innerHTML = `
                <td>${fila[COLUMNAS.GERENCIA] || ""}</td>
                <td>${fila[COLUMNAS.SECCION] || ""}</td>
                <td>${fila[COLUMNAS.SUCURSAL] || ""}</td>
                <td>${fila[COLUMNAS.MES] || ""}</td>
                <td>${fila[COLUMNAS.GRUPO] || ""}</td>
                <td>${fila[COLUMNAS.SUBGRUPO] || ""}</td>
                <td>${fila[COLUMNAS.CUENTA] || ""}</td>
                <td>${fila[COLUMNAS.CUENTA_GENERAL] || ""}</td>
                <td>${formatearPorcentaje(fila[COLUMNAS.PROM_PVTAS])}</td>
                <td>${formatearUSD(fila[COLUMNAS.PROYECCION])}</td>
                <td><input type="text" class="input-pvtas" data-index="${index}" placeholder="ej: -0,8"></td>
                <td><input type="text" class="input-usd" data-index="${index}" placeholder="ej: 3000,25"></td>
            `;
        }

        tbody.appendChild(tr);
    });
}

// =============================================
// FORMATEAR NUMEROS
// =============================================
function formatearPorcentaje(valor) {
    if (valor === null || valor === undefined || valor === "" || valor === 0) return "-";
    const num = parseFloat(valor);
    if (isNaN(num)) return "-";
    return (num * 100).toLocaleString("es-AR", { minimumFractionDigits: 3, maximumFractionDigits: 5 }) + "%";
}

function formatearUSD(valor) {
    if (valor === null || valor === undefined || valor === "") return "-";
    const num = parseFloat(valor);
    if (isNaN(num)) return "-";
    return "USD " + num.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parsearNumero(texto) {
    texto = texto.trim();
    texto = texto.replace(/\./g, "").replace(",", ".");
    const num = parseFloat(texto);
    if (isNaN(num)) return null;
    return num;
}