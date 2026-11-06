/************************************************************
 * Utilidades_Comunes.gs — funciones de apoyo (sin datos de la empresa)
 * Cada función existe UNA sola vez en todo el proyecto.
 ************************************************************/

/**
 * Normaliza un nombre para comparar sin importar mayúsculas, acentos ni orden.
 * "José Pérez López" y "PEREZ LOPEZ JOSE" → "JOSE LOPEZ PEREZ"
 */
function normalizarNombre(nombre) {
    return normalizarTexto(nombre)
        .split(" ")
        .filter(Boolean)
        .sort()
        .join(" ");
}

/** Mayúsculas, sin acentos y con espacios simples. */
function normalizarTexto(texto) {
    return (texto === null || texto === undefined ? "" : texto)
        .toString()
        .toUpperCase()
        .normalize("NFD").replace(/[̀-ͯ]/g, "")
        .replace(/\s+/g, " ")
        .trim();
}

/**
 * Lee una hoja completa.
 * Devuelve { encabezados, filas, col(nombreColumna) → índice }.
 */
function leerTabla(hoja) {
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos[0] || [];
    return {
        encabezados,
        filas: datos.slice(1),
        col: nombre => {
            const i = encabezados.indexOf(nombre);
            if (i === -1) throw new Error(`No encuentro la columna '${nombre}' en la hoja ${hoja.getName()}.`);
            return i;
        }
    };
}

/** Devuelve la etiqueta de Gmail; la crea si no existe. */
function obtenerEtiqueta(nombre) {
    return GmailApp.getUserLabelByName(nombre) || GmailApp.createLabel(nombre);
}

/** Extrae los correos de un encabezado tipo "Nombre <correo>, correo2". */
function extraerEmails(cadena) {
    if (!cadena) return [];
    return cadena
        .split(",")
        .map(s => s.trim())
        .map(s => s.replace(/^.*<([^>]+)>\s*$/, "$1"))
        .filter(Boolean);
}

/** "$1,234.50" o 1234.5 → 1234.5 */
function aNumero(v) {
    if (typeof v === "number") return v;
    return parseFloat((v || "0").toString().replace(/[^\d.-]/g, "")) || 0;
}

/** 1234.5 → "$1,234.50" */
function aMoneda(n) {
    return "$" + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/** Fecha (Date o "dd/mm/aaaa") → número para ordenar. */
function valorFecha(v) {
    if (v instanceof Date) return v.getTime();
    const m = (v || "").toString().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : 0;
}

/** Date → "dd/mm/aaaa"; texto → se deja igual. */
function formatearFecha(v) {
    if (v instanceof Date) {
        return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}/${v.getFullYear()}`;
    }
    return escaparHtml(v);
}

/** Evita que un texto del correo o de la hoja rompa el HTML. */
function escaparHtml(v) {
    return (v === null || v === undefined ? "" : v).toString()
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
