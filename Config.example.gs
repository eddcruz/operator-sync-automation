/************************************************************
 * Config.example.gs — PLANTILLA con datos ficticios.
 *
 * Cómo usarla:
 *  1. En Apps Script crea un archivo llamado "Config".
 *  2. Pega este contenido y reemplaza cada valor por el tuyo.
 *  3. Nunca subas tu Config real a GitHub (está en .gitignore).
 *
 * Todo lo propio de tu empresa vive aquí; los archivos de src/
 * no tienen correos, IDs ni nombres reales.
 ************************************************************/

const CONFIG = {
    // --- Modo prueba ---
    // true  = NO responde a RH: te manda a ti cada respuesta con "[PRUEBA]".
    // false = funcionamiento normal (responde a todos).
    // Empieza siempre en true y cámbialo cuando hayas revisado varias respuestas.
    MODO_PRUEBA: true,

    // --- Correos ---
    // Tu correo: recibe avisos, pruebas y alertas.
    CORREO_ALERTAS: "you@example.com",

    // Quién puede mandar la lista semanal de operadores activos (Excel adjunto).
    REMITENTES_LISTA_OPERADORES: [
        "payroll@example.com"
    ],
    // Texto que identifica el asunto de esos correos.
    ASUNTO_LISTA_OPERADORES: "Lista de personal activo",

    // Quién puede pedir cuentas/descuentos de operadores.
    REMITENTES_CONSULTAS: [
        "hr.one@example.com",
        "hr.two@example.com"
    ],
    // Palabras clave del asunto (sintaxis de búsqueda de Gmail).
    PALABRAS_ASUNTO_CONSULTAS: "CUENTAS OR CUENTA OR OPERADOR OR COLABORADOR OR BAJA OR FINIQUITO",

    // --- Drive y hojas ---
    // ID de la carpeta de Drive donde se guardan los Excel de RH
    // (de la URL: drive.google.com/drive/folders/ESTE_ID).
    CARPETA_DRIVE_ID: "YOUR_DRIVE_FOLDER_ID",
    // Pestañas de tu Google Sheet.
    HOJA_OPERADORES: "operators_catalog",
    HOJA_FALTANTES: "operator_missing_items",
    HOJA_BITACORA: "response_log",                 // se crea sola si no existe
    // Pestañas dentro del Excel que manda RH.
    PESTANAS_EXCEL_RH: ["Division A", "Division B"],

    // --- Etiquetas de Gmail (se crean solas) ---
    ETIQUETA_LISTA: "Sync_Operators_List",
    ETIQUETA_RESPONDIDO: "Sync_Deductions_Answered",
    ETIQUETA_REVISAR: "Sync_Deductions_Review",    // dudosos: no se responden
    ETIQUETA_PRUEBA: "Sync_Deductions_Test",       // solo en modo prueba

    // --- Bases operativas ---
    // Alias de cada base (palabras completas) para desempatar operadores
    // con el mismo nombre según la base que menciona el correo.
    BASES_OPERATIVAS: {
        NORTE: ["NORTE", "NTE"],
        CENTRO: ["CENTRO", "CTR"],
        SUR: ["SUR"]
    },

    // --- Seguridad ---
    UMBRAL_MIN_IDS_ACTIVOS: 20,   // lista con menos IDs → no se concilia
    PORCENTAJE_MAX_BAJAS: 10,     // lista que daría de baja a más del 10% de activos → no se concilia

    // --- Prueba manual (probarRespuestaFaltantes) ---
    PRUEBA_NOMBRE: "PEREZ LOPEZ JUAN",
    PRUEBA_TEXTO: "OPERADOR BASE NORTE",   // opcional: texto del correo para desempatar

    // --- Firma de las respuestas ---
    FIRMA: "Saludos,<br>Your Name"
};
