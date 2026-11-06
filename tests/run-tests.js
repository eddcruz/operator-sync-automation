// Pruebas de la lógica pura (sin Gmail ni Sheets), con datos ficticios.
// Ejecutar:  node tests/run-tests.js
const fs = require("fs");
const path = require("path");
const root = path.join(__dirname, "..");
const src = ["Config.example.gs", "src/Utilidades_Comunes.gs", "src/Procesos_Operadores.gs"]
    .map(f => fs.readFileSync(path.join(root, f), "utf8")).join("\n");
const ctx = {};
new Function("ctx", src + `
Object.assign(ctx, { normalizarNombre, resolverConsulta, detectarBase, construirRespuestaFaltantes,
                     calcularTotales, extraerNombreDeTexto, calcularConciliacion });`)(ctx);
const { normalizarNombre, resolverConsulta, detectarBase, construirRespuestaFaltantes,
        calcularTotales, extraerNombreDeTexto, calcularConciliacion } = ctx;

let ok = 0, fail = 0;
const t = (name, cond) => { cond ? ok++ : fail++; console.log((cond ? "  PASS " : "  FAIL ") + name); };
const table = (enc, rows) => ({ encabezados: enc, filas: rows, col: n => { const i = enc.indexOf(n); if (i < 0) throw Error(n); return i; } });

const encCat = ["Operador_ID", "Nombre_operador", "Posicion_Rol", "Rol", "Estatus"];
const catalogo = table(encCat, [
    [1001, "JUAN PEREZ LOPEZ", "OPERADOR SENCILLO", "PLATAFORMAS NORTE", "Activo"],
    [2001, "LUIS GOMEZ DIAZ", "OPERADOR FULL", "VOLTEOS NORTE", "Activo"],
    [2002, "LUIS GOMEZ DIAZ", "OPERADOR EN FORMACION", "PATIO SUR", "Baja"],
    [3001, "ANA MARIA RUIZ", "OPERADOR EN FORMACION", "VOLTEOS CENTRO", "Activo"]
]);
const encFal = ["Folio_ID", "Fecha_cierre", "Sem", "Mes", "Año", "Verificador", "Base_operativa", "Tracto", "Operador_ID",
    "Estatus", "Nombre_operador", "Concepto_descuento", "Movimiento", "Autoriza", "Division", "ID de artículo",
    "Descripcion_articulo", "Cantidad", "Monto_total", "Número_parcialidades"];
const row = (folio, id, name, mov, amount, date = "01/10/2026") => {
    const r = new Array(encFal.length).fill("");
    Object.assign(r, { 0: folio, 1: date, 6: "Norte", 7: "T-01", 8: id, 10: name, 11: "Falla de Operador",
                       12: mov, 16: "BANDA", 17: 1, 18: amount, 19: 2 });
    return r;
};
const faltantes = table(encFal, [
    row("A1", 1001, "JUAN PEREZ LOPEZ", "Cargo", "$326.00"),
    row("A2", 1001, "JUAN PEREZ LOPEZ", "Exento", 816),
    row("B1", 2001, "LUIS GOMEZ DIAZ", "Cargo", 100),
    row("C1", 9999, "PEDRO HISTORICO SOLO", "Cargo", 50)
]);

console.log("Names");
t("order and accents do not matter", normalizarNombre("Pérez López Juan") === normalizarNombre("JUAN PEREZ LOPEZ"));
t("empty value does not crash", normalizarNombre(null) === "" && normalizarNombre(undefined) === "");

console.log("Email parsing");
t("subject 'Cuentas de colaborador/ NAME'",
  extraerNombreDeTexto("CUENTAS DE COLABORADOR/ PEREZ LOPEZ JUAN", "Cuentas de colaborador/ PEREZ LOPEZ JUAN", "") === "PEREZ LOPEZ JUAN");
t("body table 'NOMBRE ... FECHA DE BAJA'",
  extraerNombreDeTexto("BAJA OP.", "Baja op.", "NOMBRE ANA MARIA RUIZ FECHA DE BAJA 8 DE NOVIEMBRE") === "ANA MARIA RUIZ");

t("subject 'BAJA ... OP. NAME'",
  extraerNombreDeTexto("BAJA EMPRESA CARGO OP. LUIS GERARDO MARTINEZ JAUREGUI", "Baja Empresa Cargo Op. Luis Gerardo Martinez Jauregui", "") === "LUIS GERARDO MARTINEZ JAUREGUI");
t("body 'del op. Name, quien...'",
  extraerNombreDeTexto("SALDOS", "Saldos", "SU APOYO CON LOS SALDOS DEL OP. LUIS GERARDO MARTÍNEZ JÁUREGUI, QUIEN SE ENCUENTRA EN PROCESO DE BAJA.") === "LUIS GERARDO MARTÍNEZ JÁUREGUI");
t("no name anywhere → null (goes to review)", extraerNombreDeTexto("BAJA", "Baja", "BUEN DIA, ADJUNTO REPORTE.") === null);

console.log("Bases");
t("detects base in email text", detectarBase("OPERADOR BASE NORTE") === "NORTE");
t("alias match is whole word only", detectarBase("NORTEÑO") === null);

console.log("Operator resolution");
let r = resolverConsulta("PEREZ LOPEZ JUAN", "", catalogo, faltantes);
t("unique match → 2 records by ID", r.estado === "ok" && r.filas.length === 2 && r.ids[0] === "1001");
r = resolverConsulta("LUIS GOMEZ DIAZ", "", catalogo, faltantes);
t("namesake without base → ambiguous (no reply)", r.estado === "ambiguo" && r.candidatos.length === 2);
r = resolverConsulta("LUIS GOMEZ DIAZ", "OPERADOR BASE NORTE", catalogo, faltantes);
t("namesake + base → resolved", r.estado === "ok" && r.ids[0] === "2001");
r = resolverConsulta("ANA MARIA RUIZ", "", catalogo, faltantes);
t("in catalog, no records → 'no deductions'", r.estado === "ok" && r.filas.length === 0);
r = resolverConsulta("JUAN PEREZ", "", catalogo, faltantes);
t("incomplete name → not found (never says 'no deductions')", r.estado === "no_encontrado");
r = resolverConsulta("PEDRO HISTORICO SOLO", "", catalogo, faltantes);
t("not in catalog but in history → ok via history", r.estado === "ok" && r.metodo === "historial");

console.log("Totals");
const tot = calcularTotales(faltantes.filas.slice(0, 2), 12, 18);
t("charge total excludes exempt", tot.cargo === 326 && tot.exento === 816);
const html = construirRespuestaFaltantes("PEREZ LOPEZ JUAN", faltantes.filas.slice(0, 2), encFal);
t("reply shows charge total $326.00", html.includes("Total a cargo del operador: $326.00"));
t("reply shows exempt separately", html.includes("Exento (no se cobra): $816.00"));
t("no records → 'no tiene reportes'", construirRespuestaFaltantes("X", [], encFal).includes("no tiene reportes"));

console.log("Status reconciliation guard");
const cat = Array.from({ length: 100 }, (_, i) => [i + 1, "OP " + (i + 1), "", "", "Activo"]);
const allIds = cat.map(c => c[0]);
let c = calcularConciliacion(cat, 0, 4, allIds.slice(0, 95));
t("5% deactivations → allowed (under 10%)", c.porcentajeBajas === 5 && c.bajasNuevas === 5);
c = calcularConciliacion(cat, 0, 4, allIds.slice(0, 50));
t("50% deactivations (half list missing) → blocked", c.porcentajeBajas === 50 && c.porcentajeBajas > 10);
c = calcularConciliacion([[1, "A", "", "", "Baja"], [2, "B", "", "", ""]], 0, 4, [1, 2]);
t("rehire and new operator become Activo", c.nuevaColEstatus.every(v => v[0] === "Activo"));

console.log(`\n${ok} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
