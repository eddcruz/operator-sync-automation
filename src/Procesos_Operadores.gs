/************************************************************
 * Procesos_Operadores.gs
 *
 * 1) Ingesta y conciliación de operadores (RH → Drive → Sheets)
 *    - Varios remitentes posibles (CONFIG.REMITENTES_LISTA_OPERADORES)
 *    - Inserta nuevos sin duplicar por Operador_ID (escritura en bloque)
 *    - Conciliación Activo/Baja con dos candados:
 *        a) lista con menos de UMBRAL_MIN_IDS_ACTIVOS IDs → no concilia
 *        b) si daría de baja a más del PORCENTAJE_MAX_BAJAS de los activos → no concilia
 *      En ambos casos te avisa y deja el correo SIN etiqueta para reprocesarlo.
 *
 * 2) Respuestas automáticas a consultas de faltantes
 *    - Identifica al operador en operators_catalog → Operador_ID
 *    - Si hay duda (varios o ninguno) NO responde: te avisa y etiqueta "Revisar"
 *    - Total a cargo: solo "Cargo". "Exento" se muestra aparte (no se cobra)
 *    - Detalle por folio y bitácora de cada respuesta
 *    - Modo prueba: te manda a ti la respuesta en lugar de responder a RH
 *
 * Requiere: Config.gs y Utilidades_Comunes.gs. Drive API (servicio avanzado).
 ************************************************************/

/* ==========================================================
   1) INGESTA DE OPERADORES
   ========================================================== */

function procesarCorreosOperadores() {
    const remitentes = CONFIG.REMITENTES_LISTA_OPERADORES.map(r => `from:${r}`).join(" OR ");
    const hojaBDId = SpreadsheetApp.getActiveSpreadsheet().getId();
    let nuevosOperadores = 0;

    const threads = GmailApp.search(
        `(${remitentes}) subject:(${CONFIG.ASUNTO_LISTA_OPERADORES}) newer_than:90d -label:${CONFIG.ETIQUETA_LISTA}`
    );
    const folder = DriveApp.getFolderById(CONFIG.CARPETA_DRIVE_ID);
    const etiqueta = obtenerEtiqueta(CONFIG.ETIQUETA_LISTA);

    const idsActivosDeEstaCorrida = new Set();
    if (!threads.length) return;

    threads.forEach(thread => {
        thread.getMessages().forEach(msg => {
            msg.getAttachments().forEach(file => {
                if (file.getContentType() !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") return;

                const yaExiste = folder.getFilesByName(file.getName()).hasNext();
                if (!yaExiste) {
                    const archivoGuardado = folder.createFile(file);
                    const r = importarDatosDesdeExcel(archivoGuardado.getBlob(), hojaBDId, CONFIG.HOJA_OPERADORES);
                    nuevosOperadores += r.nuevos;
                    r.idsActivos.forEach(id => idsActivosDeEstaCorrida.add(id));
                } else {
                    const r = importarDatosDesdeExcel(file.copyBlob(), hojaBDId, CONFIG.HOJA_OPERADORES, true);
                    r.idsActivos.forEach(id => idsActivosDeEstaCorrida.add(id));
                }
            });
        });
    });

    if (nuevosOperadores > 0) {
        MailApp.sendEmail({
            to: CONFIG.CORREO_ALERTAS,
            subject: "✔️ Se agregaron nuevos operadores",
            body: `Se han agregado ${nuevosOperadores} nuevos operadores a ${CONFIG.HOJA_OPERADORES}.`
        });
    }

    const totalVistos = idsActivosDeEstaCorrida.size;
    let aplicado = false;
    if (totalVistos >= CONFIG.UMBRAL_MIN_IDS_ACTIVOS) {
        aplicado = conciliarEstatusOperadores(Array.from(idsActivosDeEstaCorrida));
    } else if (totalVistos > 0) {
        MailApp.sendEmail({
            to: CONFIG.CORREO_ALERTAS,
            subject: "⚠️ Lista de operadores sospechosamente pequeña: conciliación NO aplicada",
            htmlBody: `<p>Se procesó una lista con solo <strong>${totalVistos}</strong> IDs activos (umbral: ${CONFIG.UMBRAL_MIN_IDS_ACTIVOS}).</p>
                 <p>No se actualizó la columna <strong>Estatus</strong> para evitar marcar bajas por error.</p>
                 <p>Revisa el/los archivos adjuntos de RH y vuelve a ejecutar cuando recibas la lista completa.</p>`
        });
    }

    // Solo se marca como procesado si la conciliación se aplicó.
    // Si un candado la detuvo, el correo queda sin etiqueta para volver a correrlo.
    if (aplicado) threads.forEach(t => etiqueta.addToThread(t));
}

/**
 * Convierte el Excel a Sheets (temporal), inserta operadores nuevos
 * en bloque y devuelve los IDs activos vistos.
 */
function importarDatosDesdeExcel(blob, hojaBDId, hojaBDNombre, soloRecolectar) {
    const archivoConvertido = Drive.Files.insert({ title: blob.getName(), mimeType: MimeType.GOOGLE_SHEETS }, blob);
    const archivo = SpreadsheetApp.openById(archivoConvertido.id);
    const hojaBD = SpreadsheetApp.openById(hojaBDId).getSheetByName(hojaBDNombre);

    const lastRow = hojaBD.getLastRow();
    const idExistentes = new Set(lastRow > 1
        ? hojaBD.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(id => id.toString().trim())
        : []);

    const filasNuevas = [];
    const idsActivosVistos = new Set();

    CONFIG.PESTANAS_EXCEL_RH.forEach(nombreHoja => {
        const hoja = archivo.getSheetByName(nombreHoja);
        if (!hoja) return;

        const datos = hoja.getDataRange().getValues();
        if (!datos || datos.length < 2) return;

        const enc = datos[0];
        const idIndex = enc.indexOf("IdEmpleado");
        const nombreIndex = enc.indexOf("NombrePersona");
        const posIndex = enc.indexOf("DescripcionPosicion");
        const rolIndex = enc.indexOf("valorAtributoDefault");

        for (let i = 1; i < datos.length; i++) {
            const id = (datos[i][idIndex] ?? "").toString().trim();
            if (!id) continue;
            idsActivosVistos.add(id);

            if (!soloRecolectar && !idExistentes.has(id)) {
                filasNuevas.push([id, datos[i][nombreIndex], datos[i][posIndex], datos[i][rolIndex]]);
                idExistentes.add(id);
            }
        }
    });

    // Escritura en bloque (más rápida que una fila a la vez)
    if (filasNuevas.length) {
        hojaBD.getRange(hojaBD.getLastRow() + 1, 1, filasNuevas.length, 4).setValues(filasNuevas);
        SpreadsheetApp.flush(); // guarda ya, para que la conciliación vea las filas nuevas
    }

    DriveApp.getFileById(archivoConvertido.id).setTrashed(true);
    return { nuevos: filasNuevas.length, idsActivos: Array.from(idsActivosVistos) };
}

/**
 * Activo si el ID vino en la lista de RH; Baja si no. Escritura en bloque.
 * Candado: si daría de baja a más del PORCENTAJE_MAX_BAJAS de los activos,
 * no cambia nada y te avisa. Devuelve true si se aplicó.
 */
function conciliarEstatusOperadores(idsActivos) {
    if (!idsActivos || !idsActivos.length) return false;

    const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_OPERADORES);
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return false;

    const idCol = data[0].indexOf("Operador_ID");
    const estCol = data[0].indexOf("Estatus");
    if (idCol === -1 || estCol === -1) {
        throw new Error(`No encuentro 'Operador_ID' y/o 'Estatus' en ${CONFIG.HOJA_OPERADORES}.`);
    }

    const r = calcularConciliacion(data.slice(1), idCol, estCol, idsActivos);

    if (r.porcentajeBajas > CONFIG.PORCENTAJE_MAX_BAJAS) {
        MailApp.sendEmail({
            to: CONFIG.CORREO_ALERTAS,
            subject: "⚠️ Demasiadas bajas en la lista de operadores: conciliación NO aplicada",
            htmlBody: `<p>La lista de RH daría de baja a <strong>${r.bajasNuevas}</strong> de
                <strong>${r.activosAntes}</strong> operadores activos
                (<strong>${r.porcentajeBajas.toFixed(1)}%</strong>; límite: ${CONFIG.PORCENTAJE_MAX_BAJAS}%).</p>
                <p>No se cambió ningún estatus. Posible causa: el Excel llegó incompleto
                (por ejemplo, falta una de las pestañas).</p>
                <p>Algunos que pasarían a Baja: ${escaparHtml(r.ejemplosBaja.join(", "))}</p>
                <p>Si la lista es correcta, sube temporalmente PORCENTAJE_MAX_BAJAS en Config
                y ejecuta otra vez <em>Automatización → Importar operadores</em>.</p>`
        });
        return false;
    }

    sh.getRange(2, estCol + 1, r.nuevaColEstatus.length, 1).setValues(r.nuevaColEstatus);
    return true;
}

/**
 * Cálculo puro de la conciliación (sin Sheets, para poder probarlo).
 * filas: catálogo sin encabezado. Devuelve la nueva columna y cuántos
 * activos pasarían a Baja.
 */
function calcularConciliacion(filas, idCol, estCol, idsActivos) {
    const setActivos = new Set(idsActivos.map(v => v.toString().trim()));
    let activosAntes = 0, bajasNuevas = 0;
    const ejemplosBaja = [];

    const nuevaColEstatus = filas.map(row => {
        const id = (row[idCol] ?? "").toString().trim();
        const antes = (row[estCol] ?? "").toString().trim();
        const nuevo = setActivos.has(id) ? "Activo" : "Baja";
        if (antes === "Activo") {
            activosAntes++;
            if (nuevo === "Baja") { bajasNuevas++; if (ejemplosBaja.length < 10) ejemplosBaja.push(id); }
        }
        return [nuevo];
    });

    const porcentajeBajas = activosAntes ? (bajasNuevas / activosAntes) * 100 : 0;
    return { nuevaColEstatus, activosAntes, bajasNuevas, porcentajeBajas, ejemplosBaja };
}

/* ==========================================================
   2) RESPUESTAS A CONSULTAS DE FALTANTES
   ========================================================== */

function revisarFaltantesOperador() {
    const remitentes = CONFIG.REMITENTES_CONSULTAS.map(r => `from:${r}`).join(" OR ");
    const searchQuery =
        `(${remitentes}) (subject:(${CONFIG.PALABRAS_ASUNTO_CONSULTAS})) newer_than:3d ` +
        `-label:${CONFIG.ETIQUETA_RESPONDIDO} -label:${CONFIG.ETIQUETA_REVISAR} -label:${CONFIG.ETIQUETA_PRUEBA}`;

    const threads = GmailApp.search(searchQuery);
    if (!threads.length) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalogo = leerTabla(ss.getSheetByName(CONFIG.HOJA_OPERADORES));
    const faltantes = leerTabla(ss.getSheetByName(CONFIG.HOJA_FALTANTES));
    if (!faltantes.filas.length) return;

    const miCorreo = (Session.getActiveUser().getEmail() || "").toLowerCase();

    threads.forEach(thread => {
        // Si el hilo ya se atendió, no se vuelve a tocar aunque lleguen respuestas nuevas
        const etiquetasHilo = thread.getLabels().map(l => l.getName());
        const yaAtendido = [CONFIG.ETIQUETA_RESPONDIDO, CONFIG.ETIQUETA_REVISAR, CONFIG.ETIQUETA_PRUEBA]
            .some(e => etiquetasHilo.includes(e));
        if (yaAtendido) return;

        const mensajes = thread.getMessages();
        if (!mensajes.length) return;

        const primero = mensajes[0];
        const ultimo = mensajes[mensajes.length - 1];

        const nombreBuscado = obtenerNombreOperador(primero, ultimo);
        if (!nombreBuscado) {
            // No se pudo leer el nombre: se avisa UNA vez y se etiqueta "Revisar"
            const sinNombre = { estado: "sin_nombre" };
            avisarRevision(thread, "(no se pudo leer)", sinNombre);
            obtenerEtiqueta(CONFIG.ETIQUETA_REVISAR).addToThread(thread);
            registrarBitacora(thread, "", sinNombre, "");
            return;
        }

        const textoCorreo = `${ultimo.getSubject() || ""}\n${primero.getPlainBody() || ""}`;
        const resultado = resolverConsulta(nombreBuscado, textoCorreo, catalogo, faltantes);

        // --- Caso dudoso: NO se responde ---
        if (resultado.estado !== "ok") {
            avisarRevision(thread, nombreBuscado, resultado);
            obtenerEtiqueta(CONFIG.ETIQUETA_REVISAR).addToThread(thread);
            registrarBitacora(thread, nombreBuscado, resultado, "");
            return;
        }

        const htmlBody = construirRespuestaFaltantes(nombreBuscado, resultado.filas, faltantes.encabezados);
        const cc = destinatariosFaltantes(primero, ultimo, miCorreo);

        if (CONFIG.MODO_PRUEBA) {
            MailApp.sendEmail({
                to: CONFIG.CORREO_ALERTAS,
                subject: `[PRUEBA] ${thread.getFirstMessageSubject()}`,
                htmlBody: `<p style="background:#fff3cd;padding:8px;">MODO PRUEBA: esta respuesta NO se envió a RH.<br>
                           Se habría respondido a todos + CC: ${escaparHtml(cc.join(", ") || "(ninguno extra)")}</p>` + htmlBody
            });
            obtenerEtiqueta(CONFIG.ETIQUETA_PRUEBA).addToThread(thread);
            registrarBitacora(thread, nombreBuscado, resultado, "PRUEBA (solo a mí)");
            return;
        }

        ultimo.replyAll("", { htmlBody: htmlBody, cc: cc.join(",") });
        obtenerEtiqueta(CONFIG.ETIQUETA_RESPONDIDO).addToThread(thread);
        registrarBitacora(thread, nombreBuscado, resultado, cc.join(", "));
    });
}

/**
 * Decide qué responder. Devuelve:
 *  { estado: "ok", filas, ids, metodo }       → se puede responder con seguridad
 *  { estado: "ambiguo", candidatos }          → varios operadores con ese nombre
 *  { estado: "no_encontrado" }                → nadie con ese nombre: probable error de extracción
 *
 * Solo se responde "no tiene descuentos" cuando el operador está
 * identificado SIN duda en el catálogo y no tiene registros.
 */
function resolverConsulta(nombreBuscado, textoCorreo, catalogo, faltantes) {
    const buscado = normalizarNombre(nombreBuscado);
    const cCat = catalogo.col;
    const cFal = faltantes.col;

    // 1) Buscar en el catálogo por nombre (sin importar orden ni acentos)
    let candidatos = catalogo.filas.filter(r => normalizarNombre(r[cCat("Nombre_operador")]) === buscado);

    // 2) Si hay varios, desempatar por la base mencionada en el correo
    if (candidatos.length > 1) {
        const baseCorreo = detectarBase(textoCorreo);
        if (baseCorreo) {
            const porBase = candidatos.filter(r =>
                detectarBase(`${r[cCat("Rol")] || ""} ${r[cCat("Posicion_Rol")] || ""}`) === baseCorreo);
            if (porBase.length >= 1) candidatos = porBase;
        }
    }

    if (candidatos.length > 1) {
        return { estado: "ambiguo", candidatos: candidatos.map(r => ({
            id: r[cCat("Operador_ID")], nombre: r[cCat("Nombre_operador")], rol: r[cCat("Rol")], estatus: r[cCat("Estatus")]
        })) };
    }

    if (candidatos.length === 1) {
        const id = (candidatos[0][cCat("Operador_ID")] ?? "").toString().trim();
        const filas = faltantes.filas.filter(r => {
            const idFila = (r[cFal("Operador_ID")] ?? "").toString().trim();
            if (idFila) return idFila === id;
            return normalizarNombre(r[cFal("Nombre_operador")]) === buscado; // registros viejos sin ID
        });
        return { estado: "ok", filas, ids: [id], metodo: "catalogo" };
    }

    // 3) No está en el catálogo: usar el historial de faltantes solo si el nombre
    //    coincide y todos los registros son del mismo operador
    const filasPorNombre = faltantes.filas.filter(r => normalizarNombre(r[cFal("Nombre_operador")]) === buscado);
    if (filasPorNombre.length) {
        const ids = [...new Set(filasPorNombre.map(r => (r[cFal("Operador_ID")] ?? "").toString().trim()).filter(Boolean))];
        if (ids.length <= 1) return { estado: "ok", filas: filasPorNombre, ids, metodo: "historial" };
        return { estado: "ambiguo", candidatos: ids.map(id => ({ id, nombre: nombreBuscado, rol: "", estatus: "" })) };
    }

    return { estado: "no_encontrado" };
}

/** Detecta la base operativa mencionada en un texto (o null). Alias en CONFIG.BASES_OPERATIVAS. */
function detectarBase(texto) {
    const t = normalizarTexto(texto);
    for (const base of Object.keys(CONFIG.BASES_OPERATIVAS)) {
        if (CONFIG.BASES_OPERATIVAS[base].some(alias => new RegExp(`\\b${alias}\\b`).test(t))) return base;
    }
    return null;
}

/**
 * Arma el HTML de la respuesta.
 * Total a cargo = solo "Cargo". "Exento" se muestra aparte.
 * Otros movimientos se marcan "(revisar)" y no entran al total.
 */
function construirRespuestaFaltantes(nombreBuscado, resultados, encabezados) {
    const col = n => encabezados.indexOf(n);
    const c = {
        folio: col("Folio_ID"), fecha: col("Fecha_cierre"), base: col("Base_operativa"),
        tracto: col("Tracto"), concepto: col("Concepto_descuento"), movimiento: col("Movimiento"),
        articulo: col("Descripcion_articulo"), cantidad: col("Cantidad"),
        monto: col("Monto_total"), parcialidades: col("Número_parcialidades")
    };

    const nombre = escaparHtml(nombreBuscado);
    let html = `<p>Buen día,</p>`;

    if (!resultados.length) {
        html += `<p>El operador <strong>${nombre}</strong> no tiene reportes de descuentos aplicables por equipo de sujeción.</p>`;
        return html + `<p>${CONFIG.FIRMA}</p>`;
    }

    const filas = resultados.slice().sort((a, b) => valorFecha(b[c.fecha]) - valorFecha(a[c.fecha]));
    const t = calcularTotales(filas, c.movimiento, c.monto);
    let filasHtml = "";

    filas.forEach(r => {
        const mov = (r[c.movimiento] || "").toString().trim();
        const tipo = mov.toUpperCase();
        const etiqueta = tipo === "CARGO" ? mov
            : tipo === "EXENTO" ? "Exento (no se cobra)"
            : `${mov || "Sin movimiento"} (revisar)`;

        filasHtml += `<tr>
            <td>${escaparHtml(r[c.folio])}</td>
            <td>${formatearFecha(r[c.fecha])}</td>
            <td>${escaparHtml(r[c.base])}</td>
            <td>${escaparHtml(r[c.tracto])}</td>
            <td>${escaparHtml(r[c.articulo])}</td>
            <td style="text-align:center">${escaparHtml(r[c.cantidad])}</td>
            <td>${escaparHtml(r[c.concepto])}</td>
            <td>${escaparHtml(etiqueta)}</td>
            <td style="text-align:right">${aMoneda(aNumero(r[c.monto]))}</td>
            <td style="text-align:center">${escaparHtml(r[c.parcialidades])}</td>
        </tr>`;
    });

    html += `<p>Anexo información solicitada.</p>`;
    html += `<p><strong>Historial de faltantes de equipo de sujeción</strong><br>(Operador: <strong>${nombre}</strong>)</p>`;
    html += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-size: 13px;">
        <tr style="background:#f2f2f2;">
            <th>Folio</th><th>Fecha</th><th>Base</th><th>Tracto</th><th>Artículo</th>
            <th>Cant.</th><th>Concepto</th><th>Movimiento</th><th>Monto</th><th>Parcialidades</th>
        </tr>${filasHtml}</table>`;

    html += `<p><strong>Total a cargo del operador: ${aMoneda(t.cargo)}</strong>`;
    if (t.exento > 0) html += `<br>Exento (no se cobra): ${aMoneda(t.exento)}`;
    if (t.revisar > 0) html += `<br>⚠️ Movimientos por revisar (no incluidos en el total): ${aMoneda(t.revisar)}`;
    html += `</p><p style="color:#555; font-size:12px;">Este es el historial de descuentos reportados. Liquidación valida los montos ya aplicados.</p>`;
    return html + `<p>${CONFIG.FIRMA}</p>`;
}

/** Suma por tipo de movimiento. Solo "Cargo" va al total a cobrar. */
function calcularTotales(filas, colMovimiento, colMonto) {
    const t = { cargo: 0, exento: 0, revisar: 0 };
    filas.forEach(r => {
        const tipo = (r[colMovimiento] || "").toString().trim().toUpperCase();
        const monto = aNumero(r[colMonto]);
        if (tipo === "CARGO") t.cargo += monto;
        else if (tipo === "EXENTO") t.exento += monto;
        else t.revisar += monto;
    });
    return t;
}

/** Destinatarios del primer correo que faltan en el último (como antes). */
function destinatariosFaltantes(primero, ultimo, miCorreo) {
    const juntar = m => [...extraerEmails(m.getFrom()), ...extraerEmails(m.getTo()), ...extraerEmails(m.getCc())]
        .map(e => e.toLowerCase());
    const debeIncluir = new Set(juntar(primero));
    const yaIncluidos = new Set(juntar(ultimo));
    return [...debeIncluir].filter(e => e && !yaIncluidos.has(e) && e !== miCorreo);
}

/** Te avisa de una consulta dudosa que NO se respondió. */
function avisarRevision(thread, nombreBuscado, resultado) {
    const enlace = `https://mail.google.com/mail/u/0/#all/${thread.getId()}`;
    let detalle = resultado.estado === "ambiguo"
        ? `<p>Hay <strong>varios operadores</strong> con ese nombre:</p><ul>` +
          resultado.candidatos.map(c => `<li>ID ${escaparHtml(c.id)} · ${escaparHtml(c.nombre)} · ${escaparHtml(c.rol)} · ${escaparHtml(c.estatus)}</li>`).join("") +
          `</ul>`
        : resultado.estado === "sin_nombre"
        ? `<p><strong>No pude leer el nombre del operador</strong> en el asunto ni en el cuerpo del correo.
           Puede ser un formato nuevo; si se repite, se agrega a las reglas de lectura.</p>`
        : `<p><strong>No encontré</strong> a ningún operador con ese nombre ni en el catálogo ni en el historial.
           Puede ser un error al leer el nombre del correo, o un operador que no está registrado.</p>`;

    MailApp.sendEmail({
        to: CONFIG.CORREO_ALERTAS,
        subject: `⚠️ Revisar consulta (no se respondió): ${thread.getFirstMessageSubject()}`,
        htmlBody: `<p>Nombre leído del correo: <strong>${escaparHtml(nombreBuscado)}</strong></p>${detalle}
                   <p><a href="${enlace}">Abrir el correo</a> y responder a mano.</p>`
    });
}

/** Una fila por cada consulta atendida o enviada a revisión. */
function registrarBitacora(thread, nombreBuscado, resultado, destinatarios) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(CONFIG.HOJA_BITACORA);
    if (!hoja) {
        hoja = ss.insertSheet(CONFIG.HOJA_BITACORA);
        hoja.appendRow(["Fecha", "Modo", "Asunto", "Nombre leído", "Resultado", "Operador_ID",
                        "Método", "Registros", "Total a cargo", "Exento", "CC agregados"]);
    }

    let registros = "", cargo = "", exento = "";
    if (resultado.estado === "ok") {
        const enc = ss.getSheetByName(CONFIG.HOJA_FALTANTES).getRange(1, 1, 1, ss.getSheetByName(CONFIG.HOJA_FALTANTES).getLastColumn()).getValues()[0];
        const t = calcularTotales(resultado.filas, enc.indexOf("Movimiento"), enc.indexOf("Monto_total"));
        registros = resultado.filas.length; cargo = t.cargo; exento = t.exento;
    }

    hoja.appendRow([
        new Date(),
        CONFIG.MODO_PRUEBA ? "PRUEBA" : "REAL",
        thread.getFirstMessageSubject(),
        nombreBuscado,
        resultado.estado === "ok" ? (resultado.filas.length ? "Respondido" : "Respondido: sin descuentos") :
            resultado.estado === "ambiguo" ? "Revisar: varios operadores" :
            resultado.estado === "sin_nombre" ? "Revisar: nombre no leído" : "Revisar: no encontrado",
        (resultado.ids || (resultado.candidatos || []).map(c => c.id)).join(", "),
        resultado.metodo || "",
        registros, cargo, exento,
        destinatarios
    ]);
}

/**
 * PRUEBA SEGURA: arma la respuesta de un nombre y te la manda SOLO a ti.
 * No toca ningún correo de RH. Pon el nombre en CONFIG.PRUEBA_NOMBRE y ejecuta desde el editor.
 */
function probarRespuestaFaltantes() {
    const nombrePrueba = CONFIG.PRUEBA_NOMBRE;   // se cambia en Config
    const textoPrueba = CONFIG.PRUEBA_TEXTO || "";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalogo = leerTabla(ss.getSheetByName(CONFIG.HOJA_OPERADORES));
    const faltantes = leerTabla(ss.getSheetByName(CONFIG.HOJA_FALTANTES));
    const r = resolverConsulta(nombrePrueba, textoPrueba, catalogo, faltantes);

    const cuerpo = r.estado === "ok"
        ? construirRespuestaFaltantes(nombrePrueba, r.filas, faltantes.encabezados)
        : `<p>Resultado: <strong>${r.estado}</strong> → este correo NO se respondería; te llegaría un aviso de revisión.</p>
           <pre>${escaparHtml(JSON.stringify(r.candidatos || [], null, 2))}</pre>`;

    MailApp.sendEmail({
        to: CONFIG.CORREO_ALERTAS,
        subject: `[PRUEBA] ${nombrePrueba} → ${r.estado}${r.ids ? " (ID " + r.ids.join(", ") + ")" : ""}`,
        htmlBody: cuerpo
    });
}

/* ==========================================================
   Lectura de correos (sin cambios en la lógica)
   ========================================================== */

function obtenerNombreOperador(primeroMensaje, ultimoMensaje) {
    const asuntoRaw = (ultimoMensaje.getSubject() || "").trim();
    const asunto = asuntoRaw.toUpperCase();
    const cuerpo = (primeroMensaje.getPlainBody() || "").trim().toUpperCase();
    return extraerNombreDeTexto(asunto, asuntoRaw, cuerpo);
}

/** Lógica de extracción separada para poder probarla sin Gmail. */
function extraerNombreDeTexto(asunto, asuntoRaw, cuerpo) {
    // 1) Asunto "OPERADOR: NOMBRE" o "COLABORADOR: NOMBRE"
    let m =
        asunto.match(/OPERAD[O0]R?\s*[/\-:]\s*([A-ZÁÉÍÓÚÑ.\s]+)$/i) ||
        asunto.match(/COLABORADOR\s*[/\-:]\s*([A-ZÁÉÍÓÚÑ.\s]+)$/i);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 2) Asunto "BAJA ... - NOMBRE"
    m = asuntoRaw.match(/-\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.\s]{6,})\s*$/i);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 2b) Asunto "... OP. NOMBRE" / "... OPERADOR NOMBRE" al final
    m = asunto.match(/\b(?:OP|OPERADOR|COLABORADOR)\.?\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{6,})$/);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 3) Cuerpo con campos etiquetados
    m = cuerpo.match(/NOMBRE\s+(.+?)(?=\s+FECHA\s+DE\s+BAJA|\s+TIPO\s+DE\s+BAJA|\s+CAUSA\b|\s+SUB-CAUSA\b|$)/is);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 4) Cuerpo "OPERADOR - NOMBRE"
    m = cuerpo.match(/OPERADOR[^\n]*\n[-–]\s*([A-ZÁÉÍÓÚÑ.\s]+)\b/mi);
    if (!m) m = cuerpo.match(/OPERADOR[^\-]*[-–]\s*([A-ZÁÉÍÓÚÑ.\s]+)\b/mi);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 5) Cuerpo "del Operador NOMBRE"
    m = cuerpo.match(/OPERADOR\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ.\s]{8,})(?:,|\n|$)/i);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    // 6) Cuerpo "del op. Nombre Apellido, quien..." / "del operador Nombre."
    m = cuerpo.match(/\bDEL\s+OP(?:ERADOR)?\.?\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{6,}?)(?=\s*(?:,|\.|\n|\bQUIEN\b|$))/);
    if (m && m[1]) return limpiarNombreExtraido(m[1]);

    return null;
}

function limpiarNombreExtraido(nombre) {
    return (nombre || "")
        .toString()
        .replace(/\b(FECHA DE BAJA|TIPO DE BAJA|CAUSA|SUB-CAUSA)\b.*$/i, "")
        .replace(/[^\p{L}\s.]/gu, " ")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}
