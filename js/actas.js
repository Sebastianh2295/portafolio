/* Actas en PDF: leer la fecha y los compromisos del formato FSFB (GC-GME-FT-010),
   y mostrar el PDF guardado. Usa pdf.js (lib/), que se carga solo cuando se necesita. */
/* global pdfjsLib */
(function () {
  let cargaPdfjs = null;
  function pdfjs() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (!cargaPdfjs) cargaPdfjs = new Promise((ok, falla) => {
      const s = document.createElement("script");
      s.src = new URL("lib/pdf.min.js", location.href).href;
      s.onload = () => { pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("lib/pdf.worker.min.js", location.href).href; ok(pdfjsLib); };
      s.onerror = () => falla(new Error("No se pudo cargar el lector de PDF."));
      document.head.appendChild(s);
    });
    return cargaPdfjs;
  }

  const MESES = { enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12 };
  const pad = (n) => String(n).padStart(2, "0");
  // "24 de septiembre de 2026", "29/09/2026", "2026-09-29" → "2026-09-29"; lo demás → ""
  function fechaISO(texto) {
    const t = String(texto || "").toLowerCase().replace(/\s+/g, " ").trim();
    let m = t.match(/(\d{1,2}) de ([a-záéíóúñ]+) (?:de|del) (\d{4})/);
    if (m && MESES[m[2]]) return `${m[3]}-${pad(MESES[m[2]])}-${pad(m[1])}`;
    m = t.match(/(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(\d{4})/) || t.match(/\b(\d{1,2}) (\d{1,2}) (\d{4})\b/);   // también «27 / 10 / 2026» o «27 10 2026» (PDF partido)
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1 && +m[1] <= 31) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = t.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return "";
  }

  // Plazos escritos con palabras (o fechas) que quedaron pegados al responsable: se pasan a la fecha.
  const RX_PLAZO = /\b(sin fecha(?:\s+expl[ií]cita)?|lo (?:antes|m[aá]s pronto) posible|por (?:definir|confirmar)|a m[aá]s tardar\b.*|antes de(?:l)?\b.*|inmediat[oa]|permanente|continuo|esta semana|pr[oó]xim[ao]s? (?:semana|sesi[oó]n|reuni[oó]n|comit[eé]).*|\d{1,2}\s*[/.-]\s*\d{1,2}(?:\s*[/.-]\s*\d{2,4})?|\d{1,2} \d{1,2} \d{4})/i;
  function separarPlazo(c) {
    const r = String(c.Responsable || ""), m = r.match(RX_PLAZO);
    if (!m) return c;
    const resp = (r.slice(0, m.index) + " " + r.slice(m.index + m[0].length)).replace(/\s+/g, " ").replace(/\s+(de|del|la|y)$/i, "").trim();
    const txt = `${m[0].trim()} ${c.Fecha_Texto || ""}`.trim();
    return { ...c, Responsable: resp, Fecha_Texto: txt, Fecha_Compromiso: c.Fecha_Compromiso || fechaISO(txt) };
  }

  // Extrae los fragmentos de texto de cada página con su posición (y crece hacia abajo).
  async function fragmentos(datos) {
    const lib = await pdfjs();
    const doc = await lib.getDocument({ data: datos.slice(0) }).promise;
    const paginas = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const pag = await doc.getPage(n);
      const alto = pag.getViewport({ scale: 1 }).height;
      const tc = await pag.getTextContent();
      paginas.push(tc.items.filter((i) => i.str && i.str.trim()).map((i) => ({
        t: i.str.trim(), x: i.transform[4], y: alto - i.transform[5], w: i.width, h: Math.abs(i.transform[3]) || i.height || 10,
      })));
    }
    return paginas;
  }
  const mismaLinea = (a, b) => Math.abs(a.y - b.y) < 3;
  const unir = (items) => items.sort((a, b) => a.y - b.y || a.x - b.x).map((i) => i.t).join(" ").replace(/\s+/g, " ").trim();

  function leerFecha(paginas) {
    for (const items of paginas) {
      const et = items.find((i) => /^fecha\s*:/i.test(i.t));   // «Fecha:» (no «Fecha Vigente» ni «Fecha Emisión»)
      if (!et) continue;
      const hora = items.find((i) => mismaLinea(i, et) && /^hora/i.test(i.t) && i.x > et.x);
      const resto = et.t.replace(/^fecha\s*:/i, "") + " " + unir(items.filter((i) => i !== et && mismaLinea(i, et) && i.x > et.x && (!hora || i.x < hora.x)));
      const f = fechaISO(resto);
      if (f) return f;
    }
    return "";
  }

  /* Tabla «Compromisos de la reunión»: No. | Descripción actividad | Responsable | Fecha de entrega.
     Los encabezados están centrados en sus columnas y el contenido alineado a la izquierda, así que
     los límites de columna se deducen de los centros de los encabezados. */
  // Renglones de una página: fragmentos a la misma altura, unidos de izquierda a derecha.
  function lineasDe(items) {
    const l = [];
    items.slice().sort((a, b) => a.y - b.y || a.x - b.x).forEach((i) => { const u = l[l.length - 1]; if (u && Math.abs(u.y - i.y) < 3) u.items.push(i); else l.push({ y: i.y, items: [i] }); });
    return l.map((x) => ({ y: x.y, t: unir(x.items) }));
  }
  // Agrupa los fragmentos del encabezado por columna (los que se tocan o están uno encima de otro).
  function columnasEncabezado(banda) {
    const g = [];
    banda.slice().sort((a, b) => a.x - b.x).forEach((i) => {
      const k = g.find((c) => i.x <= c.max + 6 && i.x + i.w >= c.min - 6);
      if (k) { k.min = Math.min(k.min, i.x); k.max = Math.max(k.max, i.x + i.w); k.items.push(i); } else g.push({ min: i.x, max: i.x + i.w, items: [i] });
    });
    return g;
  }
  const vigenteDe = (items) => items.filter((i) => /^\d{2}\/\d{2}\/\d{4}$/.test(i.t) && i.x > 380).sort((a, b) => a.y - b.y)[0];
  function leerCompromisos(paginas) {
    let enTabla = false, cols = null, filas = [], terminado = false, tituloVisto = false;
    for (const items of paginas) {
      if (terminado) break;
      let desde = -Infinity;
      if (!enTabla) {
        let base;
        if (!tituloVisto) {
          // El título puede venir partido en varios fragmentos: se busca en el renglón completo.
          const titulo = lineasDe(items).find((l) => /compromisos\s+de\s+la\s+reuni/i.test(l.t));
          if (!titulo) continue;
          tituloVisto = true;
          base = titulo.y;
        } else { const vig = vigenteDe(items); base = vig ? vig.y : -Infinity; }   // el encabezado de la tabla quedó en la página siguiente
        const desc = items.filter((i) => /descripci|^actividad/i.test(i.t) && i.y > base).sort((a, b) => a.y - b.y)[0];
        if (!desc) continue;
        const banda = items.filter((i) => i.y >= desc.y - 8 && i.y <= desc.y + 16);
        const resp = banda.find((i) => /responsable/i.test(i.t)), fec = banda.find((i) => /^fecha|entrega|plazo/i.test(i.t)), no = banda.find((i) => /^(no\.?|n[°º.]|#|ítem|item)$/i.test(i.t));
        if (!resp || !fec) continue;
        const grupos = columnasEncabezado(banda);
        const centro = (it) => { const k = grupos.find((c) => c.items.includes(it)); return (k.min + k.max) / 2; };
        const c = [no ? centro(no) : null, centro(desc), centro(resp), centro(fec)];
        desde = Math.max(...banda.filter((i) => grupos.some((k) => k.items.includes(i))).map((i) => i.y)) + 2;
        const corte = items.filter((i) => i.y > desde && /^(hora de finalizaci|lista de participantes|fecha\s*[–-]\s*hora|fecha emisi|propiedad intelectual)/i.test(i.t)).reduce((m, i) => Math.min(m, i.y), Infinity);
        const debajo = items.filter((i) => i.y > desde && i.y < corte - 2);   // solo lo que está dentro de la tabla
        const primerTexto = debajo.filter((i) => i.x > (c[0] ?? 0) + 8).reduce((m, i) => Math.min(m, i.x), Infinity);
        const L1 = Number.isFinite(primerTexto) ? primerTexto - 6 : (c[0] ?? c[1] - 60) + 10;
        /* Inicio de cada columna: el contenido va alineado a la izquierda, así que los x que se repiten
           marcan dónde empieza cada columna. Se toma el último inicio antes del centro de su encabezado
           (sirve con encabezados centrados o alineados a la izquierda). Si no hay datos, se estima por simetría. */
        const conteo = {};
        debajo.filter((i) => i.x > L1).forEach((i) => { const k = Math.round(i.x); conteo[k] = (conteo[k] || 0) + 1; });
        const inicios = Object.keys(conteo).map(Number).filter((x) => conteo[x] >= 2).sort((a, b) => a - b);
        const inicioAntes = (centroCol, minimo) => inicios.filter((x) => x < centroCol && x > minimo).pop();
        const r2 = inicioAntes(c[2], L1 + 20), L2 = r2 !== undefined ? r2 - 2 : 2 * c[1] - L1;
        const r3 = inicioAntes(c[3], L2 + 20), L3 = r3 !== undefined ? r3 - 2 : 2 * c[2] - L2;
        cols = { L1, L2, L3, cNo: c[0] };
        enTabla = true;
      } else {
        // Página siguiente: se salta el encabezado repetido del formato (termina en la fecha vigente).
        const vig = vigenteDe(items);
        desde = vig ? vig.y + 4 : 0;
      }
      const fin = items.filter((i) => i.y > desde && /^(hora de finalizaci|lista de participantes|fecha\s*[–-]\s*hora)/i.test(i.t)).sort((a, b) => a.y - b.y)[0];
      const pie = items.filter((i) => /^fecha emisi|propiedad intelectual/i.test(i.t)).sort((a, b) => a.y - b.y)[0];
      const hasta = Math.min(fin ? fin.y - 2 : Infinity, pie ? pie.y - 2 : Infinity);
      if (fin) terminado = true;
      const celdas = items.filter((i) => i.y > desde && i.y < hasta);
      const numeros = celdas.filter((i) => /^\d{1,3}\.?$/.test(i.t) && i.x < cols.L1).sort((a, b) => a.y - b.y);
      if (!numeros.length) {
        // Página sin números nuevos: todo es continuación de la última fila (texto largo que pasó de página).
        const colN = (i) => (i.x < cols.L1 ? 0 : i.x < cols.L2 - 2 ? 1 : i.x < cols.L3 - 2 ? 2 : 3);
        const u = filas[filas.length - 1];
        if (u && cols.centradas !== undefined) {
          const [d, r, f] = [1, 2, 3].map((c) => unir(celdas.filter((i) => colN(i) === c)));
          if (d) u.Compromiso = `${u.Compromiso} ${d}`.trim();
          if (r) u.Responsable = `${u.Responsable} ${r}`.trim();
          if (f) { u.Fecha_Texto = `${u.Fecha_Texto || ""} ${f}`.trim(); u.Fecha_Compromiso = fechaISO(u.Fecha_Texto); }
        }
        continue;
      }
      // Límite Responsable | Fecha: si en un mismo renglón hay dos textos separados dentro de «Responsable»,
      // el de la derecha ya es la columna Fecha (p. ej. «Lo antes posible» o «Sin fecha explícita» junto al nombre).
      {
        const enResp = celdas.filter((i) => i.x >= cols.L2 - 2 && i.x < cols.L3 - 2).sort((a, b) => a.y - b.y || a.x - b.x);
        const cand = [];
        enResp.forEach((i, j) => { const prev = enResp[j - 1]; if (prev && Math.abs(prev.y - i.y) < 3 && i.x - (prev.x + prev.w) > 6) cand.push(i.x); });
        const nuevo = cand.length ? Math.min(...cand) : Infinity;
        if (nuevo > cols.L2 + 30 && nuevo < cols.L3) cols.L3 = nuevo;
      }
      const col = (i) => (i.x < cols.L1 ? 0 : i.x < cols.L2 - 2 ? 1 : i.x < cols.L3 - 2 ? 2 : 3);
      const contenido = celdas.filter((i) => col(i) > 0);
      // Si en la primera página hay texto por encima del primer número, las celdas están centradas verticalmente.
      // En las páginas siguientes, el texto de arriba es la continuación de la fila que venía de la página anterior.
      const primeraPag = cols.centradas === undefined;
      if (primeraPag) cols.centradas = contenido.some((i) => i.y < numeros[0].y - 3);
      const centradas = cols.centradas;
      const cont = [[], [], []];
      if (!primeraPag && filas.length) {
        const limite = numeros[0].y - 3;
        if (centradas) {
          // Bloques que terminan antes del primer número de la página: continuación.
          for (let c = 1; c <= 3; c++) {
            const its = contenido.filter((i) => col(i) === c).sort((a, b) => a.y - b.y);
            const arriba = []; for (const i of its) { const u = arriba[arriba.length - 1]; if (i.y < limite && (!u || i.y - u.y <= Math.max(u.h, i.h) * 1.55)) arriba.push(i); else break; }
            if (arriba.length && (its[arriba.length] === undefined || its[arriba.length].y - arriba[arriba.length - 1].y > arriba[arriba.length - 1].h * 1.55)) cont[c - 1].push(...arriba);
          }
        } else contenido.filter((i) => i.y < limite).forEach((i) => cont[col(i) - 1].push(i));
        const usados = new Set(cont.flat());
        for (let c = contenido.length - 1; c >= 0; c--) if (usados.has(contenido[c])) contenido.splice(c, 1);
        if (usados.size) {
          const u = filas[filas.length - 1], [d, r, f] = cont.map(unir);
          if (d) u.Compromiso = `${u.Compromiso} ${d}`.trim();
          if (r) u.Responsable = `${u.Responsable} ${r}`.trim();
          if (f) { u.Fecha_Texto = `${u.Fecha_Texto || ""} ${f}`.trim(); u.Fecha_Compromiso = fechaISO(u.Fecha_Texto); }
        }
      }
      const grupos = numeros.map((n) => ({ n: n.t.replace(".", ""), items: [[], [], []] }));
      const cercano = (y) => numeros.reduce((mejor, n, j) => (Math.abs(n.y - y) < Math.abs(numeros[mejor].y - y) ? j : mejor), 0);
      if (centradas) {
        /* Celdas centradas verticalmente: en cada columna se arman bloques de renglones seguidos (una celda)
           y el bloque completo va a la fila cuyo número está más cerca de su centro. Así una descripción
           larga no se parte entre dos compromisos. */
        for (let c = 1; c <= 3; c++) {
          const its = contenido.filter((i) => col(i) === c).sort((a, b) => a.y - b.y || a.x - b.x);
          const lineas = [];
          its.forEach((i) => { const l = lineas[lineas.length - 1]; if (l && Math.abs(l.y - i.y) < 3) l.items.push(i); else lineas.push({ y: i.y, h: i.h, items: [i] }); });
          const bloques = [];
          lineas.forEach((l, j) => { const prev = lineas[j - 1]; if (!prev || l.y - prev.y > Math.max(prev.h, l.h) * 1.55) bloques.push([]); bloques[bloques.length - 1].push(l); });
          bloques.forEach((bl) => { const k = cercano((bl[0].y + bl[bl.length - 1].y) / 2); grupos[k].items[c - 1].push(...bl.flatMap((l) => l.items)); });
        }
      } else {
        for (const i of contenido) {
          const k = numeros.findIndex((n, j) => i.y >= n.y - 3 && (j === numeros.length - 1 || i.y < numeros[j + 1].y - 3));
          if (k >= 0) grupos[k].items[col(i) - 1].push(i);
        }
      }
      grupos.forEach((g) => {
        const [desc, resp, fechaTxt] = g.items.map(unir);
        if (desc || resp) filas.push({ Compromiso: desc, Responsable: resp, Fecha_Texto: fechaTxt, Fecha_Compromiso: fechaISO(fechaTxt) });
      });
    }
    return { encontrada: !!cols, titulo: tituloVisto, compromisos: filas };
  }

  /* ---------- Lectura genérica (cualquier formato de acta) ----------
     Busca la sección de compromisos/acuerdos/tareas y una tabla cuyo encabezado tenga columnas como
     Descripción/Actividad/Acción/Compromiso, Responsable/Encargado y Fecha/Plazo/Entrega (en cualquier orden).
     Si no hay tabla, lee la lista (viñetas o numerada) que sigue al título. */
  const RX = {
    desc: /descripci|actividad|acci[oó]n|compromiso|tarea|acuerdo|^qu[eé]\b|^detalle/i,
    resp: /responsable|encargad|qui[eé]n|due[ñn]o|a cargo/i,
    fecha: /fecha|plazo|entrega|l[ií]mite|cu[aá]ndo|vencimiento|^para\b/i,
    no: /^(no\.?|n[°º.]?|#|ítem|item|id|n[uú]m\.?)$/i,
    estado: /^estado$/i,
  };
  const RX_TIT = /(compromisos|acuerdos|tareas\s+(asignadas|pendientes)|plan\s+de\s+acci[oó]n|action\s+items|pr[oó]ximos\s+pasos|pendientes)/i;
  const RX_NO_TIT = /seguimiento\s+(a|de)\s+(los\s+)?compromisos\s+anteriores|compromisos\s+anteriores/i;
  const RX_FIN = /^(lista de participantes|participantes|asistentes|asistencia|hora de finalizaci|pr[oó]xima reuni|fecha\s*[–-]\s*hora|firmas?\b|aprobaci[oó]n del acta|observaciones|riesgos|advertencias|siendo las|se da por terminada|conclusiones|temas tratados|desarrollo de la reuni)/i;
  const RX_PIE = /^(fecha emisi|propiedad intelectual|p[aá]gina\s+\d|c[oó]digo:|versi[oó]n:|fecha vigente)/i;
  const tipoCol = (t) => (RX.no.test(t) ? "no" : RX.resp.test(t) ? "resp" : RX.fecha.test(t) ? "fecha" : RX.estado.test(t) ? "estado" : RX.desc.test(t) ? "desc" : "");
  // Fecha de entrega: también "15/10", "15-oct", "octubre 15" (usa el año del acta).
  function fechaEntrega(texto, anio) {
    const f = fechaISO(texto); if (f) return f;
    const t = String(texto || "").toLowerCase();
    let m = t.match(/(\d{1,2})[/.-](\d{1,2})(?![/.-]?\d)/);
    if (m && anio && +m[2] >= 1 && +m[2] <= 12) return `${anio}-${pad(m[2])}-${pad(m[1])}`;
    m = t.match(/(\d{1,2})\s*(?:de\s+)?([a-záéíóúñ]{3,})/);
    const mes = m && Object.keys(MESES).find((k) => k.startsWith(m[2].slice(0, 3)));
    if (m && mes && anio) return `${anio}-${pad(MESES[mes])}-${pad(m[1])}`;
    m = t.match(/([a-záéíóúñ]{3,})\s+(\d{1,2})\b/);
    const mes2 = m && Object.keys(MESES).find((k) => k.startsWith(m[1].slice(0, 3)));
    if (m && mes2 && anio) return `${anio}-${pad(MESES[mes2])}-${pad(m[2])}`;
    return "";
  }
  // Una línea de lista: "1. Juan Pérez: enviar X (20/10/2026)" → compromiso, responsable y fecha.
  const NOMBRE = "[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\\s+(?:de\\s+|del\\s+)?[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){0,4}";
  function compromisoDeLinea(txt, anio) {
    let t = String(txt || "").replace(/^\s*(\d{1,3}[.)-]|[-•●▪◦*–·✓☐]|\([a-z0-9]\))\s*/i, "").trim();
    if (t.length < 4) return null;
    let resp = "", f = "", fTxt = "";
    const fechas = [
      /(?:fecha(?:\s+de\s+entrega)?|plazo|para(?:\s+el)?|antes\s+del?|hasta\s+el)\s*:?\s*(\d{1,2}[/.-]\d{1,2}(?:[/.-]\d{2,4})?|\d{1,2}\s+de\s+[a-záéíóúñ]+(?:\s+de\s+\d{4})?)/i,
      /\(?\b(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{1,2}\s+de\s+[a-záéíóúñ]+\s+de\s+\d{4})\b\)?/i,
      /\(?\b(\d{1,2}[/.-]\d{1,2})\b(?![/.-]?\d)\)?/,
    ];
    for (const rx of fechas) { const fm = t.match(rx); if (fm) { const v = fechaEntrega(fm[1], anio); if (v) { f = v; fTxt = fm[1]; t = t.replace(fm[0], " "); break; } } }
    t = t.replace(/\(\s*\)/g, " ").replace(/\s+/g, " ").trim();
    const rx = [
      [new RegExp(`responsables?\\s*:\\s*([^;.()\\n]+)`, "i"), 1],
      [new RegExp(`\\((${NOMBRE}(?:\\s*(?:[/,;]|\\sy\\s)\\s*${NOMBRE})*)\\)`), 1],
      [new RegExp(`^(${NOMBRE}(?:\\s*(?:[/,;]|\\sy\\s)\\s*${NOMBRE})*)\\s*[:]\\s+`), 1],
      [new RegExp(`\\s[–-]\\s*(${NOMBRE}(?:\\s*(?:[/,;]|\\sy\\s)\\s*${NOMBRE})*)\\s*[–-]?\\s*$`), 1],
    ];
    for (const [r, g] of rx) { const m = t.match(r); if (m && /\s/.test(m[g].trim()) || (m && r === rx[0][0])) { resp = m[g].trim(); t = t.replace(m[0], " "); break; } }
    t = t.replace(/\(\s*\)/g, " ").replace(/\s*[–-]\s*$/, "").replace(/\s+/g, " ").replace(/^[,;:.\s–-]+|[,;:\s–-]+$/g, "").trim();
    return t.length >= 4 ? { Compromiso: t.charAt(0).toUpperCase() + t.slice(1), Responsable: resp, Fecha_Texto: fTxt, Fecha_Compromiso: f } : null;
  }
  // Renglones de todo el documento PDF, con sus fragmentos.
  function renglones(paginas) {
    const out = [];
    paginas.forEach((items, pg) => {
      const l = [];
      items.slice().sort((a, b) => a.y - b.y || a.x - b.x).forEach((i) => { const u = l[l.length - 1]; if (u && Math.abs(u.y - i.y) < 3) u.items.push(i); else l.push({ pg, y: i.y, h: i.h, items: [i] }); });
      l.forEach((x) => { x.items.sort((a, b) => a.x - b.x); x.t = unir(x.items.slice()); out.push(x); });
    });
    return out;
  }
  function leerGenerico(paginas, anio) {
    const L = renglones(paginas).filter((l) => !RX_PIE.test(l.t));
    const iTit = L.findIndex((l) => RX_TIT.test(l.t) && !RX_NO_TIT.test(l.t) && l.t.length < 90);
    // Encabezado de tabla: un renglón (o dos seguidos) con al menos dos tipos de columna, uno de ellos descripción o responsable+fecha.
    const esEnc = (items) => { const tipos = new Set(items.map((i) => tipoCol(i.t)).filter(Boolean)); return tipos.size >= 2 && (tipos.has("desc") || (tipos.has("resp") && tipos.has("fecha"))); };
    let iEnc = -1, enc = null;
    const desde = iTit >= 0 ? iTit : 0, hastaBusq = iTit >= 0 ? Math.min(L.length, iTit + 8) : L.length;
    for (let i = desde; i < hastaBusq && iEnc < 0; i++) {
      const junto = L[i + 1] && L[i + 1].pg === L[i].pg && L[i + 1].y - L[i].y < 16 ? [...L[i].items, ...L[i + 1].items] : L[i].items;
      if (esEnc(L[i].items)) { iEnc = i; enc = L[i].items; }
      else if (esEnc(junto)) { iEnc = i + 1; enc = junto; }
    }
    if (iEnc >= 0) {
      // Columnas: se agrupan los fragmentos del encabezado que se tocan; cada grupo toma el tipo de su texto.
      const g = [];
      enc.slice().sort((a, b) => a.x - b.x).forEach((i) => { const k = g.find((c) => i.x <= c.max + 6 && i.x + i.w >= c.min - 6); if (k) { k.min = Math.min(k.min, i.x); k.max = Math.max(k.max, i.x + i.w); k.t += " " + i.t; } else g.push({ min: i.x, max: i.x + i.w, t: i.t }); });
      g.forEach((c) => { c.tipo = tipoCol(c.t) || tipoCol(c.t.split(" ")[0]); c.centro = (c.min + c.max) / 2; });
      const filasDatos = [];
      for (let i = iEnc + 1; i < L.length; i++) {
        const l = L[i];
        if (RX_FIN.test(l.t) || (RX_TIT.test(l.t) && l.t.length < 60 && !esEnc(l.items))) break;
        if (esEnc(l.items)) continue;   // encabezado repetido en otra página
        filasDatos.push(l);
      }
      // Inicio real de cada columna: x de los datos que se repiten (contenido alineado a la izquierda).
      const conteo = {};
      filasDatos.forEach((l) => l.items.forEach((i) => { const k = Math.round(i.x); conteo[k] = (conteo[k] || 0) + 1; }));
      const inicios = Object.keys(conteo).map(Number).sort((a, b) => a - b);
      g.forEach((c, k) => {
        const previo = k ? g[k - 1].ini : -Infinity;
        const cand = inicios.filter((x) => x > previo + 12 && x <= c.centro + 4 && (k === 0 || x > g[k - 1].centro));
        c.ini = cand.length ? cand[0] : c.min - 4;
      });
      const colDe = (i) => { let r = g[0]; g.forEach((c) => { if (i.x >= c.ini - 4) r = c; }); return r; };
      const hayNo = g.some((c) => c.tipo === "no");
      const filas = [];
      let act = null, prevY = null, prevPg = null;
      filasDatos.forEach((l) => {
        const porCol = new Map();
        l.items.forEach((i) => { const c = colDe(i); porCol.set(c, [...(porCol.get(c) || []), i.t]); });
        const txt = (tipo) => g.filter((c) => c.tipo === tipo).map((c) => (porCol.get(c) || []).join(" ")).join(" ").trim();
        const no = txt("no");
        const salto = prevY !== null && (l.pg !== prevPg || l.y - prevY > l.h * 1.55);
        const nueva = hayNo ? /^\d{1,3}\.?$/.test(no) : !act || (salto && txt("desc"));
        if (nueva || !act) { act = { desc: [], resp: [], fecha: [] }; filas.push(act); }
        ["desc", "resp", "fecha"].forEach((k) => { const v = txt(k); if (v) act[k].push(v); });
        prevY = l.y; prevPg = l.pg;
      });
      const comps = filas.map((f) => ({ Compromiso: f.desc.join(" ").replace(/\s+/g, " ").trim(), Responsable: f.resp.join(" ").replace(/\s+/g, " ").trim(), Fecha_Texto: f.fecha.join(" "), Fecha_Compromiso: fechaEntrega(f.fecha.join(" "), anio) }))
        .filter((c) => c.Compromiso.length >= 3);
      if (comps.length) return { encontrada: true, titulo: iTit >= 0, compromisos: comps, formato: "tabla" };
    }
    // Lista bajo el título.
    if (iTit < 0) return { encontrada: false, titulo: false, compromisos: [], formato: "" };
    const items = [];
    for (let i = iTit + 1; i < L.length; i++) {
      const t = L[i].t;
      if (RX_FIN.test(t) || (RX_TIT.test(t) && t.length < 60)) break;
      if (/^\s*(\d{1,3}[.)-]|[-•●▪◦*–·✓☐])\s*/.test(t) || !items.length) items.push(t); else items[items.length - 1] += " " + t;
    }
    const comps = items.map((t) => compromisoDeLinea(t, anio)).filter(Boolean);
    return { encontrada: comps.length > 0, titulo: true, compromisos: comps, formato: "lista" };
  }

  /* ---------- Actas en Word (.docx) ---------- */
  let cargaXLSX = null;
  function libZip() {
    if (window.XLSX && window.XLSX.CFB) return Promise.resolve(window.XLSX);
    if (!cargaXLSX) cargaXLSX = new Promise((ok, falla) => { const s = document.createElement("script"); s.src = new URL("lib/xlsx.full.min.js", location.href).href; s.onload = () => ok(window.XLSX); s.onerror = () => falla(new Error("No se pudo cargar el lector de Word.")); document.head.appendChild(s); });
    return cargaXLSX;
  }
  async function leerWord(datos) {
    const X = await libZip();
    const zip = X.CFB.read(datos, { type: "array" });
    const ent = X.CFB.find(zip, "word/document.xml") || X.CFB.find(zip, "/word/document.xml");
    if (!ent) throw new Error("El archivo no parece un documento de Word (.docx).");
    const xml = new DOMParser().parseFromString(new TextDecoder("utf-8").decode(ent.content), "application/xml");
    const hijos = (n, nombre) => [...n.childNodes].filter((c) => c.localName === nombre);
    const texto = (n) => [...n.getElementsByTagNameNS("*", "p")].map((p) => [...p.getElementsByTagNameNS("*", "t")].map((t) => t.textContent).join("")).join("\n").trim();
    const cuerpo = xml.getElementsByTagNameNS("*", "body")[0];
    const bloques = [];
    [...cuerpo.childNodes].forEach((n) => {
      if (n.localName === "p") { const t = [...n.getElementsByTagNameNS("*", "t")].map((x) => x.textContent).join("").trim(); if (t) bloques.push({ p: t }); }
      else if (n.localName === "tbl") bloques.push({ tabla: hijos(n, "tr").map((tr) => hijos(tr, "tc").map(texto)) });
    });
    const todo = bloques.flatMap((b) => (b.p ? [b.p] : b.tabla.flat()));
    const etFecha = todo.find((t) => /^fecha\s*:/i.test(t) && fechaISO(t));
    const fecha = etFecha ? fechaISO(etFecha) : "";
    const anio = (fecha || new Date().toISOString()).slice(0, 4);
    let titulo = false;
    // 1) Tabla con encabezado de compromisos (la fila del título puede estar dentro de la misma tabla).
    for (let k = 0; k < bloques.length; k++) {
      const b = bloques[k];
      if (b.p && RX_TIT.test(b.p) && !RX_NO_TIT.test(b.p)) titulo = true;
      if (!b.tabla) continue;
      if (b.tabla.some((f) => f.some((c) => RX_TIT.test(c) && !RX_NO_TIT.test(c) && c.length < 90))) titulo = true;
      const iEnc = b.tabla.findIndex((f) => { const tipos = new Set(f.map((c) => tipoCol(c.split("\n")[0].trim())).filter(Boolean)); return tipos.size >= 2 && (tipos.has("desc") || (tipos.has("resp") && tipos.has("fecha"))); });
      if (iEnc < 0) continue;
      const tipos = b.tabla[iEnc].map((c) => tipoCol(c.split("\n")[0].trim()));
      const val = (f, t) => f.filter((_, j) => tipos[j] === t).join(" ").replace(/\s+/g, " ").trim();
      const comps = b.tabla.slice(iEnc + 1).filter((f) => !f.every((c) => !c.trim())).map((f) => ({ Compromiso: val(f, "desc"), Responsable: val(f, "resp"), Fecha_Texto: val(f, "fecha"), Fecha_Compromiso: fechaEntrega(val(f, "fecha"), anio) }))
        .filter((c) => c.Compromiso.length >= 3 && !RX_FIN.test(c.Compromiso));
      if (comps.length) return { fecha, encontrada: true, titulo: true, compromisos: comps, formato: "Word" };
    }
    // 2) Lista de párrafos bajo el título.
    const iT = bloques.findIndex((b) => b.p && RX_TIT.test(b.p) && !RX_NO_TIT.test(b.p) && b.p.length < 90);
    if (iT >= 0) {
      const comps = [];
      for (let k = iT + 1; k < bloques.length; k++) { const b = bloques[k]; if (!b.p) break; if (RX_FIN.test(b.p) || (RX_TIT.test(b.p) && b.p.length < 60)) break; const c = compromisoDeLinea(b.p, anio); if (c) comps.push(c); }
      return { fecha, encontrada: comps.length > 0, titulo: true, compromisos: comps, formato: "Word" };
    }
    return { fecha, encontrada: false, titulo, compromisos: [], formato: "Word" };
  }

  const TIPO_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  async function leerActa(archivo) {
    const datos = new Uint8Array(await archivo.arrayBuffer());
    const esWord = /\.docx$/i.test(archivo.name || "") || (datos[0] === 0x50 && datos[1] === 0x4b);
    if (esWord) { const w = await leerWord(datos); return { ...w, compromisos: (w.compromisos || []).map(separarPlazo), datos, tipo: TIPO_DOCX, word: true }; }
    const paginas = await fragmentos(datos);
    const fecha = leerFecha(paginas) || (() => { for (const l of renglones(paginas)) { if (/fecha/i.test(l.t)) { const f = fechaISO(l.t); if (f) return f; } } return ""; })();
    let r = leerCompromisos(paginas);   // formato FSFB (GC-GME-FT-010)
    if (r.compromisos.length) r.formato = "FSFB";
    else { const g = leerGenerico(paginas, (fecha || new Date().toISOString()).slice(0, 4)); if (g.compromisos.length || !r.titulo) r = g; }
    return { fecha, ...r, compromisos: (r.compromisos || []).map(separarPlazo), datos, tipo: "application/pdf" };
  }

  // Base64 por partes para no bloquear el navegador con archivos grandes.
  function aBase64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function deBase64(b64) {
    const s = atob(b64);
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
    return bytes;
  }

  // Dibuja el PDF dentro de un contenedor (una imagen por página).
  async function mostrar(bytes, contenedor) {
    const lib = await pdfjs();
    const doc = await lib.getDocument({ data: bytes.slice(0) }).promise;
    contenedor.innerHTML = "";
    const ancho = Math.min(contenedor.clientWidth || 800, 900);
    for (let n = 1; n <= doc.numPages; n++) {
      const pag = await doc.getPage(n);
      const base = pag.getViewport({ scale: 1 });
      const vp = pag.getViewport({ scale: (ancho / base.width) * (window.devicePixelRatio || 1) });
      const c = document.createElement("canvas");
      c.width = vp.width; c.height = vp.height;
      c.style.width = ancho + "px";
      c.className = "pagina-pdf";
      contenedor.appendChild(c);
      await pag.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
    }
  }

  window.ACTAS = { leerActa, fechaISO, aBase64, deBase64, mostrar, leerCompromisos, leerGenerico, compromisoDeLinea, TIPO_DOCX };
})();
