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
    m = t.match(/(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/);
    if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = t.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return "";
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
  function leerCompromisos(paginas) {
    let enTabla = false, cols = null, filas = [], terminado = false;
    for (const items of paginas) {
      if (terminado) break;
      let desde = -Infinity;
      if (!enTabla) {
        const titulo = items.find((i) => /compromisos de la reuni/i.test(i.t));
        if (!titulo) continue;
        const desc = items.find((i) => /descripci/i.test(i.t) && i.y > titulo.y);
        if (!desc) continue;
        const linea = items.filter((i) => mismaLinea(i, desc)).sort((a, b) => a.x - b.x);
        const idx = (re) => linea.findIndex((i) => re.test(i.t));
        const iNo = idx(/^no\.?$/i), iDesc = idx(/descripci/i), iResp = idx(/responsable/i), iFecha = idx(/^fecha/i);
        if (iDesc < 0 || iResp < 0 || iFecha < 0) continue;
        const centro = (desdeI, hastaI) => { const g = linea.slice(desdeI, hastaI); return (g[0].x + g[g.length - 1].x + g[g.length - 1].w) / 2; };
        const c = [iNo >= 0 ? centro(iNo, iDesc) : null, centro(iDesc, iResp), centro(iResp, iFecha), centro(iFecha, linea.length)];
        desde = desc.y + 2;
        const corte = items.filter((i) => i.y > desde && /^(hora de finalizaci|lista de participantes|fecha\s*[–-]\s*hora|fecha emisi|propiedad intelectual)/i.test(i.t)).reduce((m, i) => Math.min(m, i.y), Infinity);
        const debajo = items.filter((i) => i.y > desde && i.y < corte - 2);   // solo lo que está dentro de la tabla
        const primerTexto = debajo.filter((i) => i.x > (c[0] ?? 0) + 8).reduce((m, i) => Math.min(m, i.x), Infinity);
        const L1 = primerTexto - 6;
        const L2 = 2 * c[1] - L1, L3 = 2 * c[2] - L2;
        cols = { L1, L2, L3, cNo: c[0] };
        enTabla = true;
      } else {
        // Página siguiente: se salta el encabezado repetido del formato (termina en la fecha vigente).
        const vig = items.filter((i) => /^\d{2}\/\d{2}\/\d{4}$/.test(i.t) && i.x > 380).sort((a, b) => a.y - b.y)[0];
        desde = vig ? vig.y + 4 : 0;
      }
      const fin = items.filter((i) => i.y > desde && /^(hora de finalizaci|lista de participantes|fecha\s*[–-]\s*hora)/i.test(i.t)).sort((a, b) => a.y - b.y)[0];
      const pie = items.filter((i) => /^fecha emisi|propiedad intelectual/i.test(i.t)).sort((a, b) => a.y - b.y)[0];
      const hasta = Math.min(fin ? fin.y - 2 : Infinity, pie ? pie.y - 2 : Infinity);
      if (fin) terminado = true;
      const celdas = items.filter((i) => i.y > desde && i.y < hasta);
      const numeros = celdas.filter((i) => /^\d{1,3}\.?$/.test(i.t) && i.x < cols.L1).sort((a, b) => a.y - b.y);
      if (!numeros.length) continue;
      const col = (i) => (i.x < cols.L1 ? 0 : i.x < cols.L2 - 2 ? 1 : i.x < cols.L3 - 2 ? 2 : 3);
      const contenido = celdas.filter((i) => col(i) > 0);
      // Si hay texto por encima del primer número, las celdas están centradas verticalmente.
      const centradas = contenido.some((i) => i.y < numeros[0].y - 3);
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
    return { encontrada: !!cols, compromisos: filas };
  }

  async function leerActa(archivo) {
    const datos = new Uint8Array(await archivo.arrayBuffer());
    const paginas = await fragmentos(datos);
    const r = leerCompromisos(paginas);
    return { fecha: leerFecha(paginas), ...r, datos };
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

  window.ACTAS = { leerActa, fechaISO, aBase64, deBase64, mostrar, leerCompromisos };
})();
