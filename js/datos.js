/* Capa de acceso a datos: la ÚNICA parte de la app que habla con Excel.
   Tres modos de conexión, misma interfaz:
   - panel:   la app corre en el panel lateral de Excel y usa Office.js directo.
   - ventana: la app corre en la ventana grande y le pide los datos al panel (puente).
   - demo:    la app se abre fuera de Excel y usa datos de prueba en memoria. */
/* global Excel, Office */
(function () {
  // Tablas que la app espera. Si falta alguna, se crea sola (hoja + tabla con encabezados).
  const ESTRUCTURA = {
    Proyectos: null, Hitos: null, Seguimientos: null, Riesgos: null, Usuarios: null, Catalogos: null,
    Compromisos: ["ID_Compromiso", "ID_Proyecto", "ID_Seguimiento", "Compromiso", "Responsable", "Fecha_Compromiso", "Estado", "Fecha_Cierre", "Registrado_Por"],
    // Archivos (actas en PDF) guardados por partes en una hoja oculta. No se lee con leerTodo por su tamaño.
    Archivos: ["ID_Archivo", "Parte", "Total", "Nombre", "Tipo", "Datos"],
    // Hilo de comentarios de cada compromiso, hasta cerrarlo.
    Comentarios: ["ID_Comentario", "ID_Compromiso", "ID_Proyecto", "Fecha_Hora", "Autor", "Texto"],
    // Personas clave de cada proyecto (Sponsor, Líder funcional, Product Owner…). El rol sale del catálogo Rol_Stakeholder.
    Stakeholders: ["ID_Stakeholder", "ID_Proyecto", "Rol", "Nombre", "Cargo", "Area", "Correo", "Telefono", "ID_Proveedor"],
    // Maestro de proveedores; cada proyecto guarda los suyos en Proyectos.Proveedores.
    Proveedores: ["ID_Proveedor", "Nombre", "NIT", "Servicio", "Contacto", "Correo", "Telefono", "Activo"],
    // Tickets de la plataforma helpdesk asociados a cada proyecto.
    Tickets: ["ID_Ticket", "ID_Proyecto", "Numero", "Titulo", "Estado", "Prioridad", "Registrado_Por", "Fecha_Registro"],
    // Control de cambios (alcance, tiempo, costo, recursos) con aprobación.
    Cambios: ["ID_Cambio", "ID_Proyecto", "Fecha", "Tipo", "Descripcion", "Justificacion", "Impacto", "Nueva_Fecha_Fin", "Nuevo_Presupuesto", "Estado", "Solicitado_Por", "Decidido_Por", "Fecha_Decision", "Comentario_Decision"],
    // Lecciones aprendidas.
    Lecciones: ["ID_Leccion", "ID_Proyecto", "Fecha", "Categoria", "Tipo", "Situacion", "Leccion", "Recomendacion", "Registrado_Por"],
    // Matriz RACI: una fila por entregable; Asignaciones = JSON { "PM" | ID_Stakeholder: "R" | "A" | "C" | "I" }.
    RACI: ["ID_RACI", "ID_Proyecto", "Entregable", "Asignaciones"],
    // Dependencias: ID_Proyecto depende de Depende_De.
    Dependencias: ["ID_Dependencia", "ID_Proyecto", "Depende_De", "Tipo", "Descripcion"],
    // Directorio único de personas (recursos) del portafolio.
    Recursos: ["ID_Recurso", "Nombre", "Correo", "Cargo", "Area", "Telefono", "ID_Proveedor", "Capacidad", "Activo"],
    // Backlog de demanda: iniciativas asignadas por Gestión de la Demanda (llave: código Almera).
    Demandas: ["ID_Demanda", "Codigo_Almera", "Nombre", "Descripcion", "Beneficio", "Area", "Solicitante", "Sponsor", "Fecha_Recepcion", "Fecha_Deseada", "Tipo", "Categoria", "Tamano", "Obligatorio",
      "Costo_Estimado", "PM_Asignado", "Valor", "Alineacion", "Impacto_Paciente", "Urgencia", "Riesgo_Prio", "Esfuerzo", "Estado", "Fecha_Estado", "Fecha_Decision", "Decidido_Por", "Comentario_Decision", "ID_Proyecto", "Registrado_Por"],
    // Sugerencias de mejora de la app (cualquier usuario las registra; Admin/PMO las gestionan).
    Sugerencias: ["ID_Sugerencia", "Fecha_Hora", "Usuario", "Modulo", "Contexto", "ID_Proyecto", "Tipo", "Prioridad", "Descripcion", "URL_Referencia", "Adjuntos", "Estado", "Respuesta", "Version_App", "Version_Implementada", "Actualizado_Por", "Actualizado_El"],
    // Registro de auditoría: quién cambió qué y cuándo.
    Auditoria: ["Fecha_Hora", "Usuario", "Tabla", "ID_Registro", "ID_Proyecto", "Accion", "Detalle"],
    // Filtros guardados por cada usuario (y los que Admin/PMO comparten con todos).
    Filtros: ["ID_Filtro", "Usuario", "Nombre", "PM", "Proyecto", "Estado", "Metodologia", "Semaforo", "Texto", "Predeterminado", "Compartido"],
  };
  const SOLO_BAJO_DEMANDA = ["Archivos", "Auditoria"];
  const HOJAS_OCULTAS = ["Archivos"];
  const TABLAS = Object.keys(ESTRUCTURA);
  // Columnas agregadas en versiones posteriores: si faltan en el Excel, se crean al final de la tabla.
  const COLUMNAS_NUEVAS = {
    Proyectos: ["URL_Repositorio", "URL_Documentos", "Proveedores", "Codigo_Almera", "Dedicacion_PM", "Fecha_Fin_Base", "Presupuesto_Base", "Valor", "Urgencia", "Riesgo_Prio", "Esfuerzo"],
    Stakeholders: ["Dedicacion", "ID_Recurso"],
    Riesgos: ["ID_Seguimiento", "Origen", "Fecha_Identificacion"],
    Seguimientos: ["Fecha_Acta", "URL_Acta", "Acta_Archivo", "Ejecutado"],
    Compromisos: ["Correo_Responsable", "Dias_Alerta", "ID_Recurso"],
    Comentarios: ["Adjuntos", "Editado_El"],
  };
  const COLS_FECHA = ["Editado_El", "Fecha_Recepcion", "Fecha_Deseada", "Fecha_Estado", "Fecha_Identificacion", "Fecha_Fin_Base", "Nueva_Fecha_Fin", "Fecha_Decision", "Fecha", "Fecha_Registro", "Fecha_Inicio", "Fecha_Fin_Plan", "Actualizado_El", "Fecha_Plan", "Fecha_Real", "Fecha_Corte", "Fecha_Compromiso", "Fecha_Cierre", "Fecha_Acta"];

  // Excel puede convertir "2026-09-28" en número de serie; aquí se devuelve a texto ISO.
  const serialAISO = (v) => new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  function normalizar(col, v) {
    if (v === null || v === undefined) return "";
    if (COLS_FECHA.includes(col) && typeof v === "number" && v > 20000) return serialAISO(v);
    if (col === "Fecha_Hora" && typeof v === "number" && v > 20000) return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 16).replace("T", " ");
    return v;
  }
  const vacia = (fila) => fila.every((c) => c === "" || c === null);
  const letra = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const coincide = (cols, fila, criterios) => Object.entries(criterios).every(([k, v]) => String(fila[cols.indexOf(k)]) === String(v));

  async function crearTablasFaltantes(ctx, existentes) {
    const faltantes = TABLAS.filter((n) => !existentes.includes(n) && ESTRUCTURA[n]);
    if (!faltantes.length) return;
    for (const n of faltantes) {
      let hoja = ctx.workbook.worksheets.getItemOrNullObject(n);
      await ctx.sync();
      if (hoja.isNullObject) hoja = ctx.workbook.worksheets.add(n);
      const cols = ESTRUCTURA[n];
      const rango = hoja.getRange(`A1:${letra(cols.length)}1`);
      rango.values = [cols];
      const t = hoja.tables.add(rango, true);
      t.name = n;
      if (HOJAS_OCULTAS.includes(n)) hoja.visibility = "Hidden";
    }
    await ctx.sync();
  }

  async function crearColumnasFaltantes(ctx) {
    const cabeceras = Object.keys(COLUMNAS_NUEVAS).map((n) => ({ n, h: ctx.workbook.tables.getItem(n).getHeaderRowRange().load("values") }));
    await ctx.sync();
    let cambios = false;
    for (const { n, h } of cabeceras) {
      const existentes = h.values[0];
      for (const col of COLUMNAS_NUEVAS[n]) if (!existentes.includes(col)) { ctx.workbook.tables.getItem(n).columns.add(null, null, col); cambios = true; }
    }
    if (cambios) await ctx.sync();
  }

  let estructuraRevisada = false;   // tablas y columnas faltantes: se revisan solo en la primera lectura
  const OpsExcel = {
    async leerTodo() {
      return Excel.run(async (ctx) => {
        const lista = ctx.workbook.tables.load("items/name");
        await ctx.sync();
        let existentes = lista.items.map((t) => t.name);
        if (!estructuraRevisada || TABLAS.some((n) => !existentes.includes(n) && ESTRUCTURA[n])) {
          try { await crearTablasFaltantes(ctx, existentes); existentes = TABLAS; await crearColumnasFaltantes(ctx); estructuraRevisada = true; }
          catch (e) { /* usuario sin permiso de edición: se lee lo que exista */ }
        }
        const cargas = TABLAS.filter((n) => existentes.includes(n) && !SOLO_BAJO_DEMANDA.includes(n)).map((n) => {
          const t = ctx.workbook.tables.getItem(n);
          return { n, h: t.getHeaderRowRange().load("values"), b: t.getDataBodyRange().load("values") };
        });
        await ctx.sync();
        const res = Object.fromEntries(TABLAS.filter((n) => !SOLO_BAJO_DEMANDA.includes(n)).map((n) => [n, []]));
        for (const { n, h, b } of cargas) {
          const cols = h.values[0];
          res[n] = b.values.filter((r) => !vacia(r)).map((r) => Object.fromEntries(cols.map((k, i) => [k, normalizar(k, r[i])])));
        }
        return res;
      });
    },

    // Lee una tabla que no se carga con leerTodo (p. ej. Auditoria).
    async leerTabla(nombre) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItemOrNullObject(nombre);
        await ctx.sync();
        if (t.isNullObject) return [];
        const h = t.getHeaderRowRange().load("values"), b = t.getDataBodyRange().load("values");
        await ctx.sync();
        const cols = h.values[0];
        return b.values.filter((r) => !vacia(r)).map((r) => Object.fromEntries(cols.map((k, i) => [k, normalizar(k, r[i])])));
      });
    },

    // Agrega una o varias filas en una sola operación.
    async agregarFilas(tabla, objs) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem(tabla);
        const h = t.getHeaderRowRange().load("values");
        const b = t.getDataBodyRange().load("rowCount");               // solo el conteo: no se lee toda la tabla
        await ctx.sync();
        const filas = objs.map((obj) => h.values[0].map((k) => (obj[k] === undefined || obj[k] === null ? "" : obj[k])));
        let primeraVacia = false;
        if (b.rowCount === 1) { const r0 = b.getRow(0).load("values"); await ctx.sync(); primeraVacia = vacia(r0.values[0]); }
        if (primeraVacia) {                                             // tabla vacía: se usa la fila en blanco
          b.getRow(0).values = [filas[0]];
          if (filas.length > 1) t.rows.add(null, filas.slice(1));
        } else t.rows.add(null, filas);
        await ctx.sync();
        return true;
      });
    },
    async agregarFila(tabla, obj) { return OpsExcel.agregarFilas(tabla, [obj]); },

    // Actualiza varias filas buscándolas por ID (nunca por posición: la tabla pudo ser ordenada).
    async actualizarVarios(tabla, colId, cambiosPorId) {
      return Excel.run(async (ctx) => {
        // Solo se leen los encabezados y la columna ID; se escriben únicamente las celdas que cambian.
        const t = ctx.workbook.tables.getItem(tabla);
        const h = t.getHeaderRowRange().load("values");
        const ids = t.columns.getItem(colId).getDataBodyRange().load("values");
        await ctx.sync();
        const cols = h.values[0];
        const cuerpo = t.getDataBodyRange();
        for (const { id, cambios } of cambiosPorId) {
          const idx = ids.values.findIndex((r) => String(r[0]) === String(id));
          if (idx < 0) throw new Error(`No se encontró ${id} en ${tabla}. Recargue e intente de nuevo.`);
          const fila = cuerpo.getRow(idx);
          cols.forEach((k, i) => { if (k in cambios) fila.getCell(0, i).values = [[cambios[k] === null || cambios[k] === undefined ? "" : cambios[k]]]; });
        }
        await ctx.sync();
        return true;
      });
    },
    async actualizarPorId(tabla, colId, id, cambios) { return OpsExcel.actualizarVarios(tabla, colId, [{ id, cambios }]); },

    // ---- Archivos por partes (cada celda de Excel admite hasta 32.767 caracteres) ----
    async guardarParte(fila) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem("Archivos");
        const cuerpo = t.getDataBodyRange().load("rowCount");
        const h = t.getHeaderRowRange().load("values");
        await ctx.sync();
        const valores = [h.values[0].map((k) => (fila[k] === undefined ? "" : fila[k]))];
        let vacia = false;
        if (cuerpo.rowCount === 1) { const r0 = cuerpo.getRow(0).load("values"); await ctx.sync(); vacia = r0.values[0].every((c) => c === "" || c === null); }
        if (vacia) cuerpo.getRow(0).values = valores; else t.rows.add(null, valores);
        await ctx.sync();
        return true;
      });
    },
    async _filasArchivo(ctx, id) {
      const t = ctx.workbook.tables.getItem("Archivos");
      const ids = t.columns.getItem("ID_Archivo").getDataBodyRange().load("values");
      const partes = t.columns.getItem("Parte").getDataBodyRange().load("values");
      await ctx.sync();
      return { t, filas: ids.values.map((r, i) => ({ i, id: String(r[0]), parte: Number(partes.values[i][0]) })).filter((x) => x.id === String(id)) };
    },
    async borrarArchivo(id) { return OpsExcel.borrarArchivos([id]); },
    // Borra todas las partes de uno o varios archivos en una sola operación (antes era una llamada por archivo).
    async borrarArchivos(ids) {
      if (!ids || !ids.length) return 0;
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItemOrNullObject("Archivos");
        await ctx.sync();
        if (t.isNullObject) return 0;
        const col = t.columns.getItem("ID_Archivo").getDataBodyRange().load("values");
        await ctx.sync();
        const buscar = new Set(ids.map(String));
        const indices = col.values.map((r, i) => (buscar.has(String(r[0])) ? i : -1)).filter((i) => i >= 0).sort((a, b) => b - a);
        if (!indices.length) return 0;
        let quedan = col.values.length;
        for (const i of indices) {
          if (quedan === 1) { const r = t.getDataBodyRange().getRow(0); r.load("columnCount"); await ctx.sync(); r.values = [Array(r.columnCount).fill("")]; }
          else { t.rows.getItemAt(i).delete(); quedan -= 1; }
        }
        await ctx.sync();
        return indices.length;
      });
    },
    // Qué partes hay guardadas de un archivo (para verificar que quedó completo).
    async partesArchivo(id) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItemOrNullObject("Archivos");
        await ctx.sync();
        if (t.isNullObject) return [];
        const ids = t.columns.getItem("ID_Archivo").getDataBodyRange().load("values");
        const partes = t.columns.getItem("Parte").getDataBodyRange().load("values");
        const totales = t.columns.getItem("Total").getDataBodyRange().load("values");
        await ctx.sync();
        return ids.values.map((r, i) => ({ id: String(r[0]), parte: Number(partes.values[i][0]), total: Number(totales.values[i][0]) })).filter((x) => x.id === String(id)).map(({ parte, total }) => ({ parte, total }));
      });
    },
    async leerParte(id, parte) {
      return Excel.run(async (ctx) => {
        const { t, filas } = await OpsExcel._filasArchivo(ctx, id);
        const f = filas.find((x) => x.parte === Number(parte));
        if (!f) throw new Error("No se encontró el acta guardada.");
        const h = t.getHeaderRowRange().load("values");
        const r = t.getDataBodyRange().getRow(f.i).load("values");
        await ctx.sync();
        const fila = Object.fromEntries(h.values[0].map((k, i) => [k, r.values[0][i]]));
        return { Datos: String(fila.Datos || ""), Total: Number(fila.Total) || 1, Nombre: fila.Nombre, Tipo: fila.Tipo };
      });
    },

    // Borra todas las filas cuyo colId esté en ids (de abajo hacia arriba para no mover índices).
    async eliminarFilas(tabla, colId, ids) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem(tabla);
        const col = t.columns.getItem(colId).getDataBodyRange().load("values");
        await ctx.sync();
        const buscar = new Set(ids.map(String));
        const indices = col.values.map((r, i) => (buscar.has(String(r[0])) ? i : -1)).filter((i) => i >= 0).sort((a, b) => b - a);
        let quedan = col.values.length;
        for (const i of indices) {
          if (quedan === 1) { const r = t.getDataBodyRange().getRow(0); r.load("columnCount"); await ctx.sync(); r.values = [Array(r.columnCount).fill("")]; }
          else { t.rows.getItemAt(i).delete(); quedan -= 1; }
        }
        await ctx.sync();
        return indices.length;
      });
    },

    // Borra la fila cuyas columnas coinciden con todos los criterios ({ Lista: "Fase", Valor: "Cierre" }).
    async eliminarFila(tabla, criterios) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem(tabla);
        const h = t.getHeaderRowRange().load("values");
        const b = t.getDataBodyRange().load("values");
        await ctx.sync();
        const cols = h.values[0];
        const idx = b.values.findIndex((r) => coincide(cols, r, criterios));
        if (idx < 0) throw new Error(`No se encontró el registro en ${tabla}. Recargue e intente de nuevo.`);
        if (b.values.length === 1) b.values = [cols.map(() => "")]; // una tabla de Excel no puede quedar sin filas
        else t.rows.getItemAt(idx).delete();
        await ctx.sync();
        return true;
      });
    },
  };

  function demo() {
    const db = JSON.parse(JSON.stringify(window.DEMO_DATA || {}));
    TABLAS.forEach((n) => { db[n] = db[n] || []; });
    const copia = (x) => JSON.parse(JSON.stringify(x));
    const buscar = (tabla, colId, id) => {
      const r = db[tabla].find((x) => String(x[colId]) === String(id));
      if (!r) throw new Error(`No se encontró ${id} en ${tabla}.`);
      return r;
    };
    const archivos = {};
    const api = {
      async leerTodo() { const r = copia(db); SOLO_BAJO_DEMANDA.forEach((n) => delete r[n]); return r; },
      async leerTabla(nombre) { return copia(db[nombre] || []); },
      async guardarParte(f) { (archivos[f.ID_Archivo] = archivos[f.ID_Archivo] || [])[f.Parte] = copia(f); return true; },
      async borrarArchivo(id) { delete archivos[id]; return true; },
      async borrarArchivos(ids) { ids.forEach((id) => delete archivos[id]); return ids.length; },
      async partesArchivo(id) { return (archivos[id] || []).filter(Boolean).map((f) => ({ parte: Number(f.Parte), total: Number(f.Total) })); },
      async leerParte(id, parte) { const f = (archivos[id] || [])[parte]; if (!f) throw new Error("En modo demostración el acta solo existe mientras la página está abierta."); return f; },
      async agregarFilas(tabla, objs) { db[tabla].push(...copia(objs)); return true; },
      async agregarFila(tabla, obj) { return api.agregarFilas(tabla, [obj]); },
      async actualizarVarios(tabla, colId, lista) { lista.forEach(({ id, cambios }) => Object.assign(buscar(tabla, colId, id), copia(cambios))); return true; },
      async actualizarPorId(tabla, colId, id, cambios) { return api.actualizarVarios(tabla, colId, [{ id, cambios }]); },
      async eliminarFilas(tabla, colId, ids) {
        const buscar = new Set(ids.map(String));
        const antes = db[tabla].length;
        db[tabla] = db[tabla].filter((r) => !buscar.has(String(r[colId])));
        return antes - db[tabla].length;
      },
      async eliminarFila(tabla, criterios) {
        const i = db[tabla].findIndex((r) => Object.entries(criterios).every(([k, v]) => String(r[k]) === String(v)));
        if (i < 0) throw new Error(`No se encontró el registro en ${tabla}.`);
        db[tabla].splice(i, 1);
        return true;
      },
    };
    return api;
  }

  // Las operaciones contra Excel se ejecutan de a una (en fila): varias a la vez hacían que Excel en la web se quedara pegado.
  let cola = Promise.resolve();
  function enCola(op, fn) {
    const limite = op === "leerTodo" || op === "leerTabla" ? 120000 : 60000;
    const p = cola.then(() => new Promise((res, rej) => {
      const t = setTimeout(() => rej(new Error("Excel tardó demasiado en responder. Intente de nuevo; si persiste, cierre y abra el complemento.")), limite);
      Promise.resolve().then(fn).then((v) => { clearTimeout(t); res(v); }, (e) => { clearTimeout(t); rej(e); });
    }));
    cola = p.catch(() => {});
    return p;
  }
  const OPERACIONES = ["leerTodo", "leerTabla", "agregarFila", "agregarFilas", "actualizarPorId", "actualizarVarios", "eliminarFila", "eliminarFilas", "guardarParte", "borrarArchivo", "borrarArchivos", "partesArchivo", "leerParte"];

  async function puente() {
    let seq = 0;
    const pendientes = new Map();
    await new Promise((listo) => {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
        let m;
        try { m = JSON.parse(arg.message); } catch (e) { return; }
        const p = pendientes.get(m.id);
        if (!p) return;
        pendientes.delete(m.id);
        clearTimeout(p.t);
        if (m.ok) p.res(m.result); else p.rej(new Error(m.error));
      }, () => listo());
    });
    const llamar = (op, args = []) => new Promise((res, rej) => {
      const id = ++seq;
      const t = setTimeout(() => {
        pendientes.delete(id);
        rej(new Error("Excel no respondió. Verifique que el panel lateral siga abierto, o cierre esta ventana y ábrala de nuevo desde el panel."));
      }, op === "leerTodo" || op === "leerTabla" ? 150000 : 90000);
      pendientes.set(id, { res, rej, t });
      Office.context.ui.messageParent(JSON.stringify({ id, op, args }));
    });
    return Object.fromEntries(OPERACIONES.map((op) => [op, (...a) => llamar(op, a)]));
  }

  const OpsSerie = () => Object.fromEntries(OPERACIONES.map((op) => [op, (...a) => enCola(op, () => OpsExcel[op](...a))]));

  async function crear(modo) {
    if (modo === "panel") return OpsSerie();
    if (modo === "ventana") return puente();
    return demo();
  }

  window.DATOS = { crear, OpsExcel, OPERACIONES, serie: OpsSerie() };
})();
