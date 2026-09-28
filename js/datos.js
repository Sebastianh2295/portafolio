/* Capa de acceso a datos: la ÚNICA parte de la app que habla con Excel.
   Tres modos de conexión, misma interfaz (leerTodo, agregarFila, actualizarPorId):
   - panel:   la app corre en el panel lateral de Excel y usa Office.js directo.
   - ventana: la app corre en la ventana grande y le pide los datos al panel (puente).
   - demo:    la app se abre fuera de Excel y usa datos de prueba en memoria. */
/* global Excel, Office */
(function () {
  const TABLAS = ["Proyectos", "Hitos", "Seguimientos", "Riesgos", "Usuarios", "Catalogos"];
  const COLS_FECHA = ["Fecha_Inicio", "Fecha_Fin_Plan", "Actualizado_El", "Fecha_Plan", "Fecha_Real", "Fecha_Corte"];

  // Excel puede convertir "2026-09-28" en número de serie; aquí se devuelve a texto ISO.
  function serialAISO(v) {
    return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  function normalizar(col, v) {
    if (v === null || v === undefined) return "";
    if (COLS_FECHA.includes(col) && typeof v === "number" && v > 20000) return serialAISO(v);
    return v;
  }
  const vacia = (fila) => fila.every((c) => c === "" || c === null);

  const OpsExcel = {
    async leerTodo() {
      return Excel.run(async (ctx) => {
        const cargas = TABLAS.map((n) => {
          const t = ctx.workbook.tables.getItem(n);
          return { n, h: t.getHeaderRowRange().load("values"), b: t.getDataBodyRange().load("values") };
        });
        await ctx.sync();
        const res = {};
        for (const { n, h, b } of cargas) {
          const cols = h.values[0];
          res[n] = b.values.filter((r) => !vacia(r)).map((r) =>
            Object.fromEntries(cols.map((k, i) => [k, normalizar(k, r[i])])));
        }
        return res;
      });
    },

    async agregarFila(tabla, obj) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem(tabla);
        const h = t.getHeaderRowRange().load("values");
        const b = t.getDataBodyRange().load("values");
        await ctx.sync();
        const fila = h.values[0].map((k) => (obj[k] === undefined || obj[k] === null ? "" : obj[k]));
        if (b.values.length === 1 && vacia(b.values[0])) b.values = [fila]; // tabla vacía
        else t.rows.add(null, [fila]);
        await ctx.sync();
        return true;
      });
    },

    async actualizarPorId(tabla, colId, id, cambios) {
      return Excel.run(async (ctx) => {
        const t = ctx.workbook.tables.getItem(tabla);
        const h = t.getHeaderRowRange().load("values");
        const b = t.getDataBodyRange().load("values");
        await ctx.sync();
        const cols = h.values[0];
        const iId = cols.indexOf(colId);
        // Se busca por ID, nunca por posición: la tabla pudo ser ordenada.
        const idx = b.values.findIndex((r) => String(r[iId]) === String(id));
        if (idx < 0) throw new Error(`No se encontró ${id} en ${tabla}. Recargue e intente de nuevo.`);
        const fila = b.values[idx].slice();
        cols.forEach((k, i) => { if (k in cambios) fila[i] = cambios[k] === null ? "" : cambios[k]; });
        b.getRow(idx).values = [fila];
        await ctx.sync();
        return true;
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
        const idx = b.values.findIndex((r) =>
          Object.entries(criterios).every(([k, v]) => String(r[cols.indexOf(k)]) === String(v)));
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
    const copia = (x) => JSON.parse(JSON.stringify(x));
    return {
      async leerTodo() { return copia(db); },
      async agregarFila(tabla, obj) { db[tabla].push(copia(obj)); return true; },
      async actualizarPorId(tabla, colId, id, cambios) {
        const r = db[tabla].find((x) => String(x[colId]) === String(id));
        if (!r) throw new Error(`No se encontró ${id} en ${tabla}.`);
        Object.assign(r, copia(cambios));
        return true;
      },
      async eliminarFila(tabla, criterios) {
        const i = db[tabla].findIndex((r) => Object.entries(criterios).every(([k, v]) => String(r[k]) === String(v)));
        if (i < 0) throw new Error(`No se encontró el registro en ${tabla}.`);
        db[tabla].splice(i, 1);
        return true;
      },
    };
  }

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
        rej(new Error("Excel no respondió. Cierre esta ventana y vuelva a abrir el Portafolio desde Excel."));
      }, 30000);
      pendientes.set(id, { res, rej, t });
      Office.context.ui.messageParent(JSON.stringify({ id, op, args }));
    });
    return {
      leerTodo: () => llamar("leerTodo"),
      agregarFila: (...a) => llamar("agregarFila", a),
      actualizarPorId: (...a) => llamar("actualizarPorId", a),
      eliminarFila: (...a) => llamar("eliminarFila", a),
    };
  }

  async function crear(modo) {
    if (modo === "panel") return OpsExcel;
    if (modo === "ventana") return puente();
    return demo();
  }

  window.DATOS = { crear, OpsExcel, OPERACIONES: ["leerTodo", "agregarFila", "actualizarPorId", "eliminarFila"] };
})();
