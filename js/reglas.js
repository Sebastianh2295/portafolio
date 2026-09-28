/* Lógica de negocio: roles, frecuencias, estado de seguimiento, IDs y KPIs.
   No toca Excel ni pantallas; solo calcula. */
(function () {
  const DIAS_FRECUENCIA = { Semanal: 7, Quincenal: 15, Mensual: 30 };
  const DIAS_POR_VENCER = 2;

  const PERMISOS = {
    Admin:  ["crearProyecto", "editarProyecto", "editarFrecuencia", "seguimiento", "hitos", "riesgos", "usuarios", "verTodo"],
    PMO:    ["crearProyecto", "editarProyecto", "editarFrecuencia", "seguimiento", "hitos", "riesgos", "verTodo"],
    PM:     ["crearProyecto", "editarProyecto", "seguimiento", "hitos", "riesgos"],
    Lector: [],
  };

  const pad = (n, w) => String(n).padStart(w, "0");
  function hoyISO() {
    const d = new Date();
    return `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`;
  }
  const aFecha = (iso) => { const [y, m, d] = String(iso).split("-").map(Number); return new Date(y, m - 1, d); };
  const aISO = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`;
  const sumarDias = (iso, n) => { const d = aFecha(iso); d.setDate(d.getDate() + n); return aISO(d); };
  const difDias = (a, b) => Math.round((aFecha(a) - aFecha(b)) / 86400000);

  function semanaISO(iso) {
    const d = aFecha(iso);
    const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dia = t.getUTCDay() || 7;
    t.setUTCDate(t.getUTCDate() + 4 - dia);
    const inicio = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    const sem = Math.ceil(((t - inicio) / 86400000 + 1) / 7);
    return `${t.getUTCFullYear()}-S${pad(sem, 2)}`;
  }

  const puede = (u, accion) => !!u && (PERMISOS[u.Rol] || []).includes(accion);
  const asignados = (u) => String(u.Proyectos || "").split(",").map((s) => s.trim()).filter(Boolean);

  function esMio(u, p) {
    if (!u) return false;
    if (puede(u, "verTodo")) return true;
    const lista = asignados(u);
    if (lista.includes("Todos")) return true;
    return lista.includes(p.ID_Proyecto) || String(p.PM).toLowerCase() === String(u.Correo).toLowerCase();
  }
  const visibles = (u, proyectos) => proyectos.filter((p) => esMio(u, p));
  const puedeEditar = (u, p, accion) => puede(u, accion) && (puede(u, "verTodo") || esMio(u, p));

  function seguimientosDe(pid, segs) {
    return segs.filter((s) => s.ID_Proyecto === pid).sort((a, b) => (a.Fecha_Corte < b.Fecha_Corte ? -1 : 1));
  }

  function estadoSeguimiento(p, segs, hoy = hoyISO()) {
    const lista = seguimientosDe(p.ID_Proyecto, segs);
    const ultimo = lista.length ? lista[lista.length - 1].Fecha_Corte : "";
    const dias = DIAS_FRECUENCIA[p.Frecuencia_Seguimiento] || 7;
    const base = ultimo || p.Fecha_Inicio || hoy;
    const proximo = sumarDias(base, dias);
    const faltan = difDias(proximo, hoy);
    let estado = "Al día";
    if (p.Estado !== "Activo") estado = "No aplica";
    else if (faltan < 0) estado = "Vencido";
    else if (faltan <= DIAS_POR_VENCER) estado = "Por vencer";
    const hace90 = sumarDias(hoy, -90);
    return { ultimo, proximo, faltan, estado, en90: lista.filter((s) => s.Fecha_Corte >= hace90).length };
  }

  function siguienteIdProyecto(proyectos) {
    const max = proyectos.reduce((m, p) => Math.max(m, parseInt(String(p.ID_Proyecto).slice(4), 10) || 0), 0);
    return `PRY-${pad(max + 1, 4)}`;
  }
  function siguienteIdHijo(prefijo, filas, colId, pid) {
    const num = String(pid).slice(4);
    const base = `${prefijo}-${num}-`;
    const max = filas.reduce((m, f) => {
      const id = String(f[colId]);
      return id.startsWith(base) ? Math.max(m, parseInt(id.slice(base.length), 10) || 0) : m;
    }, 0);
    return `${base}${max + 1}`;
  }

  const num = (v) => Number(v) || 0;
  const prom = (arr) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);

  function kpis(proyectos, segs) {
    const activos = proyectos.filter((p) => p.Estado === "Activo");
    const pres = activos.reduce((a, p) => a + num(p.Presupuesto), 0);
    const ejec = activos.reduce((a, p) => a + num(p.Ejecutado), 0);
    const alDia = activos.filter((p) => estadoSeguimiento(p, segs).estado !== "Vencido").length;
    return {
      total: proyectos.length,
      activos: activos.length,
      avanceReal: prom(activos.map((p) => num(p.Avance_Real))),
      avancePlan: prom(activos.map((p) => num(p.Avance_Planeado))),
      rojos: activos.filter((p) => p.Semaforo === "Rojo").length,
      presupuesto: pres,
      ejecutado: ejec,
      pctEjecutado: pres ? (ejec / pres) * 100 : 0,
      pctAlDia: activos.length ? (alDia / activos.length) * 100 : 100,
    };
  }

  const calificacion = (prob, imp) => num(prob) * num(imp);
  const nivelRiesgo = (c) => (c > 12 ? "Alto" : c > 6 ? "Medio" : "Bajo");

  window.REGLAS = {
    DIAS_FRECUENCIA, hoyISO, sumarDias, difDias, semanaISO, puede, puedeEditar, esMio, visibles,
    seguimientosDe, estadoSeguimiento, siguienteIdProyecto, siguienteIdHijo, kpis, calificacion, nivelRiesgo,
  };
})();
