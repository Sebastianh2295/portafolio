/* Lógica de negocio: roles, frecuencias, estado de seguimiento, compromisos, IDs y KPIs.
   No toca Excel ni pantallas; solo calcula. */
(function () {
  const DIAS_FRECUENCIA = { Semanal: 7, Quincenal: 15, Mensual: 30 };
  const DIAS_POR_VENCER = 2;
  const DIAS_COMPROMISO_POR_VENCER = 3;

  // Un usuario puede tener varios roles ("PM, PMO"): sus permisos se suman.
  const PERMISOS = {
    Admin:  ["crearProyecto", "editarProyecto", "editarFrecuencia", "seguimiento", "hitos", "riesgos", "compromisos", "usuarios", "catalogos", "verTodo"],
    PMO:    ["crearProyecto", "editarProyecto", "editarFrecuencia", "seguimiento", "hitos", "riesgos", "compromisos", "catalogos", "verTodo"],
    PM:     ["crearProyecto", "editarProyecto", "seguimiento", "hitos", "riesgos", "compromisos"],
    Lector: [],
  };
  const ROLES = Object.keys(PERMISOS);

  const pad = (n, w) => String(n).padStart(w, "0");
  const lc = (s) => String(s || "").trim().toLowerCase();
  const aISO = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`;
  const hoyISO = () => aISO(new Date());
  const aFecha = (iso) => { const [y, m, d] = String(iso).split("-").map(Number); return new Date(y, m - 1, d); };
  const sumarDias = (iso, n) => { const d = aFecha(iso); d.setDate(d.getDate() + n); return aISO(d); };
  const difDias = (a, b) => Math.round((aFecha(a) - aFecha(b)) / 86400000);

  function semanaISO(iso) {
    const d = aFecha(iso);
    const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const dia = t.getUTCDay() || 7;
    t.setUTCDate(t.getUTCDate() + 4 - dia);
    const inicio = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    return `${t.getUTCFullYear()}-S${pad(Math.ceil(((t - inicio) / 86400000 + 1) / 7), 2)}`;
  }

  const roles = (u) => String((u && u.Rol) || "").split(/[,;]/).map((r) => r.trim()).filter((r) => ROLES.includes(r));
  const tieneRol = (u, rol) => roles(u).includes(rol);
  const puede = (u, accion) => !!u && roles(u).some((r) => PERMISOS[r].includes(accion));
  const asignados = (u) => String(u.Proyectos || "").split(",").map((s) => s.trim()).filter(Boolean);
  const esPMde = (u, p) => tieneRol(u, "PM") && lc(p.PM) === lc(u.Correo);

  // Qué proyectos ve cada quien:
  //  - Admin y PMO: todos.
  //  - PM: los proyectos donde figura como PM en la tabla Proyectos (única fuente).
  //  - Lector: los de su columna Proyectos en la tabla Usuarios (IDs o "Todos").
  function esMio(u, p) {
    if (!u) return false;
    if (puede(u, "verTodo")) return true;
    if (esPMde(u, p)) return true;
    if (tieneRol(u, "Lector")) { const l = asignados(u); return l.includes("Todos") || l.includes(p.ID_Proyecto); }
    return false;
  }
  const visibles = (u, proyectos) => proyectos.filter((p) => esMio(u, p));
  const puedeEditar = (u, p, accion) => puede(u, accion) && (puede(u, "verTodo") || esPMde(u, p));
  const soloPM = (u) => tieneRol(u, "PM") && !puede(u, "verTodo");

  function seguimientosDe(pid, segs) {
    return segs.filter((s) => s.ID_Proyecto === pid).sort((a, b) => (a.Fecha_Corte < b.Fecha_Corte ? -1 : 1));
  }

  function estadoSeguimiento(p, segs, hoy = hoyISO()) {
    const lista = seguimientosDe(p.ID_Proyecto, segs);
    const ultimo = lista.length ? lista[lista.length - 1] : null;
    const dias = DIAS_FRECUENCIA[p.Frecuencia_Seguimiento] || 7;
    const base = (ultimo && ultimo.Fecha_Corte) || p.Fecha_Inicio || hoy;
    const proximo = sumarDias(base, dias);
    const faltan = difDias(proximo, hoy);
    let estado = "Al día";
    if (p.Estado !== "Activo") estado = "No aplica";
    else if (faltan < 0) estado = "Vencido";
    else if (faltan <= DIAS_POR_VENCER) estado = "Por vencer";
    const hace90 = sumarDias(hoy, -90);
    return { ultimo, ultimaFecha: ultimo ? ultimo.Fecha_Corte : "", proximo, faltan, estado, en90: lista.filter((s) => s.Fecha_Corte >= hace90).length };
  }

  function estadoCompromiso(c, hoy = hoyISO()) {
    if (c.Estado === "Cumplido") return "Cumplido";
    if (!c.Fecha_Compromiso) return "Pendiente";
    const f = difDias(c.Fecha_Compromiso, hoy);
    if (f < 0) return "Vencido";
    if (f <= DIAS_COMPROMISO_POR_VENCER) return "Por vencer";
    return "Pendiente";
  }

  // Sugerencia de semáforo por desviación (plan - real). Es una guía: el PM decide.
  function semaforoSugerido(real, plan) {
    const d = (Number(plan) || 0) - (Number(real) || 0);
    return { semaforo: d <= 5 ? "Verde" : d <= 15 ? "Amarillo" : "Rojo", desviacion: d };
  }

  function siguienteIdProyecto(proyectos) {
    const max = proyectos.reduce((m, p) => Math.max(m, parseInt(String(p.ID_Proyecto).slice(4), 10) || 0), 0);
    return `PRY-${pad(max + 1, 4)}`;
  }
  function siguienteIdHijo(prefijo, filas, colId, pid, desplazamiento = 0) {
    const base = `${prefijo}-${String(pid).slice(4)}-`;
    const max = filas.reduce((m, f) => {
      const id = String(f[colId]);
      return id.startsWith(base) ? Math.max(m, parseInt(id.slice(base.length), 10) || 0) : m;
    }, 0);
    return `${base}${max + 1 + desplazamiento}`;
  }

  const num = (v) => Number(v) || 0;
  const prom = (arr) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);

  function kpis(proyectos, segs) {
    const activos = proyectos.filter((p) => p.Estado === "Activo");
    const pres = activos.reduce((a, p) => a + num(p.Presupuesto), 0);
    const ejec = activos.reduce((a, p) => a + num(p.Ejecutado), 0);
    const alDia = activos.filter((p) => estadoSeguimiento(p, segs).estado !== "Vencido").length;
    return {
      total: proyectos.length, activos: activos.length,
      avanceReal: prom(activos.map((p) => num(p.Avance_Real))), avancePlan: prom(activos.map((p) => num(p.Avance_Planeado))),
      rojos: activos.filter((p) => p.Semaforo === "Rojo").length,
      presupuesto: pres, ejecutado: ejec, pctEjecutado: pres ? (ejec / pres) * 100 : 0,
      pctAlDia: activos.length ? (alDia / activos.length) * 100 : 100,
    };
  }

  const calificacion = (prob, imp) => num(prob) * num(imp);
  const nivelRiesgo = (c) => (c > 12 ? "Alto" : c > 6 ? "Medio" : "Bajo");

  window.REGLAS = {
    ROLES, DIAS_FRECUENCIA, hoyISO, sumarDias, difDias, semanaISO, roles, tieneRol, puede, puedeEditar, soloPM, esPMde, esMio, visibles,
    seguimientosDe, estadoSeguimiento, estadoCompromiso, semaforoSugerido, siguienteIdProyecto, siguienteIdHijo, kpis, calificacion, nivelRiesgo,
  };
})();
