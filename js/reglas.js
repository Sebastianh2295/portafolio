/* Lógica de negocio: roles, frecuencias, estado de seguimiento, compromisos, IDs y KPIs.
   No toca Excel ni pantallas; solo calcula. */
(function () {
  const DIAS_FRECUENCIA = { Diario: 1, Semanal: 7, Quincenal: 15, Mensual: 30 };
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

  // Una «sesión» (comité, mesa técnica, reunión) deja acta y compromisos, pero no es reporte de avance.
  const esSesion = (s) => !!s && s.Tipo_Sesion === "Sesión";
  const avancesDe = (pid, segs) => seguimientosDe(pid, segs).filter((s) => !esSesion(s));
  function estadoSeguimiento(p, segs, hoy = hoyISO()) {
    const lista = avancesDe(p.ID_Proyecto, segs);
    const ultimo = lista.length ? lista[lista.length - 1] : null;
    const diario = p.Frecuencia_Seguimiento === "Diario";
    const dias = DIAS_FRECUENCIA[p.Frecuencia_Seguimiento] || 7;
    const base = (ultimo && ultimo.Fecha_Corte) || p.Fecha_Inicio || hoy;
    // Diario: el siguiente día hábil (lunes a viernes) después del último reporte.
    let proximo = sumarDias(base, dias);
    if (diario) { while ([0, 6].includes(aFecha(proximo).getDay())) proximo = sumarDias(proximo, 1); }
    const faltan = difDias(proximo, hoy);
    let estado = "Al día";
    if (p.Estado !== "Activo") estado = "No aplica";
    else if (faltan < 0) estado = "Vencido";
    else if (faltan <= (diario ? 0 : DIAS_POR_VENCER)) estado = "Por vencer";
    const hace90 = sumarDias(hoy, -90);
    return { ultimo, ultimaFecha: ultimo ? ultimo.Fecha_Corte : "", proximo, faltan, estado, en90: lista.filter((s) => s.Fecha_Corte >= hace90).length };
  }

  // Estado registrado: Pendiente → En curso → Cerrado («Cumplido» de versiones anteriores cuenta como Cerrado).
  const ESTADOS_COMPROMISO = ["Pendiente", "En curso", "Cerrado"];
  const estadoBase = (c) => (c.Estado === "Cerrado" || c.Estado === "Cumplido" ? "Cerrado" : c.Estado === "En curso" ? "En curso" : "Pendiente");
  const cerrado = (c) => estadoBase(c) === "Cerrado";
  // Estado con alerta de tiempo: Cerrado | Vencido | Por vencer | Pendiente | En curso.
  function estadoCompromiso(c, hoy = hoyISO()) {
    const base = estadoBase(c);
    if (base === "Cerrado") return "Cerrado";
    if (!c.Fecha_Compromiso) return base;
    const f = difDias(c.Fecha_Compromiso, hoy);
    if (f < 0) return "Vencido";
    const aviso = c.Dias_Alerta === "" || c.Dias_Alerta === undefined || c.Dias_Alerta === null || Number.isNaN(Number(c.Dias_Alerta)) ? DIAS_COMPROMISO_POR_VENCER : Number(c.Dias_Alerta);
    if (f <= aviso) return "Por vencer";
    return base;
  }

  // Sugerencia de semáforo por desviación (plan - real). Es una guía: el PM decide.
  // Cronograma: puntos por debajo del plan (≤5 verde, ≤15 amarillo). Costo (si hay CPI): ≥0,95 verde, ≥0,85 amarillo. Gana el peor.
  const ORDEN_SEM = { Verde: 0, Amarillo: 1, Rojo: 2 };
  function semaforoSugerido(real, plan, cpi) {
    const d = (Number(plan) || 0) - (Number(real) || 0);
    const porTiempo = d <= 5 ? "Verde" : d <= 15 ? "Amarillo" : "Rojo";
    const porCosto = cpi === null || cpi === undefined || !isFinite(cpi) ? null : cpi >= 0.95 ? "Verde" : cpi >= 0.85 ? "Amarillo" : "Rojo";
    const semaforo = porCosto && ORDEN_SEM[porCosto] > ORDEN_SEM[porTiempo] ? porCosto : porTiempo;
    return { semaforo, desviacion: d, porTiempo, porCosto };
  }
  // Valor ganado (EVM). BAC = presupuesto de la línea base (o el actual); PV = %plan·BAC; EV = %real·BAC; AC = ejecutado.
  function valorGanado(bac, real, plan, ejecutado) {
    const BAC = Number(bac) || 0, r = Number(real) || 0, pl = Number(plan) || 0, AC = Number(ejecutado) || 0;
    if (!BAC) return null;
    const PV = (pl / 100) * BAC, EV = (r / 100) * BAC;
    const SPI = PV > 0 ? EV / PV : null, CPI = AC > 0 && EV > 0 ? EV / AC : null;
    const EAC = CPI ? BAC / CPI : null;
    return { BAC, PV, EV, AC, SPI, CPI, EAC, VAC: EAC !== null ? BAC - EAC : null, SV: EV - PV, CV: AC > 0 ? EV - AC : null };
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
    DIAS_COMPROMISO_POR_VENCER, ESTADOS_COMPROMISO, estadoBase, cerrado, seguimientosDe, estadoSeguimiento, estadoCompromiso, semaforoSugerido, valorGanado, esSesion, avancesDe, siguienteIdProyecto, siguienteIdHijo, kpis, calificacion, nivelRiesgo,
  };
})();
