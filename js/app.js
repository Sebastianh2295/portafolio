/* Pantallas de la app: login, menú, módulos y formularios. */
/* global Office, Chart, DATOS, REGLAS */
(function () {
  const R = REGLAS;
  const S = {
    api: null, modo: "demo", datos: null, usuario: null, vista: "dashboard", pid: null,
    cat: {}, charts: [], semana: null, riesgoVista: "inherente",
    filtros: { texto: "", estado: "Activo", semaforo: "", pm: "" },
  };
  const COLORES = { azul: "#104994", azulClaro: "#6BBAEF", dorado: "#D6964A", Verde: "#2E8B57", Amarillo: "#E0A800", Rojo: "#C0392B" };
  const VISTAS = [
    { id: "dashboard", t: "Dashboard" },
    { id: "avances", t: "Avances de la semana" },
    { id: "seguimiento", t: "Control de seguimiento" },
    { id: "proyectos", t: "Proyectos" },
    { id: "cronograma", t: "Cronograma" },
    { id: "riesgos", t: "Riesgos" },
    { id: "catalogos", t: "Catálogos", permiso: "catalogos" },
    { id: "usuarios", t: "Usuarios", permiso: "usuarios" },
  ];

  const $ = (s, el = document) => el.querySelector(s);
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const cop = (n) => "$" + (Number(n) || 0).toLocaleString("es-CO", { maximumFractionDigits: 0 });
  const pct = (n) => `${Math.round(Number(n) || 0)}%`;
  const app = () => $("#app");

  function guardarSesion(correo) { try { sessionStorage.setItem("pmo_usuario", correo || ""); } catch (e) { /* sin almacenamiento */ } }
  function leerSesion() { try { return sessionStorage.getItem("pmo_usuario") || ""; } catch (e) { return ""; } }

  function toast(msg, error) {
    const t = document.createElement("div");
    t.className = "toast" + (error ? " error" : "");
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), error ? 7000 : 3000);
  }
  function cargando(on, msg = "Cargando…") {
    let o = $("#cargando");
    if (on) {
      if (!o) { o = document.createElement("div"); o.id = "cargando"; document.body.appendChild(o); }
      o.innerHTML = `<div class="spinner"></div><div>${esc(msg)}</div>`;
    } else if (o) o.remove();
  }

  // ---------- Datos ----------
  function armarCatalogos() {
    S.cat = {};
    for (const f of S.datos.Catalogos) (S.cat[f.Lista] = S.cat[f.Lista] || []).push(String(f.Valor));
  }
  async function recargar() {
    S.datos = await S.api.leerTodo();
    armarCatalogos();
    if (S.usuario) {
      const u = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(S.usuario.Correo) && x.Activo === "Sí");
      S.usuario = u || null;
    }
  }
  const lc = (s) => String(s || "").trim().toLowerCase();
  const visibles = () => R.visibles(S.usuario, S.datos.Proyectos);
  const proyecto = (pid) => S.datos.Proyectos.find((p) => p.ID_Proyecto === pid);
  const nombreUsuario = (correo) => (S.datos.Usuarios.find((u) => lc(u.Correo) === lc(correo)) || {}).Nombre || correo;

  async function guardar(accion, exito = "Guardado") {
    cargando(true, "Guardando en Excel…");
    try {
      await accion();
      await recargar();
      cerrarModal();
      render();
      toast(exito);
    } catch (e) {
      toast("No se pudo guardar: " + e.message, true);
    } finally {
      cargando(false);
    }
  }
  const sello = () => ({ Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() });

  // ---------- Login ----------
  function pantallaLogin(error) {
    const demo = S.modo === "demo"
      ? `<p class="nota">Modo demostración. Pruebe con <code>tu.correo@empresa.com</code> (Admin), <code>pm1@empresa.com</code> (PM) o <code>direccion@empresa.com</code> (Lector).</p>` : "";
    app().innerHTML = `
      <div class="login">
        <div class="login-caja">
          <div class="marca"><span class="logo">P</span> Portafolio de Proyectos</div>
          <p>Ingrese su correo corporativo para continuar.</p>
          <form id="f-login" novalidate>
            <label>Correo<input id="correo" type="email" autocomplete="email" required></label>
            <div class="error-campo" id="e-login">${esc(error || "")}</div>
            <button class="btn primario" type="submit">Entrar</button>
          </form>
          ${demo}
        </div>
      </div>`;
    $("#correo").focus();
    $("#correo").addEventListener("input", () => { $("#e-login").textContent = ""; });
    $("#f-login").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const correo = $("#correo").value.trim();
      if (!correo) { $("#e-login").textContent = "Escriba su correo."; return; }
      const u = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(correo));
      if (!u) { $("#e-login").textContent = "Este correo no está registrado en el portafolio. Pida acceso al administrador."; return; }
      if (u.Activo !== "Sí") { $("#e-login").textContent = "Su usuario está inactivo. Pida acceso al administrador."; return; }
      S.usuario = u;
      guardarSesion(u.Correo);
      S.vista = "dashboard";
      render();
    });
  }

  // ---------- Estructura ----------
  function render() {
    S.charts.forEach((c) => c.destroy());
    S.charts = [];
    if (!S.usuario) return pantallaLogin();
    const menu = VISTAS.filter((v) => !v.permiso || R.puede(S.usuario, v.permiso));
    const activa = S.vista === "ficha" ? "proyectos" : S.vista;
    const etiquetaModo = { demo: "Demostración: los cambios no se guardan", panel: "Conectado a Excel", ventana: "Conectado a Excel" }[S.modo];
    app().innerHTML = `
      <div class="layout">
        <aside class="lateral">
          <div class="marca"><span class="logo">P</span> Portafolio PMO</div>
          <nav>${menu.map((v) => `<button class="nav ${v.id === activa ? "activa" : ""}" data-vista="${v.id}">${esc(v.t)}</button>`).join("")}</nav>
          <select id="nav-movil" aria-label="Módulo">${menu.map((v) => `<option value="${v.id}" ${v.id === activa ? "selected" : ""}>${esc(v.t)}</option>`).join("")}</select>
        </aside>
        <main>
          <header class="barra">
            <span class="modo ${S.modo}">${esc(etiquetaModo)}</span>
            <span class="quien">${esc(S.usuario.Nombre || S.usuario.Correo)} · <b>${esc(S.usuario.Rol)}</b></span>
            <button class="btn" id="b-recargar" title="Volver a leer el Excel">Actualizar</button>
            <button class="btn" id="b-salir">Salir</button>
          </header>
          <section id="vista"></section>
        </main>
      </div>`;
    document.querySelectorAll(".nav").forEach((b) => b.addEventListener("click", () => ir(b.dataset.vista)));
    $("#nav-movil").addEventListener("change", (e) => ir(e.target.value));
    $("#b-salir").addEventListener("click", () => { S.usuario = null; guardarSesion(""); render(); });
    $("#b-recargar").addEventListener("click", async () => {
      cargando(true, "Leyendo Excel…");
      try { await recargar(); render(); toast("Datos actualizados"); } catch (e) { toast(e.message, true); } finally { cargando(false); }
    });
    const vistas = { dashboard: vDashboard, avances: vAvances, seguimiento: vSeguimiento, proyectos: vProyectos, ficha: vFicha, cronograma: vCronograma, riesgos: vRiesgos, catalogos: vCatalogos, usuarios: vUsuarios };
    (vistas[S.vista] || vDashboard)($("#vista"));
  }
  function ir(vista, pid) { S.vista = vista; if (pid) S.pid = pid; render(); window.scrollTo(0, 0); }

  const chipSemaforo = (s) => `<span class="dot" style="background:${COLORES[s] || "#999"}"></span>${esc(s)}`;
  const chipEstadoSeg = (e) => `<span class="pill ${e.replace(/\s/g, "").toLowerCase()}">${esc(e)}</span>`;
  const barraAvance = (real, plan) => `
    <div class="avance" title="Real ${pct(real)} · Planeado ${pct(plan)}">
      <div class="avance-real" style="width:${Math.min(100, Number(real) || 0)}%"></div>
      <div class="avance-plan" style="left:${Math.min(100, Number(plan) || 0)}%"></div>
    </div><span class="avance-txt">${pct(real)} / ${pct(plan)}</span>`;
  const vacio = (msg) => `<div class="vacio">${esc(msg)}</div>`;

  function grafico(canvas, config) {
    if (typeof Chart === "undefined") { canvas.replaceWith(Object.assign(document.createElement("div"), { className: "vacio", textContent: "No se pudo cargar la librería de gráficos." })); return; }
    Chart.defaults.font.family = "Calibri, Arial, sans-serif";
    S.charts.push(new Chart(canvas, config));
  }

  // ---------- Dashboard ----------
  function vDashboard(el) {
    const ps = visibles();
    const k = R.kpis(ps, S.datos.Seguimientos);
    const activos = ps.filter((p) => p.Estado === "Activo");
    const tarjeta = (titulo, valor, detalle, alerta) => `<div class="kpi ${alerta ? "alerta" : ""}"><div class="kpi-t">${titulo}</div><div class="kpi-v">${valor}</div><div class="kpi-d">${detalle}</div></div>`;
    el.innerHTML = `
      <h1>Dashboard del portafolio</h1>
      <div class="kpis">
        ${tarjeta("Proyectos activos", k.activos, `de ${k.total} en el portafolio`)}
        ${tarjeta("Avance real vs planeado", `${pct(k.avanceReal)} <small>/ ${pct(k.avancePlan)}</small>`, "promedio de proyectos activos", k.avanceReal + 5 < k.avancePlan)}
        ${tarjeta("Proyectos en rojo", k.rojos, "activos con semáforo rojo", k.rojos > 0)}
        ${tarjeta("Presupuesto ejecutado", pct(k.pctEjecutado), `${cop(k.ejecutado)} de ${cop(k.presupuesto)}`)}
        ${tarjeta("Seguimientos al día", pct(k.pctAlDia), "proyectos activos sin reporte vencido", k.pctAlDia < 80)}
      </div>
      <div class="grid2">
        <div class="card"><h2>Semáforo de proyectos activos</h2><canvas id="g-sem" height="220"></canvas></div>
        <div class="card"><h2>Proyectos por estado</h2><canvas id="g-est" height="220"></canvas></div>
      </div>
      <div class="card"><h2>Avance real vs planeado por proyecto activo</h2><div style="height:${Math.max(160, activos.length * 34 + 60)}px"><canvas id="g-av"></canvas></div></div>`;
    const sems = ["Verde", "Amarillo", "Rojo"];
    grafico($("#g-sem"), { type: "bar", data: { labels: sems, datasets: [{ data: sems.map((s) => activos.filter((p) => p.Semaforo === s).length), backgroundColor: sems.map((s) => COLORES[s]), borderRadius: 4 }] },
      options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } } });
    const estados = S.cat.Estado || ["Activo", "En pausa", "Cerrado", "Cancelado"];
    grafico($("#g-est"), { type: "bar", data: { labels: estados, datasets: [{ data: estados.map((s) => ps.filter((p) => p.Estado === s).length), backgroundColor: COLORES.azul, borderRadius: 4 }] },
      options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } } });
    grafico($("#g-av"), { type: "bar", data: { labels: activos.map((p) => p.Nombre), datasets: [
      { label: "Real", data: activos.map((p) => Number(p.Avance_Real) || 0), backgroundColor: COLORES.azul, borderRadius: 3 },
      { label: "Planeado", data: activos.map((p) => Number(p.Avance_Planeado) || 0), backgroundColor: COLORES.azulClaro, borderRadius: 3 }] },
      options: { indexAxis: "y", maintainAspectRatio: false, scales: { x: { min: 0, max: 100, ticks: { callback: (v) => v + "%" } } } } });
  }

  // ---------- Avances de la semana ----------
  function vAvances(el) {
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const segs = S.datos.Seguimientos.filter((s) => ids.has(s.ID_Proyecto));
    const actual = R.semanaISO(R.hoyISO());
    const semanas = [...new Set([actual, ...segs.map((s) => s.Semana || R.semanaISO(s.Fecha_Corte))])].sort().reverse();
    if (!S.semana) S.semana = semanas.find((w) => w <= actual && segs.some((s) => (s.Semana || R.semanaISO(s.Fecha_Corte)) === w)) || actual;
    const deSemana = segs.filter((s) => (s.Semana || R.semanaISO(s.Fecha_Corte)) === S.semana);
    const filas = deSemana.map((s) => {
      const hist = R.seguimientosDe(s.ID_Proyecto, segs);
      const i = hist.findIndex((h) => h.ID_Seguimiento === s.ID_Seguimiento);
      const ant = i > 0 ? hist[i - 1] : null;
      const variacion = ant ? (Number(s.Avance_Real) || 0) - (Number(ant.Avance_Real) || 0) : null;
      const p = proyecto(s.ID_Proyecto) || {};
      return `<tr class="clic" data-pid="${esc(s.ID_Proyecto)}">
        <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(s.ID_Proyecto)} · ${esc(nombreUsuario(p.PM))}</div></td>
        <td>${esc(s.Fecha_Corte)}</td><td>${pct(s.Avance_Real)}</td>
        <td class="${variacion > 0 ? "sube" : variacion < 0 ? "baja" : ""}">${variacion === null ? "—" : (variacion > 0 ? "+" : "") + variacion + " pts"}</td>
        <td>${chipSemaforo(s.Semaforo)}</td><td>${esc(s.Logros)}</td><td>${esc(s.Proximos_Pasos)}</td><td class="${s.Bloqueos ? "bloqueo" : ""}">${esc(s.Bloqueos || "—")}</td></tr>`;
    }).join("");
    const reportaron = new Set(deSemana.map((s) => s.ID_Proyecto));
    const sinReporte = visibles().filter((p) => p.Estado === "Activo" && p.Frecuencia_Seguimiento === "Semanal" && !reportaron.has(p.ID_Proyecto));
    el.innerHTML = `
      <h1>Avances de la semana</h1>
      <div class="filtros"><label>Semana <select id="sel-semana">${semanas.map((w) => `<option ${w === S.semana ? "selected" : ""}>${w}</option>`).join("")}</select></label>
        <span class="sub">${deSemana.length} seguimiento(s) reportado(s)</span></div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>Corte</th><th>Avance</th><th>Variación</th><th>Semáforo</th><th>Logros</th><th>Próximos pasos</th><th>Bloqueos</th></tr></thead>
        <tbody>${filas || `<tr><td colspan="8">${vacio("No hay seguimientos reportados en esta semana.")}</td></tr>`}</tbody></table></div></div>
      <div class="card"><h2>Proyectos semanales sin reporte en ${esc(S.semana)}</h2>
        ${sinReporte.length ? `<ul class="lista">${sinReporte.map((p) => `<li class="clic" data-pid="${esc(p.ID_Proyecto)}"><b>${esc(p.Nombre)}</b> · ${esc(nombreUsuario(p.PM))}</li>`).join("")}</ul>` : vacio("Todos los proyectos semanales reportaron.")}</div>`;
    $("#sel-semana").addEventListener("change", (e) => { S.semana = e.target.value; render(); });
    el.querySelectorAll(".clic").forEach((r) => r.addEventListener("click", () => ir("ficha", r.dataset.pid)));
  }

  // ---------- Control de seguimiento ----------
  function vSeguimiento(el) {
    const orden = { Vencido: 0, "Por vencer": 1, "Al día": 2 };
    const filas = visibles().filter((p) => p.Estado === "Activo")
      .map((p) => ({ p, e: R.estadoSeguimiento(p, S.datos.Seguimientos) }))
      .sort((a, b) => orden[a.e.estado] - orden[b.e.estado] || a.e.faltan - b.e.faltan);
    const cuenta = (e) => filas.filter((f) => f.e.estado === e).length;
    el.innerHTML = `
      <h1>Control de seguimiento</h1>
      <div class="kpis">
        <div class="kpi ${cuenta("Vencido") ? "alerta" : ""}"><div class="kpi-t">Vencidos</div><div class="kpi-v">${cuenta("Vencido")}</div></div>
        <div class="kpi"><div class="kpi-t">Por vencer</div><div class="kpi-v">${cuenta("Por vencer")}</div></div>
        <div class="kpi"><div class="kpi-t">Al día</div><div class="kpi-v">${cuenta("Al día")}</div></div>
      </div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>PM</th><th>Frecuencia</th><th>Último seguimiento</th><th>Próximo</th><th>Estado</th><th>Seguimientos en 90 días</th></tr></thead>
        <tbody>${filas.map(({ p, e }) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}">
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.ID_Proyecto)}</div></td><td>${esc(nombreUsuario(p.PM))}</td>
          <td>${esc(p.Frecuencia_Seguimiento)}</td><td>${esc(e.ultimo || "Sin reportes")}</td><td>${esc(e.proximo)}</td>
          <td>${chipEstadoSeg(e.estado)} <span class="sub">${e.faltan < 0 ? `${-e.faltan} día(s) de retraso` : `faltan ${e.faltan} día(s)`}</span></td>
          <td>${e.en90}</td></tr>`).join("") || `<tr><td colspan="7">${vacio("No hay proyectos activos.")}</td></tr>`}</tbody></table></div></div>
      <p class="sub">Regla: próximo seguimiento = último + frecuencia (Semanal 7, Quincenal 15, Mensual 30 días). "Por vencer" = faltan 2 días o menos.</p>`;
    el.querySelectorAll(".clic").forEach((r) => r.addEventListener("click", () => ir("ficha", r.dataset.pid)));
  }

  // ---------- Proyectos ----------
  function vProyectos(el) {
    const F = S.filtros;
    const todos = visibles();
    const pms = [...new Set(todos.map((p) => p.PM))];
    const lista = todos.filter((p) =>
      (!F.estado || p.Estado === F.estado) && (!F.semaforo || p.Semaforo === F.semaforo) && (!F.pm || p.PM === F.pm) &&
      (!F.texto || `${p.ID_Proyecto} ${p.Nombre} ${p.Cliente_Area}`.toLowerCase().includes(F.texto.toLowerCase())));
    const opts = (arr, sel, todosTxt) => `<option value="">${todosTxt}</option>` + arr.map((v) => `<option value="${esc(v)}" ${v === sel ? "selected" : ""}>${esc(v)}</option>`).join("");
    el.innerHTML = `
      <div class="titulo-fila"><h1>Proyectos</h1>${R.puede(S.usuario, "crearProyecto") ? `<button class="btn primario" id="b-nuevo">+ Nueva iniciativa</button>` : ""}</div>
      <div class="filtros">
        <input id="f-texto" placeholder="Buscar por ID, nombre o área" value="${esc(F.texto)}">
        <select id="f-estado">${opts(S.cat.Estado || [], F.estado, "Todos los estados")}</select>
        <select id="f-sem">${opts(S.cat.Semaforo || [], F.semaforo, "Todos los semáforos")}</select>
        ${pms.length > 1 ? `<select id="f-pm"><option value="">Todos los PM</option>${pms.map((m) => `<option value="${esc(m)}" ${m === F.pm ? "selected" : ""}>${esc(nombreUsuario(m))}</option>`).join("")}</select>` : ""}
        <span class="sub">${lista.length} proyecto(s)</span>
      </div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>Área</th><th>PM</th><th>Fase</th><th>Estado</th><th>Semáforo</th><th>Avance real / plan</th><th>Fin plan</th><th>Seguimiento</th></tr></thead>
        <tbody>${lista.map((p) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}">
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.ID_Proyecto)} · ${esc(p.Prioridad)}</div></td><td>${esc(p.Cliente_Area)}</td>
          <td>${esc(nombreUsuario(p.PM))}</td><td>${esc(p.Fase)}</td><td>${esc(p.Estado)}</td><td>${chipSemaforo(p.Semaforo)}</td>
          <td class="celda-avance">${barraAvance(p.Avance_Real, p.Avance_Planeado)}</td><td>${esc(p.Fecha_Fin_Plan)}</td>
          <td>${chipEstadoSeg(R.estadoSeguimiento(p, S.datos.Seguimientos).estado)}</td></tr>`).join("") || `<tr><td colspan="9">${vacio("No hay proyectos con estos filtros.")}</td></tr>`}</tbody>
      </table></div></div>`;
    const set = (k) => (e) => { F[k] = e.target.value; render(); };
    $("#f-texto").addEventListener("change", set("texto"));
    $("#f-estado").addEventListener("change", set("estado"));
    $("#f-sem").addEventListener("change", set("semaforo"));
    if ($("#f-pm")) $("#f-pm").addEventListener("change", set("pm"));
    if ($("#b-nuevo")) $("#b-nuevo").addEventListener("click", () => formProyecto());
    el.querySelectorAll(".clic").forEach((r) => r.addEventListener("click", () => ir("ficha", r.dataset.pid)));
  }

  // ---------- Ficha ----------
  function vFicha(el) {
    const p = proyecto(S.pid);
    if (!p || !R.esMio(S.usuario, p)) { el.innerHTML = vacio("Proyecto no encontrado o sin acceso."); return; }
    const u = S.usuario;
    const segs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos);
    const hitos = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto).sort((a, b) => (a.Fecha_Plan < b.Fecha_Plan ? -1 : 1));
    const riesgos = S.datos.Riesgos.filter((r) => r.ID_Proyecto === p.ID_Proyecto);
    const e = R.estadoSeguimiento(p, S.datos.Seguimientos);
    const puedeP = (a) => R.puedeEditar(u, p, a);
    const dato = (k, v) => `<div class="dato"><div class="dato-k">${k}</div><div class="dato-v">${v}</div></div>`;
    el.innerHTML = `
      <button class="btn enlace" id="b-volver">← Proyectos</button>
      <div class="titulo-fila"><div><h1>${esc(p.Nombre)}</h1><div class="sub">${esc(p.ID_Proyecto)} · ${esc(p.Cliente_Area)} · PM ${esc(nombreUsuario(p.PM))}</div></div>
        <div class="acciones">
          ${puedeP("seguimiento") && p.Estado === "Activo" ? `<button class="btn primario" id="b-seg">Registrar seguimiento</button>` : ""}
          ${puedeP("editarProyecto") ? `<button class="btn" id="b-editar">Editar</button>` : ""}
        </div></div>
      <div class="card datos">
        ${dato("Estado", esc(p.Estado))}${dato("Fase", esc(p.Fase))}${dato("Semáforo", chipSemaforo(p.Semaforo))}
        ${dato("Avance real / plan", barraAvance(p.Avance_Real, p.Avance_Planeado))}
        ${dato("Inicio", esc(p.Fecha_Inicio))}${dato("Fin plan", esc(p.Fecha_Fin_Plan))}
        ${dato("Presupuesto", cop(p.Presupuesto))}${dato("Ejecutado", cop(p.Ejecutado))}
        ${dato("Metodología", esc(p.Metodologia))}${dato("Prioridad", esc(p.Prioridad))}
        ${dato("Frecuencia", esc(p.Frecuencia_Seguimiento))}${dato("Seguimiento", `${chipEstadoSeg(e.estado)} <span class="sub">próximo ${esc(e.proximo)}</span>`)}
        <div class="dato ancho"><div class="dato-k">Comentario de estado</div><div class="dato-v">${esc(p.Comentario_Estado || "—")}</div></div>
      </div>
      <div class="card"><h2>Curva de avance</h2>${segs.length ? `<div style="height:220px"><canvas id="g-curva"></canvas></div>` : vacio("Aún no hay seguimientos.")}</div>
      <div class="card"><h2>Historial de seguimientos</h2><div class="tabla-scroll"><table>
        <thead><tr><th>Corte</th><th>Semana</th><th>Avance</th><th>Semáforo</th><th>Logros</th><th>Próximos pasos</th><th>Bloqueos</th><th>Reportó</th></tr></thead>
        <tbody>${segs.slice().reverse().map((s) => `<tr><td>${esc(s.Fecha_Corte)}</td><td>${esc(s.Semana)}</td><td>${pct(s.Avance_Real)}</td><td>${chipSemaforo(s.Semaforo)}</td>
          <td>${esc(s.Logros)}</td><td>${esc(s.Proximos_Pasos)}</td><td>${esc(s.Bloqueos || "—")}</td><td>${esc(nombreUsuario(s.Reportado_Por))}</td></tr>`).join("") || `<tr><td colspan="8">${vacio("Sin seguimientos.")}</td></tr>`}</tbody></table></div></div>
      <div class="card"><div class="titulo-fila"><h2>Hitos</h2>${puedeP("hitos") ? `<button class="btn" id="b-hito">+ Hito</button>` : ""}</div><div class="tabla-scroll"><table>
        <thead><tr><th>Hito</th><th>Fecha plan</th><th>Fecha real</th><th>Estado</th><th></th></tr></thead>
        <tbody>${hitos.map((h) => `<tr><td>${esc(h.Hito)}</td><td>${esc(h.Fecha_Plan)}</td><td>${esc(h.Fecha_Real || "—")}</td><td>${esc(h.Estado)}</td>
          <td class="derecha">${puedeP("hitos") ? `<button class="btn chico" data-hito="${esc(h.ID_Hito)}">Editar</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="5">${vacio("Sin hitos.")}</td></tr>`}</tbody></table></div></div>
      <div class="card"><div class="titulo-fila"><h2>Riesgos</h2>${puedeP("riesgos") ? `<button class="btn" id="b-riesgo">+ Riesgo</button>` : ""}</div><div class="tabla-scroll"><table>
        <thead><tr><th>Riesgo</th><th>Tipo</th><th>Estado</th><th>Inherente</th><th>Residual</th><th>Mitigación</th><th></th></tr></thead>
        <tbody>${riesgos.map((r) => `<tr><td>${esc(r.Descripcion)}</td><td>${esc(r.Tipo)}</td><td>${esc(r.Estado)}</td>
          <td>${nivel(r.Calificacion_Inherente)}</td><td>${nivel(r.Calificacion_Residual)}</td><td>${esc(r.Plan_Mitigacion)}</td>
          <td class="derecha">${puedeP("riesgos") ? `<button class="btn chico" data-riesgo="${esc(r.ID_Riesgo)}">Editar</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="7">${vacio("Sin riesgos.")}</td></tr>`}</tbody></table></div></div>`;
    $("#b-volver").addEventListener("click", () => ir("proyectos"));
    if ($("#b-seg")) $("#b-seg").addEventListener("click", () => formSeguimiento(p));
    if ($("#b-editar")) $("#b-editar").addEventListener("click", () => formProyecto(p));
    if ($("#b-hito")) $("#b-hito").addEventListener("click", () => formHito(p));
    if ($("#b-riesgo")) $("#b-riesgo").addEventListener("click", () => formRiesgo(p));
    el.querySelectorAll("[data-hito]").forEach((b) => b.addEventListener("click", () => formHito(p, hitos.find((h) => h.ID_Hito === b.dataset.hito))));
    el.querySelectorAll("[data-riesgo]").forEach((b) => b.addEventListener("click", () => formRiesgo(p, riesgos.find((r) => r.ID_Riesgo === b.dataset.riesgo))));
    if (segs.length) grafico($("#g-curva"), { type: "line", data: { labels: segs.map((s) => s.Fecha_Corte), datasets: [
      { label: "Avance real", data: segs.map((s) => Number(s.Avance_Real) || 0), borderColor: COLORES.azul, backgroundColor: COLORES.azul, tension: 0.2 }] },
      options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + "%" } } } } });
  }
  const nivel = (c) => { const n = R.nivelRiesgo(Number(c) || 0); return `<span class="pill riesgo-${n.toLowerCase()}">${n} (${Number(c) || 0})</span>`; };

  // ---------- Cronograma ----------
  function vCronograma(el) {
    const ps = visibles().filter((p) => (p.Estado === "Activo" || p.Estado === "En pausa") && p.Fecha_Inicio && p.Fecha_Fin_Plan)
      .sort((a, b) => (a.Fecha_Inicio < b.Fecha_Inicio ? -1 : 1));
    if (!ps.length) { el.innerHTML = `<h1>Cronograma</h1>${vacio("No hay proyectos activos con fechas.")}`; return; }
    const ini = ps.reduce((m, p) => (p.Fecha_Inicio < m ? p.Fecha_Inicio : m), ps[0].Fecha_Inicio).slice(0, 8) + "01";
    const finMax = ps.reduce((m, p) => (p.Fecha_Fin_Plan > m ? p.Fecha_Fin_Plan : m), ps[0].Fecha_Fin_Plan);
    const total = Math.max(1, R.difDias(finMax, ini) + 1);
    const x = (iso) => Math.max(0, Math.min(100, (R.difDias(iso, ini) / total) * 100));
    const meses = [];
    for (let d = new Date(ini + "T00:00:00"); d.toISOString().slice(0, 10) <= finMax; d.setMonth(d.getMonth() + 1)) {
      meses.push({ iso: d.toISOString().slice(0, 10), txt: d.toLocaleDateString("es-CO", { month: "short", year: "2-digit" }) });
    }
    const hoy = R.hoyISO();
    const colorHito = { Cumplido: COLORES.Verde, Atrasado: COLORES.Rojo, Pendiente: "#888" };
    el.innerHTML = `
      <h1>Cronograma</h1>
      <p class="sub">Barra = duración planeada. Rombos = hitos (verde cumplido, rojo atrasado, gris pendiente). Línea dorada = hoy.</p>
      <div class="card gantt">
        <div class="g-fila g-cab"><div class="g-nombre"></div><div class="g-pista">${meses.map((m) => `<span class="g-mes" style="left:${x(m.iso)}%">${esc(m.txt)}</span>`).join("")}</div></div>
        ${ps.map((p) => {
          const hs = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto);
          return `<div class="g-fila clic" data-pid="${esc(p.ID_Proyecto)}"><div class="g-nombre"><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.Estado)} · ${pct(p.Avance_Real)}</div></div>
            <div class="g-pista">
              <div class="g-barra" style="left:${x(p.Fecha_Inicio)}%;width:${Math.max(1, x(p.Fecha_Fin_Plan) - x(p.Fecha_Inicio))}%"><div class="g-real" style="width:${Math.min(100, Number(p.Avance_Real) || 0)}%"></div></div>
              ${hs.map((h) => `<span class="g-hito" title="${esc(h.Hito)} · ${esc(h.Fecha_Plan)} · ${esc(h.Estado)}" style="left:${x(h.Fecha_Plan)}%;background:${colorHito[h.Estado] || "#888"}"></span>`).join("")}
              <span class="g-hoy" style="left:${x(hoy)}%"></span>
            </div></div>`;
        }).join("")}
      </div>`;
    el.querySelectorAll(".clic").forEach((r) => r.addEventListener("click", () => ir("ficha", r.dataset.pid)));
  }

  // ---------- Riesgos ----------
  function vRiesgos(el) {
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const abiertos = S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto) && r.Estado === "Abierto");
    const res = S.riesgoVista === "residual";
    const P = res ? "Probabilidad_Residual" : "Probabilidad_Inherente";
    const I = res ? "Impacto_Residual" : "Impacto_Inherente";
    let celdas = "";
    for (let pr = 5; pr >= 1; pr--) {
      celdas += `<div class="hm-eje">${pr}</div>`;
      for (let im = 1; im <= 5; im++) {
        const n = abiertos.filter((r) => Number(r[P]) === pr && Number(r[I]) === im).length;
        const nv = R.nivelRiesgo(pr * im).toLowerCase();
        celdas += `<div class="hm-celda riesgo-${nv}">${n || ""}</div>`;
      }
    }
    celdas += `<div></div>${[1, 2, 3, 4, 5].map((i) => `<div class="hm-eje">${i}</div>`).join("")}`;
    const C = res ? "Calificacion_Residual" : "Calificacion_Inherente";
    const orden = abiertos.slice().sort((a, b) => (Number(b[C]) || 0) - (Number(a[C]) || 0));
    el.innerHTML = `
      <h1>Riesgos</h1>
      <div class="filtros"><div class="seg">
        <button class="btn ${!res ? "primario" : ""}" data-rv="inherente">Inherente</button>
        <button class="btn ${res ? "primario" : ""}" data-rv="residual">Residual</button></div>
        <span class="sub">${abiertos.length} riesgo(s) abierto(s)</span></div>
      <div class="grid2">
        <div class="card"><h2>Mapa de calor (${res ? "residual" : "inherente"})</h2>
          <div class="heatmap">${celdas}</div><div class="sub centro">Impacto →  ·  ↑ Probabilidad</div></div>
        <div class="card"><h2>Riesgos abiertos más críticos</h2><div class="tabla-scroll"><table>
          <thead><tr><th>Proyecto</th><th>Riesgo</th><th>Calificación</th></tr></thead>
          <tbody>${orden.slice(0, 10).map((r) => `<tr class="clic" data-pid="${esc(r.ID_Proyecto)}"><td>${esc((proyecto(r.ID_Proyecto) || {}).Nombre)}</td><td>${esc(r.Descripcion)}</td><td>${nivel(r[C])}</td></tr>`).join("") || `<tr><td colspan="3">${vacio("Sin riesgos abiertos.")}</td></tr>`}</tbody>
        </table></div></div>
      </div>`;
    el.querySelectorAll("[data-rv]").forEach((b) => b.addEventListener("click", () => { S.riesgoVista = b.dataset.rv; render(); }));
    el.querySelectorAll(".clic").forEach((r) => r.addEventListener("click", () => ir("ficha", r.dataset.pid)));
  }

  // ---------- Catálogos ----------
  // Listas editables: solo alimentan los desplegables. Las demás las usa la lógica de la app y no se editan aquí.
  const CATALOGOS_EDITABLES = [
    { lista: "Cliente_Area", t: "Cliente / área", campo: "Cliente_Area" },
    { lista: "Metodologia", t: "Metodología", campo: "Metodologia" },
    { lista: "Fase", t: "Fase", campo: "Fase" },
    { lista: "Prioridad", t: "Prioridad", campo: "Prioridad" },
  ];
  const CATALOGOS_FIJOS = { Estado: "Estado del proyecto", Semaforo: "Semáforo", Frecuencia_Seguimiento: "Frecuencia de seguimiento",
    Estado_Hito: "Estado del hito", Tipo_Riesgo: "Tipo de riesgo", Estado_Riesgo: "Estado del riesgo", Rol: "Rol" };

  function vCatalogos(el) {
    if (!R.puede(S.usuario, "catalogos")) { el.innerHTML = vacio("Sin acceso."); return; }
    const enUso = (campo, valor) => S.datos.Proyectos.filter((p) => String(p[campo]) === String(valor)).length;
    el.innerHTML = `
      <h1>Catálogos</h1>
      <p class="sub">Valores de los desplegables de los formularios. Los cambios se guardan en la tabla Catalogos del Excel y aplican para todos los usuarios.</p>
      <div class="grid2">
        ${CATALOGOS_EDITABLES.map((c) => {
          const valores = S.cat[c.lista] || [];
          return `<div class="card" data-anchor="cat-${c.lista}">
            <h2>${esc(c.t)}</h2>
            <table><thead><tr><th>Valor</th><th>Proyectos que lo usan</th><th></th></tr></thead>
              <tbody>${valores.map((v) => `<tr><td>${esc(v)}</td><td>${enUso(c.campo, v)}</td>
                <td class="derecha"><button class="btn chico" data-quitar="${esc(c.lista)}" data-valor="${esc(v)}">Quitar</button></td></tr>`).join("") || `<tr><td colspan="3">${vacio("Sin valores.")}</td></tr>`}</tbody></table>
            <form class="cat-agregar" data-lista="${esc(c.lista)}" novalidate>
              <input name="valor" placeholder="Nuevo valor" aria-label="Nuevo valor para ${esc(c.t)}">
              <button class="btn primario" type="submit">Agregar</button>
            </form>
            <div class="error-campo" data-e-cat="${esc(c.lista)}"></div>
          </div>`;
        }).join("")}
      </div>
      <div class="card"><h2>Listas fijas</h2>
        <p class="sub">Estas listas controlan reglas de la app (colores, días de seguimiento, permisos), por eso no se editan aquí. Si necesitas cambiarlas, pídelo como una mejora.</p>
        <div class="tabla-scroll"><table><thead><tr><th>Lista</th><th>Valores</th></tr></thead>
          <tbody>${Object.entries(CATALOGOS_FIJOS).map(([k, t]) => `<tr><td>${esc(t)}</td><td>${esc((S.cat[k] || []).join(", "))}</td></tr>`).join("")}</tbody></table></div>
      </div>`;
    el.querySelectorAll(".cat-agregar").forEach((f) => {
      const lista = f.dataset.lista;
      const err = el.querySelector(`[data-e-cat="${lista}"]`);
      f.valor.addEventListener("input", () => { err.textContent = ""; });
      f.addEventListener("submit", (ev) => {
        ev.preventDefault();
        const valor = f.valor.value.trim();
        if (!valor) { err.textContent = "Escriba el valor que quiere agregar."; return; }
        if ((S.cat[lista] || []).some((v) => lc(v) === lc(valor))) { err.textContent = "Ese valor ya existe en la lista."; return; }
        if (valor.length > 60) { err.textContent = "Máximo 60 caracteres."; return; }
        guardar(() => S.api.agregarFila("Catalogos", { Lista: lista, Valor: valor }), `"${valor}" agregado`);
      });
    });
    el.querySelectorAll("[data-quitar]").forEach((b) => b.addEventListener("click", () => {
      const lista = b.dataset.quitar, valor = b.dataset.valor;
      const c = CATALOGOS_EDITABLES.find((x) => x.lista === lista);
      const err = el.querySelector(`[data-e-cat="${lista}"]`);
      if ((S.cat[lista] || []).length <= 1) { err.textContent = "La lista debe tener al menos un valor."; return; }
      const n = enUso(c.campo, valor);
      confirmar(`Quitar "${valor}"`,
        n ? `${n} proyecto(s) usan este valor. Lo conservarán, pero ya no aparecerá como opción en los formularios.` : `"${valor}" dejará de aparecer en ${c.t}.`,
        "Quitar", () => guardar(() => S.api.eliminarFila("Catalogos", { Lista: lista, Valor: valor }), `"${valor}" quitado`));
    }));
  }

  function confirmar(titulo, mensaje, textoBoton, accion) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja chica" role="dialog" aria-modal="true" aria-label="${esc(titulo)}">
      <h2>${esc(titulo)}</h2><p>${esc(mensaje)}</p>
      <div class="acciones derecha"><button class="btn" id="c-no">Cancelar</button><button class="btn peligro" id="c-si">${esc(textoBoton)}</button></div></div>`;
    document.body.appendChild(m);
    $("#c-no").addEventListener("click", cerrarModal);
    $("#c-si").addEventListener("click", accion);
    $("#c-si").focus();
  }

  // ---------- Usuarios ----------
  function vUsuarios(el) {
    if (!R.puede(S.usuario, "usuarios")) { el.innerHTML = vacio("Sin acceso."); return; }
    el.innerHTML = `
      <div class="titulo-fila"><h1>Usuarios</h1><button class="btn primario" id="b-nuevo-u">+ Nuevo usuario</button></div>
      <p class="sub">Quien no esté aquí, o esté inactivo, no puede entrar. La seguridad del archivo se controla además con los permisos de SharePoint.</p>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Correo</th><th>Nombre</th><th>Rol</th><th>Proyectos</th><th>Activo</th><th></th></tr></thead>
        <tbody>${S.datos.Usuarios.map((u) => `<tr><td>${esc(u.Correo)}</td><td>${esc(u.Nombre)}</td><td>${esc(u.Rol)}</td><td>${esc(u.Proyectos)}</td><td>${esc(u.Activo)}</td>
          <td class="derecha"><button class="btn chico" data-u="${esc(u.Correo)}">Editar</button></td></tr>`).join("")}</tbody></table></div></div>`;
    $("#b-nuevo-u").addEventListener("click", () => formUsuario());
    el.querySelectorAll("[data-u]").forEach((b) => b.addEventListener("click", () => formUsuario(S.datos.Usuarios.find((u) => u.Correo === b.dataset.u))));
  }

  // ---------- Formularios ----------
  function cerrarModal() { const m = $("#modal"); if (m) m.remove(); }
  function modal(titulo, campos, valores, alGuardar, validarExtra) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    const control = (c) => {
      const v = valores[c.k] ?? c.def ?? "";
      const dis = c.bloqueado ? "disabled" : "";
      const opciones = c.tipo === "select" && v !== "" && !(c.opciones || []).some((o) => String(Array.isArray(o) ? o[0] : o) === String(v)) ? [String(v), ...(c.opciones || [])] : c.opciones;
      if (c.tipo === "select") return `<select name="${c.k}" ${dis}><option value=""></option>${(opciones || []).map((o) => { const [val, txt] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}" ${String(val) === String(v) ? "selected" : ""}>${esc(txt)}</option>`; }).join("")}</select>`;
      if (c.tipo === "textarea") return `<textarea name="${c.k}" rows="3" ${dis}>${esc(v)}</textarea>`;
      return `<input name="${c.k}" type="${c.tipo || "text"}" value="${esc(v)}" ${c.min !== undefined ? `min="${c.min}"` : ""} ${c.max !== undefined ? `max="${c.max}"` : ""} ${dis}>`;
    };
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="${esc(titulo)}">
      <div class="titulo-fila"><h2>${esc(titulo)}</h2><button class="btn enlace" id="m-cerrar" aria-label="Cerrar">✕</button></div>
      <form id="m-form" novalidate><div class="form-grid">
        ${campos.map((c) => `<label class="${c.tipo === "textarea" ? "ancho" : ""}">${esc(c.label)}${c.req ? " *" : ""}${control(c)}${c.ayuda ? `<span class="sub">${esc(c.ayuda)}</span>` : ""}<span class="error-campo" data-e="${c.k}"></span></label>`).join("")}
      </div><div class="error-campo" id="m-error"></div>
      <div class="acciones derecha"><button type="button" class="btn" id="m-cancelar">Cancelar</button><button type="submit" class="btn primario">Guardar</button></div></form></div>`;
    document.body.appendChild(m);
    $("#m-cerrar").addEventListener("click", cerrarModal);
    $("#m-cancelar").addEventListener("click", cerrarModal);
    m.querySelectorAll("input,select,textarea").forEach((i) => i.addEventListener("input", () => {
      const e = m.querySelector(`[data-e="${i.name}"]`); if (e) e.textContent = ""; $("#m-error").textContent = "";
    }));
    $("#m-form").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const fd = {};
      campos.forEach((c) => {
        const el = m.querySelector(`[name="${c.k}"]`);
        let v = el.value.trim();
        if (c.tipo === "number" && v !== "") v = Number(v);
        fd[c.k] = v;
      });
      let ok = true;
      const err = (k, msg) => { ok = false; const e = m.querySelector(`[data-e="${k}"]`); if (e) e.textContent = msg; };
      campos.forEach((c) => {
        const v = fd[c.k];
        if (c.req && (v === "" || v === null)) return err(c.k, "Campo obligatorio.");
        if (c.tipo === "number" && v !== "" && (Number.isNaN(v) || (c.min !== undefined && v < c.min) || (c.max !== undefined && v > c.max)))
          err(c.k, `Debe estar entre ${c.min} y ${c.max ?? "∞"}.`);
        if (c.tipo === "email" && v && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) err(c.k, "Correo no válido.");
      });
      if (ok && validarExtra) { const msg = validarExtra(fd); if (msg) { ok = false; $("#m-error").textContent = msg; } }
      if (!ok) return;
      alGuardar(fd);
    });
    const primero = m.querySelector("input:not([disabled]),select:not([disabled]),textarea");
    if (primero) primero.focus();
  }

  function formProyecto(p) {
    const u = S.usuario;
    const nuevo = !p;
    const esPM = u.Rol === "PM";
    const pms = S.datos.Usuarios.filter((x) => x.Activo === "Sí" && (x.Rol === "PM" || x.Rol === "PMO" || x.Rol === "Admin")).map((x) => [x.Correo, `${x.Nombre} (${x.Correo})`]);
    const campos = [
      { k: "Nombre", label: "Nombre de la iniciativa", req: true },
      { k: "Cliente_Area", label: "Cliente / área", tipo: "select", opciones: S.cat.Cliente_Area, req: true },
      { k: "PM", label: "PM responsable", tipo: "select", opciones: pms, req: true, bloqueado: esPM },
      { k: "Metodologia", label: "Metodología", tipo: "select", opciones: S.cat.Metodologia, req: true },
      { k: "Fase", label: "Fase", tipo: "select", opciones: S.cat.Fase, req: true, def: "Inicio" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado, req: true, def: "Activo" },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo, req: true, def: "Verde" },
      { k: "Prioridad", label: "Prioridad", tipo: "select", opciones: S.cat.Prioridad, req: true, def: "Media" },
      { k: "Fecha_Inicio", label: "Fecha de inicio", tipo: "date", req: true, def: R.hoyISO() },
      { k: "Fecha_Fin_Plan", label: "Fecha fin planeada", tipo: "date", req: true },
      { k: "Avance_Planeado", label: "Avance planeado (%)", tipo: "number", min: 0, max: 100, req: true, def: 0 },
      { k: "Frecuencia_Seguimiento", label: "Frecuencia de seguimiento", tipo: "select", opciones: S.cat.Frecuencia_Seguimiento, req: true, def: "Semanal",
        bloqueado: !R.puede(u, "editarFrecuencia"), ayuda: R.puede(u, "editarFrecuencia") ? "" : "La define la PMO." },
      { k: "Presupuesto", label: "Presupuesto (COP)", tipo: "number", min: 0, def: 0 },
      { k: "Ejecutado", label: "Ejecutado (COP)", tipo: "number", min: 0, def: 0 },
      { k: "Comentario_Estado", label: "Comentario de estado", tipo: "textarea" },
    ];
    const valores = p ? { ...p } : { PM: esPM ? u.Correo : "" };
    modal(nuevo ? "Nueva iniciativa" : `Editar ${p.ID_Proyecto}`, campos, valores, (fd) => {
      if (esPM) fd.PM = nuevo ? u.Correo : p.PM;
      if (!R.puede(u, "editarFrecuencia")) fd.Frecuencia_Seguimiento = nuevo ? "Semanal" : p.Frecuencia_Seguimiento;
      if (nuevo) {
        const fila = { ...fd, ID_Proyecto: R.siguienteIdProyecto(S.datos.Proyectos), Avance_Real: 0, ...sello() };
        S.filtros.estado = "";
        guardar(() => S.api.agregarFila("Proyectos", fila), `Iniciativa ${fila.ID_Proyecto} registrada`);
      } else {
        guardar(() => S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...fd, ...sello() }));
      }
    }, (fd) => (fd.Fecha_Fin_Plan < fd.Fecha_Inicio ? "La fecha fin no puede ser anterior a la fecha de inicio." : ""));
  }

  function formSeguimiento(p) {
    const campos = [
      { k: "Fecha_Corte", label: "Fecha de corte", tipo: "date", req: true, def: R.hoyISO() },
      { k: "Avance_Real", label: "Avance real (%)", tipo: "number", min: 0, max: 100, req: true, def: p.Avance_Real },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo, req: true, def: p.Semaforo },
      { k: "Avance_Planeado", label: "Avance planeado a la fecha (%)", tipo: "number", min: 0, max: 100, req: true, def: p.Avance_Planeado },
      { k: "Logros", label: "Logros del periodo", tipo: "textarea", req: true },
      { k: "Proximos_Pasos", label: "Próximos pasos", tipo: "textarea", req: true },
      { k: "Bloqueos", label: "Bloqueos (qué frena y quién debe actuar)", tipo: "textarea" },
    ];
    modal(`Seguimiento · ${p.Nombre}`, campos, {}, (fd) => {
      const fila = {
        ID_Seguimiento: R.siguienteIdHijo("SEG", S.datos.Seguimientos, "ID_Seguimiento", p.ID_Proyecto),
        ID_Proyecto: p.ID_Proyecto, Fecha_Corte: fd.Fecha_Corte, Semana: R.semanaISO(fd.Fecha_Corte),
        Avance_Real: fd.Avance_Real, Semaforo: fd.Semaforo, Logros: fd.Logros, Proximos_Pasos: fd.Proximos_Pasos,
        Bloqueos: fd.Bloqueos, Reportado_Por: S.usuario.Correo,
      };
      guardar(async () => {
        await S.api.agregarFila("Seguimientos", fila);
        await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          Avance_Real: fd.Avance_Real, Avance_Planeado: fd.Avance_Planeado, Semaforo: fd.Semaforo,
          Comentario_Estado: fd.Bloqueos ? `Bloqueo: ${fd.Bloqueos}` : fd.Logros, ...sello(),
        });
      }, "Seguimiento registrado");
    }, (fd) => (fd.Fecha_Corte > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""));
  }

  function formHito(p, h) {
    const campos = [
      { k: "Hito", label: "Hito", req: true },
      { k: "Fecha_Plan", label: "Fecha planeada", tipo: "date", req: true },
      { k: "Fecha_Real", label: "Fecha real", tipo: "date", ayuda: "Solo si ya se cumplió." },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Hito, req: true, def: "Pendiente" },
    ];
    modal(h ? "Editar hito" : `Nuevo hito · ${p.Nombre}`, campos, h || {}, (fd) => {
      if (fd.Fecha_Real) fd.Estado = "Cumplido";
      if (h) guardar(() => S.api.actualizarPorId("Hitos", "ID_Hito", h.ID_Hito, fd));
      else guardar(() => S.api.agregarFila("Hitos", { ID_Hito: R.siguienteIdHijo("HIT", S.datos.Hitos, "ID_Hito", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }), "Hito agregado");
    });
  }

  function formRiesgo(p, r) {
    const esc5 = [1, 2, 3, 4, 5];
    const campos = [
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: S.cat.Tipo_Riesgo, req: true, def: "Amenaza" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Riesgo, req: true, def: "Abierto" },
      { k: "Descripcion", label: "Descripción del riesgo", tipo: "textarea", req: true },
      { k: "Probabilidad_Inherente", label: "Probabilidad inherente (1-5)", tipo: "select", opciones: esc5, req: true },
      { k: "Impacto_Inherente", label: "Impacto inherente (1-5)", tipo: "select", opciones: esc5, req: true },
      { k: "Plan_Mitigacion", label: "Plan de mitigación", tipo: "textarea", req: true },
      { k: "Plan_Contingencia", label: "Plan de contingencia", tipo: "textarea" },
      { k: "Probabilidad_Residual", label: "Probabilidad residual (1-5)", tipo: "select", opciones: esc5, req: true },
      { k: "Impacto_Residual", label: "Impacto residual (1-5)", tipo: "select", opciones: esc5, req: true },
      { k: "Responsable", label: "Responsable", req: true, def: S.usuario.Correo },
    ];
    modal(r ? "Editar riesgo" : `Nuevo riesgo · ${p.Nombre}`, campos, r || {}, (fd) => {
      ["Probabilidad_Inherente", "Impacto_Inherente", "Probabilidad_Residual", "Impacto_Residual"].forEach((k) => { fd[k] = Number(fd[k]); });
      fd.Calificacion_Inherente = R.calificacion(fd.Probabilidad_Inherente, fd.Impacto_Inherente);
      fd.Calificacion_Residual = R.calificacion(fd.Probabilidad_Residual, fd.Impacto_Residual);
      if (r) guardar(() => S.api.actualizarPorId("Riesgos", "ID_Riesgo", r.ID_Riesgo, fd));
      else guardar(() => S.api.agregarFila("Riesgos", { ID_Riesgo: R.siguienteIdHijo("RSG", S.datos.Riesgos, "ID_Riesgo", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }), "Riesgo agregado");
    });
  }

  function formUsuario(uEdit) {
    const campos = [
      { k: "Correo", label: "Correo corporativo", tipo: "email", req: true, bloqueado: !!uEdit },
      { k: "Nombre", label: "Nombre", req: true },
      { k: "Rol", label: "Rol", tipo: "select", opciones: S.cat.Rol || ["Admin", "PMO", "PM", "Lector"], req: true },
      { k: "Proyectos", label: "Proyectos asignados", ayuda: "IDs separados por coma (PRY-0001,PRY-0004) o Todos", def: "Todos" },
      { k: "Activo", label: "Activo", tipo: "select", opciones: ["Sí", "No"], req: true, def: "Sí" },
    ];
    modal(uEdit ? "Editar usuario" : "Nuevo usuario", campos, uEdit || {}, (fd) => {
      if (uEdit) { fd.Correo = uEdit.Correo; guardar(() => S.api.actualizarPorId("Usuarios", "Correo", uEdit.Correo, fd)); }
      else guardar(() => S.api.agregarFila("Usuarios", fd), "Usuario creado");
    }, (fd) => {
      if (!uEdit && S.datos.Usuarios.some((x) => lc(x.Correo) === lc(fd.Correo))) return "Ese correo ya existe.";
      if (uEdit && lc(uEdit.Correo) === lc(S.usuario.Correo) && (fd.Rol !== "Admin" || fd.Activo !== "Sí")) return "No puede quitarse a sí mismo el rol Admin ni desactivarse.";
      return "";
    });
  }

  // ---------- Arranque ----------
  function modoDesdeURL() {
    const m = new URLSearchParams(location.search).get("modo");
    return m === "panel" || m === "ventana" ? m : "demo";
  }
  async function iniciar() {
    S.modo = typeof Office === "undefined" ? "demo" : modoDesdeURL();
    cargando(true, "Conectando con Excel…");
    try {
      S.api = await DATOS.crear(S.modo);
      await recargar();
      const correo = leerSesion();
      S.usuario = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(correo) && x.Activo === "Sí") || null;
      render();
    } catch (e) {
      app().innerHTML = `<div class="login"><div class="login-caja"><h2>No se pudo leer el Excel</h2><p>${esc(e.message)}</p>
        <p class="sub">Verifique que el archivo abierto sea Portafolio.xlsx y que tenga las tablas Proyectos, Hitos, Seguimientos, Riesgos, Usuarios y Catalogos.</p></div></div>`;
    } finally {
      cargando(false);
    }
  }

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(() => iniciar());
  else document.addEventListener("DOMContentLoaded", iniciar);
})();
