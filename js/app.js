/* Pantallas de la app: ingreso, menú, módulos, formularios y ayudas. */
/* global Office, Chart, DATOS, REGLAS, ACTAS */
(function () {
  const R = REGLAS;
  const S = {
    api: null, modo: "demo", datos: null, usuario: null, vista: "dashboard", pid: null,
    cat: {}, charts: [], semana: null, riesgoVista: "inherente",
    filtros: { pm: "", proyecto: "", estado: "", metodologia: "" },   // filtros comunes a todas las vistas
    filtrosLista: { texto: "", semaforo: "" },                        // filtros extra del listado de proyectos
  };
  const COLORES = { azul: "#104994", azulClaro: "#6BBAEF", dorado: "#D6964A", Verde: "#2E8B57", Amarillo: "#E0A800", Rojo: "#C0392B" };
  const VISTAS = [
    { id: "dashboard", t: "Dashboard" },
    { id: "avances", t: "Avances de la semana" },
    { id: "seguimiento", t: "Seguimiento y compromisos" },
    { id: "backlog", t: "Backlog de demanda" },
    { id: "proyectos", t: "Proyectos" },
    { id: "cronograma", t: "Cronograma" },
    { id: "riesgos", t: "Riesgos" },
    { id: "recursos", t: "Recursos" },
    { id: "capacidad", t: "Capacidad del equipo" },
    { id: "priorizacion", t: "Priorización" },
    { id: "lecciones", t: "Lecciones aprendidas" },
    { id: "catalogos", t: "Catálogos", permiso: "catalogos" },
    { id: "usuarios", t: "Usuarios", permiso: "usuarios" },
    { id: "auditoria", t: "Auditoría", permiso: "usuarios" },
  ];

  // Ayuda corta de cada pantalla: para qué sirve y cómo se usa.
  const AYUDAS = {
    dashboard: ["Estado del portafolio de un vistazo: salud, avance, presupuesto y distribución de los proyectos.",
      "Usa los filtros de arriba para ver un PM, un proyecto, un estado o una metodología. Con «💾 Guardar filtro» los dejas listos para la próxima vez.",
      "Haz clic en un indicador, en un color del semáforo o en un proyecto para ir al detalle."],
    avances: ["Resume lo reportado en la semana elegida: avance, logros, próximos pasos, bloqueos y compromisos.",
      "Abajo aparecen los proyectos semanales que aún no reportan.",
      "Haz clic en una fila para abrir el proyecto."],
    seguimiento: ["Controla que cada proyecto reporte según su frecuencia (semanal, quincenal o mensual).",
      "Vencido = ya pasó la fecha sin reporte. Por vencer = faltan 2 días o menos.",
      "Abajo están los compromisos abiertos; márcalos como cumplidos cuando se cierren.",
      "La campana 🔔 de arriba reúne los compromisos vencidos o por vencer y te deja enviar recordatorios por correo."],
    proyectos: ["Lista de proyectos con buscador y filtros.",
      "Usa «+ Nueva iniciativa» para registrar un proyecto.",
      "Haz clic en un proyecto para ver su ficha, registrar seguimiento, hitos, riesgos y compromisos."],
    ficha: ["Todo el proyecto organizado en pestañas: Resumen (datos, stakeholders, proveedores y curva), Compromisos, Seguimientos, Hitos y Riesgos.",
      "«Registrar seguimiento» es la acción periódica del PM: avance, logros, próximos pasos, bloqueos y compromisos.",
      "En la pestaña Compromisos creas, ordenas y filtras los compromisos; cada uno puede estar atado a una sesión o solo al proyecto. Ábrelo para comentar hasta cerrarlo."],
    cronograma: ["Línea de tiempo de los proyectos activos y en pausa.",
      "La barra oscura es el avance real; los rombos son hitos. La línea dorada es hoy."],
    riesgos: ["Todos los riesgos del portafolio: los indicadores de arriba y las casillas del mapa de calor funcionan como filtros.",
      "Cada tarjeta muestra el riesgo antes (inherente) y después de mitigar (residual); el estado se cambia ahí mismo.",
      "Filtra por estado, nivel, responsable u origen, o busca por texto."],
    catalogos: ["Valores de los desplegables de los formularios.",
      "Agrega o quita valores; los proyectos que ya usan un valor lo conservan.",
      "Aquí también configuras los roles de stakeholder (Sponsor, Líder funcional, Product Owner…) y el maestro de proveedores."],
    capacidad: ["Cuánto está dedicada cada persona sumando todos sus proyectos activos o en pausa.",
      "La dedicación del PM se pone en «Editar proyecto»; la de cada stakeholder, en su ficha o en la sección Equipo del formulario.",
      "Más de 100% = sobreasignado. Haz clic en un proyecto para ir a su ficha."],
    backlog: ["Iniciativas que asigna Gestión de la Demanda en Almera, antes de ser proyectos. El código Almera es la llave.",
      "Flujo: Recibida → En análisis → Priorizada → En comité → (Aprobada / Aplazada / Rechazada) → Convertida en proyecto.",
      "Ranking: obligatorias primero y luego por puntaje. En el tablero arrastras las tarjetas entre estados; la decisión la registra el comité."],
    recursos: ["Directorio único de personas: internas y de proveedores, sin duplicados.",
      "Al agregar stakeholders, equipo o responsables de compromisos eliges a la persona de aquí y sus datos se llenan solos.",
      "Si aparecen «Posibles duplicados», usa «Fusionar» para dejar un solo registro."],
    priorizacion: ["Ordena el portafolio por un puntaje de 0 a 100 según valor, urgencia, complejidad y esfuerzo.",
      "La calificación de cada proyecto se hace en «Editar proyecto» → Priorización.",
      "La matriz valor vs. esfuerzo ayuda a ver ganancias rápidas y proyectos a reconsiderar."],
    lecciones: ["Todas las lecciones aprendidas del portafolio en un solo lugar.",
      "Busca por palabra clave o filtra por categoría antes de arrancar un proyecto parecido.",
      "Se registran en la pestaña «Lecciones» de cada proyecto."],
    auditoria: ["Registro de quién cambió qué y cuándo en toda la app.", "Filtra por usuario, tipo de dato o texto."],
    usuarios: ["Quién entra y qué puede hacer. Un usuario puede tener varios roles.",
      "El PM ve los proyectos donde figura como PM. El Lector ve los proyectos que le asignes aquí."],
  };

  const $ = (s, el = document) => el.querySelector(s);
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const cop = (n) => "$" + (Number(n) || 0).toLocaleString("es-CO", { maximumFractionDigits: 0 });
  const pct = (n) => `${Math.round(Number(n) || 0)}%`;
  const fecha = (iso) => { const s = String(iso || ""); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? `${s.slice(8, 10)}/${s.slice(5, 7)}/${s.slice(0, 4)}` : s || "—"; };
  const lc = (s) => String(s || "").trim().toLowerCase();
  const app = () => $("#app");
  const esURL = (u) => /^https?:\/\/\S+$/i.test(String(u || "").trim());
  const enlace = (url, texto) => (esURL(url) ? `<a class="enlace-ext" href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(texto)} ↗</a>` : "");

  function leerLocal(k, def) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch (e) { return def; } }
  function guardarLocal(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } }
  function guardarSesion(correo) { try { sessionStorage.setItem("pmo_usuario", correo || ""); } catch (e) { /* sin almacenamiento */ } }
  function leerSesion() { try { return sessionStorage.getItem("pmo_usuario") || ""; } catch (e) { return ""; } }

  function toast(msg, error) {
    const t = document.createElement("div");
    t.className = "toast" + (error ? " error" : "");
    t.setAttribute("role", "status");
    t.textContent = msg;
    // Se apilan hacia arriba para no taparse entre sí.
    const alto = [...document.querySelectorAll(".toast")].reduce((m, x) => m + x.offsetHeight + 8, 0);
    t.style.bottom = 16 + alto + "px";
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
    S.cat.Estado_Compromiso = R.ESTADOS_COMPROMISO.slice();
    Object.entries(CATALOGOS_BASE).forEach(([k, v]) => { if (!S.cat[k]) S.cat[k] = v.slice(); });
    S.cat.Estado_Riesgo = [...new Set([...ESTADOS_RIESGO, ...(S.cat.Estado_Riesgo || [])])];
    S.cat.Tipo_Riesgo = [...new Set(["Amenaza", "Oportunidad", ...(S.cat.Tipo_Riesgo || [])])];
  }
  // Listas nuevas que se crean solas en la tabla Catalogos la primera vez (luego se editan desde Catálogos).
  const CATALOGOS_BASE = { Rol_Stakeholder: ["Sponsor", "Líder funcional", "Product Owner"], Estado_Ticket: ["Abierto", "En curso", "Resuelto", "Cerrado"],
    Tipo_Cambio: ["Alcance", "Tiempo", "Costo", "Recursos", "Calidad"], Tipo_Demanda: ["Proyecto nuevo", "Mejora / evolutivo", "Regulatorio / obligatorio", "Mantenimiento", "Innovación"], Categoria_Demanda: ["Estratégico", "Operativo", "Cumplimiento"],
    Categoria_Leccion: ["Planeación", "Técnica", "Proveedores", "Comunicación", "Equipo", "Calidad", "Gestión del cambio"] };
  let sembrado = false;
  // La primera vez, deja en el Excel los roles de stakeholder base para poder editarlos desde Catálogos.
  async function sembrarCatalogos() {
    // Compromisos de versiones anteriores: «Cumplido» pasa a «Cerrado».
    const viejos = S.datos.Compromisos.filter((c) => c.Estado === "Cumplido");
    if (viejos.length) {
      try { await S.api.actualizarVarios("Compromisos", "ID_Compromiso", viejos.map((c) => ({ id: c.ID_Compromiso, cambios: { Estado: "Cerrado" } }))); viejos.forEach((c) => { c.Estado = "Cerrado"; }); }
      catch (e) { /* sin permiso de edición: la app igual los trata como cerrados */ }
    }
    await migrarRecursos();
    const faltan = Object.keys(CATALOGOS_BASE).filter((k) => !S.datos.Catalogos.some((f) => f.Lista === k));
    if (sembrado || !faltan.length) return;
    sembrado = true;
    try { await S.api.agregarFilas("Catalogos", faltan.flatMap((k) => CATALOGOS_BASE[k].map((v) => ({ Lista: k, Valor: v })))); S.datos = await S.api.leerTodo(); armarCatalogos(); }
    catch (e) { /* sin permiso de edición: se usan los valores base en memoria */ }
  }
  async function recargar() {
    S.datos = await S.api.leerTodo();
    armarCatalogos();
    await sembrarCatalogos();
    if (S.usuario) S.usuario = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(S.usuario.Correo) && x.Activo === "Sí") || null;
  }
  const visibles = () => R.visibles(S.usuario, S.datos.Proyectos);
  function filtrados() {
    const F = S.filtros;
    return visibles().filter((p) => (!F.pm || p.PM === F.pm) && (!F.proyecto || p.ID_Proyecto === F.proyecto) &&
      (!F.estado || p.Estado === F.estado) && (!F.metodologia || p.Metodologia === F.metodologia));
  }
  const proyecto = (pid) => S.datos.Proyectos.find((p) => p.ID_Proyecto === pid);
  const nombreUsuario = (correo) => (S.datos.Usuarios.find((u) => lc(u.Correo) === lc(correo)) || {}).Nombre || correo || "—";
  const usuariosConRol = (rol) => S.datos.Usuarios.filter((u) => u.Activo === "Sí" && R.tieneRol(u, rol));
  const compromisosDe = (ids) => S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto));

  async function guardar(accion, exito = "Guardado") {
    cargando(true, "Guardando en Excel…");
    try {
      await accion();
      if (S.api.vaciar) await S.api.vaciar();
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

  // ---------- Ingreso ----------
  function pantallaLogin() {
    const demo = S.modo === "demo"
      ? `<div class="nota"><b>Modo demostración</b> (datos de prueba, no se guardan). Prueba con <code>tu.correo@empresa.com</code> (Admin), <code>pmo@empresa.com</code> (PMO y PM), <code>pm1@empresa.com</code> (PM) o <code>direccion@empresa.com</code> (Lector).</div>` : "";
    app().innerHTML = `
      <div class="login">
        <div class="login-caja">
          <div class="marca"><span class="logo">P</span><div><div>Portafolio de Proyectos</div><div class="sub">Oficina de Proyectos · FSFB</div></div></div>
          <p>Registra y consulta el avance de los proyectos del portafolio. Todo se guarda en el Excel del equipo.</p>
          <ol class="pasos-login"><li>Escribe tu correo corporativo.</li><li>Verás los proyectos que te corresponden según tu rol.</li></ol>
          <form id="f-login" novalidate>
            <label>Correo corporativo<input id="correo" type="email" autocomplete="email" placeholder="nombre@empresa.com" required></label>
            <div class="error-campo" id="e-login" role="alert"></div>
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
      const e = $("#e-login");
      if (!correo) { e.textContent = "Escribe tu correo."; return; }
      const u = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(correo));
      if (!u) { e.textContent = "Este correo no está registrado en el portafolio. Pide acceso al administrador."; return; }
      if (u.Activo !== "Sí") { e.textContent = "Tu usuario está inactivo. Pide acceso al administrador."; return; }
      if (!R.roles(u).length) { e.textContent = "Tu usuario no tiene un rol asignado. Pide al administrador que te asigne uno."; return; }
      S.usuario = u;
      guardarSesion(u.Correo);
      filtrosIniciales();
      S.vista = "dashboard";
      render();
    });
  }

  // ---------- Estructura ----------
  function render() {
    S.charts.forEach((c) => c.destroy());
    S.charts = [];
    if (!S.usuario) return pantallaLogin();
    recordarFiltros();
    const menu = VISTAS.filter((v) => !v.permiso || R.puede(S.usuario, v.permiso));
    const activa = S.vista === "ficha" ? "proyectos" : S.vista;
    const etiquetaModo = { demo: "Demostración: los cambios no se guardan", panel: "Conectado a Excel", ventana: "Conectado a Excel" }[S.modo];
    const u = S.usuario;
    const al = alertas(); const nAl = al.vencidos.length + al.pronto.length;
    const oscuro = document.documentElement.dataset.tema === "oscuro";
    S.navMini = leerLocal("pmo_nav_mini", false);
    const itemNav = (v) => `<button class="nav ${v.id === activa ? "activa" : ""}" data-vista="${v.id}" title="${esc(v.t)}">${ico(v.id)}<span>${esc(v.t)}</span></button>`;
    app().innerHTML = `
      <div class="shell ${S.navMini ? "mini" : ""}">
        <header class="appbar">
          <button class="ab-btn" id="b-menu" aria-label="Mostrar u ocultar el menú">${ico("menu")}</button>
          <div class="ab-marca"><span class="logo">P</span><span class="ab-nombre">Portafolio PMO</span><span class="ab-org">FSFB</span></div>
          <button class="ab-buscar" id="b-buscar" title="Búsqueda global (Ctrl+K)">${ico("buscar")}<span>Buscar proyectos, personas, riesgos…</span><kbd>Ctrl K</kbd></button>
          <div class="ab-der">
            <span class="modo ${S.modo}" title="${esc(etiquetaModo)}">${S.modo === "demo" ? "Demo" : "Excel"}</span>
            <button class="ab-btn campana ${al.vencidos.length ? "roja" : nAl ? "ambar" : ""}" id="b-alertas" title="Compromisos vencidos o por vencer">${ico("campana")}${nAl ? `<span class="ab-badge">${nAl}</span>` : ""}</button>
            <button class="ab-btn" id="b-ayuda" title="Mostrar la ayuda de esta pantalla">${ico("ayuda")}</button>
            <div class="ab-usuario"><button class="ab-avatar" id="b-usuario" aria-haspopup="true" aria-expanded="false" title="${esc(u.Nombre || u.Correo)}">${avatar(u.Nombre || u.Correo)}</button>
              <div class="menu-usuario" id="menu-usuario" hidden>
                <div class="mu-cab">${avatar(u.Nombre || u.Correo, "grande")}<div><b>${esc(u.Nombre || u.Correo)}</b><div class="sub">${esc(u.Correo)}</div><div class="sub">${esc(R.roles(u).join(", "))}</div></div></div>
                <button class="mu-item" id="b-recargar">${ico("actualizar")} Actualizar datos del Excel</button>
                <button class="mu-item" id="b-tema">${ico(oscuro ? "sol" : "luna")} ${oscuro ? "Modo claro" : "Modo oscuro"}</button>
                <div class="mu-sep"></div>
                <div class="mu-info sub">${esc(etiquetaModo)} · versión ${esc(window.APP_VERSION || "")}</div>
                <button class="mu-item" id="b-salir">${ico("salir")} Salir</button>
              </div></div>
          </div>
        </header>
        <aside class="nav-lat" aria-label="Módulos">
          ${GRUPOS_NAV.map(([g, ids]) => { const its = menu.filter((v) => ids.includes(v.id)); return its.length ? `<div class="nav-grupo"><div class="nav-titulo">${esc(g)}</div>${its.map(itemNav).join("")}</div>` : ""; }).join("")}
          <button class="nav nav-colapsar" id="b-colapsar" title="${S.navMini ? "Expandir menú" : "Contraer menú"}">${ico(S.navMini ? "expandir" : "colapsar")}<span>Contraer menú</span></button>
        </aside>
        <div class="nav-velo" id="nav-velo"></div>
        <main class="contenido"><div class="barra-migas">${migas()}${EXPORTADORES[S.vista] && S.datos ? `<button class="btn chico btn-exportar" id="b-exportar" title="Descarga en Excel lo que ves, con los filtros aplicados">${ico("exportar")}<span>${S.vista === "ficha" ? (S.fichaTab === "riesgos" ? "Exportar riesgos (formato FSFB)" : "Exportar proyecto a Excel") : "Exportar a Excel"}</span></button>` : ""}</div><section id="vista"></section></main>
      </div>`;
    document.querySelectorAll(".nav[data-vista]").forEach((b) => b.addEventListener("click", () => { document.body.classList.remove("nav-abierta"); ir(b.dataset.vista); }));
    document.querySelectorAll("[data-miga]").forEach((b) => b.addEventListener("click", () => ir(b.dataset.miga)));
    $("#b-menu").addEventListener("click", () => {
      if (window.innerWidth <= 760) document.body.classList.toggle("nav-abierta");
      else { guardarLocal("pmo_nav_mini", !S.navMini); render(); }
    });
    $("#nav-velo").addEventListener("click", () => document.body.classList.remove("nav-abierta"));
    $("#b-colapsar").addEventListener("click", () => { guardarLocal("pmo_nav_mini", !S.navMini); render(); });
    $("#b-buscar").addEventListener("click", abrirBusqueda);
    if ($("#b-exportar")) $("#b-exportar").addEventListener("click", () => EXPORTADORES[S.vista]());
    const mu = $("#menu-usuario"), bu = $("#b-usuario");
    bu.addEventListener("click", (e) => { e.stopPropagation(); mu.hidden = !mu.hidden; bu.setAttribute("aria-expanded", String(!mu.hidden)); });
    $("#b-tema").addEventListener("click", () => { aplicarTema(oscuro ? "claro" : "oscuro"); render(); });
    $("#b-salir").addEventListener("click", () => { S.usuario = null; S.avisoAlertas = false; guardarSesion(""); render(); });
    $("#b-alertas").addEventListener("click", panelAlertas);
    if (!S.avisoAlertas) {
      S.avisoAlertas = true;
      const a = alertas();
      const mios = [...a.vencidos, ...a.pronto].filter((x) => x.mio).length;
      if (a.vencidos.length || a.pronto.length) setTimeout(() => toast(`🔔 ${a.vencidos.length} compromiso(s) vencido(s) y ${a.pronto.length} por vencer${mios ? ` (${mios} a tu cargo)` : ""}. Ábrelos con la campana.`, a.vencidos.length > 0), 300);
    }
    $("#b-ayuda").addEventListener("click", () => {
      guardarLocal("pmo_ayuda_ocultas", leerLocal("pmo_ayuda_ocultas", []).filter((v) => v !== S.vista));
      render();
    });
    $("#b-recargar").addEventListener("click", async () => {
      cargando(true, "Leyendo Excel…");
      try { await recargar(); render(); toast("Datos actualizados"); } catch (e) { toast(e.message, true); } finally { cargando(false); }
    });
    const vistas = { dashboard: vDashboard, avances: vAvances, seguimiento: vSeguimiento, proyectos: vProyectos, ficha: vFicha, backlog: vBacklog, recursos: vRecursos, capacidad: vCapacidad, priorizacion: vPriorizacion, lecciones: vLecciones, auditoria: vAuditoria, cronograma: vCronograma, riesgos: vRiesgos, catalogos: vCatalogos, usuarios: vUsuarios };
    const el = $("#vista");
    (vistas[S.vista] || vDashboard)(el);
    el.insertAdjacentHTML("afterbegin", ayuda(S.vista));
    const cerrarAyuda = $("#ayuda-ok");
    if (cerrarAyuda) cerrarAyuda.addEventListener("click", () => {
      guardarLocal("pmo_ayuda_ocultas", [...new Set([...leerLocal("pmo_ayuda_ocultas", []), S.vista])]);
      $(".ayuda").remove();
    });
    enlazarClics(el);
    enlazarActas(el);
  }
  function ir(vista, pid) { S.vista = vista; if (pid) S.pid = pid; render(); window.scrollTo(0, 0); }
  // Cualquier elemento con data-pid abre la ficha del proyecto.
  function enlazarClics(el) {
    el.querySelectorAll("button[data-pid]:not(.chip-proy)").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); ir("ficha", b.dataset.pid); }));
    el.querySelectorAll("[data-pid]:not(button)").forEach((r) => r.addEventListener("click", (ev) => {
      if (ev.target.closest("button, a, input, select")) return;
      ir("ficha", r.dataset.pid);
    }));
  }

  function ayuda(vista) {
    const lineas = AYUDAS[vista];
    if (!lineas || leerLocal("pmo_ayuda_ocultas", []).includes(vista)) return "";
    return `<div class="ayuda" role="note"><div><b>Cómo usar esta pantalla</b><ul>${lineas.map((l) => `<li>${esc(l)}</li>`).join("")}</ul></div>
      <button class="btn chico" id="ayuda-ok">Entendido</button></div>`;
  }

  // ---------- Filtros guardados ----------
  const FILTRO_VACIO = () => ({ pm: "", proyecto: "", estado: "", metodologia: "" });
  const LISTA_VACIA = () => ({ texto: "", semaforo: "" });
  const esMioFiltro = (f) => lc(f.Usuario) === lc(S.usuario.Correo);
  const filtrosGuardados = () => (S.datos.Filtros || []).filter((f) => esMioFiltro(f) || f.Compartido === "Sí")
    .sort((a, b) => (esMioFiltro(b) - esMioFiltro(a)) || String(a.Nombre).localeCompare(String(b.Nombre), "es"));
  const valoresFiltro = (f) => ({
    filtros: { pm: f.PM || "", proyecto: f.Proyecto || "", estado: f.Estado || "", metodologia: f.Metodologia || "" },
    lista: { texto: f.Texto || "", semaforo: f.Semaforo || "" },
  });
  const iguales = (a, b) => Object.keys(a).every((k) => String(a[k] || "") === String(b[k] || ""));
  const filtroEnUso = () => filtrosGuardados().find((f) => { const v = valoresFiltro(f); return iguales(v.filtros, S.filtros) && iguales(v.lista, S.filtrosLista); });
  const hayFiltros = () => Object.values(S.filtros).some(Boolean) || Object.values(S.filtrosLista).some(Boolean);
  function aplicarFiltro(f) { const v = valoresFiltro(f); S.filtros = v.filtros; S.filtrosLista = v.lista; S.semana = null; }
  function recordarFiltros() { if (S.usuario) guardarLocal("pmo_filtros_" + lc(S.usuario.Correo), { filtros: S.filtros, lista: S.filtrosLista }); }
  // Al entrar: el filtro predeterminado del usuario; si no tiene, los últimos filtros que usó en este equipo.
  function filtrosIniciales() {
    S.filtros = FILTRO_VACIO(); S.filtrosLista = LISTA_VACIA();
    if (!S.usuario) return;
    const pred = (S.datos.Filtros || []).find((f) => esMioFiltro(f) && f.Predeterminado === "Sí");
    if (pred) return aplicarFiltro(pred);
    const ult = leerLocal("pmo_filtros_" + lc(S.usuario.Correo), null);
    if (ult && ult.filtros) { S.filtros = { ...FILTRO_VACIO(), ...ult.filtros }; S.filtrosLista = { ...LISTA_VACIA(), ...(ult.lista || {}) }; }
  }
  function describirFiltro(v) {
    const partes = [];
    if (v.filtros.pm) partes.push(`PM: ${nombreUsuario(v.filtros.pm)}`);
    if (v.filtros.proyecto) partes.push(`Proyecto: ${(proyecto(v.filtros.proyecto) || {}).Nombre || v.filtros.proyecto}`);
    if (v.filtros.estado) partes.push(`Estado: ${v.filtros.estado}`);
    if (v.filtros.metodologia) partes.push(`Metodología: ${v.filtros.metodologia}`);
    if (v.lista.semaforo) partes.push(`Semáforo: ${v.lista.semaforo}`);
    if (v.lista.texto) partes.push(`Búsqueda: «${v.lista.texto}»`);
    return partes.length ? partes.join(" · ") : "Sin filtros (todo el portafolio)";
  }
  // Guardar los filtros actuales (f = null) o editar/eliminar uno guardado.
  function formFiltro(f) {
    const u = S.usuario;
    const nuevo = !f;
    const v = nuevo ? { filtros: { ...S.filtros }, lista: { ...S.filtrosLista } } : valoresFiltro(f);
    const campos = [
      { k: "Nombre", label: "Nombre del filtro", placeholder: "Ej.: Mis proyectos activos ágiles", ancho: true,
        ayuda: nuevo ? "Si usas el nombre de un filtro tuyo que ya existe, se reemplaza." : "" },
      { k: "Predeterminado", label: "¿Aplicarlo al entrar?", tipo: "select", opciones: ["Sí", "No"], def: "No",
        ayuda: "Si dices Sí, la app abre con este filtro puesto." },
    ];
    if (R.puede(u, "verTodo")) campos.push({ k: "Compartido", label: "¿Compartir con todos?", tipo: "select", opciones: ["Sí", "No"], def: "No",
      ayuda: "Los demás usuarios lo verán en su lista (cada uno solo ve sus proyectos)." });
    const propios = (S.datos.Filtros || []).filter(esMioFiltro);
    modal(nuevo ? "Guardar filtro" : "Filtro guardado", campos, nuevo ? {} : f, (fd) => guardar(async () => {
      const nombre = fd.Nombre || "Mi filtro";
      const existente = nuevo ? propios.find((x) => lc(x.Nombre) === lc(nombre)) : f;
      const fila = {
        Nombre: nombre, Predeterminado: fd.Predeterminado || "No", Compartido: fd.Compartido || (existente && existente.Compartido) || "No",
        PM: v.filtros.pm, Proyecto: v.filtros.proyecto, Estado: v.filtros.estado, Metodologia: v.filtros.metodologia, Semaforo: v.lista.semaforo, Texto: v.lista.texto,
      };
      let id;
      if (existente) { id = existente.ID_Filtro; await S.api.actualizarPorId("Filtros", "ID_Filtro", id, fila); }
      else {
        const max = (S.datos.Filtros || []).reduce((m, x) => Math.max(m, parseInt(String(x.ID_Filtro).slice(4), 10) || 0), 0);
        id = `FIL-${String(max + 1).padStart(4, "0")}`;
        await S.api.agregarFila("Filtros", { ID_Filtro: id, Usuario: u.Correo, ...fila });
      }
      // Solo un predeterminado por usuario.
      if (fila.Predeterminado === "Sí") {
        const otros = propios.filter((x) => x.ID_Filtro !== id && x.Predeterminado === "Sí").map((x) => ({ id: x.ID_Filtro, cambios: { Predeterminado: "No" } }));
        if (otros.length) await S.api.actualizarVarios("Filtros", "ID_Filtro", otros);
      }
    }, nuevo ? "Filtro guardado" : "Filtro actualizado"), null, {
      antes: `<div class="resumen-filtro"><b>Criterios:</b> ${esc(describirFiltro(v))}</div>`,
      eliminar: nuevo ? null : { texto: "Eliminar filtro", mensaje: `Se eliminará el filtro «${f.Nombre}». Los datos no se tocan.`,
        accion: () => guardar(() => S.api.eliminarFilas("Filtros", "ID_Filtro", [f.ID_Filtro]), "Filtro eliminado") },
    });
  }

  // Barra de filtros común: PM, proyecto, estado y metodología, más los filtros guardados.
  function barraFiltros() {
    const F = S.filtros;
    const todos = visibles();
    const ordenar = (a, b) => String(a[1]).localeCompare(String(b[1]), "es");
    const pms = [...new Set(todos.map((p) => p.PM))].map((c) => [c, nombreUsuario(c)]).sort(ordenar);
    const proys = todos.map((p) => [p.ID_Proyecto, p.Nombre]).sort(ordenar);
    const sel = (id, etiqueta, opciones, valor, todosTxt) => `<label class="filtro"><span>${etiqueta}</span><select id="${id}">
      <option value="">${todosTxt}</option>${opciones.map(([v, t]) => `<option value="${esc(v)}" ${v === valor ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label>`;
    const activos = hayFiltros();
    const guardados = filtrosGuardados();
    const enUso = filtroEnUso();
    const etiquetaF = (f) => `${f.Predeterminado === "Sí" && esMioFiltro(f) ? "★ " : ""}${f.Nombre}${esMioFiltro(f) ? "" : " (compartido)"}`;
    const editable = enUso && (esMioFiltro(enUso) || R.tieneRol(S.usuario, "Admin"));
    return `<div class="barra-filtros" role="search" aria-label="Filtros">
      ${guardados.length ? `<label class="filtro guardados"><span>Mis filtros guardados</span><select id="fg-guardado">
        <option value="">${enUso ? "— Ninguno —" : activos ? "— Filtro sin guardar —" : "— Elegir —"}</option>
        ${guardados.map((f) => `<option value="${esc(f.ID_Filtro)}" ${enUso && enUso.ID_Filtro === f.ID_Filtro ? "selected" : ""}>${esc(etiquetaF(f))}</option>`).join("")}</select></label>` : ""}
      ${pms.length > 1 ? sel("fg-pm", "PM", pms, F.pm, "Todos") : ""}
      ${sel("fg-proy", "Proyecto", proys, F.proyecto, "Todos")}
      ${sel("fg-est", "Estado", (S.cat.Estado || []).map((v) => [v, v]), F.estado, "Todos")}
      ${sel("fg-met", "Metodología", (S.cat.Metodologia || []).map((v) => [v, v]), F.metodologia, "Todas")}
      <div class="filtro-info"><span>${filtrados().length} de ${todos.length} proyecto(s)</span>
        ${activos && !enUso ? `<button class="btn chico" id="fg-guardar" title="Guarda esta combinación para usarla con un clic">💾 Guardar filtro</button>` : ""}
        ${editable ? `<button class="btn enlace" id="fg-editar">Editar filtro</button>` : ""}
        ${activos ? `<button class="btn enlace" id="fg-limpiar">Limpiar filtros</button>` : ""}</div>
    </div>`;
  }
  function enlazarFiltros() {
    const set = (k) => (e) => { S.filtros[k] = e.target.value; S.semana = null; render(); };
    [["fg-pm", "pm"], ["fg-proy", "proyecto"], ["fg-est", "estado"], ["fg-met", "metodologia"]].forEach(([id, k]) => { const el = $("#" + id); if (el) el.addEventListener("change", set(k)); });
    const l = $("#fg-limpiar");
    if (l) l.addEventListener("click", () => { S.filtros = FILTRO_VACIO(); S.filtrosLista = LISTA_VACIA(); render(); });
    const g = $("#fg-guardado");
    if (g) g.addEventListener("change", (e) => {
      const f = filtrosGuardados().find((x) => x.ID_Filtro === e.target.value);
      if (f) aplicarFiltro(f); else { S.filtros = FILTRO_VACIO(); S.filtrosLista = LISTA_VACIA(); }
      render();
    });
    const b = $("#fg-guardar");
    if (b) b.addEventListener("click", () => formFiltro(null));
    const ed = $("#fg-editar");
    if (ed) ed.addEventListener("click", () => formFiltro(filtroEnUso()));
  }

  const chipSemaforo = (s) => (s ? `<span class="dot" style="background:${COLORES[s] || "#999"}" aria-hidden="true"></span>${esc(s)}` : "—");
  const pill = (e) => `<span class="pill ${String(e).replace(/\s/g, "").toLowerCase()}">${esc(e)}</span>`;
  // Compromiso: su estado (Pendiente / En curso / Cerrado) y, si aplica, la alerta de tiempo.
  const pillComp = (c) => { const e = R.estadoCompromiso(c); return pill(R.estadoBase(c)) + (e === "Vencido" || e === "Por vencer" ? ` ${pill(e)}` : ""); };
  const ORDEN_COMP = { Vencido: 0, "Por vencer": 1, "En curso": 2, Pendiente: 3, Cerrado: 4 };
  const barraAvance = (real, plan) => `
    <span class="avance" title="Real ${pct(real)} · Planeado ${pct(plan)}" aria-label="Avance real ${pct(real)}, planeado ${pct(plan)}">
      <span class="avance-real" style="width:${Math.min(100, Number(real) || 0)}%"></span>
      <span class="avance-plan" style="left:${Math.min(100, Number(plan) || 0)}%"></span>
    </span><span class="avance-txt">${pct(real)} / ${pct(plan)}</span>`;
  const vacio = (msg) => `<div class="vacio">${typeof ico === "function" ? ico("vacio", "vacio-ico") : ""}<div>${esc(msg)}</div></div>`;

  function grafico(canvas, config) {
    if (!canvas) return;
    if (typeof Chart === "undefined") { canvas.replaceWith(Object.assign(document.createElement("div"), { className: "vacio", textContent: "No se pudo cargar la librería de gráficos." })); return; }
    const cs = getComputedStyle(document.documentElement);
    Chart.defaults.font.family = cs.getPropertyValue("--fuente").trim() || "Segoe UI, Arial, sans-serif";
    Chart.defaults.color = cs.getPropertyValue("--texto2").trim() || "#5f6b7a";
    Chart.defaults.borderColor = cs.getPropertyValue("--borde").trim() || "#e3e8ee";
    S.charts.push(new Chart(canvas, config));
  }

  // ---------- Dashboard ----------
  // Barras horizontales en HTML: etiqueta, barra y valor siempre visibles (no depende solo del color).
  function barrasH(filas, { max, formato = (v) => v, color = COLORES.azul } = {}) {
    const tope = max || Math.max(1, ...filas.map((f) => f.valor));
    return `<div class="barras">${filas.map((f) => `<div class="barra-fila${f.pid ? " clic" : ""}"${f.pid ? ` data-pid="${esc(f.pid)}"` : ""} title="${esc(f.etiqueta)}: ${esc(formato(f.valor))}">
      <span class="barra-et">${esc(f.etiqueta)}</span>
      <span class="barra-pista"><span class="barra-val" style="width:${Math.max(f.valor ? 2 : 0, (f.valor / tope) * 100)}%;background:${f.color || color}"></span></span>
      <span class="barra-num">${esc(formato(f.valor))}</span></div>`).join("") || vacio("Sin datos en la selección.")}</div>`;
  }

  // Avance promedio del portafolio semana a semana (último reporte de cada proyecto hasta esa semana).
  function tendenciaAvance(activos, semanas = 12) {
    const hoy = R.hoyISO();
    const puntos = [];
    for (let k = semanas - 1; k >= 0; k--) {
      const corte = R.sumarDias(hoy, -7 * k);
      const valores = activos.map((p) => {
        const segs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos).filter((s) => s.Fecha_Corte && s.Fecha_Corte <= corte && s.Avance_Real !== "");
        return segs.length ? Number(segs[segs.length - 1].Avance_Real) || 0 : null;
      }).filter((v) => v !== null);
      puntos.push({ semana: R.semanaISO(corte), valor: valores.length ? Math.round(valores.reduce((a, b) => a + b, 0) / valores.length) : null, n: valores.length });
    }
    return puntos;
  }

  function vDashboard(el) {
    const ps = filtrados();
    const k = R.kpis(ps, S.datos.Seguimientos);
    const activos = ps.filter((p) => p.Estado === "Activo");
    const n = (v) => Number(v) || 0;
    const desv = Math.round(k.avanceReal - k.avancePlan);

    // Salud: barra 100 % apilada con etiquetas
    const sems = ["Verde", "Amarillo", "Rojo"];
    const conteo = sems.map((s) => ({ s, c: activos.filter((p) => p.Semaforo === s).length }));
    const sinSem = activos.length - conteo.reduce((a, x) => a + x.c, 0);
    const totalSalud = Math.max(1, activos.length);
    const salud = `<div class="salud-barra" role="img" aria-label="${conteo.map((x) => `${x.s} ${x.c}`).join(", ")}">
        ${conteo.filter((x) => x.c).map((x) => `<span class="salud-seg" style="flex:${x.c};background:${COLORES[x.s]}" title="${x.s}: ${x.c} proyecto(s)"></span>`).join("")}
        ${sinSem ? `<span class="salud-seg" style="flex:${sinSem};background:#b8c0cc" title="Sin semáforo: ${sinSem}"></span>` : ""}
      </div>
      <div class="salud-leyenda">${conteo.map((x) => `<button class="salud-item" data-sem="${x.s}"><span class="dot" style="background:${COLORES[x.s]}"></span><b>${x.c}</b> ${x.s.toLowerCase()} <span class="sub">${Math.round((x.c / totalSalud) * 100)}%</span></button>`).join("")}
        ${sinSem ? `<span class="salud-item"><span class="dot" style="background:#b8c0cc"></span><b>${sinSem}</b> sin semáforo</span>` : ""}</div>`;

    // Avance real vs planeado (bala): ordenado de mayor atraso a mayor adelanto
    const avance = activos.map((p) => ({ p, real: n(p.Avance_Real), plan: n(p.Avance_Planeado), d: n(p.Avance_Real) - n(p.Avance_Planeado) }))
      .sort((a, b) => a.d - b.d);
    const balas = avance.map(({ p, real, plan, d }) => `<div class="bala clic" data-pid="${esc(p.ID_Proyecto)}" title="${esc(p.Nombre)} · real ${real}% · planeado ${plan}%">
        <div class="bala-nombre"><b>${esc(p.Nombre)}</b><span class="sub">${esc(nombreUsuario(p.PM))}</span></div>
        <div class="bala-pista"><span class="bala-real" style="width:${Math.min(100, real)}%"></span><span class="bala-plan" style="left:${Math.min(100, plan)}%"></span></div>
        <div class="bala-num"><b>${real}%</b> <span class="sub">plan ${plan}%</span></div>
        <div class="bala-desv ${d < -15 ? "mal" : d < -5 ? "alerta" : "bien"}">${d > 0 ? "+" : ""}${d} pts</div>
        <div class="bala-sem">${chipSemaforo(p.Semaforo)}</div></div>`).join("");

    // Presupuesto: ejecutado vs presupuesto por proyecto
    const pres = activos.filter((p) => n(p.Presupuesto) > 0).map((p) => ({ p, pct: (n(p.Ejecutado) / n(p.Presupuesto)) * 100 })).sort((a, b) => b.pct - a.pct);
    const presHTML = pres.map(({ p, pct }) => `<div class="bala clic presu" data-pid="${esc(p.ID_Proyecto)}" title="${esc(p.Nombre)} · ${cop(p.Ejecutado)} de ${cop(p.Presupuesto)}">
        <div class="bala-nombre"><b>${esc(p.Nombre)}</b><span class="sub">${cop(p.Ejecutado)} de ${cop(p.Presupuesto)}</span></div>
        <div class="bala-pista"><span class="bala-real ${pct > 100 ? "excede" : ""}" style="width:${Math.min(100, pct)}%"></span><span class="bala-plan" style="left:${Math.min(100, n(p.Avance_Real))}%" title="Avance real ${n(p.Avance_Real)}%"></span></div>
        <div class="bala-num"><b class="${pct > 100 ? "texto-rojo" : ""}">${Math.round(pct)}%</b> <span class="sub">ejecutado</span></div></div>`).join("");

    // Distribuciones
    const porFase = (S.cat.Fase || []).map((f) => ({ etiqueta: f, valor: activos.filter((p) => p.Fase === f).length }));
    const porPM = [...new Set(activos.map((p) => p.PM))].map((c) => ({ etiqueta: nombreUsuario(c), valor: activos.filter((p) => p.PM === c).length })).sort((a, b) => b.valor - a.valor);
    const porArea = [...new Set(activos.map((p) => p.Cliente_Area))].map((a) => ({ etiqueta: a || "Sin área", valor: activos.filter((p) => p.Cliente_Area === a).length })).sort((a, b) => b.valor - a.valor);
    const porEstado = (S.cat.Estado || []).map((e) => ({ etiqueta: e, valor: ps.filter((p) => p.Estado === e).length }));

    const tendencia = tendenciaAvance(activos);
    const tarjeta = (titulo, valor, detalle, alerta, destino) => `<button class="kpi ${alerta ? "alerta" : ""}" data-kpi="${destino}" title="Ver detalle">
      <span class="kpi-t">${titulo}</span><span class="kpi-v">${valor}</span><span class="kpi-d">${detalle}</span></button>`;

    el.innerHTML = `
      <h1>Dashboard del portafolio</h1>
      ${barraFiltros()}
      <div class="kpis">
        ${tarjeta("Proyectos activos", k.activos, `de ${k.total} en la selección`, false, "activos")}
        ${tarjeta("Avance real promedio", pct(k.avanceReal), `plan ${pct(k.avancePlan)} · ${desv > 0 ? "+" : ""}${desv} pts`, desv < -5, "activos")}
        ${tarjeta("Proyectos en rojo", k.rojos, `${Math.round((k.rojos / Math.max(1, k.activos)) * 100)}% de los activos`, k.rojos > 0, "rojos")}
        ${tarjeta("Presupuesto ejecutado", pct(k.pctEjecutado), `${cop(k.ejecutado)} de ${cop(k.presupuesto)}`, k.pctEjecutado > 100, "activos")}
        ${tarjeta("Seguimientos al día", pct(k.pctAlDia), "activos sin reporte vencido", k.pctAlDia < 80, "seguimiento")}
      </div>

      <div class="card"><div class="titulo-fila"><h2>Salud del portafolio</h2><span class="sub">${activos.length} proyecto(s) activos por semáforo · clic para ver</span></div>${activos.length ? salud : vacio("No hay proyectos activos en la selección.")}</div>

      <div class="card"><div class="titulo-fila"><h2>Avance real vs planeado</h2>
          <span class="leyenda"><span><i class="lg-real"></i>Avance real</span><span><i class="lg-tick"></i>Planeado</span></span></div>
        <p class="sub">Ordenado del más atrasado al más adelantado. Clic en un proyecto para abrir su ficha.</p>
        ${activos.length ? `<div class="balas">${balas}</div>` : vacio("No hay proyectos activos en la selección.")}</div>

      <div class="grid2">
        <div class="card"><h2>Tendencia del avance promedio</h2><p class="sub">Últimas 12 semanas, según el último reporte de cada proyecto activo.</p>
          ${tendencia.some((t) => t.valor !== null) ? `<div class="grafico-tend"><canvas id="g-tend"></canvas></div>` : vacio("Aún no hay seguimientos para mostrar la tendencia.")}</div>
        <div class="card"><div class="titulo-fila"><h2>Presupuesto ejecutado</h2>
            <span class="leyenda"><span><i class="lg-real"></i>Ejecutado</span><span><i class="lg-tick"></i>Avance real</span></span></div>
          <p class="sub">Si la barra supera la marca, se está gastando más rápido de lo que se avanza.</p>
          ${pres.length ? `<div class="balas compacta">${presHTML}</div>` : vacio("Ningún proyecto activo tiene presupuesto registrado.")}</div>
      </div>

      <div class="grid3">
        <div class="card"><h2>Activos por fase</h2>${barrasH(porFase)}</div>
        <div class="card"><h2>Activos por PM</h2>${barrasH(porPM)}</div>
        <div class="card"><h2>Activos por área</h2>${barrasH(porArea)}</div>
      </div>
      <div class="card"><h2>Portafolio por estado</h2>${barrasH(porEstado, { color: COLORES.azulClaro })}</div>
      <div class="card"><h2>Curva S del presupuesto (proyectos activos)</h2>${(() => { S._curvaD = datosCurvaS(activos); return S._curvaD ? `<div class="grafico-alto"><canvas id="g-curvas"></canvas></div><p class="sub">Planeado acumulado según la línea base de cada proyecto contra el ejecutado reportado en los seguimientos. Si la línea real va por encima, se gasta más rápido de lo planeado.</p>` : vacio("Registra presupuesto y fechas en los proyectos activos para ver la curva S."); })()}</div>`;

    enlazarFiltros();
    el.querySelectorAll("[data-kpi]").forEach((b) => b.addEventListener("click", () => {
      const d = b.dataset.kpi;
      if (d === "seguimiento") return ir("seguimiento");
      S.filtros.estado = "Activo";
      S.filtrosLista.semaforo = d === "rojos" ? "Rojo" : "";
      ir("proyectos");
    }));
    el.querySelectorAll("[data-sem]").forEach((b) => b.addEventListener("click", () => {
      S.filtros.estado = "Activo"; S.filtrosLista.semaforo = b.dataset.sem; ir("proyectos");
    }));
    if ($("#g-curvas") && S._curvaD) graficoCurvaS($("#g-curvas"), S._curvaD);
    const datos = tendencia.map((t) => t.valor);
    grafico($("#g-tend"), { type: "line",
      data: { labels: tendencia.map((t) => t.semana.slice(5)), datasets: [{ label: "Avance promedio", data: datos, borderColor: COLORES.azul, backgroundColor: COLORES.azul, borderWidth: 2, pointRadius: 4, pointHoverRadius: 6, spanGaps: true, tension: 0.25 }] },
      options: { maintainAspectRatio: false, interaction: { mode: "index", intersect: false },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `Avance promedio ${c.parsed.y}% (${tendencia[c.dataIndex].n} proyecto(s))` } } },
        scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + "%" }, grid: { color: "#eef1f5" } }, x: { grid: { display: false } } } } });
  }

  // ---------- Avances de la semana ----------
  function vAvances(el) {
    const ids = new Set(filtrados().map((p) => p.ID_Proyecto));
    const segs = S.datos.Seguimientos.filter((s) => ids.has(s.ID_Proyecto));
    const semanaDe = (s) => s.Semana || R.semanaISO(s.Fecha_Corte);
    const actual = R.semanaISO(R.hoyISO());
    const semanas = [...new Set([actual, ...segs.map(semanaDe)])].sort().reverse();
    if (!S.semana || !semanas.includes(S.semana)) S.semana = semanas.find((w) => w <= actual && segs.some((s) => semanaDe(s) === w)) || actual;
    const deSemana = segs.filter((s) => semanaDe(s) === S.semana);
    const filas = deSemana.map((s) => {
      const hist = R.seguimientosDe(s.ID_Proyecto, segs);
      const i = hist.findIndex((h) => h.ID_Seguimiento === s.ID_Seguimiento);
      const ant = i > 0 ? hist[i - 1] : null;
      const v = ant ? (Number(s.Avance_Real) || 0) - (Number(ant.Avance_Real) || 0) : null;
      const p = proyecto(s.ID_Proyecto) || {};
      const comps = S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento);
      return `<tr class="clic" data-pid="${esc(s.ID_Proyecto)}">
        <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(nombreUsuario(p.PM))} · ${fecha(s.Fecha_Corte)}</div>${s.Acta_Archivo || esURL(s.URL_Acta) ? `<div class="celda-acta">${celdaActa(s, false)}</div>` : ""}</td>
        <td>${pct(s.Avance_Real)}<div class="${v > 0 ? "sube" : v < 0 ? "baja" : "sub"}">${v === null ? "primer reporte" : (v > 0 ? "+" : "") + v + " pts"}</div></td>
        <td>${chipSemaforo(s.Semaforo)}</td><td>${esc(s.Logros)}</td><td class="opc">${esc(s.Proximos_Pasos)}</td>
        <td class="${s.Bloqueos ? "bloqueo" : ""}">${esc(s.Bloqueos || "—")}</td>
        <td class="opc">${comps.length ? `<ul class="mini">${comps.map((c) => `<li>${esc(c.Compromiso)} <span class="sub">(${esc(c.Responsable)}, ${fecha(c.Fecha_Compromiso)})</span></li>`).join("")}</ul>` : "—"}</td></tr>`;
    }).join("");
    const reportaron = new Set(deSemana.map((s) => s.ID_Proyecto));
    const sinReporte = filtrados().filter((p) => p.Estado === "Activo" && p.Frecuencia_Seguimiento === "Semanal" && !reportaron.has(p.ID_Proyecto));
    el.innerHTML = `
      <div class="titulo-fila"><h1>Avances de la semana</h1><button class="btn primario" id="b-informe">📄 Informe semanal (PDF)</button></div>
      ${barraFiltros()}
      <div class="filtros"><label class="filtro"><span>Semana</span><select id="sel-semana">${semanas.map((w) => `<option ${w === S.semana ? "selected" : ""}>${w}</option>`).join("")}</select></label>
        <span class="sub">${deSemana.length} seguimiento(s) reportado(s) · ${sinReporte.length} proyecto(s) semanal(es) sin reporte</span></div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>Avance</th><th>Semáforo</th><th>Logros</th><th class="opc">Próximos pasos</th><th>Bloqueos</th><th class="opc">Compromisos</th></tr></thead>
        <tbody>${filas || `<tr><td colspan="7">${vacio("No hay seguimientos reportados en esta semana con los filtros actuales.")}</td></tr>`}</tbody></table></div></div>
      <div class="card"><h2>Proyectos semanales sin reporte en ${esc(S.semana)}</h2>
        ${sinReporte.length ? `<ul class="lista">${sinReporte.map((p) => `<li class="clic" data-pid="${esc(p.ID_Proyecto)}"><b>${esc(p.Nombre)}</b> · ${esc(nombreUsuario(p.PM))}</li>`).join("")}</ul>` : vacio("Todos los proyectos semanales reportaron.")}</div>`;
    $("#b-informe").addEventListener("click", () => informeSemanal(S.semana));
    enlazarFiltros();
    $("#sel-semana").addEventListener("change", (e) => { S.semana = e.target.value; render(); });
  }

  // ---------- Seguimiento y compromisos ----------
  function vSeguimiento(el) {
    const orden = { Vencido: 0, "Por vencer": 1, "Al día": 2 };
    const ps = filtrados();
    const filas = ps.filter((p) => p.Estado === "Activo").map((p) => ({ p, e: R.estadoSeguimiento(p, S.datos.Seguimientos) }))
      .sort((a, b) => orden[a.e.estado] - orden[b.e.estado] || a.e.faltan - b.e.faltan);
    const cuenta = (e) => filas.filter((f) => f.e.estado === e).length;
    const comps = compromisosDe(new Set(ps.map((p) => p.ID_Proyecto))).map((c) => ({ c, e: R.estadoCompromiso(c) }))
      .filter((x) => x.e !== "Cerrado").sort((a, b) => ORDEN_COMP[a.e] - ORDEN_COMP[b.e] || (a.c.Fecha_Compromiso < b.c.Fecha_Compromiso ? -1 : 1));
    el.innerHTML = `
      <h1>Seguimiento y compromisos</h1>
      ${barraFiltros()}
      <div class="kpis kpis-3">
        <div class="kpi estatico ${cuenta("Vencido") ? "alerta" : ""}"><span class="kpi-t">Seguimientos vencidos</span><span class="kpi-v">${cuenta("Vencido")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Por vencer</span><span class="kpi-v">${cuenta("Por vencer")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Al día</span><span class="kpi-v">${cuenta("Al día")}</span></div>
      </div>
      <div class="card"><h2>Control de seguimiento</h2><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>Frecuencia</th><th>Último</th><th>Próximo</th><th>Estado</th><th class="opc">En 90 días</th><th></th></tr></thead>
        <tbody>${filas.map(({ p, e }) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}">
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(nombreUsuario(p.PM))}</div></td><td>${esc(p.Frecuencia_Seguimiento)}</td>
          <td>${e.ultimaFecha ? fecha(e.ultimaFecha) : "Sin reportes"}</td><td>${fecha(e.proximo)}</td>
          <td>${pill(e.estado)}<div class="sub">${e.faltan < 0 ? `${-e.faltan} día(s) de retraso` : `faltan ${e.faltan} día(s)`}</div></td>
          <td class="opc">${e.en90}</td>
          <td class="derecha">${R.puedeEditar(S.usuario, p, "seguimiento") ? `<button class="btn chico" data-reportar="${esc(p.ID_Proyecto)}">Reportar</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="7">${vacio("No hay proyectos activos con los filtros actuales.")}</td></tr>`}</tbody></table></div>
        <p class="sub">Próximo seguimiento = último reporte + frecuencia (Semanal 7, Quincenal 15, Mensual 30 días).</p></div>
      <div class="card"><h2>Compromisos abiertos <span class="contador">${comps.length}</span></h2>${tablaCompromisos(comps, true)}</div>`;
    enlazarFiltros();
    el.querySelectorAll("[data-reportar]").forEach((b) => b.addEventListener("click", () => formSeguimiento(proyecto(b.dataset.reportar))));
    enlazarCompromisos(el);
  }

  // ---------- Compromisos y su hilo de comentarios ----------
  const comentariosDe = (id) => (S.datos.Comentarios || []).filter((x) => x.ID_Compromiso === id).sort((a, b) => (String(a.Fecha_Hora) < String(b.Fecha_Hora) ? -1 : 1));
  const ahora = () => { const d = new Date(); const p = (n) => String(n).padStart(2, "0"); return `${R.hoyISO()} ${p(d.getHours())}:${p(d.getMinutes())}`; };
  const fechaHora = (v) => { const s = String(v || ""); return s.length >= 16 ? `${fecha(s.slice(0, 10))} ${s.slice(11, 16)}` : fecha(s); };
  function nuevoComentario(c, texto, adjuntos) {
    const lista = S.datos.Comentarios || [];
    const max = lista.filter((x) => x.ID_Compromiso === c.ID_Compromiso).reduce((m, x) => Math.max(m, parseInt(String(x.ID_Comentario).split("-").pop(), 10) || 0), 0);
    return { ID_Comentario: `COM-${String(c.ID_Compromiso).replace(/^CMP-/, "")}-${max + 1}`, ID_Compromiso: c.ID_Compromiso, ID_Proyecto: c.ID_Proyecto, Fecha_Hora: ahora(), Autor: S.usuario.Correo, Texto: texto, Adjuntos: adjuntos && adjuntos.length ? JSON.stringify(adjuntos) : "" };
  }
  // ---------- Adjuntos de los comentarios (imágenes pegadas y archivos) ----------
  const MAX_ADJ = 5 * 1024 * 1024;
  const adjuntosDe = (x) => { try { return x.Adjuntos ? JSON.parse(x.Adjuntos) : []; } catch (e) { return []; } };
  const tamano = (n) => (n > 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);
  const esImagen = (t) => /^image\//.test(t || "");
  // Las imágenes grandes (capturas de correo) se reducen para no llenar el Excel.
  async function prepararArchivo(file) {
    let bytes = new Uint8Array(await file.arrayBuffer()), tipo = file.type || "application/octet-stream", nombre = file.name || "imagen.png";
    if (esImagen(tipo) && !/gif|svg/.test(tipo) && bytes.length > 600 * 1024) {
      try {
        const img = await createImageBitmap(file);
        const k = Math.min(1, 1600 / Math.max(img.width, img.height));
        const cv = document.createElement("canvas"); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
        cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
        const blob = await new Promise((ok) => cv.toBlob(ok, "image/jpeg", 0.85));
        if (blob && blob.size < bytes.length) { bytes = new Uint8Array(await blob.arrayBuffer()); tipo = "image/jpeg"; nombre = nombre.replace(/\.\w+$/, "") + ".jpg"; }
      } catch (e) { /* se guarda tal cual */ }
    }
    return { bytes, tipo, nombre, tam: bytes.length };
  }
  async function borrarAdjuntos(comentarios) {
    for (const x of comentarios) for (const a of adjuntosDe(x)) { try { await S.api.borrarArchivo(a.id); } catch (e) { /* ya no existe */ } }
  }
  function htmlAdjuntos(x) {
    const lista = adjuntosDe(x);
    if (!lista.length) return "";
    return `<div class="adjuntos">${lista.map((a) => esImagen(a.t)
      ? `<button class="adj-img" data-adj="${esc(a.id)}" data-n="${esc(a.n)}" data-t="${esc(a.t)}" title="Ver ${esc(a.n)}"><span class="sub">🖼 Cargando imagen…</span></button>`
      : `<button class="adj-arch" data-adj="${esc(a.id)}" data-n="${esc(a.n)}" data-t="${esc(a.t)}">📄 ${esc(a.n)} <span class="sub">${tamano(a.s || 0)}</span></button>`).join("")}</div>`;
  }
  function descargar(bytes, nombre, tipo) {
    const url = URL.createObjectURL(new Blob([bytes], { type: tipo }));
    const a = document.createElement("a"); a.href = url; a.download = nombre; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  // Carga las miniaturas y enlaza abrir/descargar dentro de un contenedor.
  function enlazarAdjuntos(cont) {
    cont.querySelectorAll(".adj-img").forEach(async (b) => {
      try {
        const f = await leerArchivo(b.dataset.adj);
        b.innerHTML = `<img alt="${esc(b.dataset.n)}" src="${URL.createObjectURL(new Blob([f.bytes], { type: f.tipo }))}">`;
      } catch (e) { b.innerHTML = `<span class="sub">No se pudo cargar ${esc(b.dataset.n)}</span>`; }
      b.addEventListener("click", () => { const img = b.querySelector("img"); if (img) verImagen(img.src, b.dataset.n); });
    });
    cont.querySelectorAll(".adj-arch").forEach((b) => b.addEventListener("click", async () => {
      const txt = b.innerHTML; b.innerHTML = "Abriendo…";
      try { const f = await leerArchivo(b.dataset.adj); descargar(f.bytes, f.nombre || b.dataset.n, f.tipo); }
      catch (e) { toast("No se pudo abrir el archivo: " + e.message, true); }
      b.innerHTML = txt;
    }));
  }
  function verImagen(src, nombre) {
    const v = document.createElement("div");
    v.className = "visor-img"; v.innerHTML = `<div class="visor-img-caja"><div class="titulo-fila"><b>${esc(nombre)}</b><div><a class="btn chico" href="${src}" download="${esc(nombre)}">Descargar</a> <button class="btn chico" id="vi-cerrar">✕</button></div></div><img src="${src}" alt="${esc(nombre)}"></div>`;
    document.body.appendChild(v);
    const cerrar = () => v.remove();
    v.addEventListener("click", (e) => { if (e.target === v) cerrar(); });
    v.querySelector("#vi-cerrar").addEventListener("click", cerrar);
  }

  const resumenCom = (x) => { const n = adjuntosDe(x).length; return [x.Texto, n ? `📎 ${n} adjunto(s)` : ""].filter(Boolean).join(" · "); };
  const origen = (c) => {
    if (!c.ID_Seguimiento) return "Registrado en el proyecto";
    const s = S.datos.Seguimientos.find((x) => x.ID_Seguimiento === c.ID_Seguimiento);
    return s ? `Sesión del ${fecha(s.Fecha_Corte)}` : "Sesión";
  };

  function tablaCompromisos(items, conProyecto) {
    if (!items.length) return vacio("No hay compromisos abiertos.");
    return `<div class="tabla-scroll"><table>
      <thead><tr>${conProyecto ? "<th>Proyecto</th>" : ""}<th>Compromiso</th><th>Responsable</th><th>Fecha</th><th>Estado</th><th class="opc">Último comentario</th><th></th></tr></thead>
      <tbody>${items.map(({ c, e }) => {
        const p = proyecto(c.ID_Proyecto) || {};
        const coms = comentariosDe(c.ID_Compromiso);
        const ult = coms[coms.length - 1];
        return `<tr ${conProyecto ? `class="clic" data-pid="${esc(c.ID_Proyecto)}"` : ""}>${conProyecto ? `<td><b>${esc(p.Nombre)}</b></td>` : ""}
          <td>${esc(c.Compromiso)}<div class="sub">${esc(origen(c))}</div></td><td>${esc(c.Responsable)}</td><td>${fecha(c.Fecha_Compromiso)}</td>
          <td>${pillComp(c)}${e === "Cerrado" && c.Fecha_Cierre ? `<div class="sub">${fecha(c.Fecha_Cierre)}</div>` : ""}</td>
          <td class="opc">${ult ? `<span class="ult-com">${esc(resumenCom(ult))}</span><div class="sub">${esc(nombreUsuario(ult.Autor))} · ${fechaHora(ult.Fecha_Hora)}</div>` : `<span class="sub">Sin comentarios</span>`}</td>
          <td class="derecha nowrap"><button class="btn chico ${coms.length ? "" : "primario-suave"}" data-abrir-comp="${esc(c.ID_Compromiso)}">${coms.length ? `Comentarios (${coms.length})` : "Abrir"}</button>
            ${R.puedeEditar(S.usuario, p, "compromisos") ? botonesEstado(c) : ""}</td></tr>`;
      }).join("")}</tbody></table></div>`;
  }
  // ---------- Alertas de compromisos ----------
  // A quién escribirle: el correo del compromiso, o el de un usuario/stakeholder con ese nombre, o el texto si ya es un correo.
  function correoDe(c) {
    if (c.Correo_Responsable) return c.Correo_Responsable;
    const r = lc(c.Responsable);
    if (!r) return "";
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r)) return c.Responsable.trim();
    const u = S.datos.Usuarios.find((x) => lc(x.Nombre) === r || lc(x.Correo) === r);
    if (u) return u.Correo;
    const st = (S.datos.Stakeholders || []).find((x) => lc(x.Nombre) === r && x.Correo);
    return st ? st.Correo : "";
  }
  const esMioComp = (c) => { const u = S.usuario; return lc(correoDe(c)) === lc(u.Correo) || (u.Nombre && lc(c.Responsable) === lc(u.Nombre)); };
  function alertas() {
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const lista = S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto) && !R.cerrado(c))
      .map((c) => ({ c, e: R.estadoCompromiso(c), mio: esMioComp(c) }))
      .sort((a, b) => (b.mio - a.mio) || String(a.c.Fecha_Compromiso || "9999").localeCompare(String(b.c.Fecha_Compromiso || "9999")));
    return { vencidos: lista.filter((x) => x.e === "Vencido"), pronto: lista.filter((x) => x.e === "Por vencer") };
  }
  // ---------- Correo de recordatorio estructurado ----------
  const validoCorreo = (c) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(c || "").trim());
  const pila = (nombre) => { const t = String(nombre || "").replace(/@.*/, "").trim().split(/\s+/)[0] || ""; return t ? t.charAt(0).toUpperCase() + t.slice(1).toLowerCase() : ""; };
  function venceTexto(c) {
    if (!c.Fecha_Compromiso) return "sin fecha definida";
    const d = R.difDias(c.Fecha_Compromiso, R.hoyISO());
    return d === 0 ? "vence hoy" : d === 1 ? "vence mañana" : d > 1 ? `vence en ${d} días` : `venció hace ${-d} día${d === -1 ? "" : "s"}`;
  }
  // Último avance real (sin las constancias automáticas de la app).
  function ultimoAvance(c) {
    const auto = /^(Recordatorio enviado|Compromiso (cerrado|en curso|reabierto)|Marcado como cumplido|Cerrado en la sesión)/i;
    const l = comentariosDe(c.ID_Compromiso).filter((x) => x.Texto && !auto.test(x.Texto));
    const u = l[l.length - 1];
    return u ? `«${u.Texto}» (${fechaHora(u.Fecha_Hora)})` : "";
  }
  const sesionDe = (c) => (c.ID_Seguimiento ? S.datos.Seguimientos.find((s) => s.ID_Seguimiento === c.ID_Seguimiento) : null);
  function correoRecordatorio(items, ccElegidos) {
    const cs = items.map((x) => x.c);
    const para = [...new Set(cs.map(correoDe).filter(validoCorreo))];
    const nombres = [...new Set(cs.map((c) => (recursoPor({ correo: correoDe(c), nombre: c.Responsable }) || {}).Nombre || c.Responsable).filter(Boolean))];
    const saludo = nombres.length === 1 ? `Hola, ${pila(nombres[0])}:` : "Hola:";
    const proys = [...new Set(cs.map((c) => c.ID_Proyecto))].map(proyecto).filter(Boolean);
    const unoP = proys.length === 1 ? proys[0] : null;
    const vencidos = cs.filter((c) => R.estadoCompromiso(c) === "Vencido");
    const u = S.usuario, yo = recursoPor({ correo: u.Correo, nombre: u.Nombre }) || {};
    const firma = ["Cordialmente,", yo.Nombre || u.Nombre || u.Correo,
      [yo.Cargo, unoP && lc(unoP.PM) === lc(u.Correo) ? `PM · Proyecto ${unoP.Nombre}` : unoP ? `Proyecto ${unoP.Nombre}` : "", "Oficina de Proyectos FSFB"].filter(Boolean).join(" · ")].join("\n");
    const detalle = (c, conProy) => [
      `Compromiso: ${c.Compromiso}`,
      conProy ? `Proyecto: ${(proyecto(c.ID_Proyecto) || {}).Nombre || c.ID_Proyecto}` : "",
      `Responsable: ${c.Responsable || "—"}`,
      `Fecha límite: ${c.Fecha_Compromiso ? fecha(c.Fecha_Compromiso) : "sin fecha"} (${venceTexto(c)})`,
      `Estado actual: ${R.estadoBase(c)}`,
      ultimoAvance(c) ? `Último avance registrado: ${ultimoAvance(c)}` : "",
    ].filter(Boolean).join("\n");
    let asunto, intro, cuerpoItems, cierre;
    if (cs.length === 1) {
      const c = cs[0], s = sesionDe(c), p = proyecto(c.ID_Proyecto) || {};
      asunto = `${p.Nombre || "Proyecto"} · Recordatorio de compromiso · ${R.estadoCompromiso(c) === "Vencido" ? `vencido desde el ${fecha(c.Fecha_Compromiso)}` : c.Fecha_Compromiso ? `vence el ${fecha(c.Fecha_Compromiso)}` : "pendiente"}`;
      intro = s ? `De acuerdo con la sesión de seguimiento del proyecto ${p.Nombre} realizada el ${fecha(s.Fecha_Corte)}${s.Fecha_Acta ? ` (acta del ${fecha(s.Fecha_Acta)})` : ""}, quedó a tu cargo el siguiente compromiso:`
        : `En el marco del seguimiento del proyecto ${p.Nombre}, quedó registrado a tu cargo el siguiente compromiso:`;
      cuerpoItems = detalle(c, false) + (s && esURL(s.URL_Acta) ? `\n\nActa de la sesión: ${s.URL_Acta}` : "");
      cierre = R.estadoCompromiso(c) === "Vencido"
        ? `A la fecha no tenemos registro de su cierre. Te agradecemos indicarnos el estado actual y una nueva fecha de entrega, para reflejarlo en el seguimiento del proyecto. Si hay algún impedimento, avísanos para revisarlo juntos.`
        : `Agradecemos tu gestión y el avance en este compromiso. Te pedimos confirmarnos cómo va y, si ya está listo, compartirnos la evidencia para cerrarlo en el próximo seguimiento. Si necesitas apoyo o ves algún riesgo para cumplir la fecha, avísanos para revisarlo.`;
    } else {
      asunto = `${unoP ? `${unoP.Nombre} · ` : ""}Recordatorio de ${cs.length} compromisos pendientes${vencidos.length ? ` (${vencidos.length} vencido${vencidos.length > 1 ? "s" : ""})` : ""}`;
      intro = `De acuerdo con los seguimientos realizados${unoP ? ` al proyecto ${unoP.Nombre}` : ""}, tienes los siguientes compromisos pendientes:`;
      const grupos = {};
      cs.forEach((c) => { const s = sesionDe(c); const k = s ? `${s.Fecha_Corte}|${s.ID_Seguimiento}` : `9999|${c.ID_Proyecto}`; (grupos[k] = grupos[k] || { s, p: proyecto(c.ID_Proyecto), l: [] }).l.push(c); });
      let n = 0;
      cuerpoItems = Object.keys(grupos).sort().map((k) => { const g = grupos[k];
        const cab = g.s ? `Sesión de seguimiento del ${fecha(g.s.Fecha_Corte)}${g.s.Fecha_Acta ? ` (acta del ${fecha(g.s.Fecha_Acta)})` : ""}${unoP ? "" : ` · ${(g.p || {}).Nombre || ""}`}:` : `Registrados en el proyecto${unoP ? "" : ` ${(g.p || {}).Nombre || ""}`}:`;
        return `${cab}\n` + g.l.map((c) => `${++n}. ${detalle(c, false).replace(/\n/g, "\n   ")}`).join("\n\n"); }).join("\n\n");
      cierre = `Agradecemos tu gestión y el avance en estos compromisos. Te pedimos confirmarnos el estado de cada uno${vencidos.length ? " y, para los vencidos, una nueva fecha de entrega" : ""}. Si necesitas apoyo, avísanos para revisarlo.`;
    }
    const cuerpo = `${saludo}\n\n${intro}\n\n${cuerpoItems}\n\n${cierre}\n\n${firma}`;
    const cc = [...new Set((ccElegidos || []).map((x) => String(x).trim()).filter((x) => validoCorreo(x) && !para.some((d) => lc(d) === lc(x))))];
    const href = `mailto:${para.join(";")}?${cc.length ? `cc=${encodeURIComponent(cc.join(";"))}&` : ""}subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
    return { para, cc, asunto, cuerpo, href };
  }
  // Candidatos a copia: lista fija (Catálogos), PM y stakeholders de los proyectos involucrados.
  function candidatosCopia(items) {
    const u = S.usuario, out = [];
    const agregar = (correo, nombre, motivo) => { if (!validoCorreo(correo) || out.some((x) => lc(x.correo) === lc(correo))) return; out.push({ correo: String(correo).trim(), nombre, motivo }); };
    (S.cat.Copia_Recordatorios || []).forEach((c) => agregar(c, (recursoPor({ correo: c }) || {}).Nombre || c, "Copia fija"));
    [...new Set(items.map((x) => x.c.ID_Proyecto))].forEach((pid) => {
      const p = proyecto(pid); if (!p) return;
      if (lc(p.PM) !== lc(u.Correo)) agregar(p.PM, nombreUsuario(p.PM), `PM · ${p.Nombre}`);
      stakeholdersDe(pid).forEach((x) => agregar(x.Correo, x.Nombre, `${x.Rol || "Stakeholder"} · ${p.Nombre}`));
    });
    return out;
  }
  // Muestra destinatarios (con copia editable) y la vista previa; abre el correo y deja constancia.
  function enviarRecordatorio(items) {
    return new Promise((resolver) => {
      const para = [...new Set(items.map((x) => correoDe(x.c)).filter(validoCorreo))];
      const cand = candidatosCopia(items).filter((x) => !para.some((d) => lc(d) === lc(x.correo)));
      const o = document.createElement("div");
      o.id = "recordatorio";
      o.innerHTML = `<div class="modal-caja rec-caja" role="dialog" aria-modal="true" aria-label="Recordatorio por correo">
        <div class="titulo-fila"><h2>✉ Recordatorio por correo</h2><button class="btn enlace" id="rc-x" aria-label="Cerrar">✕</button></div>
        <div class="rec-dest"><b>Para:</b> ${para.length ? para.map((c) => `<span class="chip-mail">${esc(c)}</span>`).join(" ") : `<span class="baja">El responsable no tiene correo registrado: escríbelo en Outlook.</span>`}</div>
        <div class="rec-dest"><b>Con copia a:</b> <span class="sub">desmarca a quien no aplique</span>
          <div class="rec-cc">${cand.map((x, i) => `<label class="check"><input type="checkbox" class="rc-cc" data-i="${i}" checked> ${esc(x.nombre)} <span class="sub">${esc(x.correo)} · ${esc(x.motivo)}</span></label>`).join("") || `<span class="sub">No hay stakeholders con correo ni copia fija. La copia fija se configura en Catálogos.</span>`}</div>
          <input id="rc-otros" placeholder="Otros correos en copia, separados por coma" aria-label="Otros correos en copia"></div>
        <details class="rec-prev" open><summary>Vista previa del correo</summary><div class="rec-asunto" id="rc-asunto"></div><pre id="rc-cuerpo"></pre></details>
        <div class="error-campo" id="rc-err"></div>
        <div class="acciones derecha"><button class="btn" id="rc-copiar" title="Por si Outlook no abre el correo">Copiar texto</button><button class="btn" id="rc-cancelar">Cancelar</button><button class="btn primario" id="rc-abrir">Abrir en Outlook</button></div>
      </div>`;
      document.body.appendChild(o);
      const elegidos = () => [...o.querySelectorAll(".rc-cc:checked")].map((c) => cand[Number(c.dataset.i)].correo)
        .concat(String(o.querySelector("#rc-otros").value || "").split(/[,;\s]+/).filter(Boolean));
      const previa = () => { const m = correoRecordatorio(items, elegidos()); o.querySelector("#rc-asunto").textContent = `Asunto: ${m.asunto}${m.cc.length ? `  ·  CC: ${m.cc.length}` : ""}`; o.querySelector("#rc-cuerpo").textContent = m.cuerpo; return m; };
      previa();
      o.querySelectorAll(".rc-cc").forEach((c) => c.addEventListener("change", previa));
      o.querySelector("#rc-otros").addEventListener("input", previa);
      const cerrar = (r) => { o.remove(); resolver(r); };
      o.querySelector("#rc-x").addEventListener("click", () => cerrar(null));
      o.querySelector("#rc-cancelar").addEventListener("click", () => cerrar(null));
      o.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrar(null); });
      o.querySelector("#rc-copiar").addEventListener("click", async () => {
        const m = previa();
        try { await navigator.clipboard.writeText(`Para: ${m.para.join("; ")}\nCC: ${m.cc.join("; ")}\nAsunto: ${m.asunto}\n\n${m.cuerpo}`); toast("Texto copiado: pégalo en un correo nuevo"); }
        catch (e) { o.querySelector("#rc-err").textContent = "No se pudo copiar automáticamente; selecciona el texto de la vista previa."; }
      });
      o.querySelector("#rc-abrir").addEventListener("click", () => {
        const otros = String(o.querySelector("#rc-otros").value || "").split(/[,;\s]+/).filter(Boolean);
        const malos = otros.filter((x) => !validoCorreo(x));
        if (malos.length) { o.querySelector("#rc-err").textContent = `Correo no válido: ${malos.join(", ")}`; return; }
        const m = previa();
        const a = document.createElement("a"); a.href = m.href; a.target = "_blank"; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove();
        // Solo se deja constancia si la persona confirma que lo envió.
        o.querySelector(".rec-caja").innerHTML = `<div class="titulo-fila"><h2>✉ ¿Enviaste el correo?</h2></div>
          <p>Se abrió el correo en Outlook. Cuando lo envíes, confírmalo aquí para dejar constancia en el seguimiento del compromiso.</p>
          <p class="sub">Para: ${esc(m.para.join(", ") || "—")}${m.cc.length ? `<br>CC: ${esc(m.cc.join(", "))}` : ""}</p>
          <div class="acciones derecha"><button class="btn" id="rc-no">No lo envié</button><button class="btn" id="rc-otra">Volver a abrir el correo</button><button class="btn primario" id="rc-si">Sí, lo envié</button></div>`;
        o.querySelector("#rc-no").addEventListener("click", () => cerrar(null));
        o.querySelector("#rc-otra").addEventListener("click", () => { const b = document.createElement("a"); b.href = m.href; b.target = "_blank"; b.rel = "noopener"; document.body.appendChild(b); b.click(); b.remove(); });
        o.querySelector("#rc-si").addEventListener("click", async () => {
          o.remove();
          try {
            await S.api.agregarFilas("Comentarios", items.map((x) => nuevoComentario(x.c, `Recordatorio enviado por correo${m.para.length ? ` a ${m.para.join(", ")}` : ""}${m.cc.length ? ` con copia a ${m.cc.join(", ")}` : ""}.`)));
            await recargar();
          } catch (e) { toast("No se pudo registrar la constancia: " + e.message, true); }
          resolver(m);
        });
        o.querySelector("#rc-si").focus();
      });
      o.querySelector("#rc-abrir").focus();
    });
  }
  function panelAlertas() {
    const a = alertas();
    const puedeComp = (c) => R.puedeEditar(S.usuario, proyecto(c.ID_Proyecto) || {}, "compromisos");
    const fila = (x) => { const p = proyecto(x.c.ID_Proyecto) || {}; const correo = correoDe(x.c); return `<li class="alerta-item ${x.e === "Vencido" ? "vencido" : "pronto"}">
      <div><b>${esc(x.c.Compromiso)}</b>${x.mio ? ` <span class="pill tuyo">A tu cargo</span>` : ""}
        <div class="sub">${esc(p.Nombre)} · ${esc(x.c.Responsable || "Sin responsable")} · ${x.c.Fecha_Compromiso ? `${x.e === "Vencido" ? "venció" : "vence"} el ${fecha(x.c.Fecha_Compromiso)}` : "sin fecha"}</div></div>
      <div class="nowrap"><button class="btn chico" data-al-abrir="${esc(x.c.ID_Compromiso)}">Abrir</button>
        ${puedeComp(x.c) ? `<button class="btn chico" data-al-mail="${esc(x.c.ID_Compromiso)}" ${correo ? `title="Se abrirá un correo a ${esc(correo)}"` : `title="No hay correo del responsable: se abrirá el correo sin destinatario"`}>✉ Recordar</button>` : ""}</div></li>`; };
    // Recordatorio agrupado por responsable
    const porResp = {};
    [...a.vencidos, ...a.pronto].filter((x) => puedeComp(x.c)).forEach((x) => { const k = correoDe(x.c) || x.c.Responsable || "(sin responsable)"; (porResp[k] = porResp[k] || []).push(x); });
    const grupos = Object.entries(porResp).filter(([, v]) => v.length > 1);
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="Alertas de compromisos">
      <div class="titulo-fila"><h2>🔔 Alertas de compromisos</h2><button class="btn enlace" id="al-cerrar" aria-label="Cerrar">✕</button></div>
      <p class="sub">Compromisos abiertos de tus proyectos que ya vencieron o vencen pronto (cada compromiso define con cuántos días avisar; por defecto ${R.DIAS_COMPROMISO_POR_VENCER}).</p>
      ${grupos.length ? `<div class="al-grupos"><span class="sub">Un solo correo por responsable:</span> ${grupos.map(([k, v], i) => `<button class="btn chico" data-al-grupo="${i}">✉ ${esc(k.includes("@") ? nombreUsuario(k) : k)} (${v.length})</button>`).join(" ")}</div>` : ""}
      <div class="form-seccion">Vencidos <span class="contador rojo">${a.vencidos.length}</span></div>
      ${a.vencidos.length ? `<ul class="alertas">${a.vencidos.map(fila).join("")}</ul>` : vacio("Nada vencido. 👏")}
      <div class="form-seccion">Por vencer <span class="contador">${a.pronto.length}</span></div>
      ${a.pronto.length ? `<ul class="alertas">${a.pronto.map(fila).join("")}</ul>` : vacio("Nada por vencer en los próximos días.")}
      <p class="sub">«✉ Recordar» abre un correo listo en tu Outlook para el responsable y deja constancia en los comentarios del compromiso.</p>
    </div>`;
    document.body.appendChild(m);
    $("#al-cerrar").addEventListener("click", () => { cerrarModal(); render(); });
    m.addEventListener("keydown", (ev) => { if (ev.key === "Escape") { cerrarModal(); render(); } });
    const buscar = (id) => [...a.vencidos, ...a.pronto].find((x) => x.c.ID_Compromiso === id);
    m.querySelectorAll("[data-al-abrir]").forEach((b) => b.addEventListener("click", () => detalleCompromiso(b.dataset.alAbrir)));
    m.querySelectorAll("[data-al-mail]").forEach((b) => b.addEventListener("click", async () => {
      const r = await enviarRecordatorio([buscar(b.dataset.alMail)]);
      if (!r) return;
      b.textContent = "✓ Correo abierto"; b.disabled = true;
      toast(r.para.length ? `Correo listo para ${r.para.join(", ")}` : "Correo abierto: escribe el destinatario (el compromiso no tiene correo del responsable).");
    }));
    m.querySelectorAll("[data-al-grupo]").forEach((b) => b.addEventListener("click", async () => {
      const r = await enviarRecordatorio(grupos[Number(b.dataset.alGrupo)][1]);
      if (!r) return;
      b.textContent = "✓ Correo abierto"; b.disabled = true;
      toast(`Correo listo para ${r.para.join(", ") || "el responsable"}`);
    }));
  }

  // Sección «Compromisos» de la ficha: filtrar, ordenar y agrupar por sesión.
  function seccionCompromisos(p, comps, segs, puede) {
    const V = S.compVista = S.compVista || { ver: "abiertos", orden: "fecha", agrupar: false };
    const FILTROS = {
      abiertos: (x) => x.e !== "Cerrado", pendientes: (x) => R.estadoBase(x.c) === "Pendiente", encurso: (x) => R.estadoBase(x.c) === "En curso",
      vencidos: (x) => x.e === "Vencido", cerrados: (x) => x.e === "Cerrado", todos: () => true };
    if (!FILTROS[V.ver]) V.ver = "abiertos";
    const cuenta = Object.fromEntries(Object.entries(FILTROS).map(([k, f]) => [k, comps.filter(f).length]));
    let lista = comps.filter(FILTROS[V.ver]);
    const ordenE = ORDEN_COMP;
    const fSes = (c) => { const s = segs.find((x) => x.ID_Seguimiento === c.ID_Seguimiento); return s ? s.Fecha_Corte : ""; };
    const porFecha = (a, b) => String(a.c.Fecha_Compromiso || "9999").localeCompare(String(b.c.Fecha_Compromiso || "9999"));
    const ORDEN = {
      fecha: porFecha,
      estado: (a, b) => ordenE[a.e] - ordenE[b.e] || porFecha(a, b),
      responsable: (a, b) => String(a.c.Responsable || "~").localeCompare(String(b.c.Responsable || "~"), "es") || porFecha(a, b),
      recientes: (a, b) => String(b.c.ID_Compromiso).localeCompare(String(a.c.ID_Compromiso), "es", { numeric: true }),
    };
    lista = lista.slice().sort(ORDEN[V.orden] || porFecha);
    const seg = (id, t) => `<button class="btn chico ${V.ver === id ? "primario" : ""}" data-cver="${id}">${t} (${cuenta[id]})</button>`;
    const fila = ({ c, e }) => {
      const coms = comentariosDe(c.ID_Compromiso);
      const ult = coms[coms.length - 1];
      return `<tr><td>${esc(c.Compromiso)}${ult ? `<div class="sub ult-com">💬 ${esc(resumenCom(ult))}</div>` : ""}</td>
        <td>${c.ID_Seguimiento ? `Sesión ${fecha(fSes(c))}` : `<span class="sub">Sin sesión</span>`}</td>
        <td>${esc(c.Responsable || "—")}</td><td class="nowrap">${fecha(c.Fecha_Compromiso)}</td>
        <td>${pillComp(c)}${e === "Cerrado" && c.Fecha_Cierre ? `<div class="sub">${fecha(c.Fecha_Cierre)}</div>` : ""}</td>
        <td class="derecha nowrap"><button class="btn chico ${coms.length ? "" : "primario-suave"}" data-abrir-comp="${esc(c.ID_Compromiso)}">${coms.length ? `Comentarios (${coms.length})` : "Abrir"}</button>
          ${puede ? botonesEstado(c) : ""}</td></tr>`;
    };
    let filas;
    if (V.agrupar) {
      const grupos = [...new Set(lista.map((x) => x.c.ID_Seguimiento || ""))]
        .sort((a, b) => (a === "" ? 1 : b === "" ? -1 : String(fSes({ ID_Seguimiento: b })).localeCompare(String(fSes({ ID_Seguimiento: a })))));
      filas = grupos.map((g) => {
        const del = lista.filter((x) => (x.c.ID_Seguimiento || "") === g);
        return `<tr class="grupo"><td colspan="6">${g ? `Sesión del ${fecha(fSes({ ID_Seguimiento: g }))}` : "Registrados en el proyecto (sin sesión)"} <span class="contador">${del.length}</span></td></tr>${del.map(fila).join("")}`;
      }).join("");
    } else filas = lista.map(fila).join("");
    return `<div class="card"><div class="titulo-fila"><h2>Compromisos del proyecto</h2>${puede ? `<button class="btn primario" id="b-comp">+ Compromiso</button>` : ""}</div>
      <p class="sub">Un compromiso puede quedar atado a una sesión de seguimiento o solo al proyecto. Ábrelo para comentar su avance hasta cerrarlo.</p>
      <div class="barra-comp">
        <div class="seg" role="group" aria-label="Qué compromisos ver">${seg("abiertos", "Abiertos")}${seg("pendientes", "Pendientes")}${seg("encurso", "En curso")}${seg("vencidos", "Vencidos")}${seg("cerrados", "Cerrados")}${seg("todos", "Todos")}</div>
        <label class="filtro"><span>Ordenar por</span><select id="c-orden">
          ${[["fecha", "Fecha límite"], ["estado", "Estado (vencidos primero)"], ["responsable", "Responsable"], ["recientes", "Más recientes"]].map(([v, t]) => `<option value="${v}" ${V.orden === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="check"><input type="checkbox" id="c-agrupar" ${V.agrupar ? "checked" : ""}> Agrupar por sesión</label>
      </div>
      ${lista.length ? `<div class="tabla-scroll"><table class="tabla-comp"><thead><tr><th>Compromiso</th><th>Sesión</th><th>Responsable</th><th>Fecha límite</th><th>Estado</th><th></th></tr></thead>
        <tbody>${filas}</tbody></table></div>`
        : vacio(comps.length ? "No hay compromisos en esta vista. Prueba «Todos»." : `Aún no hay compromisos.${puede ? " Crea el primero con «+ Compromiso»." : ""}`)}</div>`;
  }
  function enlazarSeccionCompromisos(el) {
    el.querySelectorAll("[data-cver]").forEach((b) => b.addEventListener("click", () => { S.compVista.ver = b.dataset.cver; render(); }));
    const o = el.querySelector("#c-orden"); if (o) o.addEventListener("change", () => { S.compVista.orden = o.value; render(); });
    const g = el.querySelector("#c-agrupar"); if (g) g.addEventListener("change", () => { S.compVista.agrupar = g.checked; render(); });
  }
  function enlazarCompromisos(el) {
    el.querySelectorAll("[data-abrir-comp]").forEach((b) => b.addEventListener("click", () => detalleCompromiso(b.dataset.abrirComp)));
    el.querySelectorAll("[data-cestado]").forEach((b) => b.addEventListener("click", () => {
      const c = S.datos.Compromisos.find((x) => x.ID_Compromiso === b.dataset.cestado);
      const nuevo = b.dataset.a;
      if (nuevo === "Cerrado") confirmar("Cerrar compromiso", `«${c.Compromiso}» (${c.Responsable || "sin responsable"}) quedará cerrado con fecha de hoy.`, "Cerrar",
        () => guardar(() => cambiarEstado(c, "Cerrado"), "Compromiso cerrado"), "primario");
      else guardar(() => cambiarEstado(c, nuevo), `Compromiso ${nuevo === "En curso" ? "en curso" : "pendiente"}`);
    }));
  }
  // Botones rápidos según el estado: Pendiente → «Iniciar», En curso → «Cerrar».
  function botonesEstado(c) {
    const b = R.estadoBase(c);
    if (b === "Pendiente") return `<button class="btn chico" data-cestado="${esc(c.ID_Compromiso)}" data-a="En curso" title="Pasar a En curso">▶ Iniciar</button> <button class="btn chico" data-cestado="${esc(c.ID_Compromiso)}" data-a="Cerrado">✓ Cerrar</button>`;
    if (b === "En curso") return `<button class="btn chico" data-cestado="${esc(c.ID_Compromiso)}" data-a="Cerrado">✓ Cerrar</button>`;
    return "";
  }
  const TEXTO_ESTADO = { Pendiente: "Compromiso reabierto: queda Pendiente.", "En curso": "Compromiso en curso.", Cerrado: "Compromiso cerrado." };
  // Cambia el estado, maneja la fecha de cierre y deja constancia en el hilo.
  async function cambiarEstado(c, nuevo, texto) {
    await S.api.actualizarPorId("Compromisos", "ID_Compromiso", c.ID_Compromiso, { Estado: nuevo, Fecha_Cierre: nuevo === "Cerrado" ? R.hoyISO() : "" });
    await S.api.agregarFila("Comentarios", nuevoComentario(c, texto || TEXTO_ESTADO[nuevo]));
  }

  // Ficha del compromiso: datos, hilo de comentarios y acciones (comentar, cerrar, reabrir, editar).
  function detalleCompromiso(id) {
    const c = S.datos.Compromisos.find((x) => x.ID_Compromiso === id);
    if (!c) return;
    const p = proyecto(c.ID_Proyecto) || {};
    const puede = R.puedeEditar(S.usuario, p, "compromisos");
    const e = R.estadoCompromiso(c);
    const coms = comentariosDe(id);
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="Compromiso">
      <div class="titulo-fila"><div><h2>${esc(c.Compromiso)}</h2><div class="sub">${esc(p.Nombre)} · ${esc(origen(c))}</div></div><button class="btn enlace" id="d-cerrar" aria-label="Cerrar">✕</button></div>
      <div class="contexto">
        <div><span class="sub">Estado</span><b>${pillComp(c)}</b></div>
        <div><span class="sub">Responsable</span><b>${esc(c.Responsable || "—")}</b></div>
        <div><span class="sub">Fecha límite</span><b>${fecha(c.Fecha_Compromiso)}</b></div>
        ${c.Fecha_Cierre ? `<div><span class="sub">Cerrado el</span><b>${fecha(c.Fecha_Cierre)}</b></div>` : ""}
      </div>
      <div class="form-seccion">Seguimiento del compromiso <span class="contador">${coms.length}</span></div>
      <div class="hilo">${coms.map((x) => `<div class="com"><div class="com-cab"><b>${esc(nombreUsuario(x.Autor))}</b><span class="sub">${fechaHora(x.Fecha_Hora)}${lc(x.Autor) === lc(S.usuario.Correo) || R.tieneRol(S.usuario, "Admin") ? ` <button class="btn enlace com-borrar" data-borrar-com="${esc(x.ID_Comentario)}" title="Eliminar este comentario">Eliminar</button>` : ""}</span></div>${x.Texto ? `<div class="com-texto">${esc(x.Texto)}</div>` : ""}${htmlAdjuntos(x)}</div>`).join("") || `<div class="vacio">Aún no hay comentarios. Escribe el primero: avances, dudas o bloqueos de este compromiso.</div>`}</div>
      ${puede ? `<textarea id="d-texto" rows="3" placeholder="Escribe un comentario: qué se avanzó, qué falta, quién debe actuar… (puedes pegar aquí una imagen con Ctrl+V)"></textarea>
        <div class="adj-barra"><button type="button" class="btn chico" id="d-adjuntar">📎 Adjuntar archivo</button><span class="sub">o pega una imagen del correo con Ctrl+V · máx. 5 MB por archivo</span></div>
        <div id="d-pend" class="adj-pend"></div>
        <div class="error-campo" id="d-error" role="alert"></div>
        <div class="acciones derecha">
          <button class="btn peligro-suave" id="d-editar">Editar compromiso</button>
          ${e !== "Cerrado" ? `<button class="btn" id="d-recordar" title="${esc(correoDe(c) ? `Correo a ${correoDe(c)}` : "Sin correo del responsable")}">✉ Recordar</button>` : ""}
          ${e === "Cerrado" ? `<button class="btn" id="d-reabrir">Reabrir</button>` : `${R.estadoBase(c) === "Pendiente" ? `<button class="btn" id="d-encurso">▶ Pasar a En curso</button>` : ""}<button class="btn" id="d-comentar-cerrar">Comentar y cerrar</button>`}
          <button class="btn primario" id="d-comentar">Comentar</button></div>` : `<p class="sub">Solo el PM del proyecto, la PMO o el Admin pueden comentar.</p>`}
    </div>`;
    document.body.appendChild(m);
    $("#d-cerrar").addEventListener("click", cerrarModal);
    m.addEventListener("keydown", (ev) => { if (ev.key === "Escape") cerrarModal(); });
    const h = $(".hilo"); h.scrollTop = h.scrollHeight;
    m.querySelectorAll("[data-borrar-com]").forEach((b) => b.addEventListener("click", () => {
      const x = coms.find((y) => y.ID_Comentario === b.dataset.borrarCom);
      confirmar("Eliminar comentario", `Se eliminará el comentario «${String(x.Texto || "").slice(0, 90)}»${adjuntosDe(x).length ? " y sus adjuntos" : ""}.`, "Eliminar", async () => {
        await guardar(async () => { await borrarAdjuntos([x]); await S.api.eliminarFilas("Comentarios", "ID_Comentario", [x.ID_Comentario]); }, "Comentario eliminado");
        detalleCompromiso(id);
      });
    }));
    enlazarAdjuntos(h);
    if (!puede) return;
    const txt = $("#d-texto");
    txt.focus();
    txt.addEventListener("input", () => { $("#d-error").textContent = ""; });
    // Archivos por adjuntar al próximo comentario.
    const pend = [];
    const pintarPend = () => {
      $("#d-pend").innerHTML = pend.map((a, i) => `<span class="adj-chip">${esImagen(a.tipo) ? `<img src="${a.url}" alt="">` : "📄"} ${esc(a.nombre)} <span class="sub">${tamano(a.tam)}</span><button type="button" class="btn enlace" data-quitar-adj="${i}" aria-label="Quitar">✕</button></span>`).join("");
      document.querySelectorAll("[data-quitar-adj]").forEach((b) => b.addEventListener("click", () => { pend.splice(Number(b.dataset.quitarAdj), 1); pintarPend(); }));
    };
    const agregarArchivos = async (files) => {
      for (const f of files) {
        if (f.size > MAX_ADJ * 3) { $("#d-error").textContent = `«${f.name}» pesa más de 5 MB.`; continue; }
        const a = await prepararArchivo(f);
        if (a.tam > MAX_ADJ) { $("#d-error").textContent = `«${a.nombre}» pesa más de 5 MB.`; continue; }
        a.url = esImagen(a.tipo) ? URL.createObjectURL(new Blob([a.bytes], { type: a.tipo })) : "";
        pend.push(a);
      }
      pintarPend();
    };
    txt.addEventListener("paste", (ev) => {
      const files = [...(ev.clipboardData ? ev.clipboardData.items : [])].filter((i) => i.kind === "file").map((i) => i.getAsFile()).filter(Boolean)
        .map((f, k) => (f.name && f.name !== "image.png" ? f : new File([f], `imagen-pegada-${Date.now()}-${k + 1}.${(f.type.split("/")[1] || "png")}`, { type: f.type })));
      if (files.length) { ev.preventDefault(); agregarArchivos(files); }
    });
    $("#d-adjuntar").addEventListener("click", () => {
      const i = document.createElement("input"); i.type = "file"; i.multiple = true;
      i.addEventListener("change", () => agregarArchivos([...i.files])); i.click();
    });
    const accion = async (texto, estado, exito) => {
      await guardar(async () => {
        const com = nuevoComentario(c, texto);
        const adj = [];
        for (let k = 0; k < pend.length; k++) {
          const a = pend[k], id = `ADJ-${com.ID_Comentario.replace(/^COM-/, "")}-${k + 1}`;
          await guardarArchivo(id, a.bytes, a.nombre, a.tipo, `el adjunto ${k + 1} de ${pend.length}`);
          adj.push({ id, n: a.nombre, t: a.tipo, s: a.tam });
        }
        if (adj.length) com.Adjuntos = JSON.stringify(adj);
        if (estado) await S.api.actualizarPorId("Compromisos", "ID_Compromiso", id, { Estado: estado, Fecha_Cierre: estado === "Cerrado" ? R.hoyISO() : "" });
        await S.api.agregarFila("Comentarios", com);
      }, exito);
      detalleCompromiso(id);   // se vuelve a abrir para seguir la conversación
    };
    $("#d-comentar").addEventListener("click", () => {
      const t = txt.value.trim();
      if (!t && !pend.length) { $("#d-error").textContent = "Escribe el comentario o adjunta un archivo."; return; }
      accion(t, null, "Comentario agregado");
    });
    if ($("#d-comentar-cerrar")) $("#d-comentar-cerrar").addEventListener("click", () => accion(txt.value.trim() || "Compromiso cerrado.", "Cerrado", "Compromiso cerrado"));
    if ($("#d-encurso")) $("#d-encurso").addEventListener("click", () => accion(txt.value.trim() || "Compromiso en curso.", "En curso", "Compromiso en curso"));
    if ($("#d-reabrir")) $("#d-reabrir").addEventListener("click", () => accion(txt.value.trim() || "Compromiso reabierto.", "Pendiente", "Compromiso reabierto"));
    $("#d-editar").addEventListener("click", () => formCompromiso(c));
    if ($("#d-recordar")) $("#d-recordar").addEventListener("click", async () => { const r = await enviarRecordatorio([{ c }]); if (!r) return; render(); detalleCompromiso(id); toast("Correo de recordatorio abierto en tu Outlook"); });
  }

  // ---------- Proyectos ----------
  function vProyectos(el) {
    const FL = S.filtrosLista;
    const lista = filtrados().filter((p) => (!FL.semaforo || p.Semaforo === FL.semaforo) &&
      (!FL.texto || `${p.ID_Proyecto} ${p.Nombre} ${p.Cliente_Area} ${p.Codigo_Almera || ""}`.toLowerCase().includes(FL.texto.toLowerCase())))
      .sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
    el.innerHTML = `
      <div class="titulo-fila"><h1>Proyectos</h1>${R.puede(S.usuario, "crearProyecto") ? `<button class="btn primario" id="b-nuevo">+ Nueva iniciativa</button>` : ""}</div>
      ${barraFiltros()}
      <div class="filtros">
        <label class="filtro crece"><span>Buscar</span><input id="f-texto" type="search" placeholder="ID, nombre, área o código Almera" value="${esc(FL.texto)}"></label>
        <label class="filtro"><span>Semáforo</span><select id="f-sem"><option value="">Todos</option>${(S.cat.Semaforo || []).map((v) => `<option ${v === FL.semaforo ? "selected" : ""}>${esc(v)}</option>`).join("")}</select></label>
        <span class="sub">${lista.length} resultado(s)</span>
      </div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th class="opc">Área</th><th>PM</th><th class="opc">Metodología</th><th class="opc">Fase</th><th>Estado</th><th>Semáforo</th><th>Avance real / plan</th><th class="opc">Fin plan</th><th>Seguimiento</th></tr></thead>
        <tbody>${lista.map((p) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}" tabindex="0">
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.ID_Proyecto)} · prioridad ${esc(String(p.Prioridad).toLowerCase())}</div></td><td class="opc">${esc(p.Cliente_Area)}</td>
          <td>${persona(nombreUsuario(p.PM))}</td><td class="opc">${esc(p.Metodologia)}</td><td class="opc">${esc(p.Fase)}</td><td>${esc(p.Estado)}</td><td>${chipSemaforo(p.Semaforo)}</td>
          <td class="celda-avance">${barraAvance(p.Avance_Real, p.Avance_Planeado)}</td><td class="opc">${fecha(p.Fecha_Fin_Plan)}</td>
          <td>${pill(R.estadoSeguimiento(p, S.datos.Seguimientos).estado)}</td></tr>`).join("") || `<tr><td colspan="10">${vacio(visibles().length ? "No hay proyectos con estos filtros. Prueba «Limpiar filtros»." : "Aún no tienes proyectos. Registra el primero con «+ Nueva iniciativa».")}</td></tr>`}</tbody>
      </table></div></div>`;
    enlazarFiltros();
    $("#f-texto").addEventListener("change", (e) => { FL.texto = e.target.value; render(); });
    $("#f-sem").addEventListener("change", (e) => { FL.semaforo = e.target.value; render(); });
    if ($("#b-nuevo")) $("#b-nuevo").addEventListener("click", () => formProyecto());
    el.querySelectorAll("tr[data-pid]").forEach((r) => r.addEventListener("keydown", (e) => { if (e.key === "Enter") ir("ficha", r.dataset.pid); }));
  }

  // ---------- Ficha ----------
  function vFicha(el) {
    const p = proyecto(S.pid);
    if (!p || !R.esMio(S.usuario, p)) { el.innerHTML = vacio("Proyecto no encontrado o sin acceso."); return; }
    const u = S.usuario;
    const segs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos);
    const hitos = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto).sort((a, b) => (a.Fecha_Plan < b.Fecha_Plan ? -1 : 1));
    const riesgos = S.datos.Riesgos.filter((r) => r.ID_Proyecto === p.ID_Proyecto);
    const comps = S.datos.Compromisos.filter((c) => c.ID_Proyecto === p.ID_Proyecto).map((c) => ({ c, e: R.estadoCompromiso(c) }))
      .sort((a, b) => ORDEN_COMP[a.e] - ORDEN_COMP[b.e] || (a.c.Fecha_Compromiso < b.c.Fecha_Compromiso ? -1 : 1));
    const e = R.estadoSeguimiento(p, S.datos.Seguimientos);
    const puedeP = (a) => R.puedeEditar(u, p, a);
    const dato = (k, v) => `<div class="dato"><div class="dato-k">${k}</div><div class="dato-v">${v}</div></div>`;
    const hoy = R.hoyISO();

    // Pestañas de la ficha: cada sección en su propia vista.
    if (S.fichaPid !== p.ID_Proyecto) { S.fichaPid = p.ID_Proyecto; S.fichaTab = "resumen"; }
    const abiertosN = comps.filter((x) => x.e !== "Cerrado").length;
    const vencidosN = comps.filter((x) => x.e === "Vencido").length;
    const TABS = [["resumen", "Resumen"], ["compromisos", `Compromisos <span class="contador ${vencidosN ? "rojo" : ""}">${abiertosN}</span>`],
      ["seguimientos", `Seguimientos <span class="contador">${segs.length}</span>`], ["hitos", `Hitos <span class="contador">${hitos.length}</span>`], ["riesgos", `Riesgos <span class="contador">${riesgos.filter(riesgoActivo).length}</span>`],
      ["tickets", `Tickets <span class="contador">${ticketsDe(p.ID_Proyecto).filter(ticketAbierto).length}</span>`],
      ["cambios", `Cambios${cambiosDe(p.ID_Proyecto).some((c) => c.Estado === "Solicitado") ? ` <span class="contador rojo">${cambiosDe(p.ID_Proyecto).filter((c) => c.Estado === "Solicitado").length}</span>` : ""}`],
      ["raci", "RACI"], ["lecciones", `Lecciones <span class="contador">${leccionesDe(p.ID_Proyecto).length}</span>`], ["auditoria", "Historial de cambios"]];
    const tab = TABS.some((t) => t[0] === S.fichaTab) ? S.fichaTab : "resumen";
    const pestanas = `<div class="pestanas" role="tablist">${TABS.map(([id, t]) => `<button class="pestana ${id === tab ? "activa" : ""}" role="tab" aria-selected="${id === tab}" data-tab="${id}">${t}</button>`).join("")}</div>`;
    const cuerpo = {
      resumen: () => `      <div class="card datos">
        ${dato("Código Almera", p.Codigo_Almera ? esc(p.Codigo_Almera) : `<span class="sub">Sin registrar</span>`)}${dato("Estado", esc(p.Estado))}${dato("Fase", esc(p.Fase))}${dato("Semáforo", chipSemaforo(p.Semaforo))}
        ${dato("Avance real / plan", barraAvance(p.Avance_Real, p.Avance_Planeado))}
        ${dato("Inicio", fecha(p.Fecha_Inicio))}${dato("Fin planeado", fecha(p.Fecha_Fin_Plan))}
        ${dato("Presupuesto", cop(p.Presupuesto))}${dato("Ejecutado", cop(p.Ejecutado))}
        ${dato("Metodología", esc(p.Metodologia))}${dato("Prioridad", esc(p.Prioridad))}
        ${dato("Frecuencia", esc(p.Frecuencia_Seguimiento))}${dato("Seguimiento", `${pill(e.estado)} <span class="sub">próximo ${fecha(e.proximo)}</span>`)}
        ${dato("Repositorio", enlace(p.URL_Repositorio, "Abrir") || `<span class="sub">Sin registrar</span>`)}${dato("Documentos", enlace(p.URL_Documentos, "Abrir") || `<span class="sub">Sin registrar</span>`)}
        <div class="dato ancho"><div class="dato-k">Comentario de estado</div><div class="dato-v">${esc(p.Comentario_Estado || "—")}</div></div>
      </div>
      ${cardStakeholders(p, puedeP("editarProyecto"))}
      ${cardDependencias(p, puedeP("editarProyecto"))}
      <div class="card"><h2>Curva S del presupuesto</h2>${(() => { S._curvaP = datosCurvaS([p]); return S._curvaP ? `<div class="grafico-alto"><canvas id="g-curvas-p"></canvas></div><p class="sub">Planeado: ${num0(p.Presupuesto_Base) !== null ? "presupuesto de la línea base" : "presupuesto"} repartido entre inicio y fin. Real: ejecutado reportado en cada seguimiento.</p>` : vacio("Registra presupuesto y fechas para ver la curva S."); })()}</div>
      <div class="card"><h2>Curva de avance</h2>${segs.length ? `<div class="grafico-alto"><canvas id="g-curva"></canvas></div>` : vacio("Aún no hay seguimientos. La curva aparece con el primer reporte.")}</div>
`,
      compromisos: () => seccionCompromisos(p, comps, segs, puedeP("compromisos")),
      tickets: () => seccionTickets(p, puedeP("compromisos")),
      cambios: () => seccionCambios(p, puedeP("editarProyecto"), R.puede(u, "verTodo")),
      raci: () => seccionRACI(p, puedeP("editarProyecto")),
      lecciones: () => seccionLecciones(p, puedeP("editarProyecto")),
      auditoria: () => seccionAuditoria(p),
      seguimientos: () => `      <div class="card"><div class="titulo-fila"><h2>Historial de sesiones de seguimiento</h2><div class="acciones"><button class="btn chico" id="b-expandir">Expandir compromisos</button><button class="btn chico" id="b-contraer">Contraer</button>${puedeP("seguimiento") && p.Estado === "Activo" ? `<button class="btn" id="b-seg2">+ Registrar seguimiento</button>` : ""}</div></div>
        <p class="sub">Cada fila es una sesión: su avance, lo que pasó, los compromisos que se acordaron y el acta.</p><div class="tabla-scroll"><table class="historial">
        <thead><tr><th>Sesión</th><th>Avance</th><th>Logros y próximos pasos</th><th>Bloqueos</th><th>Compromisos acordados</th><th>Acta</th></tr></thead>
        <tbody>${segs.slice().reverse().map((s) => {
          const acordados = S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento);
          return `<tr><td><b>${fecha(s.Fecha_Corte)}</b><div class="sub">${esc(s.Semana)} · ${esc(nombreUsuario(s.Reportado_Por))}</div></td>
          <td>${pct(s.Avance_Real)}<div>${chipSemaforo(s.Semaforo)}</div></td>
          <td>${esc(s.Logros)}<div class="sub">Sigue: ${esc(s.Proximos_Pasos || "—")}</div></td><td>${esc(s.Bloqueos || "—")}</td>
          <td>${acordados.length ? (() => {
            const ab = acordados.filter((c) => !R.cerrado(c)).length, ve = acordados.filter((c) => R.estadoCompromiso(c) === "Vencido").length;
            return `<details class="acordados" data-seg-det="${esc(s.ID_Seguimiento)}" ${(S.segAbiertos || new Set()).has(s.ID_Seguimiento) ? "open" : ""}>
              <summary><b>${acordados.length} compromiso(s)</b> <span class="sub">· ${ab} abierto(s)${ve ? ` · <span class="baja">${ve} vencido(s)</span>` : ""}</span></summary>
              <ul class="mini">${acordados.map((c) => `<li><button class="btn enlace" data-abrir-comp="${esc(c.ID_Compromiso)}">${esc(c.Compromiso)}</button> <span class="sub">${esc(c.Responsable)} · ${fecha(c.Fecha_Compromiso)}</span> ${pillComp(c)}</li>`).join("")}</ul></details>`;
          })() : `<span class="sub">Sin compromisos</span>`}
            ${puedeP("compromisos") ? `<div><button class="btn chico" data-comp-seg="${esc(s.ID_Seguimiento)}">+ Compromiso de esta sesión</button></div>` : ""}</td>
          <td>${s.Fecha_Acta ? fecha(s.Fecha_Acta) : `<span class="sub">Sin fecha</span>`}<div class="celda-acta">${celdaActa(s, puedeP("seguimiento"))}</div>
            ${puedeP("seguimiento") ? `<button class="btn chico enlace-edicion" data-edit-seg="${esc(s.ID_Seguimiento)}">Editar sesión</button>` : ""}</td></tr>`;
        }).join("") || `<tr><td colspan="6">${vacio("Sin seguimientos.")}</td></tr>`}</tbody></table></div></div>
`,
      hitos: () => `      <div class="card"><div class="titulo-fila"><h2>Hitos</h2>${puedeP("hitos") ? `<button class="btn" id="b-hito">+ Hito</button>` : ""}</div><div class="tabla-scroll"><table>
        <thead><tr><th>Hito</th><th>Fecha plan</th><th>Fecha real</th><th>Estado</th><th></th></tr></thead>
        <tbody>${hitos.map((h) => { const est = h.Estado !== "Cumplido" && h.Fecha_Plan < hoy ? "Atrasado" : h.Estado; return `<tr><td>${esc(h.Hito)}</td><td>${fecha(h.Fecha_Plan)}</td><td>${h.Fecha_Real ? fecha(h.Fecha_Real) : "—"}</td><td>${pill(est)}</td>
          <td class="derecha">${puedeP("hitos") ? `<button class="btn chico" data-hito="${esc(h.ID_Hito)}">Editar</button>` : ""}</td></tr>`; }).join("") || `<tr><td colspan="5">${vacio("Sin hitos. Agrega los hitos clave del proyecto con «+ Hito».")}</td></tr>`}</tbody></table></div></div>
`,
      riesgos: () => seccionRiesgos(p, puedeP("riesgos")),
    }[tab]();
    el.innerHTML = `
      <button class="btn enlace" id="b-volver">← Volver a proyectos</button>
      <div class="titulo-fila"><div><h1>${esc(p.Nombre)}</h1><div class="sub">${esc(p.ID_Proyecto)}${p.Codigo_Almera ? ` · Almera ${esc(p.Codigo_Almera)}` : ""} · ${esc(p.Cliente_Area)} · PM ${esc(nombreUsuario(p.PM))}</div></div>
        <div class="acciones">
          ${enlace(p.URL_Repositorio, "Repositorio")}${enlace(p.URL_Documentos, "Documentos")}
          ${puedeP("seguimiento") && p.Estado === "Activo" ? `<button class="btn primario" id="b-seg">Registrar seguimiento</button>` : ""}
          <button class="btn" id="b-hv" title="Documento con todo el historial del proyecto, listo para guardar como PDF">📄 Hoja de vida</button>
          ${puedeP("editarProyecto") ? `<button class="btn" id="b-editar">Editar proyecto</button>` : ""}
        </div></div>
      ${p.Estado === "Activo" && e.estado !== "Al día" ? `<div class="aviso ${e.estado === "Vencido" ? "rojo" : ""}">${e.estado === "Vencido" ? `El seguimiento está vencido desde el ${fecha(e.proximo)}.` : `El próximo seguimiento vence el ${fecha(e.proximo)}.`}${puedeP("seguimiento") ? " Usa «Registrar seguimiento»." : ""}</div>` : ""}
      ${pestanas}
      ${cuerpo}`
    $("#b-volver").addEventListener("click", () => ir("proyectos"));
    if ($("#b-seg")) $("#b-seg").addEventListener("click", () => formSeguimiento(p));
    if ($("#b-editar")) $("#b-editar").addEventListener("click", () => formProyecto(p));
    if ($("#b-hito")) $("#b-hito").addEventListener("click", () => formHito(p));
    if ($("#b-comp")) $("#b-comp").addEventListener("click", () => formCompromiso(null, p));
    if ($("#b-seg2")) $("#b-seg2").addEventListener("click", () => formSeguimiento(p));
    if ($("#b-tk")) $("#b-tk").addEventListener("click", () => formTicket(p));
    $("#b-hv").addEventListener("click", () => hojaDeVida(p));
    if ($("#g-curvas-p") && S._curvaP) graficoCurvaS($("#g-curvas-p"), S._curvaP);
    if ($("#b-cambio")) $("#b-cambio").addEventListener("click", () => formCambio(p));
    if ($("#b-lb")) $("#b-lb").addEventListener("click", () => fijarLineaBase(p));
    el.querySelectorAll("[data-cambio]").forEach((b) => b.addEventListener("click", () => formCambio(p, cambiosDe(p.ID_Proyecto).find((c) => c.ID_Cambio === b.dataset.cambio))));
    el.querySelectorAll("[data-decidir]").forEach((b) => b.addEventListener("click", () => decidirCambio(p, cambiosDe(p.ID_Proyecto).find((c) => c.ID_Cambio === b.dataset.decidir))));
    if ($("#b-leccion")) $("#b-leccion").addEventListener("click", () => formLeccion(p));
    el.querySelectorAll("[data-leccion]").forEach((b) => b.addEventListener("click", () => formLeccion(p, leccionesDe(p.ID_Proyecto).find((l) => l.ID_Leccion === b.dataset.leccion))));
    if ($("#b-dep")) $("#b-dep").addEventListener("click", () => formDependencia(p));
    el.querySelectorAll("[data-dep]").forEach((b) => b.addEventListener("click", () => formDependencia(p, (S.datos.Dependencias || []).find((d) => d.ID_Dependencia === b.dataset.dep))));
    if (tab === "raci") enlazarRACI(el, p);
    if (tab === "auditoria") {
      const fa = (a) => a.ID_Proyecto === p.ID_Proyecto;
      llenarAuditoria(el, fa, false);
      $("#aud-recargar").addEventListener("click", async () => { await cargarAuditoria(true); llenarAuditoria(el, fa, false); });
    }
    el.querySelectorAll("[data-tk]").forEach((b) => b.addEventListener("click", () => formTicket(p, ticketsDe(p.ID_Proyecto).find((t) => t.ID_Ticket === b.dataset.tk))));
    el.querySelectorAll("[data-tkver]").forEach((b) => b.addEventListener("click", () => { S.tkVer = b.dataset.tkver; render(); }));
    S.segAbiertos = S.segAbiertos || new Set();
    el.querySelectorAll("[data-seg-det]").forEach((d) => d.addEventListener("toggle", () => { if (d.open) S.segAbiertos.add(d.dataset.segDet); else S.segAbiertos.delete(d.dataset.segDet); }));
    const todos = (abrir) => el.querySelectorAll("[data-seg-det]").forEach((d) => { d.open = abrir; });
    if ($("#b-expandir")) $("#b-expandir").addEventListener("click", () => todos(true));
    if ($("#b-contraer")) $("#b-contraer").addEventListener("click", () => todos(false));
    el.querySelectorAll("[data-tab]").forEach((b) => b.addEventListener("click", () => { S.fichaTab = b.dataset.tab; render(); }));
    el.querySelectorAll("[data-comp-seg]").forEach((b) => b.addEventListener("click", () => formCompromiso(null, p, b.dataset.compSeg)));
    enlazarSeccionCompromisos(el);
    if ($("#b-riesgo")) $("#b-riesgo").addEventListener("click", () => formRiesgo(p));
    if ($("#b-rg-carga")) $("#b-rg-carga").addEventListener("click", () => cargarRiesgosExcel(p));
    if (tab === "riesgos") enlazarPanelRiesgos(el, "p:" + p.ID_Proyecto, null);
    if ($("#b-stk")) $("#b-stk").addEventListener("click", () => formStakeholder(p));
    el.querySelectorAll("[data-stk]").forEach((b) => b.addEventListener("click", () => formStakeholder(p, (S.datos.Stakeholders || []).find((x) => x.ID_Stakeholder === b.dataset.stk))));
    el.querySelectorAll("[data-rol-stk]").forEach((b) => b.addEventListener("click", () => formStakeholder(p, null, b.dataset.rolStk)));
    el.querySelectorAll("[data-hito]").forEach((b) => b.addEventListener("click", () => formHito(p, hitos.find((h) => h.ID_Hito === b.dataset.hito))));
    el.querySelectorAll("[data-riesgo]").forEach((b) => b.addEventListener("click", () => formRiesgo(p, riesgos.find((r) => r.ID_Riesgo === b.dataset.riesgo))));
    enlazarCompromisos(el);
    el.querySelectorAll("[data-edit-seg]").forEach((b) => b.addEventListener("click", () => formEditarSeguimiento(segs.find((s) => s.ID_Seguimiento === b.dataset.editSeg))));
    if (segs.length && $("#g-curva")) grafico($("#g-curva"), { type: "line", data: { labels: segs.map((s) => fecha(s.Fecha_Corte)), datasets: [
      { label: "Avance real", data: segs.map((s) => Number(s.Avance_Real) || 0), borderColor: COLORES.azul, backgroundColor: COLORES.azul, tension: 0.2 }] },
      options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + "%" } } } } });
  }
  const nivel = (c) => { const n = R.nivelRiesgo(Number(c) || 0); return `<span class="pill riesgo-${n.toLowerCase()}">${n} (${Number(c) || 0})</span>`; };

  // ---------- Cronograma ----------
  function vCronograma(el) {
    const hoy = R.hoyISO();
    const RANGOS = { "12m": "12 meses (3 atrás, 9 adelante)", anio: `Año ${hoy.slice(0, 4)}`, todo: "Todo el portafolio" };
    const rango = RANGOS[S.cronoRango] ? S.cronoRango : "12m";
    const todos = filtrados().filter((p) => (p.Estado === "Activo" || p.Estado === "En pausa") && p.Fecha_Inicio && p.Fecha_Fin_Plan);
    let ini, fin;
    if (rango === "12m") { ini = R.sumarDias(hoy, -92).slice(0, 8) + "01"; fin = R.sumarDias(ini, 366); fin = fin.slice(0, 8) + "01"; fin = R.sumarDias(fin, -1); }
    else if (rango === "anio") { ini = `${hoy.slice(0, 4)}-01-01`; fin = `${hoy.slice(0, 4)}-12-31`; }
    else if (todos.length) {
      ini = todos.reduce((m, p) => (p.Fecha_Inicio < m ? p.Fecha_Inicio : m), todos[0].Fecha_Inicio).slice(0, 8) + "01";
      fin = todos.reduce((m, p) => (p.Fecha_Fin_Plan > m ? p.Fecha_Fin_Plan : m), todos[0].Fecha_Fin_Plan);
    }
    const ps = todos.filter((p) => p.Fecha_Fin_Plan >= ini && p.Fecha_Inicio <= fin).sort((a, b) => (a.Fecha_Inicio < b.Fecha_Inicio ? -1 : 1));
    const selRango = `<div class="filtros"><label class="filtro"><span>Periodo</span><select id="crono-rango">${Object.entries(RANGOS).map(([k, t]) => `<option value="${k}" ${k === rango ? "selected" : ""}>${t}</option>`).join("")}</select></label></div>`;
    if (!ps.length) {
      el.innerHTML = `<h1>Cronograma</h1>${barraFiltros()}${selRango}${vacio("No hay proyectos activos o en pausa con fechas en este periodo.")}`;
      enlazarFiltros(); $("#crono-rango").addEventListener("change", (e) => { S.cronoRango = e.target.value; render(); }); return;
    }
    const total = Math.max(1, R.difDias(fin, ini) + 1);
    const x = (iso) => Math.max(0, Math.min(100, (R.difDias(iso, ini) / total) * 100));
    const MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    const meses = [];
    for (let y = Number(ini.slice(0, 4)), m = Number(ini.slice(5, 7)) - 1; `${y}-${String(m + 1).padStart(2, "0")}-01` <= fin; m === 11 ? (y++, m = 0) : m++) {
      meses.push({ iso: `${y}-${String(m + 1).padStart(2, "0")}-01`, m, y });
    }
    // Menos etiquetas cuando hay muchos meses, para que no se monten.
    const paso = meses.length <= 13 ? 1 : meses.length <= 26 ? 2 : meses.length <= 40 ? 3 : 6;
    const anios = [...new Set(meses.map((m) => m.y))].map((y) => {
      const a = x(`${y}-01-01` < ini ? ini : `${y}-01-01`), b = x(`${y}-12-31` > fin ? fin : `${y}-12-31`);
      return { y, a, w: Math.max(0, b - a) };
    });
    const colorHito = (h) => (h.Estado === "Cumplido" ? COLORES.Verde : h.Fecha_Plan < hoy ? COLORES.Rojo : "#888");
    const recorte = (p) => (p.Fecha_Inicio < ini ? " corta-ini" : "") + (p.Fecha_Fin_Plan > fin ? " corta-fin" : "");
    el.innerHTML = `
      <h1>Cronograma</h1>
      ${barraFiltros()}
      ${selRango}
      <div class="leyenda"><span><i class="lg-real"></i>Avance real</span><span><i class="lg-plan"></i>Duración planeada</span>
        <span><i class="lg-hito" style="background:${COLORES.Verde}"></i>Hito cumplido</span><span><i class="lg-hito" style="background:${COLORES.Rojo}"></i>Hito atrasado</span>
        <span><i class="lg-hito" style="background:#888"></i>Hito pendiente</span><span><i class="lg-hoy"></i>Hoy</span></div>
      <div class="card gantt">
        <div class="g-fila g-cab"><div class="g-nombre"></div><div class="g-pista g-escala">
          <div class="g-anios">${anios.map((a) => `<span class="g-anio" style="left:${a.a}%;width:${a.w}%">${a.y}</span>`).join("")}</div>
          <div class="g-meses">${meses.map((m, i) => `<span class="g-tick ${m.m === 0 ? "enero" : ""}" style="left:${x(m.iso)}%"></span>${i % paso === 0 ? `<span class="g-mes" style="left:${x(m.iso)}%">${MES[m.m]}</span>` : ""}`).join("")}</div>
        </div></div>
        ${ps.map((p) => {
          const hs = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto && h.Fecha_Plan >= ini && h.Fecha_Plan <= fin);
          return `<div class="g-fila clic" data-pid="${esc(p.ID_Proyecto)}"><div class="g-nombre"><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.Estado)} · ${pct(p.Avance_Real)} · ${fecha(p.Fecha_Inicio)} → ${fecha(p.Fecha_Fin_Plan)}</div>${(() => {
            const ds = depsDe(p.ID_Proyecto); if (!ds.length) return "";
            const mal = ds.filter(conflictoDep);
            return `<div class="sub ${mal.length ? "baja" : ""}" title="${esc(ds.map((d) => `Depende de ${(proyecto(d.Depende_De) || {}).Nombre || d.Depende_De}${conflictoDep(d) ? " ⚠ " + conflictoDep(d) : ""}`).join("\n"))}">⛓ Depende de ${ds.length}${mal.length ? " ⚠" : ""}</div>`;
          })()}</div>
            <div class="g-pista">
              ${meses.map((m) => `<span class="g-rejilla ${m.m === 0 ? "enero" : ""}" style="left:${x(m.iso)}%"></span>`).join("")}
              <div class="g-barra${recorte(p)}" style="left:${x(p.Fecha_Inicio)}%;width:${Math.max(1, x(p.Fecha_Fin_Plan) - x(p.Fecha_Inicio))}%" title="${esc(p.Nombre)}: ${fecha(p.Fecha_Inicio)} → ${fecha(p.Fecha_Fin_Plan)}"><div class="g-real" style="width:${Math.min(100, Number(p.Avance_Real) || 0)}%"></div></div>
              ${hs.map((h) => `<span class="g-hito" title="${esc(h.Hito)} · ${fecha(h.Fecha_Plan)} · ${esc(h.Estado)}" style="left:${x(h.Fecha_Plan)}%;background:${colorHito(h)}"></span>`).join("")}
              ${hoy >= ini && hoy <= fin ? `<span class="g-hoy" style="left:${x(hoy)}%"></span>` : ""}
            </div></div>`;
        }).join("")}
      </div>
      ${rango !== "todo" && ps.some((p) => recorte(p)) ? `<p class="sub">Las barras con borde punteado siguen antes o después del periodo mostrado. Elige «Todo el portafolio» para verlas completas.</p>` : ""}`;
    enlazarFiltros();
    $("#crono-rango").addEventListener("change", (e) => { S.cronoRango = e.target.value; render(); });
  }

  // ---------- Riesgos ----------
  function vRiesgos(el) {
    const ids = new Set(filtrados().map((p) => p.ID_Proyecto));
    const todos = S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto));
    el.innerHTML = `<h1>Riesgos del portafolio</h1>${barraFiltros()}
      ${panelRiesgos("global", todos, { conProyecto: true, puedeFn: (r) => R.puedeEditar(S.usuario, proyecto(r.ID_Proyecto) || {}, "riesgos") })}`;
    enlazarFiltros();
    enlazarPanelRiesgos(el, "global", (r) => formRiesgo(proyecto(r.ID_Proyecto), r));
  }

  // ---------- Catálogos ----------
  // Listas editables: solo alimentan los desplegables. Las demás las usa la lógica de la app y no se editan aquí.
  const CATALOGOS_EDITABLES = [
    { lista: "Cliente_Area", t: "Cliente / área", campo: "Cliente_Area" },
    { lista: "Metodologia", t: "Metodología", campo: "Metodologia" },
    { lista: "Fase", t: "Fase", campo: "Fase" },
    { lista: "Prioridad", t: "Prioridad", campo: "Prioridad" },
    { lista: "Rol_Stakeholder", t: "Rol del stakeholder", campo: "Rol", tabla: "Stakeholders", idCol: "ID_Stakeholder" },
    { lista: "Estado_Ticket", t: "Estado del ticket (helpdesk)", campo: "Estado", tabla: "Tickets", idCol: "ID_Ticket" },
    { lista: "Tipo_Cambio", t: "Tipo de cambio", campo: "Tipo", tabla: "Cambios", idCol: "ID_Cambio" },
    { lista: "Categoria_Leccion", t: "Categoría de lección aprendida", campo: "Categoria", tabla: "Lecciones", idCol: "ID_Leccion" },
    { lista: "Tipo_Demanda", t: "Tipo de demanda", campo: "Tipo", tabla: "Demandas", idCol: "ID_Demanda" },
    { lista: "Categoria_Demanda", t: "Categoría de demanda", campo: "Categoria", tabla: "Demandas", idCol: "ID_Demanda" },
    { lista: "Copia_Recordatorios", t: "Copia fija de recordatorios (correos)", campo: "", tabla: "__ninguna", correo: true },
  ];
  const filasCat = (c) => (S.datos[c.tabla || "Proyectos"] || []);
  const CATALOGOS_FIJOS = { Estado: "Estado del proyecto", Semaforo: "Semáforo", Frecuencia_Seguimiento: "Frecuencia de seguimiento",
    Estado_Hito: "Estado del hito", Tipo_Riesgo: "Tipo de riesgo", Estado_Riesgo: "Estado del riesgo", Estado_Compromiso: "Estado del compromiso", Rol: "Rol" };

  function vCatalogos(el) {
    if (!R.puede(S.usuario, "catalogos")) { el.innerHTML = vacio("Sin acceso."); return; }
    const enUso = (c, valor) => filasCat(c).filter((p) => String(p[c.campo]) === String(valor)).length;
    const usosProv = (id) => S.datos.Proyectos.filter((p) => idsLista(p.Proveedores).includes(id)).length;
    el.innerHTML = `
      <h1>Catálogos</h1>
      <p class="sub">Los cambios se guardan en la tabla Catalogos del Excel y aplican para todos los usuarios.</p>
      <div class="grid2">
        ${CATALOGOS_EDITABLES.map((c) => {
          const valores = S.cat[c.lista] || [];
          return `<div class="card" data-anchor="cat-${c.lista}">
            <h2>${esc(c.t)}</h2>
            <p class="sub">${c.correo ? "Estos correos quedan en copia (y se pueden desmarcar) en cada recordatorio de compromisos." : c.tabla ? "Entre paréntesis: registros que usan cada valor." : "Entre paréntesis: proyectos que usan cada valor."}</p>
            <ul class="cat-lista">${valores.map((v) => `<li><span class="cat-valor">${esc(v)} ${c.correo ? "" : `<span class="contador" title="En uso">${enUso(c, v)}</span>`}</span>
              <span class="cat-acc"><button class="btn chico" data-renombrar="${esc(c.lista)}" data-valor="${esc(v)}" title="Renombrar">✎ Renombrar</button><button class="btn chico" data-quitar="${esc(c.lista)}" data-valor="${esc(v)}" title="Quitar">✕</button></span></li>`).join("") || `<li>${vacio("Sin valores.")}</li>`}</ul>
            <form class="cat-agregar" data-lista="${esc(c.lista)}" novalidate>
              <input name="valor" type="${c.correo ? "email" : "text"}" placeholder="${c.correo ? "nombre@fsfb.org.co" : "Nuevo valor"}" aria-label="Nuevo valor para ${esc(c.t)}">
              <button class="btn primario" type="submit">Agregar</button>
            </form>
            <div class="error-campo" data-e-cat="${esc(c.lista)}" role="alert"></div>
          </div>`;
        }).join("")}
      </div>
      <div class="card" id="cat-proveedores"><div class="titulo-fila"><h2>Proveedores</h2><button class="btn primario" id="b-nuevo-prov">+ Proveedor</button></div>
        <p class="sub">Se guardan en la tabla Proveedores. Luego los asignas a cada proyecto con «Editar proyecto».</p>
        <div class="tabla-scroll"><table><thead><tr><th>Proveedor</th><th>Servicio</th><th>Contacto</th><th>Proyectos</th><th>Activo</th><th></th></tr></thead>
          <tbody>${proveedores().map((x) => `<tr><td><b>${esc(x.Nombre)}</b>${x.NIT ? `<div class="sub">NIT ${esc(x.NIT)}</div>` : ""}</td><td>${esc(x.Servicio || "—")}</td>
            <td>${esc(x.Contacto || "")}${x.Contacto ? "<br>" : ""}${contacto(x.Correo, x.Telefono)}</td><td>${usosProv(x.ID_Proveedor)}</td><td>${esc(x.Activo || "Sí")}</td>
            <td class="derecha"><button class="btn chico" data-prov="${esc(x.ID_Proveedor)}">Editar</button></td></tr>`).join("") || `<tr><td colspan="6">${vacio("Aún no hay proveedores. Crea el primero con «+ Proveedor».")}</td></tr>`}</tbody></table></div>
      </div>
      <div class="card"><h2>Listas fijas</h2>
        <p class="sub">Estas listas controlan reglas de la app (colores, días de seguimiento, permisos), por eso no se editan aquí. Si necesitas cambiarlas, pídelo como una mejora.</p>
        <div class="tabla-scroll"><table><thead><tr><th>Lista</th><th>Valores</th></tr></thead>
          <tbody>${Object.entries(CATALOGOS_FIJOS).map(([k, t]) => `<tr><td>${esc(t)}</td><td>${esc((S.cat[k] || []).join(", "))}</td></tr>`).join("")}</tbody></table></div>
      </div>`;
    $("#b-nuevo-prov").addEventListener("click", () => formProveedor());
    el.querySelectorAll("[data-prov]").forEach((b) => b.addEventListener("click", () => formProveedor(proveedor(b.dataset.prov))));
    el.querySelectorAll(".cat-agregar").forEach((f) => {
      const lista = f.dataset.lista;
      const err = el.querySelector(`[data-e-cat="${lista}"]`);
      f.valor.addEventListener("input", () => { err.textContent = ""; });
      f.addEventListener("submit", (ev) => {
        ev.preventDefault();
        const valor = f.valor.value.trim();
        if (!valor) { err.textContent = "Escribe el valor que quieres agregar."; return; }
        if ((S.cat[lista] || []).some((v) => lc(v) === lc(valor))) { err.textContent = "Ese valor ya existe en la lista."; return; }
        const esCorreo = (CATALOGOS_EDITABLES.find((x) => x.lista === lista) || {}).correo;
        if (esCorreo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(valor)) { err.textContent = "Escribe un correo válido."; return; }
        if (valor.length > (esCorreo ? 120 : 60)) { err.textContent = "Demasiado largo."; return; }
        guardar(() => S.api.agregarFila("Catalogos", { Lista: lista, Valor: valor }), `«${valor}» agregado`);
      });
    });
    el.querySelectorAll("[data-renombrar]").forEach((b) => b.addEventListener("click", () => {
      const lista = b.dataset.renombrar, valor = b.dataset.valor;
      const c = CATALOGOS_EDITABLES.find((x) => x.lista === lista);
      const usan = filasCat(c).filter((p) => String(p[c.campo]) === valor);
      const que = c.tabla ? "registro(s)" : "proyecto(s)";
      modal(`Renombrar «${valor}»`, [{ k: "Nuevo", label: "Nuevo nombre", def: valor, ancho: true,
        ayuda: usan.length ? `También se actualizará en ${usan.length} ${que} que lo usan.` : "Nadie usa este valor todavía." }], {}, (fd) => {
        guardar(async () => {
          await S.api.eliminarFila("Catalogos", { Lista: lista, Valor: valor });
          await S.api.agregarFila("Catalogos", { Lista: lista, Valor: fd.Nuevo });
          const idCol = c.idCol || "ID_Proyecto";
          if (usan.length) await S.api.actualizarVarios(c.tabla || "Proyectos", idCol, usan.map((p) => ({ id: p[idCol], cambios: { [c.campo]: fd.Nuevo } })));
        }, `«${valor}» ahora es «${fd.Nuevo}»`);
      }, (fd) => {
        if (!fd.Nuevo) return "Escribe el nuevo nombre.";
        if (fd.Nuevo === valor) return "Es el mismo nombre.";
        if ((S.cat[lista] || []).some((v) => lc(v) === lc(fd.Nuevo) && v !== valor)) return "Ese valor ya existe en la lista.";
        return "";
      });
    }));
    el.querySelectorAll("[data-quitar]").forEach((b) => b.addEventListener("click", () => {
      const lista = b.dataset.quitar, valor = b.dataset.valor;
      const c = CATALOGOS_EDITABLES.find((x) => x.lista === lista);
      const err = el.querySelector(`[data-e-cat="${lista}"]`);
      if (!c.correo && (S.cat[lista] || []).length <= 1) { err.textContent = "La lista debe tener al menos un valor."; return; }
      const n = enUso(c, valor);
      confirmar(`Quitar «${valor}»`,
        n ? `${n} ${c.tabla ? "registro(s)" : "proyecto(s)"} usan este valor. Lo conservarán, pero ya no aparecerá como opción en los formularios.` : `«${valor}» dejará de aparecer en ${c.t}.`,
        "Quitar", () => guardar(() => S.api.eliminarFila("Catalogos", { Lista: lista, Valor: valor }), `«${valor}» quitado`));
    }));
  }

  // ---------- Usuarios ----------
  function vUsuarios(el) {
    if (!R.puede(S.usuario, "usuarios")) { el.innerHTML = vacio("Sin acceso."); return; }
    const comoPM = (u) => S.datos.Proyectos.filter((p) => lc(p.PM) === lc(u.Correo)).map((p) => p.ID_Proyecto);
    el.innerHTML = `
      <div class="titulo-fila"><h1>Usuarios</h1><button class="btn primario" id="b-nuevo-u">+ Nuevo usuario</button></div>
      <div class="card roles-guia"><h2>Qué puede hacer cada rol</h2><div class="tabla-scroll"><table>
        <thead><tr><th>Rol</th><th>Ve</th><th>Puede</th></tr></thead><tbody>
        <tr><td><b>Admin</b></td><td>Todos los proyectos</td><td>Todo, incluidos usuarios y catálogos</td></tr>
        <tr><td><b>PMO</b></td><td>Todos los proyectos</td><td>Crear y editar proyectos, frecuencias, seguimientos, compromisos, hitos, riesgos y catálogos</td></tr>
        <tr><td><b>PM</b></td><td>Los proyectos donde figura como PM</td><td>Registrar iniciativas y gestionar seguimientos, compromisos, hitos y riesgos de sus proyectos</td></tr>
        <tr><td><b>Lector</b></td><td>Los proyectos asignados aquí (o Todos)</td><td>Solo consultar</td></tr>
        </tbody></table></div><p class="sub">Un usuario puede tener varios roles; sus permisos se suman.</p></div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Usuario</th><th>Roles</th><th>Proyectos como PM</th><th>Consulta (Lector)</th><th>Activo</th><th></th></tr></thead>
        <tbody>${S.datos.Usuarios.slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es")).map((u) => {
          const pm = comoPM(u);
          return `<tr><td><b>${esc(u.Nombre)}</b><div class="sub">${esc(u.Correo)}</div></td>
          <td>${R.roles(u).map((r) => `<span class="pill rol">${esc(r)}</span>`).join(" ") || `<span class="pill vencido">Sin rol</span>`}</td>
          <td>${R.tieneRol(u, "PM") ? (pm.length ? esc(pm.join(", ")) : `<span class="sub">Sin proyectos</span>`) : "—"}</td>
          <td>${R.tieneRol(u, "Lector") ? esc(u.Proyectos || "—") : "—"}</td><td>${esc(u.Activo)}</td>
          <td class="derecha"><button class="btn chico" data-u="${esc(u.Correo)}">Editar</button></td></tr>`;
        }).join("")}</tbody></table></div>
        <p class="sub">Los proyectos de un PM salen de la tabla Proyectos (columna PM). Para cambiarlos, edita el PM responsable en cada proyecto.</p></div>`;
    $("#b-nuevo-u").addEventListener("click", () => formUsuario());
    el.querySelectorAll("[data-u]").forEach((b) => b.addEventListener("click", () => formUsuario(S.datos.Usuarios.find((u) => u.Correo === b.dataset.u))));
  }

  // ---------- Actas en PDF guardadas en el Excel ----------
  const MAX_ACTA = 5 * 1024 * 1024;           // 5 MB por acta
  const TAM_PARTE = 30000;                     // caracteres por celda (límite de Excel: 32.767)
  // Guarda cualquier archivo en la hoja oculta Archivos, en partes de 30.000 caracteres.
  async function guardarArchivo(id, bytes, nombre, tipo, etiqueta = "el archivo") {
    const b64 = ACTAS.aBase64(bytes);
    const total = Math.max(1, Math.ceil(b64.length / TAM_PARTE));
    await S.api.borrarArchivo(id);
    for (let i = 0; i < total; i++) {
      cargando(true, `Guardando ${etiqueta} en Excel… ${i + 1} de ${total}`);
      await S.api.guardarParte({ ID_Archivo: id, Parte: i, Total: total, Nombre: nombre, Tipo: tipo || "application/octet-stream", Datos: b64.slice(i * TAM_PARTE, (i + 1) * TAM_PARTE) });
    }
  }
  const cacheArchivos = {};
  async function leerArchivo(id, avance) {
    if (cacheArchivos[id]) return cacheArchivos[id];
    const primera = await S.api.leerParte(id, 0);
    let b64 = primera.Datos;
    for (let i = 1; i < primera.Total; i++) { if (avance) avance(i + 1, primera.Total); b64 += (await S.api.leerParte(id, i)).Datos; }
    return (cacheArchivos[id] = { bytes: ACTAS.deBase64(b64), nombre: primera.Nombre, tipo: primera.Tipo });
  }
  async function subirActa(idSeg, bytes, nombre) { await guardarArchivo(idSeg, bytes, nombre, "application/pdf", "el acta"); }
  async function verActa(idSeg, nombre) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja visor" role="dialog" aria-modal="true" aria-label="Acta">
      <div class="titulo-fila"><h2>${esc(nombre || "Acta")}</h2><div class="acciones"><a class="btn" id="v-descargar" hidden>Descargar</a><button class="btn enlace" id="v-cerrar" aria-label="Cerrar">✕</button></div></div>
      <div id="v-paginas" class="visor-paginas"><div class="vacio">Abriendo el acta…</div></div></div>`;
    document.body.appendChild(m);
    $("#v-cerrar").addEventListener("click", cerrarModal);
    m.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarModal(); });
    try {
      const primera = await S.api.leerParte(idSeg, 0);
      let b64 = primera.Datos;
      for (let i = 1; i < primera.Total; i++) {
        $("#v-paginas").innerHTML = `<div class="vacio">Abriendo el acta… ${i + 1} de ${primera.Total}</div>`;
        b64 += (await S.api.leerParte(idSeg, i)).Datos;
      }
      const bytes = ACTAS.deBase64(b64);
      const d = $("#v-descargar");
      if (d) { d.href = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" })); d.download = nombre || "acta.pdf"; d.hidden = false; }
      await ACTAS.mostrar(bytes, $("#v-paginas"));
    } catch (e) {
      const c = $("#v-paginas");
      if (c) c.innerHTML = vacio("No se pudo abrir el acta: " + e.message);
    }
  }
  function elegirPDF(alElegir) {
    const i = document.createElement("input");
    i.type = "file"; i.accept = "application/pdf,.pdf";
    i.addEventListener("change", () => { if (i.files[0]) alElegir(i.files[0]); });
    i.click();
  }
  // Adjuntar el acta a una sesión ya registrada; si trae compromisos y la sesión no tiene, ofrece agregarlos.
  function adjuntarActa(idSeg) {
    const s = S.datos.Seguimientos.find((x) => x.ID_Seguimiento === idSeg);
    if (!s) return;
    elegirPDF(async (archivo) => {
      if (archivo.size > MAX_ACTA) { toast("El acta pesa más de 5 MB. Guarda una versión más liviana del PDF.", true); return; }
      cargando(true, "Leyendo el acta…");
      let res;
      try { res = await ACTAS.leerActa(archivo); } catch (e) { cargando(false); toast("No se pudo leer el PDF: " + e.message, true); return; }
      cargando(false);
      const yaTiene = S.datos.Compromisos.some((c) => c.ID_Seguimiento === idSeg);
      await guardar(async () => {
        await subirActa(idSeg, res.datos, archivo.name);
        await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", idSeg, { Acta_Archivo: archivo.name, ...(s.Fecha_Acta ? {} : { Fecha_Acta: res.fecha }) });
      }, "Acta adjuntada");
      if (!yaTiene && res.compromisos.length) {
        confirmar("Compromisos encontrados en el acta", `Leí ${res.compromisos.length} compromiso(s) en el acta. ¿Los agrego a esta sesión?`, "Agregar compromisos", () => guardar(() => S.api.agregarFilas("Compromisos",
          res.compromisos.map((c, i) => ({ ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", s.ID_Proyecto, i), ID_Proyecto: s.ID_Proyecto, ID_Seguimiento: idSeg,
            Compromiso: c.Compromiso, Responsable: c.Responsable, Fecha_Compromiso: c.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo }))),
        `${res.compromisos.length} compromiso(s) agregados`), "primario");
      }
    });
  }
  function enlazarActas(el) {
    el.querySelectorAll("[data-ver-acta]").forEach((b) => b.addEventListener("click", () => verActa(b.dataset.verActa, b.dataset.nombre)));
    el.querySelectorAll("[data-adjuntar]").forEach((b) => b.addEventListener("click", () => adjuntarActa(b.dataset.adjuntar)));
  }
  function celdaActa(s, puedeAdjuntar) {
    const partes = [];
    if (s.Acta_Archivo) partes.push(`<button class="btn chico" data-ver-acta="${esc(s.ID_Seguimiento)}" data-nombre="${esc(s.Acta_Archivo)}">Ver acta</button>`);
    if (esURL(s.URL_Acta)) partes.push(enlace(s.URL_Acta, "Enlace"));
    if (!s.Acta_Archivo && puedeAdjuntar) partes.push(`<button class="btn chico" data-adjuntar="${esc(s.ID_Seguimiento)}">Adjuntar acta</button>`);
    return partes.join(" ") || `<span class="sub">Sin acta</span>`;
  }

  // ---------- Ventanas modales ----------
  function cerrarModal() { const m = $("#modal"); if (m) m.remove(); }
  function confirmar(titulo, mensaje, textoBoton, accion, estilo = "peligro") {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja chica" role="dialog" aria-modal="true" aria-label="${esc(titulo)}">
      <h2>${esc(titulo)}</h2><p>${esc(mensaje)}</p>
      <div class="acciones derecha"><button class="btn" id="c-no">Cancelar</button><button class="btn ${estilo}" id="c-si">${esc(textoBoton)}</button></div></div>`;
    document.body.appendChild(m);
    $("#c-no").addEventListener("click", cerrarModal);
    $("#c-si").addEventListener("click", accion);
    $("#c-si").focus();
  }

  /* Formulario genérico. campos: [{ k, label, tipo, opciones, req, min, max, def, bloqueado, ayuda, placeholder, ancho } | { seccion, ayuda }]
     extra (opcional): { antes, despues, init(m), recoger(m, fd) → { error } | { datos } } para partes a la medida. */
  function modal(titulo, campos, valores, alGuardar, validarExtra, extra = {}) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    const control = (c) => {
      const v = valores[c.k] ?? c.def ?? "";
      const dis = c.bloqueado ? "disabled" : "";
      if (c.tipo === "checks") {
        const marcados = String(v).split(/[,;]/).map((x) => x.trim());
        return `<div class="checks" data-checks="${c.k}">${c.opciones.map((o) => { const [val, txt] = Array.isArray(o) ? o : [o, o]; return `<label class="check"><input type="checkbox" value="${esc(val)}" ${marcados.includes(String(val)) ? "checked" : ""} ${dis}> ${esc(txt)}</label>`; }).join("") || `<span class="sub">${esc(c.vacio || "Sin opciones.")}</span>`}</div>`;
      }
      if (c.tipo === "select") {
        const ops = c.opciones || [];
        const existe = ops.some((o) => String(Array.isArray(o) ? o[0] : o) === String(v));
        const todas = v !== "" && !existe ? [String(v), ...ops] : ops;   // un valor ya quitado del catálogo se conserva
        return `<select name="${c.k}" ${dis}><option value="">${esc(c.textoVacio || "Seleccione…")}</option>${todas.map((o) => { const [val, txt] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}" ${String(val) === String(v) ? "selected" : ""}>${esc(txt)}</option>`; }).join("")}</select>`;
      }
      if (c.tipo === "textarea") return `<textarea name="${c.k}" rows="3" placeholder="${esc(c.placeholder || "")}" ${dis}>${esc(v)}</textarea>`;
      const dl = c.sugerencias ? `<datalist id="dl-${c.k}">${c.sugerencias.map((o) => `<option value="${esc(o)}">`).join("")}</datalist>` : "";
      return dl + `<input name="${c.k}" ${c.recurso ? "data-recurso" : ""} ${c.sugerencias ? `list="dl-${c.k}" autocomplete="off"` : ""} type="${c.tipo || "text"}" value="${esc(v)}" placeholder="${esc(c.placeholder || "")}" ${c.min !== undefined ? `min="${c.min}"` : ""} ${c.max !== undefined ? `max="${c.max}"` : ""} ${dis}>`;
    };
    const cuerpo = campos.map((c) => (c.html ? `<div class="ancho form-html">${c.html}</div>` : c.seccion
      ? `<div class="form-seccion">${esc(c.seccion)}${c.ayuda ? `<span class="sub"> · ${esc(c.ayuda)}</span>` : ""}</div>`
      : `<label class="${c.tipo === "textarea" || c.tipo === "checks" || c.ancho ? "ancho" : ""}"><span class="etq">${esc(c.label)}</span>${control(c)}${c.ayuda ? `<span class="sub">${esc(c.ayuda)}</span>` : ""}<span class="error-campo" data-e="${c.k}"></span></label>`)).join("");
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="${esc(titulo)}">
      <div class="titulo-fila"><h2>${esc(titulo)}</h2><button class="btn enlace" id="m-cerrar" aria-label="Cerrar">✕</button></div>
      <form id="m-form" novalidate>${extra.antes || ""}<div class="form-grid">${cuerpo}</div>${extra.despues || ""}
      <div class="error-campo" id="m-error" role="alert"></div>
      <div class="acciones derecha">${extra.eliminar ? `<button type="button" class="btn peligro-suave" id="m-eliminar">${esc(extra.eliminar.texto)}</button>` : ""}<button type="button" class="btn" id="m-cancelar">Cancelar</button><button type="submit" class="btn primario">Guardar</button></div></form></div>`;
    document.body.appendChild(m);
    $("#m-cerrar").addEventListener("click", cerrarModal);
    $("#m-cancelar").addEventListener("click", cerrarModal);
    if (extra.eliminar) $("#m-eliminar").addEventListener("click", () => confirmar(extra.eliminar.titulo || extra.eliminar.texto, extra.eliminar.mensaje, "Eliminar", extra.eliminar.accion));
    m.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarModal(); });
    m.querySelectorAll("input,select,textarea").forEach((i) => i.addEventListener("input", () => {
      const e = i.name && m.querySelector(`[data-e="${i.name}"]`); if (e) e.textContent = ""; $("#m-error").textContent = "";
    }));
    if (extra.init) extra.init(m);
    $("#m-form").addEventListener("submit", (ev) => {
      ev.preventDefault();
      const fd = {};
      campos.filter((c) => !c.seccion && !c.html).forEach((c) => {
        if (c.tipo === "checks") { fd[c.k] = [...m.querySelectorAll(`[data-checks="${c.k}"] input:checked`)].map((i) => i.value).join(", "); return; }
        let v = m.querySelector(`[name="${c.k}"]`).value.trim();
        if (c.tipo === "number" && v !== "") v = Number(v);
        fd[c.k] = v;
      });
      let ok = true;
      const err = (k, msg) => { ok = false; const e = m.querySelector(`[data-e="${k}"]`); if (e) e.textContent = msg; };
      campos.filter((c) => !c.seccion && !c.html).forEach((c) => {
        const v = fd[c.k];
        if (c.tipo === "number" && v !== "" && (Number.isNaN(v) || (c.min !== undefined && v < c.min) || (c.max !== undefined && v > c.max)))
          err(c.k, c.max !== undefined ? `Debe estar entre ${c.min} y ${c.max}.` : `Debe ser mayor o igual a ${c.min}.`);
        if (c.tipo === "email" && v && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) err(c.k, "Correo no válido.");
        if (c.tipo === "url" && v && !esURL(v)) err(c.k, "Escribe la dirección completa, empezando por https://");
      });
      let extraDatos = null;
      if (ok && validarExtra) { const msg = validarExtra(fd); if (msg) { ok = false; $("#m-error").textContent = msg; } }
      if (ok && extra.recoger) { const r = extra.recoger(m, fd); if (r.error) { ok = false; $("#m-error").textContent = r.error; } else extraDatos = r.datos; }
      if (!ok) { const primero = m.querySelector(".error-campo:not(:empty)"); if (primero) primero.scrollIntoView({ block: "center" }); return; }
      alGuardar(fd, extraDatos);
    });
    const primero = m.querySelector("input:not([disabled]):not([type=checkbox]),select:not([disabled]),textarea");
    if (primero) primero.focus();
  }

  // ---------- Hoja de vida del proyecto (PDF con marca FSFB) ----------
  function hojaDeVida(p) {
    const pid = p.ID_Proyecto;
    const segs = R.seguimientosDe(pid, S.datos.Seguimientos);
    const hitos = S.datos.Hitos.filter((h) => h.ID_Proyecto === pid).sort((a, b) => String(a.Fecha_Plan).localeCompare(String(b.Fecha_Plan)));
    const riesgos = S.datos.Riesgos.filter((r) => r.ID_Proyecto === pid);
    const comps = S.datos.Compromisos.filter((c) => c.ID_Proyecto === pid).sort((a, b) => String(a.Fecha_Compromiso || "9999").localeCompare(String(b.Fecha_Compromiso || "9999")));
    const coms = (S.datos.Comentarios || []).filter((x) => x.ID_Proyecto === pid);
    const stks = stakeholdersDe(pid);
    const provs = idsLista(p.Proveedores).map(proveedor).filter(Boolean);
    const tks = ticketsDe(pid);
    const e = R.estadoSeguimiento(p, S.datos.Seguimientos);
    const t = (v) => esc(v === 0 ? "0" : v || "—");
    const fila = (k, v) => `<tr><th>${k}</th><td>${v}</td></tr>`;
    const tabla = (cab, filas, vacioTxt) => filas.length ? `<table class="hv-t"><thead><tr>${cab.map((c) => `<th>${c}</th>`).join("")}</tr></thead><tbody>${filas.join("")}</tbody></table>` : `<p class="hv-vacio">${vacioTxt}</p>`;
    // Cronología: todo lo que ha pasado, en orden.
    const eventos = [];
    const ev = (f, tipo, txt) => { if (f) eventos.push({ f: String(f).slice(0, 16), tipo, txt }); };
    ev(p.Fecha_Inicio, "Inicio", `Inicio del proyecto · PM ${nombreUsuario(p.PM)}`);
    segs.forEach((x) => ev(x.Fecha_Corte, "Seguimiento", `Avance ${pct(x.Avance_Real)} · semáforo ${x.Semaforo || "—"}${x.Logros ? ` · ${x.Logros}` : ""}${x.Bloqueos ? ` · Bloqueo: ${x.Bloqueos}` : ""}`));
    hitos.forEach((h) => { if (h.Fecha_Real) ev(h.Fecha_Real, "Hito", `Cumplido: ${h.Hito}`); });
    comps.forEach((c) => { if (R.cerrado(c) && c.Fecha_Cierre) ev(c.Fecha_Cierre, "Compromiso", `Cerrado: ${c.Compromiso}`); });
    coms.forEach((x) => { const c = comps.find((y) => y.ID_Compromiso === x.ID_Compromiso); ev(x.Fecha_Hora, "Comentario", `${nombreUsuario(x.Autor)} en «${c ? c.Compromiso : x.ID_Compromiso}»: ${resumenCom(x)}`); });
    tks.forEach((k) => ev(k.Fecha_Registro, "Ticket", `Ticket ${k.Numero || ""} registrado: ${k.Titulo || ""} (${k.Estado || "Abierto"})`));
    eventos.sort((a, b) => a.f.localeCompare(b.f));
    const fechaEv = (f) => (f.length > 10 ? `${fecha(f.slice(0, 10))} ${f.slice(11, 16)}` : fecha(f));
    const abiertos = comps.filter((c) => !R.cerrado(c)).length;
    const doc = `
      ${cabDocumento("Hoja de vida del proyecto", p.Nombre, `${esc(pid)}${p.Codigo_Almera ? ` · Almera ${esc(p.Codigo_Almera)}` : ""} · ${esc(p.Cliente_Area || "")}`)}
      <section class="hv-kpis">
        <div><span>Estado</span><b>${t(p.Estado)}</b></div><div><span>Fase</span><b>${t(p.Fase)}</b></div>
        <div><span>Semáforo</span><b>${t(p.Semaforo)}</b></div><div><span>Avance real / plan</span><b>${pct(p.Avance_Real)} / ${pct(p.Avance_Planeado)}</b></div>
        <div><span>Presupuesto ejecutado</span><b>${Number(p.Presupuesto) ? Math.round((Number(p.Ejecutado) || 0) / Number(p.Presupuesto) * 100) + "%" : "—"}</b></div>
        <div><span>Compromisos abiertos</span><b>${abiertos} de ${comps.length}</b></div>
      </section>
      <h2>1. Datos generales</h2>
      <table class="hv-ficha">${[
        fila("PM responsable", t(nombreUsuario(p.PM))), fila("Metodología", t(p.Metodologia)), fila("Prioridad", t(p.Prioridad)),
        fila("Fecha de inicio", fecha(p.Fecha_Inicio)), fila("Fin planeado", fecha(p.Fecha_Fin_Plan)), fila("Frecuencia de seguimiento", t(p.Frecuencia_Seguimiento)),
        fila("Estado del seguimiento", `${t(e.estado)} · próximo ${fecha(e.proximo)}`), fila("Presupuesto", cop(p.Presupuesto)), fila("Ejecutado", cop(p.Ejecutado)),
        fila("Repositorio", t(p.URL_Repositorio)), fila("Documentos", t(p.URL_Documentos)), fila("Comentario de estado", t(p.Comentario_Estado)),
      ].join("")}</table>
      <h2>2. Stakeholders y proveedores</h2>
      ${tabla(["Rol", "Nombre", "Cargo / área", "Empresa", "Correo"], stks.map((x) => `<tr><td>${t(x.Rol)}</td><td>${t(x.Nombre)}</td><td>${t([x.Cargo, x.Area].filter(Boolean).join(" · "))}</td><td>${x.ID_Proveedor ? t((proveedor(x.ID_Proveedor) || {}).Nombre) : "Interno"}</td><td>${t(x.Correo)}</td></tr>`), "Sin stakeholders registrados.")}
      <p class="hv-p"><b>Proveedores:</b> ${provs.length ? provs.map((x) => `${esc(x.Nombre)}${x.Servicio ? ` (${esc(x.Servicio)})` : ""}`).join(" · ") : "ninguno"}</p>
      <h2>3. Cronología del proyecto</h2>
      ${tabla(["Fecha", "Tipo", "Qué pasó"], eventos.map((x) => `<tr><td class="hv-nw">${fechaEv(x.f)}</td><td>${x.tipo}</td><td>${esc(x.txt)}</td></tr>`), "Sin eventos registrados.")}
      <h2>4. Sesiones de seguimiento</h2>
      ${tabla(["Fecha", "Avance", "Semáforo", "Logros", "Próximos pasos", "Bloqueos", "Acta"], segs.slice().reverse().map((x) => `<tr><td class="hv-nw">${fecha(x.Fecha_Corte)}</td><td>${pct(x.Avance_Real)}</td><td>${t(x.Semaforo)}</td><td>${t(x.Logros)}</td><td>${t(x.Proximos_Pasos)}</td><td>${t(x.Bloqueos)}</td><td>${x.Fecha_Acta ? fecha(x.Fecha_Acta) : "—"}${x.Acta_Archivo ? " · PDF" : ""}</td></tr>`), "Sin seguimientos.")}
      <h2>5. Compromisos</h2>
      ${tabla(["Compromiso", "Origen", "Responsable", "Fecha límite", "Estado", "Seguimiento"], comps.map((c) => {
        const hilo = comentariosDe(c.ID_Compromiso);
        return `<tr><td>${t(c.Compromiso)}</td><td>${esc(origen(c))}</td><td>${t(c.Responsable)}</td><td class="hv-nw">${fecha(c.Fecha_Compromiso)}</td>
          <td>${esc(R.estadoCompromiso(c))}${R.cerrado(c) && c.Fecha_Cierre ? `<br><small>cerrado ${fecha(c.Fecha_Cierre)}</small>` : ""}</td>
          <td>${hilo.length ? hilo.map((x) => `<div class="hv-com"><small>${fechaEv(String(x.Fecha_Hora))} · ${esc(nombreUsuario(x.Autor))}</small><br>${esc(resumenCom(x))}</div>`).join("") : "—"}</td></tr>`;
      }), "Sin compromisos.")}
      <h2>6. Hitos</h2>
      ${tabla(["Hito", "Fecha plan", "Fecha real", "Estado"], hitos.map((h) => `<tr><td>${t(h.Hito)}</td><td>${fecha(h.Fecha_Plan)}</td><td>${h.Fecha_Real ? fecha(h.Fecha_Real) : "—"}</td><td>${t(h.Estado)}</td></tr>`), "Sin hitos.")}
      <h2>7. Riesgos</h2>
      ${tabla(["No.", "Riesgo", "Responsable", "Inherente", "Planes", "Residual", "Origen"], riesgos.slice().sort((a, b) => numRiesgo(a) - numRiesgo(b)).map((r) => `<tr><td>${numRiesgo(r)}</td><td><small>${t(r.Tipo)} · ${t(r.Estado)}</small><br>${t(r.Descripcion)}</td><td>${t(r.Responsable)}</td>
        <td class="hv-nw">${califTxt(r.Probabilidad_Inherente, r.Impacto_Inherente)}</td><td>${r.Plan_Mitigacion ? `<b>Mitigación:</b> ${t(r.Plan_Mitigacion)}` : ""}${r.Plan_Contingencia ? `<br><b>Contingencia:</b> ${t(r.Plan_Contingencia)}` : ""}</td>
        <td class="hv-nw">${califTxt(r.Probabilidad_Residual, r.Impacto_Residual)}</td><td>${esc(origenRiesgo(r))}</td></tr>`), "Sin riesgos.")}
      <h2>8. Control de cambios</h2>
      ${(() => { const lb = lineaBase(p); return lb.tiene ? `<p class="hv-p"><b>Línea base:</b> fin ${fecha(lb.fb)} → actual ${fecha(p.Fecha_Fin_Plan)}${lb.dias !== null ? ` (${lb.dias > 0 ? "+" : ""}${lb.dias} días)` : ""} · presupuesto ${lb.pb !== null ? cop(lb.pb) : "—"} → actual ${cop(p.Presupuesto)}${lb.pctCosto !== null ? ` (${lb.pctCosto > 0 ? "+" : ""}${lb.pctCosto}%)` : ""}</p>` : ""; })()}
      ${tabla(["Fecha", "Tipo", "Cambio", "Impacto", "Estado"], cambiosDe(pid).map((c) => `<tr><td class="hv-nw">${fecha(c.Fecha)}</td><td>${t(c.Tipo)}</td><td>${t(c.Descripcion)}</td><td>${t(c.Impacto)}${c.Nueva_Fecha_Fin ? `<br><small>Nueva fecha fin ${fecha(c.Nueva_Fecha_Fin)}</small>` : ""}${num0(c.Nuevo_Presupuesto) !== null ? `<br><small>Nuevo presupuesto ${cop(c.Nuevo_Presupuesto)}</small>` : ""}</td><td>${t(c.Estado)}${c.Fecha_Decision ? `<br><small>${fecha(c.Fecha_Decision)}</small>` : ""}</td></tr>`), "Sin cambios registrados.")}
      <h2>9. Dependencias</h2>
      ${tabla(["Relación", "Proyecto", "Tipo", "Alerta"], [...depsDe(pid).map((d) => `<tr><td>Depende de</td><td>${esc((proyecto(d.Depende_De) || {}).Nombre || d.Depende_De)}</td><td>${t(d.Tipo)}</td><td>${esc(conflictoDep(d) || "—")}</td></tr>`),
        ...bloqueaA(pid).map((d) => `<tr><td>Bloquea a</td><td>${esc((proyecto(d.ID_Proyecto) || {}).Nombre || d.ID_Proyecto)}</td><td>${t(d.Tipo)}</td><td>${esc(conflictoDep(d) || "—")}</td></tr>`)], "Sin dependencias.")}
      <h2>10. Lecciones aprendidas</h2>
      ${tabla(["Tipo", "Categoría", "Qué pasó", "Lección", "Recomendación"], leccionesDe(pid).map((l) => `<tr><td>${t(l.Tipo)}</td><td>${t(l.Categoria)}</td><td>${t(l.Situacion)}</td><td>${t(l.Leccion)}</td><td>${t(l.Recomendacion)}</td></tr>`), "Sin lecciones registradas.")}
      <h2>11. Tickets del helpdesk</h2>
      ${tabla(["N.º", "Título", "Estado", "Prioridad", "Registrado"], tks.map((k) => `<tr><td>${t(k.Numero)}</td><td>${t(k.Titulo)}</td><td>${t(k.Estado)}</td><td>${t(k.Prioridad)}</td><td>${fecha(k.Fecha_Registro)}</td></tr>`), "Sin tickets.")}
      <footer class="hv-pie">Fundación Santa Fe de Bogotá · Oficina de Proyectos · Hoja de vida ${esc(pid)} · Documento generado desde el Portafolio PMO</footer>`;
    mostrarDocumento(`Hoja de vida · ${p.Nombre}`, doc, `Hoja de vida ${pid} - ${p.Nombre}`);
  }
  // Vista de documento imprimible (hoja de vida, informe semanal): se guarda como PDF desde Imprimir.
  function mostrarDocumento(titulo, html, nombreArchivo) {
    cerrarModal();
    const o = document.createElement("div");
    o.className = "hv-overlay";
    o.innerHTML = `<div class="hv-barra no-print"><b>${esc(titulo)}</b><div class="acciones">
      <button class="btn primario" id="hv-imprimir">🖨 Imprimir / Guardar como PDF</button><button class="btn" id="hv-cerrar">Cerrar</button></div>
      <div class="sub">En la ventana de impresión elige «Guardar como PDF» como destino.</div></div>
      <div class="hv-hoja"><article class="hv">${html}</article></div>`;
    document.body.appendChild(o);
    document.body.classList.add("con-hv");
    const cerrar = () => { o.remove(); document.body.classList.remove("con-hv"); };
    o.querySelector("#hv-cerrar").addEventListener("click", cerrar);
    o.querySelector("#hv-imprimir").addEventListener("click", () => {
      const tituloAnt = document.title;
      document.title = nombreArchivo;
      window.print();
      setTimeout(() => { document.title = tituloAnt; }, 1000);
    });
  }
  const cabDocumento = (tipo, titulo, sub) => `<div class="hv-marca" aria-hidden="true"><span>FSFB</span></div>
      <header class="hv-cab"><div class="hv-logo">Fundación<br>Santa Fe de Bogotá</div>
        <div class="hv-cab-t"><div class="hv-tipo">${tipo}</div><h1>${esc(titulo)}</h1><div class="hv-sub">${sub}</div></div>
        <div class="hv-gen">Generado el ${fecha(R.hoyISO())}<br>por ${esc(S.usuario.Nombre || S.usuario.Correo)}</div></header>`;

  // ---------- Control de cambios y línea base ----------
  const cambiosDe = (pid) => (S.datos.Cambios || []).filter((c) => c.ID_Proyecto === pid).sort((a, b) => String(b.Fecha).localeCompare(String(a.Fecha)) || String(b.ID_Cambio).localeCompare(String(a.ID_Cambio), "es", { numeric: true }));
  function lineaBase(p) {
    const fb = p.Fecha_Fin_Base || "", pb = num0(p.Presupuesto_Base);
    return {
      tiene: !!(fb || pb !== null), fb, pb,
      dias: fb && p.Fecha_Fin_Plan ? R.difDias(p.Fecha_Fin_Plan, fb) : null,
      pctCosto: pb ? Math.round(((Number(p.Presupuesto) || 0) - pb) / pb * 100) : null,
    };
  }
  function seccionCambios(p, puedeSolicitar, puedeDecidir) {
    const lb = lineaBase(p);
    const lista = cambiosDe(p.ID_Proyecto);
    const pend = lista.filter((c) => c.Estado === "Solicitado").length;
    const desvio = (v, u) => (v === null ? "—" : `<span class="${v > 0 ? "baja" : v < 0 ? "sube" : ""}">${v > 0 ? "+" : ""}${v}${u}</span>`);
    return `<div class="card"><div class="titulo-fila"><h2>Línea base</h2>${puedeDecidir ? `<button class="btn chico" id="b-lb">${lb.tiene ? "Volver a fijar línea base" : "Fijar línea base"}</button>` : ""}</div>
        ${lb.tiene ? `<div class="datos">
          <div class="dato"><div class="dato-k">Fin planeado original</div><div class="dato-v">${fecha(lb.fb)}</div></div>
          <div class="dato"><div class="dato-k">Fin planeado actual</div><div class="dato-v">${fecha(p.Fecha_Fin_Plan)} · ${desvio(lb.dias, " días")}</div></div>
          <div class="dato"><div class="dato-k">Presupuesto original</div><div class="dato-v">${lb.pb !== null ? cop(lb.pb) : "—"}</div></div>
          <div class="dato"><div class="dato-k">Presupuesto actual</div><div class="dato-v">${cop(p.Presupuesto)} · ${desvio(lb.pctCosto, "%")}</div></div></div>`
        : `<p class="sub">Aún no hay línea base. Se fija sola con el primer cambio aprobado (toma la fecha fin y el presupuesto de ese momento)${puedeDecidir ? ", o puedes fijarla ahora" : ""}.</p>`}
      </div>
      <div class="card"><div class="titulo-fila"><h2>Control de cambios ${pend ? `<span class="contador rojo">${pend} por decidir</span>` : ""}</h2>${puedeSolicitar ? `<button class="btn primario" id="b-cambio">+ Solicitar cambio</button>` : ""}</div>
        <p class="sub">Cambios de alcance, tiempo, costo o recursos. El PM los solicita y la PMO o el Admin los aprueban; al aprobar, la fecha fin y el presupuesto nuevos se aplican al proyecto.</p>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Cambio</th><th>Impacto</th><th>Estado</th><th></th></tr></thead><tbody>
          ${lista.map((c) => `<tr><td class="nowrap">${fecha(c.Fecha)}<div class="sub">${esc(nombreUsuario(c.Solicitado_Por))}</div></td><td>${esc(c.Tipo)}</td>
            <td><b>${esc(c.Descripcion)}</b>${c.Justificacion ? `<div class="sub">Por qué: ${esc(c.Justificacion)}</div>` : ""}</td>
            <td>${esc(c.Impacto || "")}${c.Nueva_Fecha_Fin ? `<div class="sub">Nueva fecha fin: ${fecha(c.Nueva_Fecha_Fin)}</div>` : ""}${num0(c.Nuevo_Presupuesto) !== null ? `<div class="sub">Nuevo presupuesto: ${cop(c.Nuevo_Presupuesto)}</div>` : ""}</td>
            <td>${pill(c.Estado || "Solicitado")}${c.Fecha_Decision ? `<div class="sub">${fecha(c.Fecha_Decision)} · ${esc(nombreUsuario(c.Decidido_Por))}</div>` : ""}${c.Comentario_Decision ? `<div class="sub">${esc(c.Comentario_Decision)}</div>` : ""}</td>
            <td class="derecha nowrap">${c.Estado === "Solicitado" && puedeDecidir ? `<button class="btn chico primario" data-decidir="${esc(c.ID_Cambio)}">Decidir</button> ` : ""}${c.Estado === "Solicitado" && puedeSolicitar ? `<button class="btn chico" data-cambio="${esc(c.ID_Cambio)}">Editar</button>` : ""}</td></tr>`).join("")}
        </tbody></table></div>` : vacio("Sin cambios registrados.")}</div>`;
  }
  function formCambio(p, c) {
    const nuevo = !c;
    const campos = [
      { k: "Tipo", label: "Tipo de cambio", tipo: "select", opciones: S.cat.Tipo_Cambio, def: "Alcance" },
      { k: "Fecha", label: "Fecha de la solicitud", tipo: "date", def: R.hoyISO() },
      { k: "Descripcion", label: "¿Qué cambia?", tipo: "textarea", placeholder: "Describe el cambio" },
      { k: "Justificacion", label: "¿Por qué?", tipo: "textarea", placeholder: "Motivo o necesidad" },
      { k: "Impacto", label: "Impacto", tipo: "textarea", placeholder: "Efecto en alcance, calidad, equipo, riesgos…" },
      { seccion: "Si cambia la fecha o el costo", ayuda: "se aplican al proyecto cuando se aprueba" },
      { k: "Nueva_Fecha_Fin", label: "Nueva fecha fin", tipo: "date", ayuda: `Hoy: ${fecha(p.Fecha_Fin_Plan)}` },
      { k: "Nuevo_Presupuesto", label: "Nuevo presupuesto (COP)", tipo: "number", min: 0, ayuda: `Hoy: ${cop(p.Presupuesto)}` },
    ];
    modal(nuevo ? `Solicitar cambio · ${p.Nombre}` : "Editar solicitud de cambio", campos, nuevo ? {} : c, (fd) => {
      if (nuevo) {
        const id = R.siguienteIdHijo("CAM", S.datos.Cambios || [], "ID_Cambio", p.ID_Proyecto);
        guardar(() => S.api.agregarFila("Cambios", { ID_Cambio: id, ID_Proyecto: p.ID_Proyecto, ...fd, Estado: "Solicitado", Solicitado_Por: S.usuario.Correo }), "Cambio solicitado");
      } else guardar(() => S.api.actualizarPorId("Cambios", "ID_Cambio", c.ID_Cambio, fd), "Solicitud actualizada");
    }, null, nuevo ? {} : { eliminar: { texto: "Eliminar solicitud", mensaje: "Se borrará esta solicitud de cambio.",
      accion: () => guardar(() => S.api.eliminarFilas("Cambios", "ID_Cambio", [c.ID_Cambio]), "Solicitud eliminada") } });
  }
  function decidirCambio(p, c) {
    const campos = [
      { k: "Decision", label: "Decisión", tipo: "select", opciones: ["Aprobado", "Rechazado"], def: "Aprobado" },
      { k: "Comentario_Decision", label: "Comentario (opcional)", tipo: "textarea", placeholder: "Condiciones, instancia que aprobó…" },
    ];
    const efectos = [c.Nueva_Fecha_Fin ? `fecha fin ${fecha(p.Fecha_Fin_Plan)} → ${fecha(c.Nueva_Fecha_Fin)}` : "", num0(c.Nuevo_Presupuesto) !== null ? `presupuesto ${cop(p.Presupuesto)} → ${cop(c.Nuevo_Presupuesto)}` : ""].filter(Boolean);
    modal(`Decidir cambio · ${c.Tipo}`, campos, {}, (fd) => guardar(async () => {
      const aprobado = fd.Decision !== "Rechazado";
      if (aprobado) {
        const lb = lineaBase(p);
        const cambiosP = {};
        if (!lb.tiene) { cambiosP.Fecha_Fin_Base = p.Fecha_Fin_Plan; cambiosP.Presupuesto_Base = p.Presupuesto; }
        if (c.Nueva_Fecha_Fin) cambiosP.Fecha_Fin_Plan = c.Nueva_Fecha_Fin;
        if (num0(c.Nuevo_Presupuesto) !== null) cambiosP.Presupuesto = Number(c.Nuevo_Presupuesto);
        if (Object.keys(cambiosP).length) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...cambiosP, ...sello() });
      }
      await S.api.actualizarPorId("Cambios", "ID_Cambio", c.ID_Cambio, { Estado: aprobado ? "Aprobado" : "Rechazado", Decidido_Por: S.usuario.Correo, Fecha_Decision: R.hoyISO(), Comentario_Decision: fd.Comentario_Decision });
    }, fd.Decision === "Rechazado" ? "Cambio rechazado" : "Cambio aprobado y aplicado"), null,
    { antes: `<div class="resumen-filtro"><b>${esc(c.Descripcion)}</b>${efectos.length ? `<br>Si se aprueba: ${esc(efectos.join(" · "))}` : "<br>No cambia fechas ni presupuesto."}</div>` });
  }
  function fijarLineaBase(p) {
    confirmar("Fijar línea base", `La línea base quedará con fin ${fecha(p.Fecha_Fin_Plan)} y presupuesto ${cop(p.Presupuesto)}. Los cambios futuros se medirán contra estos valores.`, "Fijar",
      () => guardar(() => S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { Fecha_Fin_Base: p.Fecha_Fin_Plan, Presupuesto_Base: p.Presupuesto }), "Línea base fijada"), "primario");
  }

  // ---------- Lecciones aprendidas ----------
  const leccionesDe = (pid) => (S.datos.Lecciones || []).filter((l) => l.ID_Proyecto === pid);
  const tarjetaLeccion = (l, conProyecto, puede) => `<div class="leccion ${l.Tipo === "Positiva" ? "positiva" : "mejorar"}">
      <div class="titulo-fila"><div>${pill(l.Tipo || "A mejorar")} <b>${esc(l.Categoria || "General")}</b>${conProyecto ? ` · <button class="btn enlace" data-pid="${esc(l.ID_Proyecto)}">${esc((proyecto(l.ID_Proyecto) || {}).Nombre || l.ID_Proyecto)}</button>` : ""}</div>
        <div><span class="sub">${fecha(l.Fecha)} · ${esc(nombreUsuario(l.Registrado_Por))}</span>${puede ? ` <button class="btn chico" data-leccion="${esc(l.ID_Leccion)}">Editar</button>` : ""}</div></div>
      ${l.Situacion ? `<div><span class="sub">Qué pasó:</span> ${esc(l.Situacion)}</div>` : ""}
      <div><span class="sub">Lección:</span> <b>${esc(l.Leccion)}</b></div>
      ${l.Recomendacion ? `<div><span class="sub">Recomendación:</span> ${esc(l.Recomendacion)}</div>` : ""}</div>`;
  function seccionLecciones(p, puede) {
    const lista = leccionesDe(p.ID_Proyecto).sort((a, b) => String(b.Fecha).localeCompare(String(a.Fecha)));
    return `<div class="card"><div class="titulo-fila"><h2>Lecciones aprendidas</h2>${puede ? `<button class="btn primario" id="b-leccion">+ Lección</button>` : ""}</div>
      <p class="sub">Qué funcionó y qué haríamos distinto. Se pueden consultar desde el menú «Lecciones aprendidas» para proyectos parecidos.</p>
      ${lista.length ? lista.map((l) => tarjetaLeccion(l, false, puede)).join("") : vacio(`Aún no hay lecciones.${puede ? " Regístralas durante el proyecto y, sobre todo, al cerrarlo." : ""}`)}</div>`;
  }
  function formLeccion(p, l) {
    const nuevo = !l;
    const campos = [
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: ["Positiva", "A mejorar"], def: "A mejorar" },
      { k: "Categoria", label: "Categoría", tipo: "select", opciones: S.cat.Categoria_Leccion },
      { k: "Situacion", label: "¿Qué pasó?", tipo: "textarea", placeholder: "La situación o el hecho" },
      { k: "Leccion", label: "Lección aprendida", tipo: "textarea", placeholder: "Qué aprendimos" },
      { k: "Recomendacion", label: "Recomendación para otros proyectos", tipo: "textarea" },
      { k: "Fecha", label: "Fecha", tipo: "date", def: R.hoyISO() },
    ];
    modal(nuevo ? `Nueva lección · ${p.Nombre}` : "Editar lección", campos, nuevo ? {} : l, (fd) => {
      if (nuevo) {
        const id = R.siguienteIdHijo("LEC", S.datos.Lecciones || [], "ID_Leccion", p.ID_Proyecto);
        guardar(() => S.api.agregarFila("Lecciones", { ID_Leccion: id, ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo }), "Lección registrada");
      } else guardar(() => S.api.actualizarPorId("Lecciones", "ID_Leccion", l.ID_Leccion, fd), "Lección actualizada");
    }, null, nuevo ? {} : { eliminar: { texto: "Eliminar lección", mensaje: "Se borrará esta lección.",
      accion: () => guardar(() => S.api.eliminarFilas("Lecciones", "ID_Leccion", [l.ID_Leccion]), "Lección eliminada") } });
  }
  function vLecciones(el) {
    const ids = new Set(filtrados().map((p) => p.ID_Proyecto));
    const F = S.lecF = S.lecF || { texto: "", cat: "", tipo: "" };
    const lista = (S.datos.Lecciones || []).filter((l) => ids.has(l.ID_Proyecto) && (!F.cat || l.Categoria === F.cat) && (!F.tipo || l.Tipo === F.tipo) &&
      (!F.texto || `${l.Situacion} ${l.Leccion} ${l.Recomendacion} ${(proyecto(l.ID_Proyecto) || {}).Nombre}`.toLowerCase().includes(F.texto.toLowerCase())))
      .sort((a, b) => String(b.Fecha).localeCompare(String(a.Fecha)));
    el.innerHTML = `<h1>Lecciones aprendidas</h1>${barraFiltros()}
      <div class="filtros"><label class="filtro crece"><span>Buscar</span><input id="lec-txt" type="search" value="${esc(F.texto)}" placeholder="Palabra clave: proveedor, integración, pruebas…"></label>
        <label class="filtro"><span>Categoría</span><select id="lec-cat"><option value="">Todas</option>${(S.cat.Categoria_Leccion || []).map((c) => `<option ${c === F.cat ? "selected" : ""}>${esc(c)}</option>`).join("")}</select></label>
        <label class="filtro"><span>Tipo</span><select id="lec-tipo"><option value="">Todas</option>${["Positiva", "A mejorar"].map((c) => `<option ${c === F.tipo ? "selected" : ""}>${c}</option>`).join("")}</select></label></div>
      <div class="card"><h2>${lista.length} lección(es)</h2>${lista.length ? lista.map((l) => tarjetaLeccion(l, true, false)).join("") : vacio("No hay lecciones con estos filtros. Se registran en la pestaña «Lecciones» de cada proyecto.")}</div>`;
    enlazarFiltros();
    $("#lec-txt").addEventListener("change", (e) => { F.texto = e.target.value; render(); });
    $("#lec-cat").addEventListener("change", (e) => { F.cat = e.target.value; render(); });
    $("#lec-tipo").addEventListener("change", (e) => { F.tipo = e.target.value; render(); });
  }

  // ---------- Matriz RACI ----------
  const raciDe = (pid) => (S.datos.RACI || []).filter((r) => r.ID_Proyecto === pid);
  const asignRaci = (r) => { try { return r.Asignaciones ? JSON.parse(r.Asignaciones) : {}; } catch (e) { return {}; } };
  function seccionRACI(p, puede) {
    const filas = raciDe(p.ID_Proyecto);
    const cols = [{ k: "PM", t: "PM", n: nombreUsuario(p.PM) }, ...stakeholdersDe(p.ID_Proyecto).map((x) => ({ k: x.ID_Stakeholder, t: x.Rol, n: x.Nombre }))];
    const aviso = (a) => { const v = Object.values(a); const nA = v.filter((x) => x === "A").length, nR = v.filter((x) => x === "R").length;
      return nA !== 1 ? (nA ? "Debe haber un solo A" : "Falta el A (quien aprueba)") : !nR ? "Falta al menos un R (quien ejecuta)" : ""; };
    return `<div class="card"><div class="titulo-fila"><h2>Matriz RACI</h2>${puede ? `<div class="acciones"><button class="btn primario" id="raci-guardar" disabled>Guardar matriz</button></div>` : ""}</div>
      <p class="sub"><b>R</b> Responsable (ejecuta) · <b>A</b> Aprueba (uno por entregable) · <b>C</b> Consultado · <b>I</b> Informado. Las columnas son el PM y los stakeholders del proyecto.</p>
      ${filas.length ? `<div class="tabla-scroll"><table class="raci"><thead><tr><th>Entregable</th>${cols.map((c) => `<th title="${esc(c.n)}">${esc(c.t)}<div class="raci-n">${esc(c.n)}</div></th>`).join("")}${puede ? "<th></th>" : ""}</tr></thead><tbody>
        ${filas.map((r) => { const a = asignRaci(r); const w = aviso(a); return `<tr data-raci="${esc(r.ID_RACI)}"><td><b>${esc(r.Entregable)}</b>${w ? `<div class="sub baja">⚠ ${w}</div>` : ""}</td>
          ${cols.map((c) => `<td class="centro">${puede ? `<select class="raci-sel v-${esc(a[c.k] || "")}" data-col="${esc(c.k)}" aria-label="${esc(r.Entregable)} · ${esc(c.t)}">${["", "R", "A", "C", "I"].map((v) => `<option ${a[c.k] === v || (!a[c.k] && !v) ? "selected" : ""}>${v}</option>`).join("")}</select>` : `<b class="raci-v v-${esc(a[c.k] || "")}">${esc(a[c.k] || "")}</b>`}</td>`).join("")}
          ${puede ? `<td><button class="btn chico" data-raci-del="${esc(r.ID_RACI)}" title="Quitar entregable">✕</button></td>` : ""}</tr>`; }).join("")}</tbody></table></div>` : vacio("Aún no hay entregables en la matriz.")}
      ${puede ? `<form class="cat-agregar" id="raci-nuevo" novalidate><input name="ent" placeholder="Nuevo entregable (ej.: Documento de requerimientos)" aria-label="Nuevo entregable"><button class="btn" type="submit">+ Entregable</button></form>
        ${cols.length === 1 ? `<p class="sub">Agrega stakeholders al proyecto (Resumen o «Editar proyecto») para tener más columnas.</p>` : ""}` : ""}</div>`;
  }
  function enlazarRACI(el, p) {
    const btn = el.querySelector("#raci-guardar");
    el.querySelectorAll(".raci-sel").forEach((s) => s.addEventListener("change", () => { s.className = `raci-sel v-${s.value}`; if (btn) { btn.disabled = false; btn.textContent = "Guardar matriz •"; } }));
    if (btn) btn.addEventListener("click", () => {
      const cambios = [...el.querySelectorAll("tr[data-raci]")].map((tr) => {
        const a = {}; tr.querySelectorAll(".raci-sel").forEach((s) => { if (s.value) a[s.dataset.col] = s.value; });
        return { id: tr.dataset.raci, cambios: { Asignaciones: JSON.stringify(a) } };
      });
      guardar(() => S.api.actualizarVarios("RACI", "ID_RACI", cambios), "Matriz RACI guardada");
    });
    const f = el.querySelector("#raci-nuevo");
    if (f) f.addEventListener("submit", (ev) => {
      ev.preventDefault();
      const ent = f.ent.value.trim();
      if (!ent) return;
      guardar(() => S.api.agregarFila("RACI", { ID_RACI: R.siguienteIdHijo("RACI", S.datos.RACI || [], "ID_RACI", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, Entregable: ent, Asignaciones: "{}" }), "Entregable agregado");
    });
    el.querySelectorAll("[data-raci-del]").forEach((b) => b.addEventListener("click", () => confirmar("Quitar entregable", "Se quitará esta fila de la matriz RACI.", "Quitar",
      () => guardar(() => S.api.eliminarFilas("RACI", "ID_RACI", [b.dataset.raciDel]), "Entregable quitado"))));
  }

  // ---------- Dependencias entre proyectos ----------
  const TIPOS_DEP = ["Fin → Inicio", "Inicio → Inicio", "Fin → Fin"];
  const depsDe = (pid) => (S.datos.Dependencias || []).filter((d) => d.ID_Proyecto === pid);
  const bloqueaA = (pid) => (S.datos.Dependencias || []).filter((d) => d.Depende_De === pid);
  // ¿La dependencia está en riesgo? Devuelve el motivo o "".
  function conflictoDep(d) {
    const suc = proyecto(d.ID_Proyecto), pre = proyecto(d.Depende_De);
    if (!suc || !pre || pre.Estado === "Cerrado") return "";
    const motivos = [];
    const tipo = d.Tipo || TIPOS_DEP[0];
    if (tipo === "Fin → Inicio" && pre.Fecha_Fin_Plan && suc.Fecha_Inicio && pre.Fecha_Fin_Plan > suc.Fecha_Inicio) motivos.push(`«${pre.Nombre}» termina el ${fecha(pre.Fecha_Fin_Plan)}, después del inicio de «${suc.Nombre}» (${fecha(suc.Fecha_Inicio)})`);
    if (tipo === "Fin → Fin" && pre.Fecha_Fin_Plan && suc.Fecha_Fin_Plan && pre.Fecha_Fin_Plan > suc.Fecha_Fin_Plan) motivos.push(`«${pre.Nombre}» termina después que «${suc.Nombre}»`);
    if (tipo === "Inicio → Inicio" && pre.Fecha_Inicio && suc.Fecha_Inicio && pre.Fecha_Inicio > suc.Fecha_Inicio) motivos.push(`«${pre.Nombre}» inicia después que «${suc.Nombre}»`);
    if (pre.Semaforo === "Rojo") motivos.push(`«${pre.Nombre}» está en rojo`);
    return motivos.join(" · ");
  }
  function cardDependencias(p, puede) {
    const deps = depsDe(p.ID_Proyecto), bloq = bloqueaA(p.ID_Proyecto);
    const ids = new Set(visibles().map((x) => x.ID_Proyecto));
    const item = (d, otroId, dir) => { const o = proyecto(otroId) || {}; const c = conflictoDep(d);
      return `<li class="${c ? "dep-alerta" : ""}">${dir} ${ids.has(otroId) ? `<button class="btn enlace" data-pid="${esc(otroId)}">${esc(o.Nombre || otroId)}</button>` : esc(o.Nombre || otroId)}
        <span class="sub">· ${esc(d.Tipo || TIPOS_DEP[0])} · ${esc(o.Estado || "")} · fin ${fecha(o.Fecha_Fin_Plan)}</span>${d.Descripcion ? `<div class="sub">${esc(d.Descripcion)}</div>` : ""}
        ${c ? `<div class="baja">⚠ ${esc(c)}</div>` : ""}${puede ? ` <button class="btn chico" data-dep="${esc(d.ID_Dependencia)}">Editar</button>` : ""}</li>`; };
    return `<div class="card"><div class="titulo-fila"><h2>Dependencias</h2>${puede ? `<button class="btn" id="b-dep">+ Dependencia</button>` : ""}</div>
      ${deps.length || bloq.length ? `<div class="grid2">
        <div><b>Este proyecto depende de</b><ul class="deps">${deps.map((d) => item(d, d.Depende_De, "⬅")).join("") || `<li class="sub">Ninguno</li>`}</ul></div>
        <div><b>Bloquea a</b><ul class="deps">${bloq.map((d) => item(d, d.ID_Proyecto, "➡")).join("") || `<li class="sub">Ninguno</li>`}</ul></div></div>`
        : `<p class="sub">Sin dependencias. Regístralas cuando este proyecto necesite que otro termine (o empiece) primero.</p>`}</div>`;
  }
  function formDependencia(p, d) {
    const nuevo = !d;
    const ya = new Set(depsDe(p.ID_Proyecto).map((x) => x.Depende_De));
    const opciones = S.datos.Proyectos.filter((x) => x.ID_Proyecto !== p.ID_Proyecto && (!ya.has(x.ID_Proyecto) || (d && d.Depende_De === x.ID_Proyecto)))
      .sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es")).map((x) => [x.ID_Proyecto, `${x.Nombre} (${x.ID_Proyecto})`]);
    const campos = [
      { k: "Depende_De", label: "Depende del proyecto", tipo: "select", opciones, ancho: true },
      { k: "Tipo", label: "Tipo de dependencia", tipo: "select", opciones: TIPOS_DEP, def: TIPOS_DEP[0], ayuda: "Fin → Inicio: el otro debe terminar antes de que este empiece." },
      { k: "Descripcion", label: "Qué se necesita del otro proyecto", tipo: "textarea", placeholder: "Ej.: la integración con el ERP debe estar en producción" },
    ];
    modal(nuevo ? `Nueva dependencia · ${p.Nombre}` : "Editar dependencia", campos, nuevo ? {} : d, (fd) => {
      if (nuevo) {
        const id = R.siguienteIdHijo("DEP", S.datos.Dependencias || [], "ID_Dependencia", p.ID_Proyecto);
        guardar(() => S.api.agregarFila("Dependencias", { ID_Dependencia: id, ID_Proyecto: p.ID_Proyecto, ...fd }), "Dependencia registrada");
      } else guardar(() => S.api.actualizarPorId("Dependencias", "ID_Dependencia", d.ID_Dependencia, fd), "Dependencia actualizada");
    }, (fd) => (!fd.Depende_De ? "Elige el proyecto del que depende." : depsDe(fd.Depende_De).some((x) => x.Depende_De === p.ID_Proyecto) ? "Ese proyecto ya depende de este: se formaría un ciclo." : ""),
    nuevo ? {} : { eliminar: { texto: "Quitar dependencia", mensaje: "Se quitará esta dependencia.", accion: () => guardar(() => S.api.eliminarFilas("Dependencias", "ID_Dependencia", [d.ID_Dependencia]), "Dependencia quitada") } });
  }

  // ---------- Priorización del portafolio ----------
  const CRITERIOS = [
    { k: "Valor", t: "Valor para la institución", peso: 0.35, inv: false },
    { k: "Urgencia", t: "Urgencia", peso: 0.25, inv: false },
    { k: "Riesgo_Prio", t: "Complejidad / riesgo de ejecución", peso: 0.2, inv: true },
    { k: "Esfuerzo", t: "Esfuerzo / costo", peso: 0.2, inv: true },
  ];
  // Puntaje 0-100: más valor y urgencia suben; más riesgo y esfuerzo bajan.
  function puntaje(p) {
    if (CRITERIOS.some((c) => !num0(p[c.k]))) return null;
    const v = CRITERIOS.reduce((a, c) => a + c.peso * (c.inv ? 6 - Number(p[c.k]) : Number(p[c.k])), 0);
    return Math.round(((v - 1) / 4) * 100);
  }
  const cuadrante = (p) => { const v = Number(p.Valor), e = Number(p.Esfuerzo); return v >= 3 ? (e <= 3 ? "Ganancia rápida" : "Proyecto mayor") : (e <= 3 ? "Relleno" : "Reconsiderar"); };
  function vPriorizacion(el) {
    const ps = filtrados();
    const cal = ps.filter((p) => puntaje(p) !== null).sort((a, b) => puntaje(b) - puntaje(a));
    const sin = ps.filter((p) => puntaje(p) === null);
    const puede = (p) => R.puedeEditar(S.usuario, p, "editarProyecto");
    el.innerHTML = `<h1>Priorización del portafolio</h1>${barraFiltros()}
      <div class="card"><h2>Cómo se calcula</h2><p class="sub">Cada proyecto se califica de 1 a 5 en «Editar proyecto» → sección Priorización. Puntaje (0 a 100) = ${CRITERIOS.map((c) => `${c.t} ${Math.round(c.peso * 100)}%${c.inv ? " (a menor, mejor)" : ""}`).join(" · ")}.</p></div>
      <div class="grid2">
        <div class="card"><h2>Ranking</h2>${cal.length ? `<div class="tabla-scroll"><table><thead><tr><th>#</th><th>Proyecto</th><th>Puntaje</th><th class="opc">V · U · R · E</th><th>Cuadrante</th></tr></thead><tbody>
          ${cal.map((p, i) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}"><td>${i + 1}</td><td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.Estado)} · ${esc(nombreUsuario(p.PM))}</div></td>
            <td class="nowrap"><span class="cap-barra"><span class="cap-lleno" style="width:${puntaje(p)}%;background:var(--azul)"></span></span> <b>${puntaje(p)}</b></td>
            <td class="opc">${CRITERIOS.map((c) => esc(p[c.k])).join(" · ")}</td><td>${pill(cuadrante(p))}</td></tr>`).join("")}</tbody></table></div>` : vacio("Aún no hay proyectos calificados.")}</div>
        <div class="card"><h2>Valor vs. esfuerzo</h2>${cal.length ? `<div class="grafico-alto"><canvas id="g-prio"></canvas></div><p class="sub">Arriba a la izquierda: ganancias rápidas (mucho valor, poco esfuerzo). El tamaño del punto es la urgencia.</p>` : vacio("Califica proyectos para ver la matriz.")}</div>
      </div>
      ${sin.length ? `<div class="card"><h2>Sin calificar <span class="contador">${sin.length}</span></h2><ul class="lista">${sin.map((p) => `<li>${esc(p.Nombre)} ${puede(p) ? `<button class="btn chico" data-calificar="${esc(p.ID_Proyecto)}">Calificar</button>` : ""}</li>`).join("")}</ul></div>` : ""}`;
    enlazarFiltros();
    el.querySelectorAll("[data-calificar]").forEach((b) => b.addEventListener("click", () => formProyecto(proyecto(b.dataset.calificar))));
    if (cal.length) {
      const jit = (i) => ((i % 5) - 2) * 0.06;
      grafico($("#g-prio"), { type: "bubble", data: { datasets: [{ label: "Proyectos", data: cal.map((p, i) => ({ x: Number(p.Esfuerzo) + jit(i), y: Number(p.Valor) + jit(i + 2), r: 4 + Number(p.Urgencia) * 2.5, nombre: p.Nombre, pts: puntaje(p) })),
        backgroundColor: "rgba(16,73,148,.55)", borderColor: COLORES.azul }] },
        options: { maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `${c.raw.nombre}: ${c.raw.pts} pts` } } },
          scales: { x: { min: 0.5, max: 5.5, title: { display: true, text: "Esfuerzo / costo →" }, ticks: { stepSize: 1 } }, y: { min: 0.5, max: 5.5, title: { display: true, text: "Valor →" }, ticks: { stepSize: 1 } } } },
        plugins: [{ id: "cuadrantes", beforeDraw: (ch) => { const { ctx, scales: { x, y } } = ch; const cx = x.getPixelForValue(3.5), cy = y.getPixelForValue(2.5);
          ctx.save(); ctx.strokeStyle = "#ccc"; ctx.setLineDash([4, 4]); ctx.beginPath(); ctx.moveTo(cx, y.top); ctx.lineTo(cx, y.bottom); ctx.moveTo(x.left, cy); ctx.lineTo(x.right, cy); ctx.stroke();
          ctx.setLineDash([]); ctx.fillStyle = "#999"; ctx.font = "11px Calibri, Arial"; ctx.fillText("Ganancias rápidas", x.left + 6, y.top + 14); ctx.fillText("Proyectos mayores", cx + 6, y.top + 14);
          ctx.fillText("Rellenos", x.left + 6, y.bottom - 6); ctx.fillText("Reconsiderar", cx + 6, y.bottom - 6); ctx.restore(); } }] });
    }
  }

  // ---------- Curva S de presupuesto ----------
  const MESES_C = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const mesesEntre = (a, b) => { const r = []; let y = Number(a.slice(0, 4)), m = Number(a.slice(5, 7)); const yb = Number(b.slice(0, 4)), mb = Number(b.slice(5, 7));
    while (y < yb || (y === yb && m <= mb)) { r.push(`${y}-${String(m).padStart(2, "0")}`); m === 12 ? (y++, m = 1) : m++; if (r.length > 120) break; } return r; };
  const finDeMes = (ym) => { const [y, m] = ym.split("-").map(Number); return `${ym}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`; };
  const etiquetaMes = (ym) => `${MESES_C[Number(ym.slice(5, 7)) - 1]} ${ym.slice(2, 4)}`;
  // Plan: presupuesto de la línea base (o actual) repartido parejo entre inicio y fin; real: ejecutado reportado en los seguimientos.
  function planAcumulado(p, hasta) {
    const pres = num0(p.Presupuesto_Base) ?? (Number(p.Presupuesto) || 0);
    const fin = p.Fecha_Fin_Base || p.Fecha_Fin_Plan;
    if (!p.Fecha_Inicio || !fin || !pres) return 0;
    if (hasta < p.Fecha_Inicio) return 0;
    if (hasta >= fin) return pres;
    return pres * (R.difDias(hasta, p.Fecha_Inicio) + 1) / (R.difDias(fin, p.Fecha_Inicio) + 1);
  }
  function realAcumulado(p, hasta) {
    const hoy = R.hoyISO();
    const segs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos).filter((s) => num0(s.Ejecutado) !== null && s.Fecha_Corte <= hasta);
    if (hasta >= hoy) return Number(p.Ejecutado) || 0;
    return segs.length ? Number(segs[segs.length - 1].Ejecutado) : null;
  }
  function datosCurvaS(ps) {
    const con = ps.filter((p) => p.Fecha_Inicio && (p.Fecha_Fin_Base || p.Fecha_Fin_Plan) && (num0(p.Presupuesto_Base) || Number(p.Presupuesto)));
    if (!con.length) return null;
    const ini = con.reduce((m, p) => (p.Fecha_Inicio < m ? p.Fecha_Inicio : m), con[0].Fecha_Inicio);
    const fin = con.reduce((m, p) => { const f = p.Fecha_Fin_Base > p.Fecha_Fin_Plan ? p.Fecha_Fin_Base : p.Fecha_Fin_Plan; return f > m ? f : m; }, "0000");
    const meses = mesesEntre(ini, fin);
    const hoyM = R.hoyISO().slice(0, 7);
    const plan = meses.map((ym) => Math.round(con.reduce((a, p) => a + planAcumulado(p, finDeMes(ym)), 0)));
    const real = meses.map((ym) => {
      if (ym > hoyM) return null;
      const vals = con.map((p) => realAcumulado(p, ym === hoyM ? R.hoyISO() : finDeMes(ym)));
      return vals.every((v) => v === null) ? null : Math.round(vals.reduce((a, v) => a + (v || 0), 0));
    });
    return { labels: meses.map(etiquetaMes), plan, real };
  }
  function graficoCurvaS(canvas, d) {
    grafico(canvas, { type: "line", data: { labels: d.labels, datasets: [
      { label: "Planeado acumulado", data: d.plan, borderColor: COLORES.azulClaro || "#6BBAEF", backgroundColor: "#6BBAEF", borderDash: [6, 4], pointRadius: 0, tension: 0.2 },
      { label: "Ejecutado real", data: d.real, borderColor: COLORES.azul, backgroundColor: COLORES.azul, spanGaps: true, tension: 0.2 }] },
      options: { maintainAspectRatio: false, plugins: { tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${cop(c.raw)}` } } },
        scales: { y: { beginAtZero: true, ticks: { callback: (v) => (v >= 1e6 ? `$${Math.round(v / 1e6)} M` : cop(v)) } } } } });
  }

  // ---------- Auditoría ----------
  const ID_COLS = { Proyectos: "ID_Proyecto", Hitos: "ID_Hito", Seguimientos: "ID_Seguimiento", Riesgos: "ID_Riesgo", Usuarios: "Correo", Compromisos: "ID_Compromiso",
    Comentarios: "ID_Comentario", Stakeholders: "ID_Stakeholder", Proveedores: "ID_Proveedor", Recursos: "ID_Recurso", Tickets: "ID_Ticket", Cambios: "ID_Cambio", Lecciones: "ID_Leccion", RACI: "ID_RACI", Dependencias: "ID_Dependencia" };
  const CAMPO_DESC = ["Nombre", "Compromiso", "Titulo", "Hito", "Descripcion", "Entregable", "Leccion", "Texto", "Valor", "Logros"];
  const NO_AUDITAR = new Set(["Auditoria", "Archivos", "Filtros"]);
  const IGNORAR_CAMPOS = new Set(["Actualizado_Por", "Actualizado_El"]);
  function conAuditoria(api) {
    const pend = [];
    let timer = null;
    const buscar = (tabla, id) => (S.datos && S.datos[tabla] || []).find((r) => String(r[ID_COLS[tabla]]) === String(id));
    const desc = (o) => { const k = CAMPO_DESC.find((c) => o && o[c]); return k ? String(o[k]).slice(0, 140) : ""; };
    const corto = (v) => { const t = String(v ?? ""); return t.length > 80 ? t.slice(0, 80) + "…" : t; };
    const reg = (tabla, id, pid, accion, detalle) => {
      if (NO_AUDITAR.has(tabla) || !S.usuario) return;
      pend.push({ Fecha_Hora: ahora(), Usuario: S.usuario.Correo, Tabla: tabla, ID_Registro: String(id || ""), ID_Proyecto: pid || "", Accion: accion, Detalle: String(detalle || "").slice(0, 1500) });
      clearTimeout(timer); timer = setTimeout(vaciar, 1500);
    };
    async function vaciar() {
      clearTimeout(timer);
      if (!pend.length) return;
      const lote = pend.splice(0);
      try { await api.agregarFilas("Auditoria", lote); S.auditoria = null; } catch (e) { /* la auditoría nunca bloquea el trabajo */ }
    }
    const w = { ...api, vaciar };
    w.agregarFilas = async (tabla, objs) => {
      const r = await api.agregarFilas(tabla, objs);
      objs.forEach((o) => reg(tabla, o[ID_COLS[tabla]] || o.Valor || "", tabla === "Proyectos" ? o.ID_Proyecto : o.ID_Proyecto, "Creó", desc(o) || (tabla === "Catalogos" ? `${o.Lista}: ${o.Valor}` : "")));
      return r;
    };
    w.agregarFila = (tabla, obj) => w.agregarFilas(tabla, [obj]);
    w.actualizarVarios = async (tabla, colId, lista) => {
      const antes = lista.map(({ id }) => ({ ...(buscar(tabla, id) || {}) }));
      const r = await api.actualizarVarios(tabla, colId, lista);
      lista.forEach(({ id, cambios }, i) => {
        const a = antes[i];
        const dif = Object.entries(cambios).filter(([k, v]) => !IGNORAR_CAMPOS.has(k) && String(a[k] ?? "") !== String(v ?? ""))
          .map(([k, v]) => (k === "Asignaciones" || k === "Adjuntos" ? `${k} actualizado` : `${k}: «${corto(a[k])}» → «${corto(v)}»`));
        if (dif.length) reg(tabla, id, tabla === "Proyectos" ? id : a.ID_Proyecto, "Modificó", `${desc(a) ? desc(a) + " · " : ""}${dif.join(" · ")}`);
      });
      return r;
    };
    w.actualizarPorId = (tabla, colId, id, cambios) => w.actualizarVarios(tabla, colId, [{ id, cambios }]);
    w.eliminarFilas = async (tabla, colId, ids) => {
      const antes = ids.map((id) => buscar(tabla, id) || {});
      const r = await api.eliminarFilas(tabla, colId, ids);
      ids.forEach((id, i) => reg(tabla, id, tabla === "Proyectos" ? id : antes[i].ID_Proyecto, "Eliminó", desc(antes[i])));
      return r;
    };
    w.eliminarFila = async (tabla, criterios) => {
      const r = await api.eliminarFila(tabla, criterios);
      reg(tabla, "", "", "Eliminó", Object.entries(criterios).map(([k, v]) => `${k}: ${v}`).join(" · "));
      return r;
    };
    return w;
  }
  async function cargarAuditoria(forzar) {
    if (S.auditoria && !forzar) return S.auditoria;
    try { S.auditoria = (await S.api.leerTabla("Auditoria")).sort((a, b) => String(b.Fecha_Hora).localeCompare(String(a.Fecha_Hora))); }
    catch (e) { S.auditoria = []; }
    return S.auditoria;
  }
  const NOMBRE_TABLA = { Proyectos: "Proyecto", Hitos: "Hito", Seguimientos: "Seguimiento", Riesgos: "Riesgo", Usuarios: "Usuario", Compromisos: "Compromiso", Comentarios: "Comentario",
    Stakeholders: "Stakeholder", Proveedores: "Proveedor", Recursos: "Recurso", Tickets: "Ticket", Cambios: "Cambio", Lecciones: "Lección", RACI: "RACI", Dependencias: "Dependencia", Catalogos: "Catálogo" };
  const tablaAuditoria = (lista, conProyecto) => lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Fecha</th><th>Usuario</th>${conProyecto ? "<th>Proyecto</th>" : ""}<th>Qué</th><th>Detalle</th></tr></thead><tbody>
    ${lista.map((a) => `<tr><td class="nowrap">${fechaHora(a.Fecha_Hora)}</td><td>${esc(nombreUsuario(a.Usuario))}</td>${conProyecto ? `<td>${esc((proyecto(a.ID_Proyecto) || {}).Nombre || a.ID_Proyecto || "—")}</td>` : ""}
      <td class="nowrap">${pill(a.Accion)} ${esc(NOMBRE_TABLA[a.Tabla] || a.Tabla)}</td><td class="aud-det">${esc(a.Detalle)}</td></tr>`).join("")}</tbody></table></div>` : vacio("Sin cambios registrados todavía. El registro empieza desde esta versión.");
  function seccionAuditoria(p) {
    return `<div class="card"><div class="titulo-fila"><h2>Historial de cambios</h2><button class="btn chico" id="aud-recargar">Actualizar</button></div>
      <p class="sub">Quién cambió qué y cuándo en este proyecto (datos, seguimientos, compromisos, riesgos, stakeholders…).</p><div id="aud-cont">${vacio("Cargando historial…")}</div></div>`;
  }
  async function llenarAuditoria(el, filtro, conProyecto) {
    const c = el.querySelector("#aud-cont");
    if (!c) return;
    const lista = (await cargarAuditoria()).filter(filtro).slice(0, 500);
    if (el.querySelector("#aud-cont")) el.querySelector("#aud-cont").innerHTML = tablaAuditoria(lista, conProyecto);
  }
  function vAuditoria(el) {
    if (!R.puede(S.usuario, "usuarios")) { el.innerHTML = vacio("Sin acceso."); return; }
    const F = S.audF = S.audF || { usuario: "", tabla: "", texto: "" };
    el.innerHTML = `<div class="titulo-fila"><h1>Auditoría</h1><button class="btn" id="aud-recargar">Actualizar</button></div>
      <div class="filtros"><label class="filtro"><span>Usuario</span><select id="aud-u"><option value="">Todos</option>${S.datos.Usuarios.map((u) => `<option value="${esc(u.Correo)}" ${u.Correo === F.usuario ? "selected" : ""}>${esc(u.Nombre || u.Correo)}</option>`).join("")}</select></label>
        <label class="filtro"><span>Tipo de dato</span><select id="aud-t"><option value="">Todos</option>${Object.entries(NOMBRE_TABLA).map(([k, t]) => `<option value="${k}" ${k === F.tabla ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="filtro crece"><span>Buscar</span><input id="aud-x" type="search" value="${esc(F.texto)}" placeholder="Texto del detalle"></label></div>
      <div class="card"><p class="sub">Últimos 500 movimientos. Se guardan en la hoja Auditoria del Excel.</p><div id="aud-cont">${vacio("Cargando…")}</div></div>`;
    const filtro = (a) => (!F.usuario || lc(a.Usuario) === lc(F.usuario)) && (!F.tabla || a.Tabla === F.tabla) && (!F.texto || lc(`${a.Detalle} ${a.ID_Registro}`).includes(lc(F.texto)));
    $("#aud-u").addEventListener("change", (e) => { F.usuario = e.target.value; render(); });
    $("#aud-t").addEventListener("change", (e) => { F.tabla = e.target.value; render(); });
    $("#aud-x").addEventListener("change", (e) => { F.texto = e.target.value; render(); });
    $("#aud-recargar").addEventListener("click", async () => { await cargarAuditoria(true); render(); });
    llenarAuditoria(el, filtro, true);
  }

  // ---------- Informe semanal del portafolio ----------
  function rangoSemana(sem) {
    const [y, w] = sem.split("-S").map(Number);
    const ene4 = new Date(y, 0, 4), dia = ene4.getDay() || 7;
    const lunes = new Date(y, 0, 4 - dia + 1 + (w - 1) * 7);
    const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    const dom = new Date(lunes); dom.setDate(lunes.getDate() + 6);
    return { ini: iso(lunes), fin: iso(dom) };
  }
  function informeSemanal(sem) {
    const ps = filtrados();
    const ids = new Set(ps.map((p) => p.ID_Proyecto));
    const { ini, fin } = rangoSemana(sem);
    const enSemana = (f) => f && String(f).slice(0, 10) >= ini && String(f).slice(0, 10) <= fin;
    const k = R.kpis(ps, S.datos.Seguimientos);
    const activos = ps.filter((p) => p.Estado === "Activo");
    const t = (v) => esc(v === 0 ? "0" : v || "—");
    const tabla = (cab, filas, vacioTxt) => filas.length ? `<table class="hv-t"><thead><tr>${cab.map((c) => `<th>${c}</th>`).join("")}</tr></thead><tbody>${filas.join("")}</tbody></table>` : `<p class="hv-vacio">${vacioTxt}</p>`;
    const segsSem = S.datos.Seguimientos.filter((s) => ids.has(s.ID_Proyecto) && enSemana(s.Fecha_Corte));
    const variacion = (p) => {
      const hist = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos).filter((s) => s.Fecha_Corte <= fin);
      const ult = hist[hist.length - 1], ant = hist.filter((s) => s.Fecha_Corte < ini).pop();
      return ult && ant ? (Number(ult.Avance_Real) || 0) - (Number(ant.Avance_Real) || 0) : null;
    };
    const comps = S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto));
    const vencidos = comps.filter((c) => R.estadoCompromiso(c) === "Vencido");
    const cerradosSem = comps.filter((c) => R.cerrado(c) && enSemana(c.Fecha_Cierre));
    const acordadosSem = comps.filter((c) => segsSem.some((s) => s.ID_Seguimiento === c.ID_Seguimiento));
    const riesgosAltos = S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto) && riesgoActivo(r) && R.nivelRiesgo(Number(r.Calificacion_Residual) || Number(r.Calificacion_Inherente) || 0) === "Alto");
    const cambiosSem = (S.datos.Cambios || []).filter((c) => ids.has(c.ID_Proyecto) && (enSemana(c.Fecha) || enSemana(c.Fecha_Decision)));
    const sobre = capacidad().filter((x) => x.total > x.cap && x.asign.some((a) => ids.has(a.p.ID_Proyecto)));
    const depsRiesgo = (S.datos.Dependencias || []).filter((d) => ids.has(d.ID_Proyecto) && conflictoDep(d));
    const nomP = (id) => esc((proyecto(id) || {}).Nombre || id);
    const conteo = ["Verde", "Amarillo", "Rojo"].map((s) => `${s}: ${activos.filter((p) => p.Semaforo === s).length}`).join(" · ");
    const doc = `${cabDocumento("Informe semanal del portafolio", `Semana ${sem.split("-S")[1]} de ${sem.slice(0, 4)}`, `Del ${fecha(ini)} al ${fecha(fin)} · ${ps.length} proyecto(s) en la selección`)}
      <section class="hv-kpis">
        <div><span>Proyectos activos</span><b>${k.activos}</b></div><div><span>Avance real / plan (promedio)</span><b>${pct(k.avanceReal)} / ${pct(k.avancePlan)}</b></div>
        <div><span>Semáforo</span><b style="font-size:13px">${conteo}</b></div><div><span>Seguimiento al día</span><b>${pct(k.pctAlDia)}</b></div>
        <div><span>Presupuesto ejecutado</span><b>${pct(k.pctEjecutado)}</b></div><div><span>Compromisos vencidos</span><b>${vencidos.length}</b></div>
      </section>
      <h2>1. Estado de los proyectos</h2>
      ${tabla(["Proyecto", "PM", "Semáforo", "Avance", "Δ semana", "Seguimiento", "Novedad"], activos.sort((a, b) => ({ Rojo: 0, Amarillo: 1, Verde: 2 }[a.Semaforo] ?? 3) - ({ Rojo: 0, Amarillo: 1, Verde: 2 }[b.Semaforo] ?? 3)).map((p) => {
        const v = variacion(p), e = R.estadoSeguimiento(p, S.datos.Seguimientos);
        return `<tr><td><b>${t(p.Nombre)}</b>${p.Codigo_Almera ? `<br><small>Almera ${esc(p.Codigo_Almera)}</small>` : ""}</td><td>${t(nombreUsuario(p.PM))}</td><td>${t(p.Semaforo)}</td>
          <td>${pct(p.Avance_Real)} / ${pct(p.Avance_Planeado)}</td><td>${v === null ? "—" : `${v > 0 ? "+" : ""}${v} pts`}</td><td>${t(e.estado)}</td><td>${t(p.Comentario_Estado)}</td></tr>`;
      }), "No hay proyectos activos en la selección.")}
      <h2>2. Seguimientos de la semana</h2>
      ${tabla(["Proyecto", "Fecha", "Avance", "Logros", "Próximos pasos", "Bloqueos"], segsSem.map((s) => `<tr><td>${nomP(s.ID_Proyecto)}</td><td class="hv-nw">${fecha(s.Fecha_Corte)}</td><td>${pct(s.Avance_Real)}</td><td>${t(s.Logros)}</td><td>${t(s.Proximos_Pasos)}</td><td>${t(s.Bloqueos)}</td></tr>`), "No se reportaron seguimientos esta semana.")}
      <h2>3. Compromisos</h2>
      <p class="hv-p"><b>${acordadosSem.length}</b> acordados en la semana · <b>${cerradosSem.length}</b> cerrados en la semana · <b>${vencidos.length}</b> vencidos a la fecha.</p>
      ${tabla(["Compromiso vencido", "Proyecto", "Responsable", "Fecha límite", "Estado"], vencidos.sort((a, b) => String(a.Fecha_Compromiso).localeCompare(String(b.Fecha_Compromiso))).map((c) => `<tr><td>${t(c.Compromiso)}</td><td>${nomP(c.ID_Proyecto)}</td><td>${t(c.Responsable)}</td><td class="hv-nw">${fecha(c.Fecha_Compromiso)}</td><td>${esc(R.estadoBase(c))}</td></tr>`), "Sin compromisos vencidos.")}
      ${cerradosSem.length ? tabla(["Cerrado en la semana", "Proyecto", "Responsable", "Cerrado"], cerradosSem.map((c) => `<tr><td>${t(c.Compromiso)}</td><td>${nomP(c.ID_Proyecto)}</td><td>${t(c.Responsable)}</td><td class="hv-nw">${fecha(c.Fecha_Cierre)}</td></tr>`), "") : ""}
      <h2>4. Riesgos altos activos</h2>
      ${tabla(["Riesgo", "Proyecto", "Nivel", "Mitigación"], riesgosAltos.map((r) => `<tr><td>${t(r.Descripcion)}</td><td>${nomP(r.ID_Proyecto)}</td><td>${R.nivelRiesgo(Number(r.Calificacion_Residual) || Number(r.Calificacion_Inherente) || 0)}</td><td>${t(r.Plan_Mitigacion)}</td></tr>`), "Sin riesgos altos abiertos.")}
      <h2>5. Control de cambios</h2>
      ${tabla(["Cambio", "Proyecto", "Tipo", "Estado"], cambiosSem.map((c) => `<tr><td>${t(c.Descripcion)}</td><td>${nomP(c.ID_Proyecto)}</td><td>${t(c.Tipo)}</td><td>${t(c.Estado)}</td></tr>`), "Sin cambios solicitados o decididos en la semana.")}
      <h2>6. Alertas del portafolio</h2>
      ${tabla(["Alerta", "Detalle"], [
        ...sobre.map((x) => `<tr><td>Persona sobreasignada</td><td>${esc(x.nombre)}: ${x.total}% (${x.asign.map((a) => `${esc(a.p.Nombre)} ${a.ded}%`).join(", ")})</td></tr>`),
        ...depsRiesgo.map((d) => `<tr><td>Dependencia en riesgo</td><td>${esc(conflictoDep(d))}</td></tr>`),
      ], "Sin alertas de capacidad ni de dependencias.")}
      <footer class="hv-pie">Fundación Santa Fe de Bogotá · Oficina de Proyectos · Informe semanal ${esc(sem)} · Generado desde el Portafolio PMO</footer>`;
    mostrarDocumento(`Informe semanal ${sem}`, doc, `Informe semanal portafolio ${sem}`);
  }

  // ---------- Recursos: directorio único de personas (sin duplicados) ----------
  const normTxt = (t) => String(t || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim();
  const recursos = () => (S.datos.Recursos || []).slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
  const recursoId = (id) => (S.datos.Recursos || []).find((r) => r.ID_Recurso === id);
  // Busca por ID, luego por correo y luego por nombre (sin tildes ni mayúsculas).
  function recursoPor({ id, correo, nombre }) {
    const L = S.datos.Recursos || [];
    return (id && L.find((r) => r.ID_Recurso === id)) || (correo && L.find((r) => lc(r.Correo) === lc(correo))) || (nombre && L.find((r) => normTxt(r.Nombre) === normTxt(nombre))) || null;
  }
  const empresaDe = (r) => (r && r.ID_Proveedor ? (proveedor(r.ID_Proveedor) || {}).Nombre || "Proveedor" : "Interno (FSFB)");
  const datalistRecursos = (id) => `<datalist id="${id}">${recursos().filter((r) => r.Activo !== "No").map((r) => `<option value="${esc(r.Nombre)}" label="${esc([r.Cargo, empresaDe(r), r.Correo].filter(Boolean).join(" · "))}">`).join("")}</datalist>`;
  const nuevoIdRecurso = (desp = 0) => `REC-${String((S.datos.Recursos || []).reduce((m, r) => Math.max(m, parseInt(String(r.ID_Recurso).slice(4), 10) || 0), 0) + 1 + desp).padStart(4, "0")}`;
  // Devuelve el recurso existente o lo crea (una sola vez) con los datos disponibles.
  async function asegurarRecurso(d) {
    if (!d || !String(d.Nombre || "").trim()) return null;
    const ya = recursoPor({ id: d.ID_Recurso, correo: d.Correo, nombre: d.Nombre });
    if (ya) {
      const faltan = {};
      ["Correo", "Cargo", "Area", "Telefono", "ID_Proveedor"].forEach((k) => { if (d[k] && !ya[k]) faltan[k] = d[k]; });
      if (Object.keys(faltan).length) { await S.api.actualizarPorId("Recursos", "ID_Recurso", ya.ID_Recurso, faltan); Object.assign(ya, faltan); }
      return ya;
    }
    const r = { ID_Recurso: nuevoIdRecurso(), Nombre: String(d.Nombre).trim(), Correo: d.Correo || "", Cargo: d.Cargo || "", Area: d.Area || "", Telefono: d.Telefono || "",
      ID_Proveedor: d.ID_Proveedor || "", Capacidad: 100, Activo: "Sí" };
    await S.api.agregarFila("Recursos", r);
    (S.datos.Recursos = S.datos.Recursos || []).push(r);
    return r;
  }
  // Primera vez: arma el directorio con las personas que ya existen (usuarios, stakeholders, responsables y contactos de proveedores).
  let recursosMigrados = false;
  async function migrarRecursos() {
    if (recursosMigrados || !S.datos.Recursos || S.datos.Recursos.length) return;
    recursosMigrados = true;
    const mapa = new Map();
    const agregar = (d) => {
      const nombre = String(d.Nombre || "").trim(), correo = String(d.Correo || "").trim();
      if (!nombre && !correo) return;
      if (correo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo) && !nombre) return;
      const k = [...mapa.keys()].find((x) => (correo && mapa.get(x).Correo && lc(mapa.get(x).Correo) === lc(correo)) || (nombre && normTxt(mapa.get(x).Nombre) === normTxt(nombre)));
      if (k) { const r = mapa.get(k); ["Nombre", "Correo", "Cargo", "Area", "Telefono", "ID_Proveedor"].forEach((c) => { if (!r[c] && d[c]) r[c] = d[c]; }); return; }
      mapa.set(correo || normTxt(nombre), { Nombre: nombre || correo, Correo: correo, Cargo: d.Cargo || "", Area: d.Area || "", Telefono: d.Telefono || "", ID_Proveedor: d.ID_Proveedor || "" });
    };
    S.datos.Usuarios.forEach((u) => agregar({ Nombre: u.Nombre, Correo: u.Correo }));
    (S.datos.Stakeholders || []).forEach((x) => agregar(x));
    (S.datos.Proveedores || []).forEach((v) => { if (v.Contacto) agregar({ Nombre: v.Contacto, Correo: v.Correo, Telefono: v.Telefono, ID_Proveedor: v.ID_Proveedor }); });
    S.datos.Compromisos.forEach((c) => { if (c.Correo_Responsable) agregar({ Nombre: c.Responsable, Correo: c.Correo_Responsable }); });
    const filas = [...mapa.values()].map((r, i) => ({ ID_Recurso: `REC-${String(i + 1).padStart(4, "0")}`, ...r, Capacidad: 100, Activo: "Sí" }));
    if (!filas.length) return;
    try {
      await S.api.agregarFilas("Recursos", filas);
      S.datos.Recursos = filas;
      const enlaces = (S.datos.Stakeholders || []).map((x) => { const r = recursoPor({ correo: x.Correo, nombre: x.Nombre }); return r ? { id: x.ID_Stakeholder, cambios: { ID_Recurso: r.ID_Recurso } } : null; }).filter(Boolean);
      if (enlaces.length) await S.api.actualizarVarios("Stakeholders", "ID_Stakeholder", enlaces);
    } catch (e) { S.datos.Recursos = []; /* sin permiso de edición */ }
  }
  // Cambios en un recurso se reflejan donde aparece (stakeholders y responsables de compromisos).
  async function propagarRecurso(antes, r) {
    const stk = (S.datos.Stakeholders || []).filter((x) => x.ID_Recurso === r.ID_Recurso || (!x.ID_Recurso && ((antes.Correo && lc(x.Correo) === lc(antes.Correo)) || normTxt(x.Nombre) === normTxt(antes.Nombre))));
    if (stk.length) await S.api.actualizarVarios("Stakeholders", "ID_Stakeholder", stk.map((x) => ({ id: x.ID_Stakeholder,
      cambios: { ID_Recurso: r.ID_Recurso, Nombre: r.Nombre, Correo: r.Correo, Cargo: r.Cargo, Area: r.Area, Telefono: r.Telefono, ID_Proveedor: r.ID_Proveedor } })));
    const cmp = S.datos.Compromisos.filter((c) => (antes.Correo && lc(c.Correo_Responsable) === lc(antes.Correo)) || (antes.Nombre && normTxt(c.Responsable) === normTxt(antes.Nombre)));
    if (cmp.length && (antes.Nombre !== r.Nombre || antes.Correo !== r.Correo)) await S.api.actualizarVarios("Compromisos", "ID_Compromiso", cmp.map((c) => ({ id: c.ID_Compromiso, cambios: { Responsable: r.Nombre, Correo_Responsable: r.Correo } })));
  }
  const dupRecurso = (fd, excluir) => (S.datos.Recursos || []).find((r) => r.ID_Recurso !== excluir && ((fd.Correo && lc(r.Correo) === lc(fd.Correo)) || normTxt(r.Nombre) === normTxt(fd.Nombre)));
  function formRecurso(r, alCrear) {
    const nuevo = !r;
    const empresas = [["", "Interno (FSFB)"], ...proveedoresActivos(r && r.ID_Proveedor).map((v) => [v.ID_Proveedor, v.Nombre])];
    const usos = nuevo ? { stk: [], cmp: [] } : {
      stk: (S.datos.Stakeholders || []).filter((x) => x.ID_Recurso === r.ID_Recurso || (r.Correo && lc(x.Correo) === lc(r.Correo))),
      cmp: S.datos.Compromisos.filter((c) => r.Correo && lc(c.Correo_Responsable) === lc(r.Correo)),
    };
    const campos = [
      { k: "Nombre", label: "Nombre y apellido", ancho: true },
      { k: "Correo", label: "Correo", tipo: "email", placeholder: "nombre@fsfb.org.co" },
      { k: "Telefono", label: "Teléfono" },
      { k: "Cargo", label: "Cargo" },
      { k: "Area", label: "Área", tipo: "select", opciones: S.cat.Cliente_Area },
      { k: "ID_Proveedor", label: "Empresa", tipo: "select", opciones: empresas, ayuda: "Interno o el proveedor al que pertenece." },
      { k: "Capacidad", label: "Capacidad disponible (%)", tipo: "number", min: 0, max: 100, def: 100, ayuda: "100 = tiempo completo; 50 = medio tiempo. Se usa en «Capacidad del equipo»." },
      { k: "Activo", label: "¿Activo?", tipo: "select", opciones: ["Sí", "No"], def: "Sí" },
    ];
    modal(nuevo ? "Nuevo recurso" : `Editar ${r.Nombre}`, campos, nuevo ? {} : r, (fd) => {
      if (nuevo) {
        const fila = { ID_Recurso: nuevoIdRecurso(), ...fd, Capacidad: fd.Capacidad === "" ? 100 : fd.Capacidad };
        guardar(() => S.api.agregarFila("Recursos", fila), `Recurso «${fila.Nombre}» creado`).then(() => { if (alCrear) alCrear(fila); });
      } else guardar(async () => { await S.api.actualizarPorId("Recursos", "ID_Recurso", r.ID_Recurso, fd); await propagarRecurso(r, { ...r, ...fd }); }, "Recurso actualizado");
    }, (fd) => {
      if (!fd.Nombre) return "Escribe el nombre.";
      const d = dupRecurso(fd, r && r.ID_Recurso);
      return d ? `Ya existe «${d.Nombre}»${d.Correo ? ` (${d.Correo})` : ""} en Recursos. Usa ese registro para no duplicar.` : "";
    }, nuevo ? {} : { eliminar: { texto: "Eliminar recurso",
      mensaje: usos.stk.length || usos.cmp.length ? `Aparece en ${usos.stk.length} proyecto(s) como stakeholder y en ${usos.cmp.length} compromiso(s); esos registros se conservan con su nombre. Si ya no trabaja con ustedes, mejor márcalo como inactivo.` : `Se eliminará «${r.Nombre}» del directorio.`,
      accion: () => guardar(async () => {
        if (usos.stk.length) await S.api.actualizarVarios("Stakeholders", "ID_Stakeholder", usos.stk.map((x) => ({ id: x.ID_Stakeholder, cambios: { ID_Recurso: "" } })));
        await S.api.eliminarFilas("Recursos", "ID_Recurso", [r.ID_Recurso]);
      }, "Recurso eliminado") } });
  }
  // Posibles duplicados: mismo correo o mismo nombre (sin tildes); también nombres donde uno contiene al otro.
  function gruposDuplicados() {
    const L = recursos(), vistos = new Set(), grupos = [];
    L.forEach((a) => {
      if (vistos.has(a.ID_Recurso)) return;
      const g = L.filter((b) => b.ID_Recurso !== a.ID_Recurso && !vistos.has(b.ID_Recurso) && ((a.Correo && lc(a.Correo) === lc(b.Correo)) || normTxt(a.Nombre) === normTxt(b.Nombre) ||
        (normTxt(a.Nombre).length > 5 && normTxt(b.Nombre).length > 5 && (normTxt(a.Nombre).includes(normTxt(b.Nombre)) || normTxt(b.Nombre).includes(normTxt(a.Nombre))))));
      if (g.length) { const todos = [a, ...g]; todos.forEach((x) => vistos.add(x.ID_Recurso)); grupos.push(todos); }
    });
    return grupos;
  }
  function fusionar(grupo) {
    const campos = [{ k: "Queda", label: "¿Cuál registro se conserva?", tipo: "select", ancho: true, opciones: grupo.map((r) => [r.ID_Recurso, `${r.Nombre}${r.Correo ? ` · ${r.Correo}` : ""}${r.Cargo ? ` · ${r.Cargo}` : ""}`]), def: grupo[0].ID_Recurso }];
    modal("Fusionar recursos duplicados", campos, {}, (fd) => guardar(async () => {
      const queda = grupo.find((r) => r.ID_Recurso === fd.Queda) || grupo[0];
      const otros = grupo.filter((r) => r !== queda);
      const completo = { ...queda };
      otros.forEach((o) => ["Correo", "Cargo", "Area", "Telefono", "ID_Proveedor"].forEach((k) => { if (!completo[k] && o[k]) completo[k] = o[k]; }));
      await S.api.actualizarPorId("Recursos", "ID_Recurso", queda.ID_Recurso, completo);
      for (const o of otros) await propagarRecurso(o, completo);
      await S.api.eliminarFilas("Recursos", "ID_Recurso", otros.map((o) => o.ID_Recurso));
    }, "Recursos fusionados"), null, { antes: `<p class="sub">Los stakeholders y compromisos de los demás pasan al que conserves; se completan los datos que le falten y se eliminan los repetidos.</p>` });
  }
  function vRecursos(el) {
    const puede = R.puede(S.usuario, "editarProyecto");
    const q = S.recQ || "";
    const cap = capacidad();
    const asig = (r) => cap.find((x) => (r.Correo && lc(x.correo) === lc(r.Correo)) || normTxt(x.nombre) === normTxt(r.Nombre));
    const lista = recursos().filter((r) => !q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${r.Area} ${empresaDe(r)}`).includes(normTxt(q)));
    const dups = gruposDuplicados();
    el.innerHTML = `<div class="titulo-fila"><h1>Recursos</h1>${puede ? `<button class="btn primario" id="b-rec">+ Recurso</button>` : ""}</div>
      <div class="filtros"><label class="filtro crece"><span>Buscar</span><input id="rec-q" type="search" value="${esc(q)}" placeholder="Nombre, correo, cargo, área o empresa"></label></div>
      ${dups.length && puede ? `<div class="card aviso-dup"><h2>Posibles duplicados <span class="contador rojo">${dups.length}</span></h2>
        <ul class="lista">${dups.map((g, i) => `<li>${g.map((r) => `<b>${esc(r.Nombre)}</b>${r.Correo ? ` <span class="sub">${esc(r.Correo)}</span>` : ""}`).join(" · ")} <button class="btn chico" data-fusionar="${i}">Fusionar</button></li>`).join("")}</ul></div>` : ""}
      <div class="card"><p class="sub">Directorio único de personas del portafolio. En stakeholders, equipo y responsables de compromisos se eligen de aquí y sus datos se llenan solos; si escribes a alguien nuevo, se agrega automáticamente.</p>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Persona</th><th>Empresa</th><th>Área</th><th>Contacto</th><th>Asignado</th><th></th></tr></thead><tbody>
          ${lista.map((r) => { const a = asig(r), c = num0(r.Capacidad) ?? 100; return `<tr class="${r.Activo === "No" ? "inactivo" : ""}"><td>${persona(r.Nombre, r.Cargo)}${r.Activo === "No" ? ` <span class="pill noaplica">Inactivo</span>` : ""}</td>
            <td>${esc(empresaDe(r))}</td><td>${esc(r.Area || "—")}</td><td>${contacto(r.Correo, r.Telefono)}</td>
            <td class="nowrap">${a ? `${a.total}% de ${c}%${a.total > c ? ` <span class="baja">⚠</span>` : ""}<div class="sub">${a.asign.length} asignación(es)</div>` : `<span class="sub">— de ${c}%</span>`}</td>
            <td class="derecha">${puede ? `<button class="btn chico" data-rec="${esc(r.ID_Recurso)}">Editar</button>` : ""}</td></tr>`; }).join("")}</tbody></table></div>`
        : vacio(q ? "Nadie coincide con la búsqueda." : "Aún no hay recursos. Se crean solos con las personas que ya existen o con «+ Recurso».")}</div>`;
    const qi = $("#rec-q"); qi.addEventListener("change", () => { S.recQ = qi.value; render(); });
    if ($("#b-rec")) $("#b-rec").addEventListener("click", () => formRecurso());
    el.querySelectorAll("[data-rec]").forEach((b) => b.addEventListener("click", () => formRecurso(recursoId(b.dataset.rec))));
    el.querySelectorAll("[data-fusionar]").forEach((b) => b.addEventListener("click", () => fusionar(dups[Number(b.dataset.fusionar)])));
  }
  // Conecta un campo de nombre a Recursos: al elegir a alguien llena sus datos; muestra si es nuevo.
  function selectorPersona(m, nombreSel, relleno, avisoSel) {
    const no = m.querySelector(nombreSel);
    if (!no) return;
    no.setAttribute("data-recurso", ""); no.setAttribute("autocomplete", "off");
    const aviso = !avisoSel ? null : typeof avisoSel === "string" ? m.querySelector(avisoSel) : avisoSel;
    const act = () => {
      const r = recursoPor({ nombre: no.value });
      if (r) Object.entries(relleno).forEach(([campo, sel]) => { const i = m.querySelector(sel); if (i && r[campo] !== undefined && r[campo] !== "") i.value = r[campo]; });
      if (aviso) aviso.innerHTML = !no.value.trim() ? "" : r ? `✓ Del directorio de recursos${r.Cargo ? ` · ${esc(r.Cargo)}` : ""}` : `Persona nueva: se agregará a Recursos al guardar.`;
      no.dispatchEvent(new Event("recurso"));
    };
    no.addEventListener("change", act); no.addEventListener("input", () => { if (recursoPor({ nombre: no.value })) act(); else if (aviso) aviso.innerHTML = no.value.trim() ? "Persona nueva: se agregará a Recursos al guardar." : ""; });
    if (no.value) act();
  }

  // ---------- Riesgos: formato FSFB, origen y carga masiva desde Excel ----------
  const PROB_TXT = { 1: "1 (Muy baja)", 2: "2 (Baja)", 3: "3 (Moderada)", 4: "4 (Alta)", 5: "5 (Muy alta)" };
  const IMP_TXT = { 1: "1 (Insignificante)", 2: "2 (Menor)", 3: "3 (Moderado)", 4: "4 (Mayor)", 5: "5 (Catastrófico)" };
  const ESTADOS_RIESGO = ["Abierto", "En seguimiento", "Materializado", "Cerrado"];
  const riesgoActivo = (r) => r.Estado !== "Cerrado";
  const riesgosDe = (pid) => S.datos.Riesgos.filter((r) => r.ID_Proyecto === pid);
  const numRiesgo = (r) => parseInt(String(r.ID_Riesgo).split("-").pop(), 10) || 0;
  const califTxt = (p, i) => { const c = R.calificacion(p, i); return c ? `${c} · ${R.nivelRiesgo(c)}` : "—"; };
  const origenRiesgo = (r) => {
    const s = r.ID_Seguimiento && S.datos.Seguimientos.find((x) => x.ID_Seguimiento === r.ID_Seguimiento);
    return [r.Origen, s ? `Sesión del ${fecha(s.Fecha_Corte)}` : "", !s && r.Fecha_Identificacion ? fecha(r.Fecha_Identificacion) : ""].filter(Boolean).join(" · ") || "Sin origen registrado";
  };
  function seccionRiesgos(p, puede) {
    return `<div class="card"><div class="titulo-fila"><h2>Registro de riesgos</h2>${puede ? `<div class="acciones">
        <a class="btn" href="assets/Plantilla_Riesgos.xlsx" download="Plantilla_Riesgos.xlsx" title="Formato FSFB con listas desplegables">⬇ Plantilla Excel</a>
        <button class="btn" id="b-rg-carga">📥 Cargar desde Excel</button><button class="btn primario" id="b-riesgo">+ Riesgo</button></div>` : ""}</div>
      <p class="sub">Formato FSFB: probabilidad × impacto (1 a 5) antes y después de mitigar. Usa los indicadores y el mapa para filtrar; cambia el estado directamente en cada tarjeta.</p></div>
      ${panelRiesgos("p:" + p.ID_Proyecto, riesgosDe(p.ID_Proyecto), { conProyecto: false, puedeFn: () => puede })}`;
  }
  // Campos de origen comunes al formulario y a la carga masiva.
  const opcionesSesion = (pid) => R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse().map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)}${s.Fecha_Acta && s.Fecha_Acta !== s.Fecha_Corte ? ` (acta ${fecha(s.Fecha_Acta)})` : ""}`]);

  // Lee el Excel del formato FSFB (o parecido): encabezados en cualquier fila de las primeras 15.
  let cargaXlsx = null;
  function libXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!cargaXlsx) cargaXlsx = new Promise((ok, falla) => {
      const s = document.createElement("script");
      s.src = new URL("lib/xlsx.full.min.js", location.href).href;
      s.onload = () => ok(window.XLSX); s.onerror = () => falla(new Error("No se pudo cargar el lector de Excel."));
      document.head.appendChild(s);
    });
    return cargaXlsx;
  }
  const COLS_RIESGO = [
    ["No", /^no\.?$|^n[°º]|^numero|^#$/], ["Tipo", /^tipo/], ["Estado", /^estado/], ["Responsable", /^responsable/], ["Descripcion", /descripcion|^riesgo$/],
    ["Probabilidad_Inherente", /probabilidad.*inherente|^probabilidad$/], ["Impacto_Inherente", /impacto.*inherente|^impacto$/],
    ["Plan_Mitigacion", /mitigacion/], ["Plan_Contingencia", /contingencia/],
    ["Probabilidad_Residual", /probabilidad.*residual/], ["Impacto_Residual", /impacto.*residual/],
  ];
  const num15 = (v) => { const m = String(v ?? "").match(/[1-5]/); return m ? Number(m[0]) : ""; };
  async function leerExcelRiesgos(file) {
    const X = await libXlsx();
    const wb = X.read(new Uint8Array(await file.arrayBuffer()), { type: "array" });
    const hoja = wb.Sheets[wb.SheetNames.find((n) => /riesgo/i.test(n)) || wb.SheetNames[0]];
    const filas = X.utils.sheet_to_json(hoja, { header: 1, raw: false, defval: "" });
    const iCab = filas.slice(0, 15).findIndex((f) => f.some((c) => /descripci/i.test(normTxt(c))));
    if (iCab < 0) throw new Error("No encontré la columna «Descripción del riesgo». Usa la plantilla del formato.");
    const cab = filas[iCab].map(normTxt);
    const idx = {};
    COLS_RIESGO.forEach(([k, re]) => { const i = cab.findIndex((c, j) => re.test(c) && !Object.values(idx).includes(j)); if (i >= 0) idx[k] = i; });
    const titulo = filas.slice(0, iCab).map((f) => f.filter(Boolean).join(" ")).join(" ").trim();
    const riesgos = filas.slice(iCab + 1).map((f) => {
      const g = (k) => (idx[k] === undefined ? "" : String(f[idx[k]] ?? "").trim());
      if (!g("Descripcion")) return null;
      return {
        Tipo: /oportunidad/i.test(g("Tipo")) ? "Oportunidad" : "Amenaza",
        Estado: ESTADOS_RIESGO.find((e) => normTxt(e) === normTxt(g("Estado"))) || g("Estado") || "Abierto",
        Responsable: g("Responsable"), Descripcion: g("Descripcion"),
        Probabilidad_Inherente: num15(g("Probabilidad_Inherente")), Impacto_Inherente: num15(g("Impacto_Inherente")),
        Plan_Mitigacion: g("Plan_Mitigacion"), Plan_Contingencia: g("Plan_Contingencia"),
        Probabilidad_Residual: num15(g("Probabilidad_Residual")), Impacto_Residual: num15(g("Impacto_Residual")),
      };
    }).filter(Boolean);
    // Del título «Registro de riesgos — Comité Operativo – Proyecto Kardex — 29/09/2026» salen la instancia y la fecha.
    const fechaT = ACTAS.fechaISO(titulo);
    const partes = titulo.replace(/registro de riesgos/i, "").split(/[—–-]/).map((t) => t.trim()).filter((t) => t && !/^proyecto\b/i.test(t) && !/\d{1,2}\/\d{1,2}\/\d{4}/.test(t) && !/^\[/.test(t));
    return { titulo, riesgos, fecha: fechaT, instancia: partes[0] || "" };
  }
  function cargarRiesgosExcel(p) {
    const i = document.createElement("input");
    i.type = "file"; i.accept = ".xlsx,.xls,.csv";
    i.addEventListener("change", async () => {
      const f = i.files[0];
      if (!f) return;
      cargando(true, "Leyendo el Excel…");
      let datos;
      try { datos = await leerExcelRiesgos(f); } catch (e) { cargando(false); toast(e.message, true); return; }
      cargando(false);
      if (!datos.riesgos.length) { toast("El archivo no tiene riesgos con descripción.", true); return; }
      previsualizarRiesgos(p, datos, f.name);
    });
    i.click();
  }
  function previsualizarRiesgos(p, datos, nombreArchivo) {
    const existentes = riesgosDe(p.ID_Proyecto);
    const dupDe = (r) => existentes.find((x) => normTxt(x.Descripcion) === normTxt(r.Descripcion));
    const sesiones = opcionesSesion(p.ID_Proyecto);
    const segFecha = datos.fecha && R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos).find((s) => s.Fecha_Corte === datos.fecha || s.Fecha_Acta === datos.fecha);
    const campos = [
      { k: "ID_Seguimiento", label: "¿De qué sesión salieron?", tipo: "select", opciones: sesiones, textoVacio: "No salieron de una sesión registrada", ancho: true,
        ayuda: segFecha ? "Encontré una sesión con la misma fecha del archivo y la dejé elegida." : "Si no está la sesión, déjalo así e indica la instancia." },
      { k: "Origen", label: "Instancia donde se identificaron", placeholder: "Ej.: Comité Operativo, Kickoff, Comité directivo" },
      { k: "Fecha_Identificacion", label: "Fecha de identificación", tipo: "date" },
    ];
    const filas = datos.riesgos.map((r, k) => {
      const d = dupDe(r);
      return `<tr><td><input type="checkbox" class="rg-ok" data-k="${k}" ${d ? "" : "checked"} aria-label="Importar riesgo ${k + 1}"></td>
        <td>${pill(r.Tipo)} ${pill(r.Estado)}<div class="rg-desc">${esc(r.Descripcion)}</div><div class="sub">${r.Responsable ? `${esc(r.Responsable)} ${recursoPor({ nombre: r.Responsable }) ? "· ✓ en Recursos" : "· se agregará a Recursos"}` : "Sin responsable"}</div>
          ${d ? `<div class="baja">Ya existe en el proyecto (${esc(d.ID_Riesgo)}). <label class="check"><input type="checkbox" class="rg-act" data-k="${k}"> Actualizar el existente</label></div>` : ""}</td>
        <td class="nowrap">${califTxt(r.Probabilidad_Inherente, r.Impacto_Inherente)}</td><td class="nowrap">${califTxt(r.Probabilidad_Residual, r.Impacto_Residual)}</td></tr>`;
    }).join("");
    const antes = `<div class="resumen-filtro"><b>${esc(nombreArchivo)}</b>${datos.titulo ? `<br><span class="sub">${esc(datos.titulo)}</span>` : ""}<br>${datos.riesgos.length} riesgo(s) encontrados. Desmarca los que no quieras cargar.</div>`;
    const despues = `<div class="tabla-scroll carga-riesgos"><table><thead><tr><th></th><th>Riesgo</th><th>Inherente</th><th>Residual</th></tr></thead><tbody>${filas}</tbody></table></div>`;
    modal(`Cargar riesgos · ${p.Nombre}`, campos, { ID_Seguimiento: segFecha ? segFecha.ID_Seguimiento : "", Origen: datos.instancia, Fecha_Identificacion: datos.fecha || R.hoyISO() }, (fd, extra) => {
      const nuevos = [], actualizar = [];
      extra.sel.forEach((k) => {
        const r = datos.riesgos[k];
        const fila = { ...r, ID_Seguimiento: fd.ID_Seguimiento, Origen: fd.Origen, Fecha_Identificacion: fd.Fecha_Identificacion,
          Calificacion_Inherente: R.calificacion(r.Probabilidad_Inherente, r.Impacto_Inherente) || "", Calificacion_Residual: R.calificacion(r.Probabilidad_Residual, r.Impacto_Residual) || "" };
        const d = dupDe(r);
        if (d && extra.act.includes(k)) actualizar.push({ id: d.ID_Riesgo, cambios: fila });
        else if (!d || !extra.act.includes(k)) nuevos.push(fila);
      });
      guardar(async () => {
        for (const n of [...new Set([...nuevos, ...actualizar.map((a) => a.cambios)].map((x) => x.Responsable).filter(Boolean))]) {
          const rec = await asegurarRecurso({ Nombre: n });
          if (rec) [...nuevos, ...actualizar.map((a) => a.cambios)].forEach((x) => { if (normTxt(x.Responsable) === normTxt(n)) x.Responsable = rec.Nombre; });
        }
        if (nuevos.length) await S.api.agregarFilas("Riesgos", nuevos.map((r, j) => ({ ID_Riesgo: R.siguienteIdHijo("RSG", S.datos.Riesgos, "ID_Riesgo", p.ID_Proyecto, j), ID_Proyecto: p.ID_Proyecto, ...r })));
        if (actualizar.length) await S.api.actualizarVarios("Riesgos", "ID_Riesgo", actualizar);
      }, `${nuevos.length} riesgo(s) cargado(s)${actualizar.length ? ` y ${actualizar.length} actualizado(s)` : ""}`);
    }, null, { antes, despues,
      init: (m) => m.querySelectorAll(".rg-act").forEach((c) => c.addEventListener("change", () => { if (c.checked) m.querySelector(`.rg-ok[data-k="${c.dataset.k}"]`).checked = true; })),
      recoger: (m) => {
        const sel = [...m.querySelectorAll(".rg-ok:checked")].map((c) => Number(c.dataset.k));
        if (!sel.length) return { error: "Marca al menos un riesgo para cargar." };
        const act = [...m.querySelectorAll(".rg-act:checked")].map((c) => Number(c.dataset.k));
        const soloDup = sel.filter((k) => dupDe(datos.riesgos[k]) && !act.includes(k));
        if (soloDup.length) return { error: `${soloDup.length} riesgo(s) marcado(s) ya existen: desmárcalos o elige «Actualizar el existente».` };
        return { datos: { sel, act } };
      } });
  }

  // ---------- Selector de recursos (lista desplegable con búsqueda) ----------
  // Cualquier <input data-recurso> se vuelve un buscador de la tabla Recursos al enfocarlo.
  function comboRecurso(input) {
    if (input._combo) return;
    input._combo = true;
    input.setAttribute("autocomplete", "off");
    input.removeAttribute("list");
    const caja = document.createElement("div");
    caja.className = "combo";
    input.parentNode.insertBefore(caja, input);
    caja.appendChild(input);
    const lista = document.createElement("ul");
    lista.className = "combo-lista"; lista.hidden = true; lista.setAttribute("role", "listbox");
    caja.appendChild(lista);
    let items = [], activo = -1;
    const elegir = (r, nuevo) => {
      input.value = r ? r.Nombre : nuevo;
      input.dataset.idRecurso = r ? r.ID_Recurso : "";
      lista.hidden = true;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    };
    const pintar = () => {
      const q = normTxt(input.value);
      const rs = recursos().filter((r) => r.Activo !== "No" && (!q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${empresaDe(r)}`).includes(q))).slice(0, 8);
      const exacto = rs.some((r) => normTxt(r.Nombre) === q);
      items = [...rs.map((r) => ({ r })), ...(q && !exacto ? [{ nuevo: input.value.trim() }] : [])];
      activo = items.length ? 0 : -1;
      lista.innerHTML = items.map((it, i) => it.r
        ? `<li role="option" data-i="${i}" class="${i === activo ? "activo" : ""}"><b>${esc(it.r.Nombre)}</b><span class="sub">${esc([it.r.Cargo, empresaDe(it.r), it.r.Correo].filter(Boolean).join(" · "))}</span></li>`
        : `<li role="option" data-i="${i}" class="nuevo ${i === activo ? "activo" : ""}">+ Agregar «${esc(it.nuevo)}» como recurso nuevo</li>`).join("")
        || `<li class="vacio-combo">No hay recursos${q ? " con ese texto" : ""}. Escribe el nombre para crearlo.</li>`;
      lista.hidden = false;
    };
    const marcar = () => lista.querySelectorAll("li[data-i]").forEach((li) => li.classList.toggle("activo", Number(li.dataset.i) === activo));
    input.addEventListener("focus", pintar);
    input.addEventListener("input", pintar);
    input.addEventListener("keydown", (e) => {
      if (lista.hidden) return;
      if (e.key === "ArrowDown") { e.preventDefault(); activo = Math.min(items.length - 1, activo + 1); marcar(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); activo = Math.max(0, activo - 1); marcar(); }
      else if (e.key === "Enter" && activo >= 0) { e.preventDefault(); const it = items[activo]; elegir(it.r, it.nuevo); }
      else if (e.key === "Escape") { e.stopPropagation(); lista.hidden = true; }
    });
    lista.addEventListener("mousedown", (e) => {
      const li = e.target.closest("li[data-i]");
      if (!li) return;
      e.preventDefault();
      const it = items[Number(li.dataset.i)];
      elegir(it.r, it.nuevo);
    });
    input.addEventListener("blur", () => setTimeout(() => { lista.hidden = true; }, 150));
  }
  document.addEventListener("focusin", (e) => { if (e.target.matches && e.target.matches("input[data-recurso]")) comboRecurso(e.target); });

  // ---------- Riesgos interactivos: indicadores, mapa de calor que filtra, tarjetas y filtros ----------
  const califActual = (r) => Number(r.Calificacion_Residual) || Number(r.Calificacion_Inherente) || 0;
  const nivelActual = (r) => (califActual(r) ? R.nivelRiesgo(califActual(r)) : "Sin calificar");
  function panelRiesgos(clave, todos, { conProyecto, puedeFn }) {
    S.rgF = S.rgF || {};
    const F = (S.rgF[clave] = S.rgF[clave] || { estado: "activos", nivel: "", resp: "", origen: "", q: "", celda: "", vista: "inherente", orden: "calif", extra: "" });
    const res = F.vista === "residual";
    const P = res ? "Probabilidad_Residual" : "Probabilidad_Inherente", I = res ? "Impacto_Residual" : "Impacto_Inherente";
    const base = todos.filter((r) => F.estado === "todos" || (F.estado === "activos" ? riesgoActivo(r) : r.Estado === F.estado));
    const kpi = {
      activos: todos.filter(riesgoActivo).length,
      altos: todos.filter((r) => riesgoActivo(r) && nivelActual(r) === "Alto").length,
      mater: todos.filter((r) => r.Estado === "Materializado").length,
      sinPlan: todos.filter((r) => riesgoActivo(r) && !String(r.Plan_Mitigacion || "").trim()).length,
      sinResp: todos.filter((r) => riesgoActivo(r) && !String(r.Responsable || "").trim()).length,
    };
    let lista = base.filter((r) => (!F.nivel || nivelActual(r) === F.nivel) && (!F.resp || normTxt(r.Responsable) === normTxt(F.resp)) &&
      (!F.origen || origenRiesgo(r) === F.origen) && (!F.q || normTxt(`${r.Descripcion} ${r.Plan_Mitigacion} ${r.Plan_Contingencia} ${r.Responsable}`).includes(normTxt(F.q))) &&
      (!F.celda || `${Number(r[P])}-${Number(r[I])}` === F.celda) &&
      (F.extra !== "sinPlan" || !String(r.Plan_Mitigacion || "").trim()) && (F.extra !== "sinResp" || !String(r.Responsable || "").trim()));
    const ORD = { calif: (a, b) => califActual(b) - califActual(a) || numRiesgo(a) - numRiesgo(b), num: (a, b) => (conProyecto ? String(a.ID_Proyecto).localeCompare(String(b.ID_Proyecto)) : 0) || numRiesgo(a) - numRiesgo(b),
      reciente: (a, b) => String(b.Fecha_Identificacion || "").localeCompare(String(a.Fecha_Identificacion || "")) || numRiesgo(b) - numRiesgo(a) };
    lista = lista.slice().sort(ORD[F.orden] || ORD.calif);
    // Mapa de calor sobre los riesgos del estado elegido (sin el filtro de celda).
    const enMapa = base;
    let celdas = "";
    for (let pr = 5; pr >= 1; pr--) {
      celdas += `<div class="hm-eje" title="${esc(PROB_TXT[pr])}">${pr}</div>`;
      for (let im = 1; im <= 5; im++) {
        const n = enMapa.filter((r) => Number(r[P]) === pr && Number(r[I]) === im).length;
        const k = `${pr}-${im}`;
        celdas += `<button class="hm-celda riesgo-${R.nivelRiesgo(pr * im).toLowerCase()} ${F.celda === k ? "sel" : ""} ${n ? "" : "cero"}" data-celda="${k}" title="Probabilidad ${PROB_TXT[pr]} × Impacto ${IMP_TXT[im]} = ${pr * im}: ${n} riesgo(s)" ${n ? "" : "disabled"}>${n || ""}</button>`;
      }
    }
    celdas += `<div></div>${[1, 2, 3, 4, 5].map((i) => `<div class="hm-eje" title="${esc(IMP_TXT[i])}">${i}</div>`).join("")}`;
    const responsables = [...new Set(todos.map((r) => r.Responsable).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
    const origenes = [...new Set(todos.map(origenRiesgo))].sort();
    const chip = (grupo, v, t, n) => `<button class="btn chico ${F[grupo] === v ? "primario" : ""}" data-rgf="${grupo}" data-v="${esc(v)}">${t}${n !== undefined ? ` (${n})` : ""}</button>`;
    const kpiB = (id, t, n, cls) => `<button class="kpi rg-kpi ${cls || ""} ${F.extra === id || (id === "altos" && F.nivel === "Alto") ? "sel" : ""}" data-rgkpi="${id}"><span class="kpi-t">${t}</span><span class="kpi-v">${n}</span></button>`;
    const tarjeta = (r) => {
      const ci = R.calificacion(r.Probabilidad_Inherente, r.Impacto_Inherente), cr = R.calificacion(r.Probabilidad_Residual, r.Impacto_Residual);
      const baja = ci && cr ? Math.round((1 - cr / ci) * 100) : null;
      const puede = puedeFn(r);
      const p = proyecto(r.ID_Proyecto) || {};
      return `<div class="rg-card nivel-${normTxt(nivelActual(r)).replace(/\s/g, "")}">
        <div class="rg-top"><span class="rg-num">#${numRiesgo(r)}</span>${pill(r.Tipo || "Amenaza")}
          ${puede ? `<select class="rg-estado" data-rgest="${esc(r.ID_Riesgo)}" aria-label="Estado del riesgo">${S.cat.Estado_Riesgo.map((e) => `<option ${e === (r.Estado || "Abierto") ? "selected" : ""}>${esc(e)}</option>`).join("")}</select>` : pill(r.Estado || "Abierto")}
          ${conProyecto ? `<button class="btn enlace rg-proy" data-pid="${esc(r.ID_Proyecto)}">${esc(p.Nombre || r.ID_Proyecto)}</button>` : ""}
          <span class="rg-acc">${puede ? `<button class="btn chico" data-riesgo="${esc(r.ID_Riesgo)}" data-rpid="${esc(r.ID_Proyecto)}">Editar</button>` : ""}</span></div>
        <div class="rg-desc">${esc(r.Descripcion)}</div>
        <div class="rg-evol">
          <span class="rg-cal ${ci ? R.nivelRiesgo(ci).toLowerCase() : ""}" title="Antes de mitigar: probabilidad ${esc(PROB_TXT[r.Probabilidad_Inherente] || "—")} × impacto ${esc(IMP_TXT[r.Impacto_Inherente] || "—")}">Inherente <b>${ci || "—"}</b> ${ci ? R.nivelRiesgo(ci) : ""}</span>
          <span class="rg-flecha">→</span>
          <span class="rg-cal ${cr ? R.nivelRiesgo(cr).toLowerCase() : "vacia"}" title="Después de mitigar: probabilidad ${esc(PROB_TXT[r.Probabilidad_Residual] || "—")} × impacto ${esc(IMP_TXT[r.Impacto_Residual] || "—")}">Residual <b>${cr || "—"}</b> ${cr ? R.nivelRiesgo(cr) : ""}</span>
          ${baja !== null ? `<span class="sub">${baja > 0 ? `la mitigación lo baja ${baja}%` : baja < 0 ? "⚠ el residual es mayor" : "sin reducción"}</span>` : ""}
        </div>
        <div class="rg-meta"><span>${r.Responsable ? persona(r.Responsable) : `<span class="baja">Sin responsable</span>`}</span><span>📍 ${esc(origenRiesgo(r))}</span></div>
        ${r.Plan_Mitigacion || r.Plan_Contingencia ? `<details class="rg-planes"><summary>Planes de mitigación y contingencia</summary>
          ${r.Plan_Mitigacion ? `<div><b>Mitigación:</b> ${esc(r.Plan_Mitigacion)}</div>` : ""}${r.Plan_Contingencia ? `<div><b>Contingencia:</b> ${esc(r.Plan_Contingencia)}</div>` : ""}</details>`
          : riesgoActivo(r) ? `<div class="baja sub">⚠ Sin plan de mitigación</div>` : ""}
      </div>`;
    };
    const hayFiltro = F.nivel || F.resp || F.origen || F.q || F.celda || F.extra;
    return `
      <div class="kpis rg-kpis">${kpiB("activos", "Activos", kpi.activos)}${kpiB("altos", "Nivel alto (hoy)", kpi.altos, kpi.altos ? "alerta" : "")}${kpiB("mater", "Materializados", kpi.mater, kpi.mater ? "alerta" : "")}${kpiB("sinPlan", "Sin plan de mitigación", kpi.sinPlan, kpi.sinPlan ? "alerta" : "")}${kpiB("sinResp", "Sin responsable", kpi.sinResp, kpi.sinResp ? "alerta" : "")}</div>
      <div class="rg-layout">
        <div class="card rg-mapa"><div class="titulo-fila"><h2>Mapa de calor</h2>
          <div class="seg">${chip("vista", "inherente", "Inherente")}${chip("vista", "residual", "Residual")}</div></div>
          <div class="heatmap">${celdas}</div>
          <div class="sub centro">→ Impacto · ↑ Probabilidad. Haz clic en una casilla para ver esos riesgos.</div>
          <div class="hm-ley"><span class="riesgo-bajo">Bajo ≤ 6</span><span class="riesgo-medio">Medio 7–12</span><span class="riesgo-alto">Alto &gt; 12</span></div>
        </div>
        <div class="card rg-lista">
          <div class="rg-filtros">
            <div class="seg">${chip("estado", "activos", "Activos", todos.filter(riesgoActivo).length)}${S.cat.Estado_Riesgo.filter((e) => todos.some((r) => r.Estado === e)).map((e) => chip("estado", e, e, todos.filter((r) => r.Estado === e).length)).join("")}${chip("estado", "todos", "Todos", todos.length)}</div>
            <div class="seg">${["Alto", "Medio", "Bajo"].map((n) => chip("nivel", n, n)).join("")}</div>
            <div class="rg-filtros-2">
              <label class="filtro crece"><span>Buscar</span><input type="search" data-rgin="q" value="${esc(F.q)}" placeholder="Texto del riesgo o de los planes"></label>
              <label class="filtro"><span>Responsable</span><select data-rgin="resp"><option value="">Todos</option>${responsables.map((x) => `<option ${x === F.resp ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
              <label class="filtro"><span>Origen</span><select data-rgin="origen"><option value="">Todos</option>${origenes.map((x) => `<option ${x === F.origen ? "selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
              <label class="filtro"><span>Ordenar</span><select data-rgin="orden"><option value="calif" ${F.orden === "calif" ? "selected" : ""}>Más críticos primero</option><option value="num" ${F.orden === "num" ? "selected" : ""}>Por número</option><option value="reciente" ${F.orden === "reciente" ? "selected" : ""}>Más recientes</option></select></label>
            </div>
            <div class="sub">${lista.length} de ${todos.length} riesgo(s)${F.celda ? ` · casilla P${F.celda.split("-")[0]} × I${F.celda.split("-")[1]}` : ""} ${hayFiltro ? `<button class="btn enlace" data-rglimpiar="1">Quitar filtros</button>` : ""}</div>
          </div>
          ${lista.length ? `<div class="rg-cards">${lista.map(tarjeta).join("")}</div>` : vacio(todos.length ? "Ningún riesgo coincide con los filtros." : "Aún no hay riesgos registrados.")}
        </div>
      </div>`;
  }
  function enlazarPanelRiesgos(el, clave, alEditar) {
    const F = S.rgF[clave];
    el.querySelectorAll("[data-rgf]").forEach((b) => b.addEventListener("click", () => {
      const g = b.dataset.rgf, v = b.dataset.v;
      F[g] = g === "nivel" && F.nivel === v ? "" : v;
      if (g === "vista") F.celda = "";
      render();
    }));
    el.querySelectorAll("[data-celda]").forEach((b) => b.addEventListener("click", () => { F.celda = F.celda === b.dataset.celda ? "" : b.dataset.celda; render(); }));
    el.querySelectorAll("[data-rgkpi]").forEach((b) => b.addEventListener("click", () => {
      const k = b.dataset.rgkpi;
      if (k === "activos") Object.assign(F, { estado: "activos", nivel: "", extra: "", celda: "" });
      else if (k === "altos") Object.assign(F, { estado: "activos", nivel: F.nivel === "Alto" ? "" : "Alto", extra: "" });
      else if (k === "mater") Object.assign(F, { estado: "Materializado", nivel: "", extra: "" });
      else Object.assign(F, { estado: "activos", extra: F.extra === k ? "" : k });
      render();
    }));
    el.querySelectorAll("[data-rgin]").forEach((i) => i.addEventListener("change", () => { F[i.dataset.rgin] = i.value; render(); }));
    const l = el.querySelector("[data-rglimpiar]");
    if (l) l.addEventListener("click", () => { Object.assign(F, { nivel: "", resp: "", origen: "", q: "", celda: "", extra: "" }); render(); });
    el.querySelectorAll("[data-rgest]").forEach((s) => s.addEventListener("change", () => {
      const r = S.datos.Riesgos.find((x) => x.ID_Riesgo === s.dataset.rgest);
      guardar(() => S.api.actualizarPorId("Riesgos", "ID_Riesgo", r.ID_Riesgo, { Estado: s.value }), `Riesgo #${numRiesgo(r)}: ${s.value}`);
    }));
    if (alEditar) el.querySelectorAll(".rg-card [data-riesgo]").forEach((b) => b.addEventListener("click", () => alEditar(S.datos.Riesgos.find((x) => x.ID_Riesgo === b.dataset.riesgo))));
  }

  // ---------- Íconos (trazos estilo Lucide, dibujados en línea) ----------
  const ICONOS = {
    dashboard: '<rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/>',
    avances: '<polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/>',
    seguimiento: '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="m9 14 2 2 4-4"/>',
    proyectos: '<path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
    cronograma: '<path d="M3 3v18h18"/><rect x="7" y="6" width="9" height="3" rx="1"/><rect x="10" y="11" width="8" height="3" rx="1"/><rect x="6" y="16" width="6" height="3" rx="1"/>',
    riesgos: '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    recursos: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    capacidad: '<path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>',
    priorizacion: '<path d="M3 3v18h18"/><path d="M7 16h4"/><path d="M7 11h8"/><path d="M7 6h12"/>',
    lecciones: '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
    catalogos: '<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M2 14h4M10 8h4M18 16h4"/>',
    usuarios: '<circle cx="12" cy="8" r="4"/><path d="M6 21v-2a6 6 0 0 1 12 0v2"/>',
    auditoria: '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l4 2"/>',
    buscar: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>',
    campana: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
    ayuda: '<circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><path d="M12 17h.01"/>',
    actualizar: '<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>',
    salir: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" x2="9" y1="12" y2="12"/>',
    luna: '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
    sol: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/>',
    menu: '<path d="M4 6h16M4 12h16M4 18h16"/>',
    colapsar: '<path d="m11 17-5-5 5-5"/><path d="m18 17-5-5 5-5"/>',
    expandir: '<path d="m6 17 5-5-5-5"/><path d="m13 17 5-5-5-5"/>',
    vacio: '<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    compromiso: '<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>',
    ticket: '<path d="M2 9a3 3 0 0 1 0 6v2a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-2a3 3 0 0 1 0-6V7a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/><path d="M13 5v2M13 17v2M13 11v2"/>',
    ir: '<path d="M5 12h14"/><path d="m12 5 7 7-7 7"/>',
    backlog: '<path d="M3 5h18M3 12h18M3 19h12"/><circle cx="19" cy="19" r="2"/>',
    exportar: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/>',
  };
  const ico = (n, cls = "") => `<svg class="ico ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONOS[n] || ""}</svg>`;

  // Avatar con iniciales y un color estable por persona.
  function avatar(nombre, tam = "") {
    const n = String(nombre || "?").replace(/@.*/, "").trim();
    const partes = n.split(/[\s._-]+/).filter(Boolean);
    const ini = ((partes[0] || "?")[0] + (partes.length > 1 ? partes[partes.length - 1][0] : "")).toUpperCase();
    let h = 0; for (const ch of n) h = (h * 31 + ch.charCodeAt(0)) % 360;
    return `<span class="avatar ${tam}" style="--av:${h}" title="${esc(nombre)}" aria-hidden="true">${esc(ini)}</span>`;
  }
  const persona = (nombre, sub) => `<span class="persona">${avatar(nombre)}<span><span class="persona-n">${esc(nombre || "—")}</span>${sub ? `<span class="sub">${esc(sub)}</span>` : ""}</span></span>`;

  // Tema claro / oscuro (se recuerda en este equipo).
  function aplicarTema(t) { document.documentElement.dataset.tema = t; guardarLocal("pmo_tema", t); }
  aplicarTema(leerLocal("pmo_tema", "claro"));

  // Menú agrupado como en las suites empresariales.
  const GRUPOS_NAV = [
    ["Inicio", ["dashboard", "avances", "seguimiento"]],
    ["Portafolio", ["backlog", "proyectos", "cronograma", "riesgos", "priorizacion"]],
    ["Personas", ["recursos", "capacidad"]],
    ["Conocimiento", ["lecciones"]],
    ["Administración", ["catalogos", "usuarios", "auditoria"]],
  ];
  const NOMBRE_TAB = { resumen: "Resumen", compromisos: "Compromisos", seguimientos: "Seguimientos", hitos: "Hitos", riesgos: "Riesgos", tickets: "Tickets", cambios: "Cambios", raci: "RACI", lecciones: "Lecciones", auditoria: "Historial de cambios" };
  function migas() {
    const v = VISTAS.find((x) => x.id === (S.vista === "ficha" ? "proyectos" : S.vista));
    const grupo = (GRUPOS_NAV.find((g) => g[1].includes(v ? v.id : "")) || ["Inicio"])[0];
    const partes = [`<span>${esc(grupo)}</span>`, S.vista === "ficha" ? `<button class="btn enlace" data-miga="proyectos">Proyectos</button>` : `<span>${esc(v ? v.t : "")}</span>`];
    if (S.vista === "ficha") { const p = proyecto(S.pid); if (p) partes.push(`<span>${esc(p.Nombre)}</span>`, `<span>${esc(NOMBRE_TAB[S.fichaTab] || "Resumen")}</span>`); }
    return `<nav class="migas" aria-label="Ubicación">${partes.join('<span class="sep">›</span>')}</nav>`;
  }

  // Ordenar cualquier tabla con clic en el encabezado (ignora tablas agrupadas o de edición).
  document.addEventListener("click", (e) => {
    const th = e.target.closest && e.target.closest("#vista table:not(.raci):not(.no-orden) thead th");
    if (!th || e.target.closest("button, a, input, select")) return;
    const tabla = th.closest("table"), tb = tabla.tBodies[0];
    if (!tb || tb.querySelector("tr.grupo, td[colspan]")) return;
    const col = [...th.parentNode.children].indexOf(th);
    const dir = th.dataset.orden === "asc" ? "desc" : "asc";
    tabla.querySelectorAll("thead th").forEach((x) => delete x.dataset.orden);
    th.dataset.orden = dir;
    const val = (tr) => { const t = (tr.children[col] ? tr.children[col].innerText : "").trim(); const f = t.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (f) return `${f[3]}${f[2]}${f[1]}`; const n = parseFloat(t.replace(/[$.\s]/g, "").replace(",", ".")); return /^[$\d]/.test(t) && !Number.isNaN(n) ? n : normTxt(t); };
    const filas = [...tb.rows].sort((a, b) => { const x = val(a), y = val(b); const r = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "es", { numeric: true }); return dir === "asc" ? r : -r; });
    filas.forEach((f) => tb.appendChild(f));
  });

  // ---------- Búsqueda global (Ctrl+K) ----------
  function indiceBusqueda() {
    const out = [];
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const nom = (pid) => (proyecto(pid) || {}).Nombre || pid;
    VISTAS.filter((v) => !v.permiso || R.puede(S.usuario, v.permiso)).forEach((v) => out.push({ tipo: "Ir a", ico: v.id, t: v.t, sub: "Pantalla", ir: () => ir(v.id) }));
    visibles().forEach((p) => out.push({ tipo: "Proyectos", ico: "proyectos", t: p.Nombre, sub: [p.ID_Proyecto, p.Codigo_Almera ? `Almera ${p.Codigo_Almera}` : "", nombreUsuario(p.PM), p.Estado].filter(Boolean).join(" · "), ir: () => { S.fichaTab = "resumen"; ir("ficha", p.ID_Proyecto); } }));
    demandas().forEach((d) => out.push({ tipo: "Demandas", ico: "backlog", t: `${d.Codigo_Almera} · ${d.Nombre}`, sub: `${d.Estado || "Recibida"} · ${d.Area || ""} · puntaje ${puntajeDemanda(d) ?? "—"}`, ir: () => { ir("backlog"); formDemanda(d); } }));
    recursos().forEach((r) => out.push({ tipo: "Personas", ico: "recursos", t: r.Nombre, sub: [r.Cargo, empresaDe(r), r.Correo].filter(Boolean).join(" · "), ir: () => { S.recQ = r.Nombre; ir("recursos"); } }));
    S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto)).forEach((c) => out.push({ tipo: "Compromisos", ico: "compromiso", t: c.Compromiso, sub: `${nom(c.ID_Proyecto)} · ${c.Responsable || "sin responsable"} · ${R.estadoCompromiso(c)}`, peso: R.cerrado(c) ? 1 : 0,
      ir: () => { S.fichaTab = "compromisos"; ir("ficha", c.ID_Proyecto); detalleCompromiso(c.ID_Compromiso); } }));
    S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto)).forEach((r) => out.push({ tipo: "Riesgos", ico: "riesgos", t: r.Descripcion, sub: `${nom(r.ID_Proyecto)} · #${numRiesgo(r)} · ${r.Estado} · ${nivelActual(r)}`, peso: riesgoActivo(r) ? 0 : 1,
      ir: () => { S.fichaTab = "riesgos"; ir("ficha", r.ID_Proyecto); formRiesgo(proyecto(r.ID_Proyecto), r); } }));
    (S.datos.Tickets || []).filter((t) => ids.has(t.ID_Proyecto)).forEach((t) => out.push({ tipo: "Tickets", ico: "ticket", t: `${t.Numero || ""} ${t.Titulo || ""}`.trim(), sub: `${nom(t.ID_Proyecto)} · ${t.Estado}`, ir: () => { S.fichaTab = "tickets"; ir("ficha", t.ID_Proyecto); } }));
    (S.datos.Lecciones || []).filter((l) => ids.has(l.ID_Proyecto)).forEach((l) => out.push({ tipo: "Lecciones", ico: "lecciones", t: l.Leccion, sub: `${nom(l.ID_Proyecto)} · ${l.Categoria || ""}`, ir: () => { S.fichaTab = "lecciones"; ir("ficha", l.ID_Proyecto); } }));
    return out;
  }
  function abrirBusqueda() {
    if (!S.usuario || $("#paleta")) return;
    const indice = indiceBusqueda();
    const o = document.createElement("div");
    o.id = "paleta";
    o.innerHTML = `<div class="paleta-caja" role="dialog" aria-modal="true" aria-label="Búsqueda global">
      <div class="paleta-in">${ico("buscar")}<input id="paleta-q" placeholder="Buscar proyectos, personas, compromisos, riesgos, tickets o pantallas…" autocomplete="off"><kbd>Esc</kbd></div>
      <div id="paleta-res" class="paleta-res" role="listbox"></div>
      <div class="paleta-pie"><span><kbd>↑</kbd><kbd>↓</kbd> moverse</span><span><kbd>Enter</kbd> abrir</span><span><kbd>Esc</kbd> cerrar</span></div></div>`;
    document.body.appendChild(o);
    const q = $("#paleta-q"), res = $("#paleta-res");
    let vis = [], act = 0;
    const cerrar = () => o.remove();
    const pintar = () => {
      const terms = normTxt(q.value).split(" ").filter(Boolean);
      const orden = ["Ir a", "Proyectos", "Demandas", "Personas", "Compromisos", "Riesgos", "Tickets", "Lecciones"];
      vis = (terms.length ? indice.filter((x) => { const h = normTxt(`${x.t} ${x.sub}`); return terms.every((t) => h.includes(t)); }) : indice.filter((x) => x.tipo === "Ir a" || x.tipo === "Proyectos"))
        .sort((a, b) => orden.indexOf(a.tipo) - orden.indexOf(b.tipo) || (a.peso || 0) - (b.peso || 0));
      const porTipo = {}; vis.forEach((x) => { (porTipo[x.tipo] = porTipo[x.tipo] || []).push(x); });
      vis = Object.values(porTipo).flatMap((l) => l.slice(0, terms.length ? 6 : 12));
      act = Math.min(act, Math.max(0, vis.length - 1));
      let i = -1, html = "";
      Object.keys(porTipo).forEach((tipo) => {
        const l = vis.filter((x) => x.tipo === tipo);
        if (!l.length) return;
        html += `<div class="paleta-grupo">${esc(tipo)}</div>` + l.map((x) => { i++; return `<div class="paleta-item ${i === act ? "activo" : ""}" data-i="${i}" role="option">${ico(x.ico)}<div><div class="paleta-t">${esc(x.t)}</div><div class="sub">${esc(x.sub)}</div></div>${ico("ir", "paleta-ir")}</div>`; }).join("");
      });
      res.innerHTML = html || `<div class="vacio">Sin resultados para «${esc(q.value)}».</div>`;
      const a = res.querySelector(".activo"); if (a) a.scrollIntoView({ block: "nearest" });
    };
    const abrir = (i) => { const x = vis[i]; if (!x) return; cerrar(); x.ir(); };
    q.addEventListener("input", () => { act = 0; pintar(); });
    q.addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown") { e.preventDefault(); act = Math.min(vis.length - 1, act + 1); pintar(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); act = Math.max(0, act - 1); pintar(); }
      else if (e.key === "Enter") { e.preventDefault(); abrir(act); }
      else if (e.key === "Escape") cerrar();
    });
    res.addEventListener("click", (e) => { const it = e.target.closest(".paleta-item"); if (it) abrir(Number(it.dataset.i)); });
    o.addEventListener("mousedown", (e) => { if (e.target === o) cerrar(); });
    pintar(); q.focus();
  }
  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") { e.preventDefault(); abrirBusqueda(); }
  });
  document.addEventListener("click", (e) => { const mu = document.getElementById("menu-usuario"); if (mu && !mu.hidden && !e.target.closest(".ab-usuario")) mu.hidden = true; });

  // ---------- Exportar a Excel (con formato) por módulo, respetando filtros ----------
  let cargaExcelJS = null;
  function libExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!cargaExcelJS) cargaExcelJS = new Promise((ok, falla) => {
      const s = document.createElement("script");
      s.src = new URL("lib/exceljs.min.js", location.href).href;
      s.onload = () => ok(window.ExcelJS); s.onerror = () => falla(new Error("No se pudo cargar el generador de Excel."));
      document.head.appendChild(s);
    });
    return cargaExcelJS;
  }
  // Riesgos tal como se ven con los filtros del panel (clave "global" o "p:<ID>").
  function riesgosVisibles(clave, todos) {
    const F = (S.rgF && S.rgF[clave]) || { estado: "activos", vista: "inherente" };
    const P = F.vista === "residual" ? "Probabilidad_Residual" : "Probabilidad_Inherente", I = F.vista === "residual" ? "Impacto_Residual" : "Impacto_Inherente";
    return todos.filter((r) => (F.estado === "todos" || (F.estado === "activos" ? riesgoActivo(r) : r.Estado === F.estado)) &&
      (!F.nivel || nivelActual(r) === F.nivel) && (!F.resp || normTxt(r.Responsable) === normTxt(F.resp)) && (!F.origen || origenRiesgo(r) === F.origen) &&
      (!F.q || normTxt(`${r.Descripcion} ${r.Plan_Mitigacion} ${r.Plan_Contingencia} ${r.Responsable}`).includes(normTxt(F.q))) &&
      (!F.celda || `${Number(r[P])}-${Number(r[I])}` === F.celda) &&
      (F.extra !== "sinPlan" || !String(r.Plan_Mitigacion || "").trim()) && (F.extra !== "sinResp" || !String(r.Responsable || "").trim()))
      .sort((a, b) => String(a.ID_Proyecto).localeCompare(String(b.ID_Proyecto)) || numRiesgo(a) - numRiesgo(b));
  }
  function describirFiltrosRiesgo(clave) {
    const F = (S.rgF && S.rgF[clave]) || {};
    return [F.estado && F.estado !== "todos" ? `Estado: ${F.estado}` : "", F.nivel ? `Nivel: ${F.nivel}` : "", F.resp ? `Responsable: ${F.resp}` : "", F.origen ? `Origen: ${F.origen}` : "",
      F.q ? `Búsqueda: «${F.q}»` : "", F.celda ? `Casilla del mapa P${F.celda.replace("-", " × I")} (${F.vista || "inherente"})` : "", F.extra === "sinPlan" ? "Sin plan de mitigación" : "", F.extra === "sinResp" ? "Sin responsable" : ""].filter(Boolean);
  }
  const filtrosGenerales = () => { const v = describirFiltro({ filtros: S.filtros, lista: S.filtrosLista }); return v.startsWith("Sin filtros") ? [] : [v]; };
  // Colores por valor (semáforo, nivel, estado) para resaltar celdas.
  const COLOR_CELDA = {
    Verde: "C6EFCE", Amarillo: "FFEB9C", Rojo: "FFC7CE", Alto: "FFC7CE", Medio: "FFEB9C", Bajo: "C6EFCE",
    Vencido: "FFC7CE", "Por vencer": "FFEB9C", Cerrado: "C6EFCE", Cumplido: "C6EFCE", "En curso": "DDEBF7", Pendiente: "F2F2F2",
    Sobreasignado: "FFC7CE", "Al límite": "FFEB9C", Disponible: "C6EFCE", Aprobado: "C6EFCE", Rechazado: "FFC7CE", Solicitado: "FFEB9C",
    Materializado: "FFC7CE", "En seguimiento": "DDEBF7", Abierto: "FFEB9C", Resuelto: "C6EFCE", Atrasado: "FFC7CE",
  };
  const aFechaX = (v) => { const m = String(v || "").match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/); return m ? new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0))) : v || ""; };
  /* hojas: [{ nombre, columnas: [{ t, v: (fila) => valor, tipo?: "fecha"|"fechahora"|"cop"|"pct"|"num"|"texto", ancho?, color?: true }], filas }] */
  async function exportarExcel(modulo, hojas, filtros) {
    cargando(true, "Generando Excel…");
    try {
      const X = await libExcelJS();
      const wb = new X.Workbook();
      wb.creator = "Portafolio PMO · FSFB"; wb.created = new Date();
      const info = wb.addWorksheet("Info");
      info.columns = [{ width: 26 }, { width: 90 }];
      const filasInfo = [["Portafolio de Proyectos · FSFB", ""], ["Módulo", modulo], ["Generado", new Date().toLocaleString("es-CO")], ["Por", `${S.usuario.Nombre || ""} <${S.usuario.Correo}>`],
        ["Filtros aplicados", filtros.length ? filtros.join(" · ") : "Ninguno (todo lo visible para tu rol)"], ["Contenido", hojas.map((h) => `${h.nombre}: ${h.filas.length} fila(s)`).join(" · ")]];
      filasInfo.forEach((f) => info.addRow(f));
      info.getCell("A1").font = { bold: true, size: 14, color: { argb: "FF104994" } };
      for (let r = 2; r <= filasInfo.length; r++) { info.getCell(`A${r}`).font = { bold: true, color: { argb: "FF5F6B7A" } }; info.getCell(`B${r}`).alignment = { wrapText: true, vertical: "top" }; }
      for (const h of hojas) {
        const ws = wb.addWorksheet(h.nombre.slice(0, 31), { views: [{ state: "frozen", ySplit: 1 }] });
        ws.columns = h.columnas.map((c) => ({ header: c.t, width: c.ancho || (c.tipo === "fecha" ? 12 : c.tipo === "cop" ? 16 : c.tipo === "pct" || c.tipo === "num" ? 11 : 22) }));
        h.filas.forEach((f) => ws.addRow(h.columnas.map((c) => {
          let v = c.v(f);
          if (v === undefined || v === null) v = "";
          if (c.tipo === "fecha" || c.tipo === "fechahora") return aFechaX(v);
          if (c.tipo === "cop" || c.tipo === "num") return v === "" ? "" : Number(v) || 0;
          if (c.tipo === "pct") return v === "" ? "" : (Number(v) || 0) / 100;
          return String(v);
        })));
        const cab = ws.getRow(1);
        cab.height = 22;
        cab.eachCell((cel) => { cel.font = { bold: true, color: { argb: "FFFFFFFF" } }; cel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF104994" } }; cel.alignment = { vertical: "middle", wrapText: true }; });
        h.columnas.forEach((c, i) => {
          const col = ws.getColumn(i + 1);
          if (c.tipo === "fecha") col.numFmt = "dd/mm/yyyy";
          if (c.tipo === "fechahora") col.numFmt = "dd/mm/yyyy hh:mm";
          if (c.tipo === "cop") col.numFmt = '"$"#,##0';
          if (c.tipo === "pct") col.numFmt = "0%";
          if ((c.ancho || 22) >= 30) col.alignment = { wrapText: true, vertical: "top" };
        });
        ws.eachRow((row, n) => {
          if (n === 1) return;
          row.alignment = { vertical: "top", wrapText: true };
          h.columnas.forEach((c, i) => {
            const cel = row.getCell(i + 1);
            cel.border = { bottom: { style: "thin", color: { argb: "FFE3E8EE" } } };
            const k = c.color && COLOR_CELDA[String(cel.value)];
            if (k) cel.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + k } };
          });
        });
        if (h.filas.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: h.columnas.length } };
      }
      const buf = await wb.xlsx.writeBuffer();
      const nombre = `Portafolio_${modulo.replace(/[^\wÁÉÍÓÚáéíóúÑñ]+/g, "_")}_${R.hoyISO()}.xlsx`;
      descargar(new Uint8Array(buf), nombre, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      toast(`Excel generado: ${nombre}`);
    } catch (e) { toast("No se pudo generar el Excel: " + e.message + (S.modo === "panel" ? " Si el panel bloquea la descarga, abre la app en ventana grande." : ""), true); }
    finally { cargando(false); }
  }
  // Riesgos en el formato FSFB (plantilla con listas desplegables), para editar y volver a cargar.
  async function exportarRiesgosFSFB(lista, titulo, filtros) {
    cargando(true, "Generando Excel de riesgos…");
    try {
      const X = await libExcelJS();
      const wb = new X.Workbook();
      await wb.xlsx.load(await (await fetch(new URL("assets/Plantilla_Riesgos.xlsx", location.href).href)).arrayBuffer());
      const ws = wb.getWorksheet("Riesgos") || wb.worksheets[0];
      ws.getCell("A1").value = titulo;
      const varios = new Set(lista.map((r) => r.ID_Proyecto)).size > 1;
      ws.getCell("P3").value = "Origen"; ws.getCell("Q3").value = "Fecha identificación";
      if (varios) ws.getCell("R3").value = "Proyecto";
      ["P3", "Q3", "R3"].forEach((k) => { const c = ws.getCell(k); if (c.value) { c.style = { ...ws.getCell("O3").style }; } });
      ws.getColumn(16).width = 34; ws.getColumn(17).width = 14; if (varios) ws.getColumn(18).width = 30;
      lista.forEach((r, i) => {
        const f = 4 + i;
        const ci = R.calificacion(r.Probabilidad_Inherente, r.Impacto_Inherente), cr = R.calificacion(r.Probabilidad_Residual, r.Impacto_Residual);
        const vals = [numRiesgo(r), r.Tipo || "Amenaza", r.Estado || "Abierto", r.Responsable || "", r.Descripcion || "", PROB_TXT[r.Probabilidad_Inherente] || "", IMP_TXT[r.Impacto_Inherente] || "",
          ci || "", ci ? R.nivelRiesgo(ci) : "", r.Plan_Mitigacion || "", r.Plan_Contingencia || "", PROB_TXT[r.Probabilidad_Residual] || "", IMP_TXT[r.Impacto_Residual] || "",
          cr || "", cr ? R.nivelRiesgo(cr) : "", origenRiesgo(r), aFechaX(r.Fecha_Identificacion), ...(varios ? [(proyecto(r.ID_Proyecto) || {}).Nombre || r.ID_Proyecto] : [])];
        vals.forEach((v, j) => {
          const c = ws.getCell(f, j + 1); c.value = v;
          if (f > 4 && ws.getCell(4, j + 1).style) c.style = { ...ws.getCell(4, j + 1).style };
          c.alignment = { wrapText: true, vertical: "top" };
          if ((j === 8 || j === 14) && COLOR_CELDA[v]) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + COLOR_CELDA[v] } };
        });
        ws.getCell(f, 17).numFmt = "dd/mm/yyyy";
      });
      const info = wb.addWorksheet("Info");
      info.columns = [{ width: 22 }, { width: 90 }];
      [["Generado", new Date().toLocaleString("es-CO")], ["Por", `${S.usuario.Nombre || ""} <${S.usuario.Correo}>`], ["Filtros", filtros.length ? filtros.join(" · ") : "Ninguno"], ["Riesgos", String(lista.length)]].forEach((x) => info.addRow(x));
      const buf = await wb.xlsx.writeBuffer();
      const nombre = `Riesgos_${titulo.replace(/.*Proyecto\s*/i, "").replace(/[^\wÁÉÍÓÚáéíóúÑñ]+/g, "_").slice(0, 60)}_${R.hoyISO()}.xlsx`;
      descargar(new Uint8Array(buf), nombre, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      toast(`Excel generado: ${nombre}`);
    } catch (e) { toast("No se pudo generar el Excel: " + e.message, true); }
    finally { cargando(false); }
  }

  // Columnas reutilizables.
  const C = {
    proy: { t: "Proyecto", v: (x) => (proyecto(x.ID_Proyecto) || {}).Nombre || x.ID_Proyecto, ancho: 30 },
    idp: { t: "ID proyecto", v: (x) => x.ID_Proyecto, ancho: 12 },
  };
  const colsProyecto = [
    { t: "ID", v: (p) => p.ID_Proyecto, ancho: 11 }, { t: "Código Almera", v: (p) => p.Codigo_Almera, ancho: 14 }, { t: "Proyecto", v: (p) => p.Nombre, ancho: 34 },
    { t: "Cliente / área", v: (p) => p.Cliente_Area }, { t: "PM", v: (p) => nombreUsuario(p.PM) }, { t: "Metodología", v: (p) => p.Metodologia, ancho: 14 },
    { t: "Fase", v: (p) => p.Fase, ancho: 14 }, { t: "Estado", v: (p) => p.Estado, ancho: 12, color: true }, { t: "Semáforo", v: (p) => p.Semaforo, ancho: 11, color: true },
    { t: "Prioridad", v: (p) => p.Prioridad, ancho: 11 }, { t: "Avance real", v: (p) => p.Avance_Real, tipo: "pct" }, { t: "Avance planeado", v: (p) => p.Avance_Planeado, tipo: "pct" },
    { t: "Inicio", v: (p) => p.Fecha_Inicio, tipo: "fecha" }, { t: "Fin planeado", v: (p) => p.Fecha_Fin_Plan, tipo: "fecha" }, { t: "Fin línea base", v: (p) => p.Fecha_Fin_Base, tipo: "fecha" },
    { t: "Presupuesto", v: (p) => p.Presupuesto, tipo: "cop" }, { t: "Ejecutado", v: (p) => p.Ejecutado, tipo: "cop" },
    { t: "% ejecutado", v: (p) => (Number(p.Presupuesto) ? Math.round((Number(p.Ejecutado) || 0) / Number(p.Presupuesto) * 100) : ""), tipo: "pct" },
    { t: "Frecuencia", v: (p) => p.Frecuencia_Seguimiento, ancho: 12 }, { t: "Seguimiento", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).estado, ancho: 12, color: true },
    { t: "Último seguimiento", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).ultimaFecha, tipo: "fecha" },
    { t: "Proveedores", v: (p) => idsLista(p.Proveedores).map((i) => (proveedor(i) || {}).Nombre || i).join(", "), ancho: 26 },
    { t: "Puntaje prioridad", v: (p) => (puntaje(p) === null ? "" : puntaje(p)), tipo: "num" },
    { t: "Comentario de estado", v: (p) => p.Comentario_Estado, ancho: 44 },
  ];
  const colsCompromiso = (conProy) => [
    ...(conProy ? [C.proy] : []), { t: "Compromiso", v: (c) => c.Compromiso, ancho: 46 },
    { t: "Sesión", v: (c) => { const s = sesionDe(c); return s ? `Sesión del ${fecha(s.Fecha_Corte)}` : "Sin sesión"; }, ancho: 18 },
    { t: "Responsable", v: (c) => c.Responsable }, { t: "Correo responsable", v: (c) => correoDe(c), ancho: 26 },
    { t: "Fecha límite", v: (c) => c.Fecha_Compromiso, tipo: "fecha" }, { t: "Estado", v: (c) => R.estadoBase(c), ancho: 12, color: true },
    { t: "Alerta", v: (c) => { const e = R.estadoCompromiso(c); return e === "Vencido" || e === "Por vencer" ? e : ""; }, ancho: 12, color: true },
    { t: "Fecha de cierre", v: (c) => c.Fecha_Cierre, tipo: "fecha" },
    { t: "Último comentario", v: (c) => { const l = comentariosDe(c.ID_Compromiso); const u = l[l.length - 1]; return u ? resumenCom(u) : ""; }, ancho: 46 },
    { t: "Fecha último comentario", v: (c) => { const l = comentariosDe(c.ID_Compromiso); const u = l[l.length - 1]; return u ? u.Fecha_Hora : ""; }, tipo: "fechahora", ancho: 16 },
  ];
  const colsSeguimiento = (conProy) => [
    ...(conProy ? [C.proy] : []), { t: "Fecha de corte", v: (s) => s.Fecha_Corte, tipo: "fecha" }, { t: "Semana", v: (s) => s.Semana, ancho: 10 },
    { t: "Avance real", v: (s) => s.Avance_Real, tipo: "pct" }, { t: "Semáforo", v: (s) => s.Semaforo, ancho: 11, color: true }, { t: "Ejecutado", v: (s) => s.Ejecutado, tipo: "cop" },
    { t: "Logros", v: (s) => s.Logros, ancho: 44 }, { t: "Próximos pasos", v: (s) => s.Proximos_Pasos, ancho: 40 }, { t: "Bloqueos", v: (s) => s.Bloqueos, ancho: 36 },
    { t: "Compromisos acordados", v: (s) => S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento).length, tipo: "num" },
    { t: "Fecha acta", v: (s) => s.Fecha_Acta, tipo: "fecha" }, { t: "Acta", v: (s) => s.Acta_Archivo || s.URL_Acta, ancho: 30 }, { t: "Reportado por", v: (s) => nombreUsuario(s.Reportado_Por) },
  ];
  const colsRiesgo = (conProy) => [
    ...(conProy ? [C.proy] : []), { t: "No.", v: (r) => numRiesgo(r), tipo: "num", ancho: 6 }, { t: "Tipo", v: (r) => r.Tipo, ancho: 12 }, { t: "Estado", v: (r) => r.Estado, ancho: 14, color: true },
    { t: "Responsable", v: (r) => r.Responsable }, { t: "Descripción", v: (r) => r.Descripcion, ancho: 50 },
    { t: "Prob. inherente", v: (r) => PROB_TXT[r.Probabilidad_Inherente] || "", ancho: 15 }, { t: "Impacto inherente", v: (r) => IMP_TXT[r.Impacto_Inherente] || "", ancho: 17 },
    { t: "Calificación", v: (r) => R.calificacion(r.Probabilidad_Inherente, r.Impacto_Inherente) || "", tipo: "num" }, { t: "Nivel", v: (r) => { const c = R.calificacion(r.Probabilidad_Inherente, r.Impacto_Inherente); return c ? R.nivelRiesgo(c) : ""; }, ancho: 10, color: true },
    { t: "Plan de mitigación", v: (r) => r.Plan_Mitigacion, ancho: 44 }, { t: "Plan de contingencia", v: (r) => r.Plan_Contingencia, ancho: 44 },
    { t: "Prob. residual", v: (r) => PROB_TXT[r.Probabilidad_Residual] || "", ancho: 15 }, { t: "Impacto residual", v: (r) => IMP_TXT[r.Impacto_Residual] || "", ancho: 17 },
    { t: "Calificación residual", v: (r) => R.calificacion(r.Probabilidad_Residual, r.Impacto_Residual) || "", tipo: "num" }, { t: "Nivel residual", v: (r) => { const c = R.calificacion(r.Probabilidad_Residual, r.Impacto_Residual); return c ? R.nivelRiesgo(c) : ""; }, ancho: 12, color: true },
    { t: "Origen", v: (r) => origenRiesgo(r), ancho: 30 }, { t: "Fecha identificación", v: (r) => r.Fecha_Identificacion, tipo: "fecha" },
  ];
  // Qué exporta cada pantalla.
  const EXPORTADORES = {
    backlog: () => exportarExcel("Backlog de demanda", [{ nombre: "Backlog", columnas: [
      { t: "#", v: (d) => demandasFiltradas().indexOf(d) + 1, tipo: "num", ancho: 6 }, { t: "Código Almera", v: (d) => d.Codigo_Almera, ancho: 16 }, { t: "Iniciativa", v: (d) => d.Nombre, ancho: 34 },
      { t: "Necesidad", v: (d) => d.Descripcion, ancho: 40 }, { t: "Beneficio", v: (d) => d.Beneficio, ancho: 36 }, { t: "Área", v: (d) => d.Area }, { t: "Solicitante", v: (d) => d.Solicitante }, { t: "Sponsor", v: (d) => d.Sponsor },
      { t: "Tipo", v: (d) => d.Tipo }, { t: "Categoría", v: (d) => d.Categoria, ancho: 14 }, { t: "Talla", v: (d) => d.Tamano, ancho: 8 }, { t: "Obligatorio", v: (d) => d.Obligatorio, ancho: 11 },
      ...CRITERIOS_DEM.map((c) => ({ t: c.t, v: (d) => d[c.k], tipo: "num", ancho: 14 })), { t: "Puntaje", v: (d) => puntajeDemanda(d) ?? "", tipo: "num" },
      { t: "Estado", v: (d) => d.Estado || "Recibida", ancho: 13 }, { t: "PM asignado", v: (d) => nombreUsuario(d.PM_Asignado) }, { t: "Costo estimado", v: (d) => d.Costo_Estimado, tipo: "cop" },
      { t: "Recibida", v: (d) => d.Fecha_Recepcion, tipo: "fecha" }, { t: "Fecha deseada", v: (d) => d.Fecha_Deseada, tipo: "fecha" }, { t: "Días de antigüedad", v: (d) => antiguedad(d), tipo: "num" },
      { t: "Decisión", v: (d) => d.Fecha_Decision, tipo: "fecha" }, { t: "Decidido por", v: (d) => d.Decidido_Por }, { t: "Observaciones", v: (d) => d.Comentario_Decision, ancho: 36 }, { t: "Proyecto", v: (d) => d.ID_Proyecto, ancho: 12 }],
      filas: demandasFiltradas() }], [(S.demF || {}).estado ? `Estado: ${S.demF.estado}` : "", (S.demF || {}).tipo ? `Tipo: ${S.demF.tipo}` : "", (S.demF || {}).area ? `Área: ${S.demF.area}` : "", (S.demF || {}).pm ? `PM: ${nombreUsuario(S.demF.pm)}` : "", (S.demF || {}).q ? `Búsqueda: «${S.demF.q}»` : ""].filter(Boolean)),
    dashboard: () => {
      const ps = filtrados(), k = R.kpis(ps, S.datos.Seguimientos);
      return exportarExcel("Dashboard", [
        { nombre: "Indicadores", columnas: [{ t: "Indicador", v: (x) => x[0], ancho: 34 }, { t: "Valor", v: (x) => x[1], ancho: 22 }], filas: [
          ["Proyectos en la selección", k.total], ["Proyectos activos", k.activos], ["Avance real promedio", `${Math.round(k.avanceReal)}%`], ["Avance planeado promedio", `${Math.round(k.avancePlan)}%`],
          ["Proyectos en rojo", k.rojos], ["Presupuesto (activos)", cop(k.presupuesto)], ["Ejecutado (activos)", cop(k.ejecutado)], ["% ejecutado", `${Math.round(k.pctEjecutado)}%`], ["Seguimientos al día", `${Math.round(k.pctAlDia)}%`]] },
        { nombre: "Proyectos", columnas: colsProyecto, filas: ps }], filtrosGenerales());
    },
    proyectos: () => {
      const FL = S.filtrosLista;
      const lista = filtrados().filter((p) => (!FL.semaforo || p.Semaforo === FL.semaforo) && (!FL.texto || `${p.ID_Proyecto} ${p.Nombre} ${p.Cliente_Area} ${p.Codigo_Almera || ""}`.toLowerCase().includes(FL.texto.toLowerCase())))
        .sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
      return exportarExcel("Proyectos", [{ nombre: "Proyectos", columnas: colsProyecto, filas: lista }], filtrosGenerales());
    },
    avances: () => {
      const ids = new Set(filtrados().map((p) => p.ID_Proyecto));
      const segs = S.datos.Seguimientos.filter((s) => ids.has(s.ID_Proyecto) && (s.Semana || R.semanaISO(s.Fecha_Corte)) === S.semana);
      const rep = new Set(segs.map((s) => s.ID_Proyecto));
      const sin = filtrados().filter((p) => p.Estado === "Activo" && p.Frecuencia_Seguimiento === "Semanal" && !rep.has(p.ID_Proyecto));
      return exportarExcel(`Avances ${S.semana}`, [{ nombre: "Seguimientos de la semana", columnas: colsSeguimiento(true), filas: segs },
        { nombre: "Sin reporte", columnas: [{ t: "Proyecto", v: (p) => p.Nombre, ancho: 34 }, { t: "PM", v: (p) => nombreUsuario(p.PM) }, { t: "Último seguimiento", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).ultimaFecha, tipo: "fecha" }], filas: sin }],
        [...filtrosGenerales(), `Semana: ${S.semana}`]);
    },
    seguimiento: () => {
      const ps = filtrados().filter((p) => p.Estado === "Activo");
      const ids = new Set(ps.map((p) => p.ID_Proyecto));
      return exportarExcel("Seguimiento y compromisos", [
        { nombre: "Control de seguimiento", columnas: [{ t: "Proyecto", v: (p) => p.Nombre, ancho: 34 }, { t: "PM", v: (p) => nombreUsuario(p.PM) }, { t: "Frecuencia", v: (p) => p.Frecuencia_Seguimiento, ancho: 12 },
          { t: "Último", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).ultimaFecha, tipo: "fecha" }, { t: "Próximo", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).proximo, tipo: "fecha" },
          { t: "Estado", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).estado, ancho: 12, color: true }, { t: "En 90 días", v: (p) => R.estadoSeguimiento(p, S.datos.Seguimientos).en90, tipo: "num" }], filas: ps },
        { nombre: "Compromisos abiertos", columnas: colsCompromiso(true), filas: S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto) && !R.cerrado(c)).sort((a, b) => String(a.Fecha_Compromiso || "9").localeCompare(String(b.Fecha_Compromiso || "9"))) }],
        filtrosGenerales());
    },
    cronograma: () => {
      const ps = filtrados().filter((p) => p.Fecha_Inicio && p.Fecha_Fin_Plan).sort((a, b) => String(a.Fecha_Inicio).localeCompare(String(b.Fecha_Inicio)));
      const ids = new Set(ps.map((p) => p.ID_Proyecto));
      return exportarExcel("Cronograma", [
        { nombre: "Proyectos", columnas: [colsProyecto[2], colsProyecto[7], colsProyecto[12], colsProyecto[13], colsProyecto[14], colsProyecto[10],
          { t: "Depende de", v: (p) => depsDe(p.ID_Proyecto).map((d) => (proyecto(d.Depende_De) || {}).Nombre || d.Depende_De).join(", "), ancho: 30 }], filas: ps },
        { nombre: "Hitos", columnas: [C.proy, { t: "Hito", v: (h) => h.Hito, ancho: 40 }, { t: "Fecha plan", v: (h) => h.Fecha_Plan, tipo: "fecha" }, { t: "Fecha real", v: (h) => h.Fecha_Real, tipo: "fecha" },
          { t: "Estado", v: (h) => (h.Estado !== "Cumplido" && h.Fecha_Plan < R.hoyISO() ? "Atrasado" : h.Estado), ancho: 12, color: true }],
          filas: S.datos.Hitos.filter((h) => ids.has(h.ID_Proyecto)).sort((a, b) => String(a.Fecha_Plan).localeCompare(String(b.Fecha_Plan))) }],
        [...filtrosGenerales(), `Periodo en pantalla: ${S.cronoRango || "12m"}`]);
    },
    riesgos: () => {
      const ids = new Set(filtrados().map((p) => p.ID_Proyecto));
      const lista = riesgosVisibles("global", S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto)));
      return exportarRiesgosFSFB(lista, `Registro de riesgos — Portafolio — ${fecha(R.hoyISO())}`, [...filtrosGenerales(), ...describirFiltrosRiesgo("global")]);
    },
    priorizacion: () => exportarExcel("Priorización", [{ nombre: "Ranking", columnas: [
      { t: "Puesto", v: (p) => filtrados().filter((x) => puntaje(x) !== null).sort((a, b) => puntaje(b) - puntaje(a)).indexOf(p) + 1 || "", tipo: "num", ancho: 8 },
      colsProyecto[2], colsProyecto[7], { t: "PM", v: (p) => nombreUsuario(p.PM) }, ...CRITERIOS.map((c) => ({ t: c.t, v: (p) => p[c.k], tipo: "num", ancho: 14 })),
      { t: "Puntaje", v: (p) => (puntaje(p) === null ? "" : puntaje(p)), tipo: "num" }, { t: "Cuadrante", v: (p) => (puntaje(p) === null ? "Sin calificar" : cuadrante(p)), ancho: 16 }],
      filas: filtrados().slice().sort((a, b) => (puntaje(b) ?? -1) - (puntaje(a) ?? -1)) }], filtrosGenerales()),
    recursos: () => {
      const q = S.recQ || "", cap = capacidad();
      const lista = recursos().filter((r) => !q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${r.Area} ${empresaDe(r)}`).includes(normTxt(q)));
      const asig = (r) => cap.find((x) => (r.Correo && lc(x.correo) === lc(r.Correo)) || normTxt(x.nombre) === normTxt(r.Nombre));
      return exportarExcel("Recursos", [{ nombre: "Recursos", columnas: [{ t: "Nombre", v: (r) => r.Nombre, ancho: 30 }, { t: "Cargo", v: (r) => r.Cargo }, { t: "Área", v: (r) => r.Area },
        { t: "Empresa", v: (r) => empresaDe(r) }, { t: "Correo", v: (r) => r.Correo, ancho: 28 }, { t: "Teléfono", v: (r) => r.Telefono, ancho: 14 },
        { t: "Capacidad", v: (r) => num0(r.Capacidad) ?? 100, tipo: "pct" }, { t: "Asignado", v: (r) => (asig(r) || {}).total || 0, tipo: "pct" }, { t: "Activo", v: (r) => r.Activo || "Sí", ancho: 8 }], filas: lista }],
        q ? [`Búsqueda: «${q}»`] : []);
    },
    capacidad: () => {
      const idsF = new Set(filtrados().map((p) => p.ID_Proyecto));
      let lista = capacidad().filter((x) => x.asign.some((a) => idsF.has(a.p.ID_Proyecto)));
      if (S.capVer && S.capVer !== "todos") lista = lista.filter((x) => nivelCap(x.total, x.cap) === S.capVer);
      const det = lista.flatMap((x) => x.asign.map((a) => ({ x, a })));
      return exportarExcel("Capacidad del equipo", [
        { nombre: "Por persona", columnas: [{ t: "Persona", v: (x) => x.nombre, ancho: 30 }, { t: "Correo", v: (x) => x.correo, ancho: 28 }, { t: "Asignado", v: (x) => x.total, tipo: "pct" },
          { t: "Capacidad", v: (x) => x.cap, tipo: "pct" }, { t: "Estado", v: (x) => nivelCap(x.total, x.cap), ancho: 14, color: true }, { t: "Proyectos", v: (x) => x.asign.length, tipo: "num" }], filas: lista },
        { nombre: "Detalle por proyecto", columnas: [{ t: "Persona", v: (d) => d.x.nombre, ancho: 30 }, { t: "Proyecto", v: (d) => d.a.p.Nombre, ancho: 32 }, { t: "Rol", v: (d) => d.a.rol }, { t: "Dedicación", v: (d) => d.a.ded, tipo: "pct" }], filas: det }],
        [...filtrosGenerales(), S.capVer && S.capVer !== "todos" ? `Estado: ${S.capVer}` : ""].filter(Boolean));
    },
    lecciones: () => {
      const ids = new Set(filtrados().map((p) => p.ID_Proyecto)), F = S.lecF || {};
      const lista = (S.datos.Lecciones || []).filter((l) => ids.has(l.ID_Proyecto) && (!F.cat || l.Categoria === F.cat) && (!F.tipo || l.Tipo === F.tipo) &&
        (!F.texto || `${l.Situacion} ${l.Leccion} ${l.Recomendacion} ${(proyecto(l.ID_Proyecto) || {}).Nombre}`.toLowerCase().includes(F.texto.toLowerCase())));
      return exportarExcel("Lecciones aprendidas", [{ nombre: "Lecciones", columnas: [C.proy, { t: "Fecha", v: (l) => l.Fecha, tipo: "fecha" }, { t: "Tipo", v: (l) => l.Tipo, ancho: 12 },
        { t: "Categoría", v: (l) => l.Categoria, ancho: 16 }, { t: "Qué pasó", v: (l) => l.Situacion, ancho: 44 }, { t: "Lección", v: (l) => l.Leccion, ancho: 44 },
        { t: "Recomendación", v: (l) => l.Recomendacion, ancho: 44 }, { t: "Registrado por", v: (l) => nombreUsuario(l.Registrado_Por) }], filas: lista }],
        [...filtrosGenerales(), F.cat ? `Categoría: ${F.cat}` : "", F.tipo ? `Tipo: ${F.tipo}` : "", F.texto ? `Búsqueda: «${F.texto}»` : ""].filter(Boolean));
    },
    auditoria: async () => {
      const F = S.audF || {};
      const lista = (await cargarAuditoria()).filter((a) => (!F.usuario || lc(a.Usuario) === lc(F.usuario)) && (!F.tabla || a.Tabla === F.tabla) && (!F.texto || lc(`${a.Detalle} ${a.ID_Registro}`).includes(lc(F.texto))));
      return exportarExcel("Auditoría", [{ nombre: "Auditoría", columnas: [{ t: "Fecha y hora", v: (a) => a.Fecha_Hora, tipo: "fechahora", ancho: 17 }, { t: "Usuario", v: (a) => nombreUsuario(a.Usuario) },
        { t: "Acción", v: (a) => a.Accion, ancho: 11 }, { t: "Tipo de dato", v: (a) => NOMBRE_TABLA[a.Tabla] || a.Tabla, ancho: 14 }, { t: "Registro", v: (a) => a.ID_Registro, ancho: 16 },
        { t: "Proyecto", v: (a) => (proyecto(a.ID_Proyecto) || {}).Nombre || a.ID_Proyecto, ancho: 28 }, { t: "Detalle", v: (a) => a.Detalle, ancho: 70 }], filas: lista }],
        [F.usuario ? `Usuario: ${F.usuario}` : "", F.tabla ? `Tipo: ${F.tabla}` : "", F.texto ? `Búsqueda: «${F.texto}»` : ""].filter(Boolean));
    },
    usuarios: () => exportarExcel("Usuarios", [{ nombre: "Usuarios", columnas: [{ t: "Correo", v: (u) => u.Correo, ancho: 30 }, { t: "Nombre", v: (u) => u.Nombre, ancho: 28 },
      { t: "Roles", v: (u) => u.Rol }, { t: "Proyectos (Lector)", v: (u) => u.Proyectos, ancho: 26 }, { t: "Activo", v: (u) => u.Activo, ancho: 8 }], filas: S.datos.Usuarios }], []),
    catalogos: () => exportarExcel("Catálogos", [{ nombre: "Catálogos", columnas: [{ t: "Lista", v: (c) => c.Lista, ancho: 24 }, { t: "Valor", v: (c) => c.Valor, ancho: 34 }], filas: S.datos.Catalogos },
      { nombre: "Proveedores", columnas: [{ t: "Proveedor", v: (v) => v.Nombre, ancho: 30 }, { t: "NIT", v: (v) => v.NIT, ancho: 14 }, { t: "Servicio", v: (v) => v.Servicio, ancho: 26 },
        { t: "Contacto", v: (v) => v.Contacto }, { t: "Correo", v: (v) => v.Correo, ancho: 28 }, { t: "Teléfono", v: (v) => v.Telefono, ancho: 14 }, { t: "Activo", v: (v) => v.Activo, ancho: 8 }], filas: S.datos.Proveedores || [] }], []),
    ficha: () => {
      const p = proyecto(S.pid);
      if (!p) return null;
      const pid = p.ID_Proyecto;
      if (S.fichaTab === "riesgos") return exportarRiesgosFSFB(riesgosVisibles("p:" + pid, riesgosDe(pid)), `Registro de riesgos — Proyecto ${p.Nombre} — ${fecha(R.hoyISO())}`, describirFiltrosRiesgo("p:" + pid));
      const resumen = colsProyecto.map((c) => ({ campo: c.t, valor: c.tipo === "cop" ? cop(c.v(p)) : c.tipo === "pct" ? (c.v(p) === "" ? "" : `${Math.round(Number(c.v(p)) || 0)}%`) : c.tipo === "fecha" ? fecha(c.v(p)) : c.v(p) }));
      return exportarExcel(`Proyecto ${p.Nombre}`, [
        { nombre: "Resumen", columnas: [{ t: "Campo", v: (x) => x.campo, ancho: 24 }, { t: "Valor", v: (x) => x.valor, ancho: 70 }], filas: resumen },
        { nombre: "Stakeholders", columnas: [{ t: "Rol", v: (x) => x.Rol }, { t: "Nombre", v: (x) => x.Nombre, ancho: 28 }, { t: "Cargo", v: (x) => x.Cargo }, { t: "Área", v: (x) => x.Area },
          { t: "Empresa", v: (x) => (x.ID_Proveedor ? (proveedor(x.ID_Proveedor) || {}).Nombre : "Interno") }, { t: "Correo", v: (x) => x.Correo, ancho: 28 }, { t: "Dedicación", v: (x) => x.Dedicacion, tipo: "pct" }],
          filas: [{ Rol: "PM", Nombre: nombreUsuario(p.PM), Cargo: "Gerente del proyecto", Area: "", ID_Proveedor: "", Correo: p.PM, Dedicacion: p.Dedicacion_PM }, ...stakeholdersDe(pid)] },
        { nombre: "Compromisos", columnas: colsCompromiso(false), filas: S.datos.Compromisos.filter((c) => c.ID_Proyecto === pid).sort((a, b) => String(a.Fecha_Compromiso || "9").localeCompare(String(b.Fecha_Compromiso || "9"))) },
        { nombre: "Seguimientos", columnas: colsSeguimiento(false), filas: R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse() },
        { nombre: "Hitos", columnas: [{ t: "Hito", v: (h) => h.Hito, ancho: 40 }, { t: "Fecha plan", v: (h) => h.Fecha_Plan, tipo: "fecha" }, { t: "Fecha real", v: (h) => h.Fecha_Real, tipo: "fecha" },
          { t: "Estado", v: (h) => (h.Estado !== "Cumplido" && h.Fecha_Plan < R.hoyISO() ? "Atrasado" : h.Estado), ancho: 12, color: true }], filas: S.datos.Hitos.filter((h) => h.ID_Proyecto === pid) },
        { nombre: "Riesgos", columnas: colsRiesgo(false), filas: riesgosDe(pid).sort((a, b) => numRiesgo(a) - numRiesgo(b)) },
        { nombre: "Tickets", columnas: [{ t: "N.º", v: (t) => t.Numero, ancho: 12 }, { t: "Título", v: (t) => t.Titulo, ancho: 44 }, { t: "Estado", v: (t) => t.Estado, ancho: 12, color: true },
          { t: "Prioridad", v: (t) => t.Prioridad, ancho: 11 }, { t: "Registrado", v: (t) => t.Fecha_Registro, tipo: "fecha" }], filas: ticketsDe(pid) },
        { nombre: "Cambios", columnas: [{ t: "Fecha", v: (c) => c.Fecha, tipo: "fecha" }, { t: "Tipo", v: (c) => c.Tipo, ancho: 12 }, { t: "Cambio", v: (c) => c.Descripcion, ancho: 44 },
          { t: "Justificación", v: (c) => c.Justificacion, ancho: 36 }, { t: "Impacto", v: (c) => c.Impacto, ancho: 36 }, { t: "Nueva fecha fin", v: (c) => c.Nueva_Fecha_Fin, tipo: "fecha" },
          { t: "Nuevo presupuesto", v: (c) => c.Nuevo_Presupuesto, tipo: "cop" }, { t: "Estado", v: (c) => c.Estado, ancho: 12, color: true }, { t: "Decisión", v: (c) => c.Fecha_Decision, tipo: "fecha" }], filas: cambiosDe(pid) },
        { nombre: "Lecciones", columnas: [{ t: "Fecha", v: (l) => l.Fecha, tipo: "fecha" }, { t: "Tipo", v: (l) => l.Tipo, ancho: 12 }, { t: "Categoría", v: (l) => l.Categoria, ancho: 16 },
          { t: "Qué pasó", v: (l) => l.Situacion, ancho: 40 }, { t: "Lección", v: (l) => l.Leccion, ancho: 40 }, { t: "Recomendación", v: (l) => l.Recomendacion, ancho: 40 }], filas: leccionesDe(pid) },
        { nombre: "RACI", columnas: [{ t: "Entregable", v: (r) => r.Entregable, ancho: 36 }, ...[{ k: "PM", t: `PM (${nombreUsuario(p.PM)})` }, ...stakeholdersDe(pid).map((x) => ({ k: x.ID_Stakeholder, t: `${x.Rol} (${x.Nombre})` }))]
          .map((c) => ({ t: c.t, v: (r) => asignRaci(r)[c.k] || "", ancho: 16 }))], filas: raciDe(pid) },
      ], [`Proyecto: ${p.Nombre} (${pid})`]);
    },
  };

  // ---------- Backlog de demanda (iniciativas asignadas por Gestión de la Demanda vía Almera) ----------
  const ESTADOS_DEMANDA = ["Recibida", "En análisis", "Priorizada", "En comité", "Aprobada", "Convertida", "Aplazada", "Rechazada"];
  const ABIERTOS_DEMANDA = ["Recibida", "En análisis", "Priorizada", "En comité", "Aprobada"];
  const DECISIONES = ["Aprobada", "Aplazada", "Rechazada"];
  // Criterios de priorización (1 a 5). "inv" = a menor valor, mejor. Los pesos se ajustan en la app.
  const CRITERIOS_DEM = [
    { k: "Valor", t: "Valor / beneficio para la institución", peso: 25 },
    { k: "Alineacion", t: "Alineación estratégica", peso: 20 },
    { k: "Impacto_Paciente", t: "Impacto en paciente / seguridad", peso: 15 },
    { k: "Urgencia", t: "Urgencia", peso: 15 },
    { k: "Riesgo_Prio", t: "Complejidad / riesgo de ejecución", peso: 10, inv: true },
    { k: "Esfuerzo", t: "Esfuerzo / costo", peso: 15, inv: true },
  ];
  const pesosDemanda = () => {
    const guardados = {};
    (S.datos.Catalogos || []).filter((c) => c.Lista === "Peso_Prioridad").forEach((c) => { const [k, v] = String(c.Valor).split(":"); if (k && !Number.isNaN(Number(v))) guardados[k] = Number(v); });
    return Object.fromEntries(CRITERIOS_DEM.map((c) => [c.k, guardados[c.k] ?? c.peso]));
  };
  // Puntaje 0 a 100; lo regulatorio / obligatorio va primero en el ranking.
  function puntajeDemanda(d) {
    const w = pesosDemanda(), total = Object.values(w).reduce((a, b) => a + b, 0) || 1;
    if (CRITERIOS_DEM.some((c) => !num0(d[c.k]))) return null;
    const v = CRITERIOS_DEM.reduce((a, c) => a + w[c.k] * (c.inv ? 6 - Number(d[c.k]) : Number(d[c.k])), 0) / total;
    return Math.round(((v - 1) / 4) * 100);
  }
  const demandas = () => (S.datos.Demandas || []);
  const obligatoria = (d) => d.Obligatorio === "Sí";
  const diasEn = (d) => R.difDias(R.hoyISO(), String(d.Fecha_Estado || d.Fecha_Recepcion || R.hoyISO()).slice(0, 10));
  const antiguedad = (d) => R.difDias(R.hoyISO(), String(d.Fecha_Recepcion || R.hoyISO()).slice(0, 10));
  const ordenRanking = (a, b) => (obligatoria(b) - obligatoria(a)) || ((puntajeDemanda(b) ?? -1) - (puntajeDemanda(a) ?? -1)) || String(a.Fecha_Recepcion).localeCompare(String(b.Fecha_Recepcion));
  const puedeDemanda = () => R.puede(S.usuario, "crearProyecto");
  const decideDemanda = () => R.puede(S.usuario, "verTodo");
  function demandasFiltradas() {
    const F = S.demF = S.demF || { estado: "abiertas", tipo: "", area: "", pm: "", q: "" };
    return demandas().filter((d) => (F.estado === "todas" || (F.estado === "abiertas" ? ABIERTOS_DEMANDA.includes(d.Estado || "Recibida") : (d.Estado || "Recibida") === F.estado)) &&
      (!F.tipo || d.Tipo === F.tipo) && (!F.area || d.Area === F.area) && (!F.pm || lc(d.PM_Asignado) === lc(F.pm)) &&
      (!F.q || normTxt(`${d.Codigo_Almera} ${d.Nombre} ${d.Descripcion} ${d.Solicitante} ${d.Sponsor}`).includes(normTxt(F.q))))
      .sort(ordenRanking);
  }
  function vBacklog(el) {
    const F = S.demF = S.demF || { estado: "abiertas", tipo: "", area: "", pm: "", q: "" };
    const vista = S.demVista || "ranking";
    const todas = demandas(), lista = demandasFiltradas();
    const abiertas = todas.filter((d) => ABIERTOS_DEMANDA.includes(d.Estado || "Recibida"));
    const decididas = todas.filter((d) => ["Aprobada", "Convertida", "Rechazada"].includes(d.Estado));
    const tasa = decididas.length ? Math.round(decididas.filter((d) => d.Estado !== "Rechazada").length / decididas.length * 100) : null;
    const viejas = abiertas.filter((d) => antiguedad(d) > 30).length;
    const pms = usuariosConRol("PM").map((u) => [u.Correo, u.Nombre || u.Correo]);
    const kpi = (t, v, d, cls) => `<div class="kpi estatico ${cls || ""}"><span class="kpi-t">${t}</span><span class="kpi-v">${v}</span>${d ? `<span class="kpi-d">${d}</span>` : ""}</div>`;
    const tarjeta = (d) => { const pts = puntajeDemanda(d); return `<div class="dem-card ${obligatoria(d) ? "oblig" : ""}" draggable="${puedeDemanda()}" data-dem="${esc(d.ID_Demanda)}">
      <div class="dem-top"><span class="dem-almera">${esc(d.Codigo_Almera || "Sin código")}</span>${obligatoria(d) ? `<span class="pill materializado">Obligatorio</span>` : ""}<span class="dem-pts" title="Puntaje de priorización">${pts === null ? "—" : pts}</span></div>
      <div class="dem-nombre">${esc(d.Nombre || "(sin nombre)")}</div>
      <div class="sub">${esc([d.Tipo, d.Area, d.Tamano ? `Talla ${d.Tamano}` : ""].filter(Boolean).join(" · "))}</div>
      <div class="dem-pie">${d.PM_Asignado ? persona(nombreUsuario(d.PM_Asignado)) : `<span class="sub">Sin PM</span>`}<span class="sub ${antiguedad(d) > 30 && ABIERTOS_DEMANDA.includes(d.Estado || "Recibida") ? "baja" : ""}" title="Días desde que llegó">${antiguedad(d)} d</span></div></div>`; };
    const ranking = lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>#</th><th>Almera</th><th>Iniciativa</th><th>Tipificación</th><th>Estado</th><th>Puntaje</th><th>PM</th><th>Antigüedad</th></tr></thead><tbody>
      ${lista.map((d, i) => { const pts = puntajeDemanda(d); return `<tr class="clic" data-dem="${esc(d.ID_Demanda)}"><td>${i + 1}</td><td><b>${esc(d.Codigo_Almera)}</b></td>
        <td><b>${esc(d.Nombre)}</b><div class="sub">${esc(d.Solicitante ? `Solicita: ${d.Solicitante}` : "")}${d.Sponsor ? ` · Sponsor: ${esc(d.Sponsor)}` : ""}</div></td>
        <td>${esc(d.Tipo || "—")}<div class="sub">${esc([d.Categoria, d.Area, d.Tamano ? `Talla ${d.Tamano}` : ""].filter(Boolean).join(" · "))}</div>${obligatoria(d) ? `<span class="pill materializado">Obligatorio</span>` : ""}</td>
        <td>${pill(d.Estado || "Recibida")}${d.ID_Proyecto ? `<div><button class="btn enlace" data-pid="${esc(d.ID_Proyecto)}">Ver proyecto</button></div>` : ""}</td>
        <td class="nowrap">${pts === null ? `<span class="sub">Sin calificar</span>` : `<span class="cap-barra"><span class="cap-lleno" style="width:${pts}%;background:var(--azul)"></span></span> <b>${pts}</b>`}</td>
        <td>${d.PM_Asignado ? persona(nombreUsuario(d.PM_Asignado)) : `<span class="sub">—</span>`}</td>
        <td class="${antiguedad(d) > 30 && ABIERTOS_DEMANDA.includes(d.Estado || "Recibida") ? "baja" : ""}">${antiguedad(d)} días<div class="sub">${diasEn(d)} en este estado</div></td></tr>`; }).join("")}</tbody></table></div>`
      : vacio(todas.length ? "Ninguna demanda coincide con los filtros." : "Aún no hay demandas. Registra la primera con «+ Nueva demanda» cuando Gestión de la Demanda te asigne una iniciativa en Almera.");
    const cols = ESTADOS_DEMANDA.filter((e) => F.estado === "todas" || e !== "Convertida" && e !== "Rechazada" || F.estado === e);
    const tablero = `<div class="kanban">${cols.map((e) => { const l = lista.filter((d) => (d.Estado || "Recibida") === e); return `<div class="kb-col" data-col="${esc(e)}"><div class="kb-cab">${pill(e)} <span class="contador">${l.length}</span></div><div class="kb-lista">${l.map(tarjeta).join("") || `<div class="kb-vacio">Arrastra aquí</div>`}</div></div>`; }).join("")}</div>`;
    el.innerHTML = `<div class="titulo-fila"><h1>Backlog de demanda</h1><div class="acciones">
        ${decideDemanda() ? `<button class="btn" id="b-pesos">⚖ Pesos de priorización</button>` : ""}
        <button class="btn" id="b-agenda">📄 Agenda del comité</button>
        ${puedeDemanda() ? `<button class="btn primario" id="b-dem">+ Nueva demanda</button>` : ""}</div></div>
      <div class="kpis">${kpi("Demandas abiertas", abiertas.length)}${kpi("En análisis", todas.filter((d) => d.Estado === "En análisis").length)}${kpi("Para comité", todas.filter((d) => d.Estado === "En comité" || d.Estado === "Priorizada").length)}
        ${kpi("Más de 30 días sin decisión", viejas, "", viejas ? "alerta" : "")}${kpi("Tasa de aprobación", tasa === null ? "—" : `${tasa}%`, `${decididas.length} decidida(s)`)}</div>
      <div class="card"><div class="dem-barra">
        <div class="seg">${["ranking", "tablero"].map((v) => `<button class="btn chico ${vista === v ? "primario" : ""}" data-demvista="${v}">${v === "ranking" ? "☰ Ranking" : "▦ Tablero"}</button>`).join("")}</div>
        <label class="filtro"><span>Estado</span><select data-demf="estado"><option value="abiertas" ${F.estado === "abiertas" ? "selected" : ""}>Abiertas</option>${ESTADOS_DEMANDA.map((e) => `<option ${F.estado === e ? "selected" : ""}>${e}</option>`).join("")}<option value="todas" ${F.estado === "todas" ? "selected" : ""}>Todas</option></select></label>
        <label class="filtro"><span>Tipo</span><select data-demf="tipo"><option value="">Todos</option>${(S.cat.Tipo_Demanda || []).map((t) => `<option ${F.tipo === t ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label>
        <label class="filtro"><span>Área</span><select data-demf="area"><option value="">Todas</option>${(S.cat.Cliente_Area || []).map((t) => `<option ${F.area === t ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label>
        <label class="filtro"><span>PM</span><select data-demf="pm"><option value="">Todos</option>${pms.map(([c, n]) => `<option value="${esc(c)}" ${lc(F.pm) === lc(c) ? "selected" : ""}>${esc(n)}</option>`).join("")}</select></label>
        <label class="filtro crece"><span>Buscar</span><input type="search" data-demf="q" value="${esc(F.q)}" placeholder="Código Almera, nombre, solicitante…"></label></div>
        <p class="sub">${lista.length} demanda(s). Orden: obligatorias primero y luego por puntaje. ${vista === "tablero" && puedeDemanda() ? "Arrastra las tarjetas para cambiar de estado." : ""}</p>
        ${vista === "tablero" ? tablero : ranking}</div>`;
    el.querySelectorAll("[data-demvista]").forEach((b) => b.addEventListener("click", () => { S.demVista = b.dataset.demvista; render(); }));
    el.querySelectorAll("[data-demf]").forEach((i) => i.addEventListener("change", () => { F[i.dataset.demf] = i.value; render(); }));
    el.querySelectorAll("tr[data-dem], .dem-card").forEach((r) => r.addEventListener("click", (e) => { if (e.target.closest("button")) return; formDemanda(demandas().find((d) => d.ID_Demanda === r.dataset.dem)); }));
    if ($("#b-dem")) $("#b-dem").addEventListener("click", () => formDemanda());
    if ($("#b-pesos")) $("#b-pesos").addEventListener("click", formPesos);
    $("#b-agenda").addEventListener("click", agendaComite);
    // Arrastrar tarjetas entre columnas del tablero.
    el.querySelectorAll(".dem-card[draggable=true]").forEach((c) => c.addEventListener("dragstart", (e) => { e.dataTransfer.setData("text/plain", c.dataset.dem); c.classList.add("arrastrando"); }));
    el.querySelectorAll(".kb-col").forEach((col) => {
      col.addEventListener("dragover", (e) => { e.preventDefault(); col.classList.add("sobre"); });
      col.addEventListener("dragleave", () => col.classList.remove("sobre"));
      col.addEventListener("drop", (e) => {
        e.preventDefault(); col.classList.remove("sobre");
        const d = demandas().find((x) => x.ID_Demanda === e.dataTransfer.getData("text/plain"));
        if (d) moverDemanda(d, col.dataset.col);
      });
    });
  }
  // Cambio de estado desde el tablero, con las reglas del flujo.
  function moverDemanda(d, estado) {
    if (!d || (d.Estado || "Recibida") === estado) return;
    if (estado === "Convertida") return d.Estado === "Aprobada" ? convertirDemanda(d) : toast("Solo una demanda aprobada se convierte en proyecto.", true);
    if (DECISIONES.includes(estado)) return decideDemanda() ? decidirDemanda(d, estado) : toast("Aprobar, aplazar o rechazar lo decide el comité (rol PMO o Admin).", true);
    guardar(() => S.api.actualizarPorId("Demandas", "ID_Demanda", d.ID_Demanda, { Estado: estado, Fecha_Estado: R.hoyISO() }), `${d.Codigo_Almera}: ${estado}`);
  }
  function formDemanda(d) {
    const nuevo = !d;
    const esc5 = [[1, "1 · Muy bajo"], [2, "2 · Bajo"], [3, "3 · Medio"], [4, "4 · Alto"], [5, "5 · Muy alto"]];
    const pms = usuariosConRol("PM").map((u) => [u.Correo, u.Nombre || u.Correo]);
    const estadosEditables = ESTADOS_DEMANDA.filter((e) => !DECISIONES.includes(e) && e !== "Convertida");
    const campos = [
      { seccion: "1. Iniciativa" },
      { k: "Codigo_Almera", label: "Código Almera *", placeholder: "Código de la demanda en Almera" },
      { k: "Fecha_Recepcion", label: "Fecha de asignación", tipo: "date", def: R.hoyISO() },
      { k: "Nombre", label: "Nombre de la iniciativa", ancho: true },
      { k: "Descripcion", label: "Necesidad / problema", tipo: "textarea" },
      { k: "Beneficio", label: "Beneficio esperado", tipo: "textarea" },
      { k: "Area", label: "Área solicitante", tipo: "select", opciones: S.cat.Cliente_Area },
      { k: "Solicitante", label: "Solicitante", recurso: true, placeholder: "Busca en Recursos" },
      { k: "Sponsor", label: "Sponsor", recurso: true, placeholder: "Busca en Recursos" },
      { k: "Fecha_Deseada", label: "Fecha deseada por el solicitante", tipo: "date" },
      { seccion: "2. Tipificación" },
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: S.cat.Tipo_Demanda },
      { k: "Categoria", label: "Categoría", tipo: "select", opciones: S.cat.Categoria_Demanda },
      { k: "Tamano", label: "Talla estimada", tipo: "select", opciones: [["S", "S · Menos de 1 mes"], ["M", "M · 1 a 3 meses"], ["L", "L · 3 a 6 meses"], ["XL", "XL · Más de 6 meses"]] },
      { k: "Obligatorio", label: "¿Regulatorio / obligatorio?", tipo: "select", opciones: ["No", "Sí"], def: "No", ayuda: "Las obligatorias van primero en el ranking." },
      { k: "Costo_Estimado", label: "Costo estimado (COP)", tipo: "number", min: 0 },
      { k: "PM_Asignado", label: "PM asignado", tipo: "select", opciones: pms, ayuda: " " },
      { seccion: "3. Priorización", ayuda: "de 1 a 5; el puntaje se calcula con los pesos del comité" },
      ...CRITERIOS_DEM.map((c) => ({ k: c.k, label: `${c.t}${c.inv ? " (a menor, mejor)" : ""}`, tipo: "select", opciones: esc5 })),
      { html: `<div class="calif-viva" id="dem-pts"></div>` },
      { seccion: "4. Estado" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: nuevo || estadosEditables.includes(d.Estado || "Recibida") ? estadosEditables : [d.Estado], def: "Recibida", bloqueado: !nuevo && !estadosEditables.includes(d.Estado || "Recibida"),
        ayuda: "Aprobar, aplazar o rechazar se hace con «Decisión del comité»." },
    ];
    const historial = nuevo ? "" : `<div class="form-seccion">Decisión del comité</div>
      <div class="dem-decision">${d.Fecha_Decision ? `${pill(d.Estado)} el ${fecha(d.Fecha_Decision)} por ${esc(d.Decidido_Por || "—")}${d.Comentario_Decision ? `<div class="sub">${esc(d.Comentario_Decision)}</div>` : ""}` : `<span class="sub">Sin decisión todavía.</span>`}
        <div class="acciones">${decideDemanda() && !["Convertida"].includes(d.Estado) ? `<button type="button" class="btn" id="dem-decidir">⚖ Decisión del comité</button>` : ""}
        ${d.Estado === "Aprobada" && R.puede(S.usuario, "crearProyecto") ? `<button type="button" class="btn primario" id="dem-convertir">🚀 Convertir en proyecto</button>` : ""}
        ${d.ID_Proyecto ? `<button type="button" class="btn" id="dem-proy">Ver proyecto ${esc(d.ID_Proyecto)}</button>` : ""}</div></div>
      <div class="form-seccion">Historial</div><div id="dem-hist" class="sub">Cargando…</div>`;
    const init = (m) => {
      const calc = () => {
        const x = {}; CRITERIOS_DEM.forEach((c) => { x[c.k] = m.querySelector(`[name="${c.k}"]`).value; });
        const pts = puntajeDemanda(x);
        m.querySelector("#dem-pts").innerHTML = pts === null ? "Califica los 6 criterios para ver el puntaje." : `Puntaje de priorización: <b>${pts} / 100</b>${m.querySelector('[name="Obligatorio"]').value === "Sí" ? " · <b>Obligatoria</b> (va primero)" : ""}`;
      };
      m.querySelectorAll("select").forEach((s) => s.addEventListener("change", calc)); calc();
      const pmSel = m.querySelector('[name="PM_Asignado"]'), pmAy = pmSel.parentElement.querySelector(".sub");
      const cap = () => { const c = capacidad().find((x) => lc(x.correo) === lc(pmSel.value)); pmAy.innerHTML = pmSel.value ? `Hoy tiene asignado <b class="${c && c.total > c.cap ? "baja" : ""}">${c ? c.total : 0}%</b> de ${c ? c.cap : 100}% en proyectos activos.` : "Quien hace el análisis de la demanda."; };
      pmSel.addEventListener("change", cap); cap();
      if (!nuevo) {
        if (m.querySelector("#dem-decidir")) m.querySelector("#dem-decidir").addEventListener("click", () => decidirDemanda(d));
        if (m.querySelector("#dem-convertir")) m.querySelector("#dem-convertir").addEventListener("click", () => convertirDemanda(d));
        if (m.querySelector("#dem-proy")) m.querySelector("#dem-proy").addEventListener("click", () => { cerrarModal(); ir("ficha", d.ID_Proyecto); });
        cargarAuditoria().then((a) => { const h = m.querySelector("#dem-hist"); if (!h) return; const l = a.filter((x) => x.Tabla === "Demandas" && x.ID_Registro === d.ID_Demanda);
          h.innerHTML = l.length ? `<ul class="lista">${l.map((x) => `<li>${fechaHora(x.Fecha_Hora)} · ${esc(nombreUsuario(x.Usuario))} · ${esc(x.Accion)}: ${esc(x.Detalle)}</li>`).join("")}</ul>` : "Sin movimientos registrados."; });
      }
    };
    modal(nuevo ? "Nueva demanda" : `Demanda ${d.Codigo_Almera}`, campos, nuevo ? {} : d, (fd) => {
      const cambios = { ...fd, Codigo_Almera: String(fd.Codigo_Almera).trim() };
      if (nuevo) {
        const max = demandas().reduce((mx, x) => Math.max(mx, parseInt(String(x.ID_Demanda).slice(4), 10) || 0), 0);
        const fila = { ID_Demanda: `DEM-${String(max + 1).padStart(4, "0")}`, ...cambios, Fecha_Estado: R.hoyISO(), Registrado_Por: S.usuario.Correo };
        guardar(async () => { for (const n of [fd.Solicitante, fd.Sponsor].filter(Boolean)) await asegurarRecurso({ Nombre: n }); await S.api.agregarFila("Demandas", fila); }, `Demanda ${fila.Codigo_Almera} registrada`);
      } else {
        if ((cambios.Estado || "") !== (d.Estado || "")) cambios.Fecha_Estado = R.hoyISO();
        if (!estadosEditables.includes(d.Estado || "Recibida")) delete cambios.Estado;
        guardar(async () => { for (const n of [fd.Solicitante, fd.Sponsor].filter(Boolean)) await asegurarRecurso({ Nombre: n }); await S.api.actualizarPorId("Demandas", "ID_Demanda", d.ID_Demanda, cambios); }, "Demanda actualizada");
      }
    }, (fd) => {
      const cod = String(fd.Codigo_Almera || "").trim();
      if (!cod) return "El código Almera es obligatorio: es la llave de la demanda.";
      const dup = demandas().find((x) => lc(x.Codigo_Almera) === lc(cod) && (!d || x.ID_Demanda !== d.ID_Demanda));
      if (dup) return `Ya existe la demanda ${dup.Codigo_Almera} («${dup.Nombre}»).`;
      const dupP = S.datos.Proyectos.find((p) => lc(p.Codigo_Almera) === lc(cod) && (!d || p.ID_Proyecto !== d.ID_Proyecto));
      if (dupP) return `El código ${cod} ya está en el proyecto «${dupP.Nombre}».`;
      return "";
    }, { init, despues: historial, eliminar: nuevo || d.Estado === "Convertida" || !decideDemanda() ? null : { texto: "Eliminar demanda", mensaje: `Se borrará la demanda ${d.Codigo_Almera}. Si solo no avanzó, mejor recházala o aplázala para que quede la trazabilidad.`,
      accion: () => guardar(() => S.api.eliminarFilas("Demandas", "ID_Demanda", [d.ID_Demanda]), "Demanda eliminada") } });
  }
  function decidirDemanda(d, sugerida) {
    const ultimo = leerLocal("pmo_decide", "");
    modal(`Decisión del comité · ${d.Codigo_Almera}`, [
      { k: "Decision", label: "Decisión", tipo: "select", opciones: DECISIONES, def: sugerida || "Aprobada" },
      { k: "Fecha_Decision", label: "Fecha del comité", tipo: "date", def: R.hoyISO() },
      { k: "Decidido_Por", label: "Aprobado / decidido por", recurso: true, def: ultimo, placeholder: "Quien aprueba en el comité" },
      { k: "Comentario_Decision", label: "Observaciones / condiciones", tipo: "textarea", ancho: true },
    ], {}, (fd) => {
      guardarLocal("pmo_decide", fd.Decidido_Por || "");
      guardar(async () => {
        if (fd.Decidido_Por) await asegurarRecurso({ Nombre: fd.Decidido_Por });
        await S.api.actualizarPorId("Demandas", "ID_Demanda", d.ID_Demanda, { Estado: fd.Decision, Fecha_Decision: fd.Fecha_Decision, Decidido_Por: fd.Decidido_Por, Comentario_Decision: fd.Comentario_Decision, Fecha_Estado: R.hoyISO() });
      }, `Demanda ${d.Codigo_Almera}: ${fd.Decision}`);
    }, null, { antes: `<div class="resumen-filtro"><b>${esc(d.Nombre)}</b><br>Puntaje ${puntajeDemanda(d) ?? "sin calificar"}${obligatoria(d) ? " · Obligatoria" : ""} · ${esc(d.Tipo || "")} · ${esc(d.Area || "")}</div>` });
  }
  // Crea el proyecto con los datos de la demanda y deja ambos enlazados.
  function convertirDemanda(d) {
    const pref = { Nombre: d.Nombre, Cliente_Area: d.Area, PM: d.PM_Asignado, Codigo_Almera: d.Codigo_Almera, Estado: "Activo", Fase: "Inicio", Presupuesto: d.Costo_Estimado || 0,
      Valor: d.Valor, Urgencia: d.Urgencia, Riesgo_Prio: d.Riesgo_Prio, Esfuerzo: d.Esfuerzo, Comentario_Estado: `Viene de la demanda ${d.Codigo_Almera}. ${d.Descripcion || ""}`.trim() };
    formProyecto(null, pref, async (idProyecto) => {
      await S.api.actualizarPorId("Demandas", "ID_Demanda", d.ID_Demanda, { Estado: "Convertida", ID_Proyecto: idProyecto, Fecha_Estado: R.hoyISO() });
      const personas = [["Sponsor", d.Sponsor], ["Solicitante", d.Solicitante]].filter(([, n]) => n);
      if (personas.length) {
        const recs = [];
        for (const [, n] of personas) recs.push(await asegurarRecurso({ Nombre: n }));
        await S.api.agregarFilas("Stakeholders", personas.map(([rol, n], i) => { const r = recs[i] || {}; return { ID_Stakeholder: R.siguienteIdHijo("STK", S.datos.Stakeholders || [], "ID_Stakeholder", idProyecto, i), ID_Proyecto: idProyecto,
          Rol: rol, Nombre: r.Nombre || n, Cargo: r.Cargo || "", Area: r.Area || "", Correo: r.Correo || "", Telefono: r.Telefono || "", ID_Proveedor: r.ID_Proveedor || "", ID_Recurso: r.ID_Recurso || "" }; }));
      }
    });
  }
  function formPesos() {
    const w = pesosDemanda();
    modal("Pesos de priorización", CRITERIOS_DEM.map((c) => ({ k: c.k, label: `${c.t} (%)`, tipo: "number", min: 0, max: 100, def: w[c.k] })), {}, (fd) => guardar(async () => {
      const viejos = (S.datos.Catalogos || []).filter((c) => c.Lista === "Peso_Prioridad");
      for (const v of viejos) await S.api.eliminarFila("Catalogos", { Lista: "Peso_Prioridad", Valor: v.Valor });
      await S.api.agregarFilas("Catalogos", CRITERIOS_DEM.map((c) => ({ Lista: "Peso_Prioridad", Valor: `${c.k}:${Number(fd[c.k]) || 0}` })));
    }, "Pesos actualizados"), (fd) => { const t = CRITERIOS_DEM.reduce((a, c) => a + (Number(fd[c.k]) || 0), 0); return t !== 100 ? `Los pesos deben sumar 100% (hoy suman ${t}%).` : ""; },
    { antes: `<p class="sub">Cuánto pesa cada criterio en el puntaje (deben sumar 100%). Aplica a todas las demandas; lo regulatorio u obligatorio siempre va primero.</p>` });
  }
  function agendaComite() {
    const lista = demandas().filter((d) => ["En comité", "Priorizada"].includes(d.Estado)).sort(ordenRanking);
    const t = (v) => esc(v === 0 ? "0" : v || "—");
    const doc = `${cabDocumento("Agenda del comité de priorización", "Backlog de demanda", `${lista.length} demanda(s) para decisión · ${fecha(R.hoyISO())}`)}
      <h2>Demandas para decisión (orden sugerido)</h2>
      ${lista.length ? `<table class="hv-t"><thead><tr><th>#</th><th>Almera</th><th>Iniciativa</th><th>Tipificación</th><th>Puntaje</th><th>Solicitante / sponsor</th><th>PM</th><th>Decisión</th></tr></thead><tbody>
        ${lista.map((d, i) => `<tr><td>${i + 1}</td><td class="hv-nw"><b>${t(d.Codigo_Almera)}</b></td><td><b>${t(d.Nombre)}</b><br><small>${t(d.Descripcion)}</small>${d.Beneficio ? `<br><small><b>Beneficio:</b> ${t(d.Beneficio)}</small>` : ""}</td>
          <td>${t(d.Tipo)}<br><small>${t(d.Categoria)} · ${t(d.Area)}${d.Tamano ? ` · Talla ${esc(d.Tamano)}` : ""}${obligatoria(d) ? " · <b>Obligatoria</b>" : ""}</small></td>
          <td class="hv-nw"><b>${puntajeDemanda(d) ?? "—"}</b></td><td>${t(d.Solicitante)}<br><small>${t(d.Sponsor)}</small></td><td>${t(nombreUsuario(d.PM_Asignado))}</td><td style="min-width:90px">☐ Aprobar<br>☐ Aplazar<br>☐ Rechazar</td></tr>`).join("")}</tbody></table>`
        : `<p class="hv-vacio">No hay demandas en estado «Priorizada» o «En comité».</p>`}
      <h2>Criterios y pesos</h2><p class="hv-p">${CRITERIOS_DEM.map((c) => `${esc(c.t)} ${pesosDemanda()[c.k]}%${c.inv ? " (a menor, mejor)" : ""}`).join(" · ")}. Las demandas regulatorias u obligatorias van primero.</p>
      <footer class="hv-pie">Fundación Santa Fe de Bogotá · Oficina de Proyectos · Agenda del comité de priorización</footer>`;
    mostrarDocumento("Agenda del comité", doc, `Agenda comite priorizacion ${R.hoyISO()}`);
  }

  // ---------- Tickets del helpdesk ----------
  const ticketsDe = (pid) => (S.datos.Tickets || []).filter((t) => t.ID_Proyecto === pid);
  const ticketAbierto = (t) => !/^(resuelto|cerrado)$/i.test(String(t.Estado || ""));
  function seccionTickets(p, puede) {
    const ver = S.tkVer || "abiertos";
    const todos = ticketsDe(p.ID_Proyecto);
    const ordenP = { Alta: 0, Media: 1, Baja: 2 };
    const lista = todos.filter((t) => ver === "todos" || ticketAbierto(t))
      .sort((a, b) => (ticketAbierto(b) - ticketAbierto(a)) || ((ordenP[a.Prioridad] ?? 9) - (ordenP[b.Prioridad] ?? 9)) || String(a.Numero).localeCompare(String(b.Numero), "es", { numeric: true }));
    const n = todos.filter(ticketAbierto).length;
    return `<div class="card"><div class="titulo-fila"><h2>Tickets del helpdesk</h2>${puede ? `<button class="btn primario" id="b-tk">+ Ticket</button>` : ""}</div>
      <p class="sub">Los tickets que se están manejando en la plataforma helpdesk para este proyecto. Los estados se configuran en Catálogos.</p>
      <div class="seg barra-comp"><button class="btn chico ${ver === "abiertos" ? "primario" : ""}" data-tkver="abiertos">Abiertos (${n})</button><button class="btn chico ${ver === "todos" ? "primario" : ""}" data-tkver="todos">Todos (${todos.length})</button></div>
      ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>N.º ticket</th><th>Título</th><th>Estado</th><th>Prioridad</th><th class="opc">Registrado</th><th></th></tr></thead>
        <tbody>${lista.map((t) => `<tr><td><b>${esc(t.Numero || "—")}</b></td><td>${esc(t.Titulo)}</td><td>${pill(t.Estado || "Abierto")}</td><td>${esc(t.Prioridad || "—")}</td>
          <td class="opc sub">${t.Fecha_Registro ? fecha(t.Fecha_Registro) : ""} ${esc(nombreUsuario(t.Registrado_Por))}</td>
          <td class="derecha">${puede ? `<button class="btn chico" data-tk="${esc(t.ID_Ticket)}">Editar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`
        : vacio(todos.length ? "No hay tickets abiertos. Mira «Todos»." : `Aún no hay tickets.${puede ? " Registra el primero con «+ Ticket»." : ""}`)}</div>`;
  }
  function formTicket(p, t) {
    const nuevo = !t;
    const campos = [
      { k: "Numero", label: "N.º de ticket", placeholder: "Ej.: 45872" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Ticket, def: (S.cat.Estado_Ticket || [])[0] },
      { k: "Titulo", label: "Título", placeholder: "Asunto del ticket en el helpdesk", ancho: true },
      { k: "Prioridad", label: "Prioridad", tipo: "select", opciones: S.cat.Prioridad, def: "Media" },
    ];
    modal(nuevo ? `Nuevo ticket · ${p.Nombre}` : `Ticket ${t.Numero || ""}`, campos, nuevo ? {} : t, (fd) => {
      if (nuevo) {
        const id = R.siguienteIdHijo("TK", S.datos.Tickets || [], "ID_Ticket", p.ID_Proyecto);
        guardar(() => S.api.agregarFila("Tickets", { ID_Ticket: id, ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo, Fecha_Registro: R.hoyISO() }), "Ticket registrado");
      } else guardar(() => S.api.actualizarPorId("Tickets", "ID_Ticket", t.ID_Ticket, fd), "Ticket actualizado");
    }, (fd) => (nuevo && fd.Numero && ticketsDe(p.ID_Proyecto).some((x) => String(x.Numero) === String(fd.Numero)) ? "Ese número de ticket ya está registrado en este proyecto." : ""),
    nuevo ? {} : { eliminar: { texto: "Eliminar ticket", mensaje: `Se quitará el ticket ${t.Numero || ""} de este proyecto.`,
      accion: () => guardar(() => S.api.eliminarFilas("Tickets", "ID_Ticket", [t.ID_Ticket]), "Ticket eliminado") } });
  }

  // ---------- Capacidad del equipo ----------
  // Cada persona suma su dedicación (%) en los proyectos activos o en pausa: PM (Dedicacion_PM) y stakeholders (Dedicacion).
  const CUENTAN = (p) => p.Estado === "Activo" || p.Estado === "En pausa";
  const clavePersona = (correo, nombre) => (correo ? lc(correo) : "n:" + lc(nombre));
  const num0 = (v) => (v === "" || v === null || v === undefined || Number.isNaN(Number(v)) ? null : Number(v));
  function capacidad() {
    const personas = {};
    const sumar = (correo, nombre, p, rol, ded, idStk) => {
      if (ded === null || !(correo || nombre)) return;
      const k = clavePersona(correo, nombre);
      const x = (personas[k] = personas[k] || { k, nombre: nombre || nombreUsuario(correo), correo: correo || "", total: 0, asign: [] });
      if (!x.correo && correo) x.correo = correo;
      x.total += ded; x.asign.push({ p, rol, ded, idStk });
    };
    S.datos.Proyectos.filter(CUENTAN).forEach((p) => {
      sumar(p.PM, nombreUsuario(p.PM), p, "PM", num0(p.Dedicacion_PM));
      stakeholdersDe(p.ID_Proyecto).forEach((x) => sumar(x.Correo, x.Nombre, p, x.Rol, num0(x.Dedicacion), x.ID_Stakeholder));
    });
    Object.values(personas).forEach((x) => { const r = recursoPor({ correo: x.correo, nombre: x.nombre }); x.cap = (r && num0(r.Capacidad)) ?? 100; if (r) x.nombre = r.Nombre; });
    return Object.values(personas).sort((a, b) => b.total / b.cap - a.total / a.cap || String(a.nombre).localeCompare(String(b.nombre), "es"));
  }
  const nivelCap = (t, cap = 100) => (t > cap ? "Sobreasignado" : t >= cap * 0.85 ? "Al límite" : "Disponible");
  const totalDe = (correo, nombre) => (capacidad().find((x) => x.k === clavePersona(correo, nombre)) || { total: 0, asign: [] });
  const barraCap = (t, cap = 100) => `<span class="cap-barra" title="${t}% de ${cap}%"><span class="cap-lleno ${t > cap ? "rojo" : t >= cap * 0.85 ? "ambar" : ""}" style="width:${Math.min(100, (t / 150) * 100)}%"></span><span class="cap-100" style="left:${(cap / 150) * 100}%"></span></span> <b>${t}%</b>${cap !== 100 ? ` <span class="sub">de ${cap}%</span>` : ""}`;

  function vCapacidad(el) {
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const idsF = new Set(filtrados().map((p) => p.ID_Proyecto));
    const ver = S.capVer || "todos";
    // Se muestran las personas de los proyectos filtrados; su total cuenta todos sus proyectos.
    let lista = capacidad().filter((x) => x.asign.some((a) => idsF.has(a.p.ID_Proyecto)));
    const n = (nv) => lista.filter((x) => nivelCap(x.total, x.cap) === nv).length;
    if (ver !== "todos") lista = lista.filter((x) => nivelCap(x.total, x.cap) === ver);
    el.innerHTML = `
      <h1>Capacidad del equipo</h1>
      ${barraFiltros()}
      <div class="kpis kpis-3">
        <div class="kpi estatico ${n("Sobreasignado") ? "alerta" : ""}"><span class="kpi-t">Sobreasignados (sobre su capacidad)</span><span class="kpi-v">${n("Sobreasignado")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Al límite (85% a 100% de su capacidad)</span><span class="kpi-v">${n("Al límite")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Con capacidad disponible</span><span class="kpi-v">${n("Disponible")}</span></div>
      </div>
      <div class="card"><div class="titulo-fila"><h2>Dedicación por persona</h2>
        <div class="seg">${["todos", "Sobreasignado", "Al límite", "Disponible"].map((v) => `<button class="btn chico ${ver === v ? "primario" : ""}" data-capver="${v}">${v === "todos" ? "Todos" : v}</button>`).join("")}</div></div>
        <p class="sub">Suma el % de dedicación de cada persona en los proyectos activos o en pausa (PM y stakeholders). Se registra en «Editar proyecto» o al agregar un stakeholder. La línea marca la capacidad de cada persona (100% salvo que en Recursos tenga otra).</p>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Persona</th><th>Total</th><th>Estado</th><th>Proyectos</th></tr></thead><tbody>
          ${lista.map((x) => `<tr><td>${persona(x.nombre, x.correo)}</td><td class="nowrap">${barraCap(x.total, x.cap)}</td><td>${pill(nivelCap(x.total, x.cap))}</td>
            <td><div class="cap-asign">${x.asign.sort((a, b) => b.ded - a.ded).map((a) => ids.has(a.p.ID_Proyecto)
              ? `<button class="chip-proy" data-pid="${esc(a.p.ID_Proyecto)}">${esc(a.p.Nombre)} · ${esc(a.rol)} · <b>${a.ded}%</b></button>`
              : `<span class="chip-proy">Otro proyecto · <b>${a.ded}%</b></span>`).join("")}</div></td></tr>`).join("")}</tbody></table></div>`
        : vacio("Aún no hay dedicaciones registradas. Pon el % de dedicación del PM en «Editar proyecto» y el de cada stakeholder en su ficha.")}
      </div>`;
    enlazarFiltros();
    el.querySelectorAll("[data-capver]").forEach((b) => b.addEventListener("click", () => { S.capVer = b.dataset.capver; render(); }));
    el.querySelectorAll(".chip-proy[data-pid]").forEach((b) => b.addEventListener("click", (e) => { e.stopPropagation(); ir("ficha", b.dataset.pid); }));
  }

  // ---------- Stakeholders y proveedores ----------
  const proveedores = () => (S.datos.Proveedores || []).slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
  const proveedor = (id) => (S.datos.Proveedores || []).find((x) => x.ID_Proveedor === id);
  const idsLista = (v) => String(v || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  // Activos, más los inactivos que el proyecto ya tenía (para no perderlos al editar).
  const proveedoresActivos = (actuales) => proveedores().filter((x) => x.Activo !== "No" || idsLista(actuales).includes(x.ID_Proveedor));
  const stakeholdersDe = (pid) => (S.datos.Stakeholders || []).filter((x) => x.ID_Proyecto === pid);
  const contacto = (correo, tel) => [correo ? `<a href="mailto:${esc(correo)}">${esc(correo)}</a>` : "", tel ? esc(tel) : ""].filter(Boolean).join("<br>") || "—";

  function celdaDedicacion(correo, nombre, ded) {
    const d = num0(ded);
    if (d === null) return `<span class="sub">Sin registrar</span>`;
    const t = totalDe(correo, nombre).total;
    return `<b>${d}%</b><div class="sub ${t > 100 ? "baja" : ""}">Total: ${t}%${t > 100 ? " ⚠ sobreasignado" : ""}</div>`;
  }
  function cardStakeholders(p, puede) {
    const roles = S.cat.Rol_Stakeholder || [];
    const orden = (r) => { const i = roles.indexOf(r); return i < 0 ? 99 : i; };
    const lista = stakeholdersDe(p.ID_Proyecto).sort((a, b) => orden(a.Rol) - orden(b.Rol) || String(a.Nombre).localeCompare(String(b.Nombre), "es"));
    const faltan = roles.filter((r) => !lista.some((x) => x.Rol === r));
    const provs = idsLista(p.Proveedores).map(proveedor).filter(Boolean);
    return `<div class="card"><div class="titulo-fila"><h2>Stakeholders y proveedores</h2>${puede ? `<button class="btn" id="b-stk">+ Stakeholder</button>` : ""}</div>
      ${faltan.length ? `<div class="faltan-stk"><span class="sub">Sin asignar:</span> ${faltan.map((r) => puede ? `<button class="btn chico" data-rol-stk="${esc(r)}">+ ${esc(r)}</button>` : `<span class="pill noaplica">${esc(r)}</span>`).join(" ")}</div>` : ""}
      <div class="tabla-scroll"><table><thead><tr><th>Rol</th><th>Nombre</th><th>Cargo / área</th><th>Empresa</th><th>Contacto</th><th>Dedicación</th><th></th></tr></thead>
        <tbody><tr class="fila-pm"><td><b>PM</b></td><td>${persona(nombreUsuario(p.PM))}</td><td>Gerente del proyecto</td><td>Interno</td><td>${contacto(p.PM, "")}</td>
          <td>${celdaDedicacion(p.PM, nombreUsuario(p.PM), p.Dedicacion_PM)}</td><td class="derecha sub">${puede ? "en «Editar proyecto»" : ""}</td></tr>
        ${lista.map((x) => `<tr><td><b>${esc(x.Rol)}</b></td><td>${persona(x.Nombre)}</td><td>${esc([x.Cargo, x.Area].filter(Boolean).join(" · ") || "—")}</td>
          <td>${x.ID_Proveedor ? esc((proveedor(x.ID_Proveedor) || {}).Nombre || x.ID_Proveedor) : "Interno"}</td><td>${contacto(x.Correo, x.Telefono)}</td>
          <td>${celdaDedicacion(x.Correo, x.Nombre, x.Dedicacion)}</td>
          <td class="derecha">${puede ? `<button class="btn chico" data-stk="${esc(x.ID_Stakeholder)}">Editar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>
      <div class="provs"><b>Proveedor(es):</b> ${provs.length ? provs.map((x) => `<span class="prov">${esc(x.Nombre)}${x.Servicio ? ` <span class="sub">· ${esc(x.Servicio)}</span>` : ""}${x.Contacto || x.Correo ? ` <span class="sub">· ${esc(x.Contacto || "")} ${x.Correo ? `<a href="mailto:${esc(x.Correo)}">${esc(x.Correo)}</a>` : ""}</span>` : ""}</span>`).join("")
        : `<span class="sub">Ninguno${puede ? " — asígnalos con «Editar proyecto»." : "."}</span>`}</div></div>`;
  }

  function formStakeholder(p, x, rolSugerido) {
    const nuevo = !x;
    const empresas = [["", "Interno (FSFB)"], ...proveedoresActivos(x && x.ID_Proveedor).map((v) => [v.ID_Proveedor, v.Nombre])];
    const campos = [
      { k: "Rol", label: "Rol en el proyecto", tipo: "select", opciones: S.cat.Rol_Stakeholder, ayuda: "Los roles se configuran en Catálogos." },
      { k: "Nombre", label: "Persona", placeholder: "Escribe y elige del directorio", ayuda: " " },
      { k: "Cargo", label: "Cargo", placeholder: "Ej.: Director médico" },
      { k: "Area", label: "Área", tipo: "select", opciones: S.cat.Cliente_Area },
      { k: "ID_Proveedor", label: "Empresa", tipo: "select", opciones: empresas, ayuda: "Si la persona es del proveedor, elígelo aquí." },
      { k: "Correo", label: "Correo", tipo: "email", placeholder: "nombre@empresa.com" },
      { k: "Telefono", label: "Teléfono", placeholder: "Opcional" },
      { k: "Dedicacion", label: "Dedicación a este proyecto (%)", tipo: "number", min: 0, max: 100, placeholder: "Ej.: 30", ayuda: "Porcentaje de su tiempo. Déjalo vacío si no aplica." },
    ];
    const vivo = { init: (m) => {
      const ded = m.querySelector('[name="Dedicacion"]'), co = m.querySelector('[name="Correo"]'), no = m.querySelector('[name="Nombre"]');
      selectorPersona(m, '[name="Nombre"]', { Correo: '[name="Correo"]', Cargo: '[name="Cargo"]', Area: '[name="Area"]', ID_Proveedor: '[name="ID_Proveedor"]', Telefono: '[name="Telefono"]' }, no.parentElement.querySelector(".sub"));
      no.addEventListener("recurso", () => ded.dispatchEvent(new Event("input")));
      const ay = ded.parentElement.querySelector(".sub");
      const calc = () => {
        const otros = totalDe(co.value.trim(), no.value.trim()).total - (x ? num0(x.Dedicacion) || 0 : 0);
        const t = otros + (Number(ded.value) || 0);
        ay.innerHTML = (co.value || no.value) ? `En otros proyectos: ${otros}% · <b class="${t > 100 ? "baja" : ""}">Total con este: ${t}%${t > 100 ? " ⚠ supera el 100%" : ""}</b>` : "Porcentaje de su tiempo. Déjalo vacío si no aplica.";
      };
      [ded, co, no].forEach((i) => i.addEventListener("input", calc)); calc();
    } };
    modal(nuevo ? `Nuevo stakeholder · ${p.Nombre}` : `Editar stakeholder · ${x.Nombre || x.Rol}`, campos, nuevo ? { Rol: rolSugerido || "" } : x, (fd) => {
      // La persona queda en Recursos (sin duplicar) y sus datos actualizados se reflejan en todos sus proyectos.
      const conRecurso = async () => {
        const antes = recursoPor({ id: x && x.ID_Recurso, correo: fd.Correo, nombre: fd.Nombre });
        const r = await asegurarRecurso(fd);
        if (!r) return "";
        const cambios = {};
        ["Nombre", "Correo", "Cargo", "Area", "Telefono", "ID_Proveedor"].forEach((k) => { if (fd[k] !== undefined && fd[k] !== "" && String(fd[k]) !== String(r[k] || "")) cambios[k] = fd[k]; });
        if (antes && Object.keys(cambios).length) { const previo = { ...r }; await S.api.actualizarPorId("Recursos", "ID_Recurso", r.ID_Recurso, cambios); Object.assign(r, cambios); await propagarRecurso(previo, r); }
        return r.ID_Recurso;
      };
      if (nuevo) {
        const id = R.siguienteIdHijo("STK", S.datos.Stakeholders || [], "ID_Stakeholder", p.ID_Proyecto);
        guardar(async () => { const idr = await conRecurso(); await S.api.agregarFila("Stakeholders", { ID_Stakeholder: id, ID_Proyecto: p.ID_Proyecto, ...fd, ID_Recurso: idr }); }, "Stakeholder agregado");
      } else guardar(async () => { const idr = await conRecurso(); await S.api.actualizarPorId("Stakeholders", "ID_Stakeholder", x.ID_Stakeholder, { ...fd, ID_Recurso: idr }); });
    }, null, nuevo ? vivo : { ...vivo, eliminar: { texto: "Quitar del proyecto", mensaje: `${x.Nombre || "Esta persona"} dejará de aparecer como ${x.Rol || "stakeholder"} de este proyecto.`,
      accion: () => guardar(() => S.api.eliminarFilas("Stakeholders", "ID_Stakeholder", [x.ID_Stakeholder]), "Stakeholder quitado") } });
  }

  function formProveedor(x) {
    const nuevo = !x;
    const usan = nuevo ? [] : S.datos.Proyectos.filter((p) => idsLista(p.Proveedores).includes(x.ID_Proveedor));
    const personas = nuevo ? [] : (S.datos.Stakeholders || []).filter((s) => s.ID_Proveedor === x.ID_Proveedor);
    const campos = [
      { k: "Nombre", label: "Nombre del proveedor", placeholder: "Razón social o nombre comercial", ancho: true },
      { k: "NIT", label: "NIT", placeholder: "Opcional" },
      { k: "Servicio", label: "Servicio / producto", placeholder: "Ej.: Desarrollo, licenciamiento, soporte" },
      { k: "Contacto", label: "Persona de contacto" },
      { k: "Correo", label: "Correo", tipo: "email" },
      { k: "Telefono", label: "Teléfono" },
      { k: "Activo", label: "¿Activo?", tipo: "select", opciones: ["Sí", "No"], def: "Sí", ayuda: "Si no está activo, deja de aparecer para nuevos proyectos." },
    ];
    modal(nuevo ? "Nuevo proveedor" : `Editar ${x.Nombre}`, campos, nuevo ? {} : x, (fd) => {
      if (nuevo) {
        const max = (S.datos.Proveedores || []).reduce((m, v) => Math.max(m, parseInt(String(v.ID_Proveedor).slice(4), 10) || 0), 0);
        guardar(() => S.api.agregarFila("Proveedores", { ID_Proveedor: `PRV-${String(max + 1).padStart(4, "0")}`, ...fd, Nombre: fd.Nombre || "Proveedor sin nombre" }), "Proveedor creado");
      } else guardar(() => S.api.actualizarPorId("Proveedores", "ID_Proveedor", x.ID_Proveedor, fd));
    }, (fd) => ((S.datos.Proveedores || []).some((v) => lc(v.Nombre) === lc(fd.Nombre) && (!x || v.ID_Proveedor !== x.ID_Proveedor)) ? "Ya existe un proveedor con ese nombre." : ""),
    nuevo ? {} : { eliminar: { texto: "Eliminar proveedor",
      mensaje: usan.length || personas.length
        ? `Lo usan ${usan.length} proyecto(s) y ${personas.length} stakeholder(s). Se quitará de ellos (los stakeholders quedan como internos). Si solo dejó de trabajar con ustedes, mejor márcalo como inactivo.`
        : `Se eliminará «${x.Nombre}».`,
      accion: () => guardar(async () => {
        if (usan.length) await S.api.actualizarVarios("Proyectos", "ID_Proyecto", usan.map((p) => ({ id: p.ID_Proyecto, cambios: { Proveedores: idsLista(p.Proveedores).filter((i) => i !== x.ID_Proveedor).join(", ") } })));
        if (personas.length) await S.api.actualizarVarios("Stakeholders", "ID_Stakeholder", personas.map((s) => ({ id: s.ID_Stakeholder, cambios: { ID_Proveedor: "" } })));
        await S.api.eliminarFilas("Proveedores", "ID_Proveedor", [x.ID_Proveedor]);
      }, "Proveedor eliminado") } });
  }

  function formProyecto(p, pref, alCrear) {
    const u = S.usuario;
    const nuevo = !p;
    const pmFijo = R.soloPM(u);
    const pms = usuariosConRol("PM").sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es")).map((x) => [x.Correo, x.Nombre || x.Correo]);
    const frecuenciaLibre = R.puede(u, "editarFrecuencia");
    const campos = [
      { seccion: "1. Datos generales" },
      { k: "Nombre", label: "Nombre de la iniciativa", placeholder: "Ej.: Migración del ERP", ancho: true },
      { k: "Codigo_Almera", label: "Código Almera", placeholder: "Código del proyecto en Almera" },
      { k: "Cliente_Area", label: "Cliente / área", tipo: "select", opciones: S.cat.Cliente_Area },
      { k: "PM", label: "PM responsable", tipo: "select", opciones: pms, bloqueado: pmFijo,
        ayuda: pmFijo ? "Quedas como PM de esta iniciativa." : "Solo aparecen usuarios con rol PM. El PM verá este proyecto." },
      { k: "Dedicacion_PM", label: "Dedicación del PM a este proyecto (%)", tipo: "number", min: 0, max: 100, placeholder: "Ej.: 40", ayuda: "Suma en «Capacidad del equipo»." },
      { k: "Metodologia", label: "Metodología", tipo: "select", opciones: S.cat.Metodologia },
      { k: "Prioridad", label: "Prioridad", tipo: "select", opciones: S.cat.Prioridad, def: "Media" },
      { seccion: "2. Estado" },
      { k: "Fase", label: "Fase", tipo: "select", opciones: S.cat.Fase, def: "Inicio" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado, def: "Activo" },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo, def: "Verde" },
      { k: "Frecuencia_Seguimiento", label: "Frecuencia de seguimiento", tipo: "select", opciones: S.cat.Frecuencia_Seguimiento, def: "Semanal",
        bloqueado: !frecuenciaLibre, ayuda: frecuenciaLibre ? "Cada cuánto debe reportar el PM." : "La define la PMO." },
      { seccion: "3. Fechas y avance" },
      { k: "Fecha_Inicio", label: "Fecha de inicio", tipo: "date", def: R.hoyISO() },
      { k: "Fecha_Fin_Plan", label: "Fecha fin planeada", tipo: "date" },
      { k: "Avance_Real", label: "Avance real (%)", tipo: "number", min: 0, max: 100, def: 0, ayuda: "Lo que lleva hoy. También se actualiza con cada seguimiento." },
      { k: "Avance_Planeado", label: "Avance planeado a hoy (%)", tipo: "number", min: 0, max: 100, def: 0, ayuda: "Lo que debería llevar según el plan." },
      { seccion: "4. Presupuesto", ayuda: "en pesos colombianos, sin puntos" },
      { k: "Presupuesto", label: "Presupuesto (COP)", tipo: "number", min: 0, def: 0 },
      { k: "Ejecutado", label: "Ejecutado (COP)", tipo: "number", min: 0, def: 0 },
      { seccion: "5. Priorización", ayuda: "califica de 1 (muy bajo) a 5 (muy alto); alimenta el menú Priorización" },
      ...CRITERIOS.map((c) => ({ k: c.k, label: c.t, tipo: "select", opciones: [[1, "1 · Muy bajo"], [2, "2 · Bajo"], [3, "3 · Medio"], [4, "4 · Alto"], [5, "5 · Muy alto"]] })),
      { seccion: "6. Proveedores", ayuda: "márcalos o crea uno nuevo aquí mismo" },
      { k: "Proveedores", label: "Proveedor(es) del proyecto", tipo: "checks", vacio: "Aún no hay proveedores. Créalos en Catálogos → Proveedores.",
        opciones: proveedoresActivos(p ? p.Proveedores : "").map((x) => [x.ID_Proveedor, x.Nombre]) },
      { seccion: "7. Equipo y stakeholders", ayuda: "roles, personas y su dedicación" },
      { html: `<div id="fp-equipo"></div>` },
      { seccion: "8. Enlaces", ayuda: "opcionales; pega la dirección completa" },
      { k: "URL_Repositorio", label: "URL del repositorio", tipo: "url", placeholder: "https://github.com/… o https://dev.azure.com/…" },
      { k: "URL_Documentos", label: "URL de documentos", tipo: "url", placeholder: "https://…sharepoint.com/…" },
      { seccion: "9. Comentario" },
      { k: "Comentario_Estado", label: "Comentario de estado", tipo: "textarea", placeholder: "Novedad principal del proyecto" },
    ];
    const valores = p ? { ...p } : { ...(pref || {}), PM: pmFijo ? u.Correo : (pref && pref.PM) || "" };
    // Equipo: personas nuevas se guardan con el proyecto; los roles nuevos van de una vez al catálogo.
    const pendStk = [];
    const initEquipo = (m) => {
      const cont = m.querySelector("#fp-equipo");
      const actuales = p ? stakeholdersDe(p.ID_Proyecto) : [];
      const opcionesRol = () => (S.cat.Rol_Stakeholder || []).map((r) => `<option>${esc(r)}</option>`).join("");
      cont.innerHTML = `
        <ul class="eq-lista" id="eq-lista"></ul>
        <div class="acciones"><button type="button" class="btn chico" id="eq-abrir">+ Agregar persona</button><button type="button" class="btn chico" id="rol-abrir">+ Crear rol nuevo</button></div>
        <div id="rol-form" class="eq-form" hidden><input id="rol-nombre" placeholder="Nombre del rol (ej.: Líder técnico, Arquitecto, QA)" aria-label="Nuevo rol">
          <button type="button" class="btn primario chico" id="rol-crear">Crear rol</button><button type="button" class="btn chico" id="rol-cancelar">Cancelar</button></div>
        <div id="eq-form" class="eq-form" hidden><div class="pn-grid">
          <select id="eq-rol" aria-label="Rol">${opcionesRol()}</select>
          <input id="eq-nombre" placeholder="Persona * (elige del directorio)" aria-label="Persona" autocomplete="off">
          <input id="eq-correo" type="email" placeholder="Correo" aria-label="Correo">
          <input id="eq-ded" type="number" min="0" max="100" placeholder="Dedicación %" aria-label="Dedicación"></div>
          <div class="sub" id="eq-rec"></div><div class="sub" id="eq-total"></div>
          <div class="acciones"><button type="button" class="btn primario chico" id="eq-agregar">Agregar al equipo</button><button type="button" class="btn chico" id="eq-cancelar">Cancelar</button></div></div>
        <div class="error-campo" id="eq-error"></div>`;
      const pintar = () => {
        m.querySelector("#eq-lista").innerHTML = [
          ...actuales.map((x) => `<li><b>${esc(x.Rol)}</b> · ${esc(x.Nombre)}${num0(x.Dedicacion) !== null ? ` · ${num0(x.Dedicacion)}%` : ""} <span class="sub">(se edita en la ficha)</span></li>`),
          ...pendStk.map((x, i) => `<li class="nuevo"><b>${esc(x.Rol)}</b> · ${esc(x.Nombre)}${x.Dedicacion !== "" ? ` · ${x.Dedicacion}%` : ""} <span class="pill encurso">nuevo</span> <button type="button" class="btn enlace" data-q="${i}">✕</button></li>`),
        ].join("") || `<li class="sub">Sin personas aún. Agrega Sponsor, Líder funcional, Product Owner…</li>`;
        m.querySelectorAll("[data-q]").forEach((b) => b.addEventListener("click", () => { pendStk.splice(Number(b.dataset.q), 1); pintar(); }));
      };
      pintar();
      const $m = (id) => m.querySelector(id);
      const err = $m("#eq-error");
      const gente = [...S.datos.Usuarios.map((x) => [x.Nombre, x.Correo]), ...(S.datos.Stakeholders || []).map((x) => [x.Nombre, x.Correo])];
      selectorPersona(m, "#eq-nombre", { Correo: "#eq-correo" }, "#eq-rec");
      $m("#eq-nombre").addEventListener("recurso", () => tot());
      $m("#eq-nombre").addEventListener("change", () => { const g = gente.find((x) => lc(x[0]) === lc($m("#eq-nombre").value) && x[1]); if (g && !$m("#eq-correo").value) $m("#eq-correo").value = g[1]; tot(); });
      const tot = () => {
        const n = $m("#eq-nombre").value.trim(), c = $m("#eq-correo").value.trim();
        if (!n && !c) { $m("#eq-total").textContent = ""; return; }
        const otros = totalDe(c, n).total, t = otros + (Number($m("#eq-ded").value) || 0);
        $m("#eq-total").innerHTML = `Hoy tiene asignado: ${otros}% · <b class="${t > 100 ? "baja" : ""}">Total: ${t}%${t > 100 ? " ⚠ supera el 100%" : ""}</b>`;
      };
      ["#eq-ded", "#eq-correo"].forEach((id) => $m(id).addEventListener("input", tot));
      $m("#eq-abrir").addEventListener("click", () => { $m("#eq-form").hidden = false; $m("#rol-form").hidden = true; $m("#eq-nombre").focus(); });
      $m("#eq-cancelar").addEventListener("click", () => { $m("#eq-form").hidden = true; });
      $m("#rol-abrir").addEventListener("click", () => { $m("#rol-form").hidden = false; $m("#rol-nombre").focus(); });
      $m("#rol-cancelar").addEventListener("click", () => { $m("#rol-form").hidden = true; });
      $m("#eq-agregar").addEventListener("click", () => {
        const nombre = $m("#eq-nombre").value.trim(), correo = $m("#eq-correo").value.trim(), ded = $m("#eq-ded").value.trim();
        if (!nombre) { err.textContent = "Escribe el nombre de la persona."; return; }
        if (correo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) { err.textContent = "Correo no válido."; return; }
        if (ded !== "" && (Number(ded) < 0 || Number(ded) > 100)) { err.textContent = "La dedicación va de 0 a 100."; return; }
        pendStk.push({ Rol: $m("#eq-rol").value, Nombre: nombre, Correo: correo, Dedicacion: ded === "" ? "" : Number(ded) });
        ["#eq-nombre", "#eq-correo", "#eq-ded"].forEach((id) => { $m(id).value = ""; });
        $m("#eq-total").textContent = ""; err.textContent = ""; pintar();
      });
      $m("#rol-crear").addEventListener("click", async () => {
        const rol = $m("#rol-nombre").value.trim();
        if (!rol) { err.textContent = "Escribe el nombre del rol."; return; }
        if ((S.cat.Rol_Stakeholder || []).some((r) => lc(r) === lc(rol))) { err.textContent = "Ese rol ya existe."; return; }
        cargando(true, "Creando rol…");
        try {
          await S.api.agregarFila("Catalogos", { Lista: "Rol_Stakeholder", Valor: rol });
          S.datos.Catalogos.push({ Lista: "Rol_Stakeholder", Valor: rol }); (S.cat.Rol_Stakeholder = S.cat.Rol_Stakeholder || []).push(rol);
          $m("#eq-rol").innerHTML = opcionesRol(); $m("#eq-rol").value = rol;
          $m("#rol-nombre").value = ""; $m("#rol-form").hidden = true; $m("#eq-form").hidden = false; err.textContent = "";
          toast(`Rol «${rol}» creado`);
        } catch (e) { err.textContent = "No se pudo crear el rol: " + e.message; }
        finally { cargando(false); }
      });
    };
    const guardarEquipo = async (pid) => {
      if (!pendStk.length) return;
      const base = S.datos.Stakeholders || [];
      const recs = [];
      for (const x of pendStk) recs.push(await asegurarRecurso({ Nombre: x.Nombre, Correo: x.Correo }));
      await S.api.agregarFilas("Stakeholders", pendStk.map((x, i) => { const r = recs[i] || {}; return { ID_Stakeholder: R.siguienteIdHijo("STK", base, "ID_Stakeholder", pid, i), ID_Proyecto: pid,
        Rol: x.Rol, Nombre: r.Nombre || x.Nombre, Cargo: r.Cargo || "", Area: r.Area || "", Correo: x.Correo || r.Correo || "", Telefono: r.Telefono || "", ID_Proveedor: r.ID_Proveedor || "", Dedicacion: x.Dedicacion, ID_Recurso: r.ID_Recurso || "" }; }));
    };
    // Crear un proveedor sin salir del formulario: se guarda de inmediato y queda marcado.
    const initProv = (m) => {
      const checks = m.querySelector('[data-checks="Proveedores"]');
      if (!checks) return;
      checks.insertAdjacentHTML("afterend", `<div class="prov-nuevo">
        <button type="button" class="btn chico" id="pn-abrir">+ Crear proveedor nuevo</button>
        <div id="pn-form" hidden><div class="pn-grid">
          <input id="pn-nombre" placeholder="Nombre del proveedor *" aria-label="Nombre del proveedor">
          <input id="pn-servicio" placeholder="Servicio / producto" aria-label="Servicio">
          <input id="pn-contacto" placeholder="Persona de contacto" aria-label="Contacto">
          <input id="pn-correo" type="email" placeholder="Correo" aria-label="Correo"></div>
          <div class="acciones"><button type="button" class="btn primario chico" id="pn-crear">Crear y marcar</button><button type="button" class="btn chico" id="pn-cancelar">Cancelar</button></div>
          <div class="error-campo" id="pn-error"></div></div></div>`);
      const f = m.querySelector("#pn-form"), ab = m.querySelector("#pn-abrir");
      ab.addEventListener("click", () => { f.hidden = false; ab.hidden = true; m.querySelector("#pn-nombre").focus(); });
      m.querySelector("#pn-cancelar").addEventListener("click", () => { f.hidden = true; ab.hidden = false; });
      m.querySelector("#pn-crear").addEventListener("click", async () => {
        const v = (k) => m.querySelector("#pn-" + k).value.trim();
        const err = m.querySelector("#pn-error");
        if (!v("nombre")) { err.textContent = "Escribe el nombre del proveedor."; return; }
        if ((S.datos.Proveedores || []).some((x) => lc(x.Nombre) === lc(v("nombre")))) { err.textContent = "Ya existe un proveedor con ese nombre: márcalo en la lista."; return; }
        if (v("correo") && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v("correo"))) { err.textContent = "Correo no válido."; return; }
        const max = (S.datos.Proveedores || []).reduce((mx, x) => Math.max(mx, parseInt(String(x.ID_Proveedor).slice(4), 10) || 0), 0);
        const fila = { ID_Proveedor: `PRV-${String(max + 1).padStart(4, "0")}`, Nombre: v("nombre"), NIT: "", Servicio: v("servicio"), Contacto: v("contacto"), Correo: v("correo"), Telefono: "", Activo: "Sí" };
        cargando(true, "Creando proveedor…");
        try {
          await S.api.agregarFila("Proveedores", fila);
          (S.datos.Proveedores = S.datos.Proveedores || []).push(fila);
          const vacioTxt = checks.querySelector(".sub"); if (vacioTxt) vacioTxt.remove();
          checks.insertAdjacentHTML("beforeend", `<label class="check"><input type="checkbox" value="${esc(fila.ID_Proveedor)}" checked> ${esc(fila.Nombre)}</label>`);
          ["nombre", "servicio", "contacto", "correo"].forEach((k) => { m.querySelector("#pn-" + k).value = ""; });
          err.textContent = ""; f.hidden = true; ab.hidden = false;
          toast(`Proveedor «${fila.Nombre}» creado y marcado`);
        } catch (e) { err.textContent = "No se pudo crear: " + e.message; }
        finally { cargando(false); }
      });
    };
    modal(nuevo ? "Nueva iniciativa" : `Editar ${p.ID_Proyecto}`, campos, valores, (fd) => {
      if (pmFijo) fd.PM = nuevo ? u.Correo : p.PM;
      if (!frecuenciaLibre) fd.Frecuencia_Seguimiento = nuevo ? "Semanal" : p.Frecuencia_Seguimiento;
      if (nuevo) {
        const id = R.siguienteIdProyecto(S.datos.Proyectos);
        const fila = { ...fd, Nombre: fd.Nombre || `Iniciativa ${id}`, ID_Proyecto: id, Avance_Real: fd.Avance_Real === "" ? 0 : fd.Avance_Real, ...sello() };
        S.pid = fila.ID_Proyecto;
        S.vista = "ficha";
        guardar(async () => { await S.api.agregarFila("Proyectos", fila); await guardarEquipo(id); if (alCrear) await alCrear(id); }, `Iniciativa ${fila.ID_Proyecto} registrada`);
      } else {
        const seCierra = fd.Estado === "Cerrado" && p.Estado !== "Cerrado";
        if (seCierra) S.fichaTab = "lecciones";
        guardar(async () => { await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...fd, ...sello() }); await guardarEquipo(p.ID_Proyecto); },
          seCierra ? "Proyecto cerrado. Es buen momento para registrar las lecciones aprendidas." : "Guardado");
      }
    }, (fd) => (fd.Fecha_Inicio && fd.Fecha_Fin_Plan && fd.Fecha_Fin_Plan < fd.Fecha_Inicio ? "La fecha fin no puede ser anterior a la fecha de inicio." : ""),
    !nuevo && R.puede(u, "verTodo") ? { init: (m) => { initProv(m); initEquipo(m); }, eliminar: { texto: "Eliminar proyecto", titulo: `Eliminar ${p.ID_Proyecto}`,
      mensaje: `Se borrará «${p.Nombre}» con todos sus seguimientos, compromisos, hitos, riesgos, stakeholders y actas. No se puede deshacer. Si solo terminó, mejor cambia su estado a Cerrado o Cancelado.`,
      accion: () => { S.vista = "proyectos"; guardar(() => eliminarProyecto(p.ID_Proyecto), `Proyecto ${p.ID_Proyecto} eliminado`); } } } : { init: (m) => { initProv(m); initEquipo(m); } });
  }

  function formSeguimiento(p) {
    const e = R.estadoSeguimiento(p, S.datos.Seguimientos);
    const ult = e.ultimo;
    const abiertos = S.datos.Compromisos.filter((c) => c.ID_Proyecto === p.ID_Proyecto && !R.cerrado(c))
      .sort((a, b) => (a.Fecha_Compromiso < b.Fecha_Compromiso ? -1 : 1));
    const campos = [
      { seccion: "1. Avance" },
      { k: "Fecha_Corte", label: "Fecha de corte", tipo: "date", def: R.hoyISO() },
      { k: "Avance_Real", label: "Avance real (%)", tipo: "number", min: 0, max: 100, def: p.Avance_Real },
      { k: "Avance_Planeado", label: "Avance planeado a la fecha (%)", tipo: "number", min: 0, max: 100, def: p.Avance_Planeado, ayuda: "Lo que debería llevar según el plan." },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo, def: p.Semaforo },
      { k: "Ejecutado", label: "Presupuesto ejecutado a la fecha (COP)", tipo: "number", min: 0, def: p.Ejecutado, ayuda: "Acumulado; alimenta la curva S." },
      { seccion: "2. Qué pasó en el periodo" },
      { k: "Logros", label: "Logros del periodo", tipo: "textarea", placeholder: "Qué se terminó o avanzó" },
      { k: "Proximos_Pasos", label: "Próximos pasos", tipo: "textarea", placeholder: "Qué sigue hasta el próximo reporte" },
      { k: "Bloqueos", label: "Bloqueos", tipo: "textarea", placeholder: "Qué frena el avance y quién debe actuar (déjalo vacío si no hay)" },
      { seccion: "3. Acta de la sesión" },
      { k: "Fecha_Acta", label: "Fecha del acta", tipo: "date", def: R.hoyISO() },
      { k: "URL_Acta", label: "Enlace al acta (opcional)", tipo: "url", placeholder: "https://…", ayuda: "Solo si además la tienes en SharePoint o Teams." },
    ];
    const cargaActa = `<div class="carga-acta">
      <div><b>¿Tienes el acta en PDF?</b><div class="sub">Súbela: la app lee la fecha y los compromisos, los llena abajo y guarda el acta con la sesión.</div></div>
      <button type="button" class="btn primario" id="b-acta-pdf">Subir acta (PDF)</button>
      <div id="acta-estado" class="acta-estado" aria-live="polite"></div></div>`;
    const contexto = `${datalistRecursos("dl-recursos")}<div class="contexto">
      <div><span class="sub">Último reporte</span><b>${ult ? `${fecha(ult.Fecha_Corte)} · ${pct(ult.Avance_Real)} · ${esc(ult.Semaforo)}` : "Es el primer reporte"}</b></div>
      <div><span class="sub">Plan registrado</span><b>${pct(p.Avance_Planeado)}</b></div>
      <div><span class="sub">Frecuencia</span><b>${esc(p.Frecuencia_Seguimiento)}</b></div></div>`;
    const filaComp = (c = {}) => `<div class="comp-fila">
      <input class="c-texto" placeholder="Compromiso (qué se hará)" aria-label="Compromiso" value="${esc(c.Compromiso || "")}">
      <input class="c-resp" placeholder="Responsable (de Recursos)" aria-label="Responsable" data-recurso autocomplete="off" value="${esc(c.Responsable || "")}">
      <input class="c-fecha" type="date" aria-label="Fecha límite" value="${esc(c.Fecha_Compromiso || "")}" title="${esc(c.Fecha_Texto && !c.Fecha_Compromiso ? `En el acta: ${c.Fecha_Texto}` : "")}">
      <button type="button" class="btn chico c-quitar" aria-label="Quitar compromiso">✕</button></div>`;
    const despues = `
      <div class="sugerencia" id="sug-sem" aria-live="polite"></div>
      ${abiertos.length ? `<div class="form-seccion">4. Compromisos anteriores<span class="sub"> · marca los que ya se cumplieron</span></div>
        <div class="comp-previos">${abiertos.map((c) => `<label class="check"><input type="checkbox" value="${esc(c.ID_Compromiso)}"> <span>${esc(c.Compromiso)}
          <span class="sub">${esc(c.Responsable)} · ${fecha(c.Fecha_Compromiso)}</span> ${R.estadoCompromiso(c) === "Vencido" ? pill("Vencido") : ""}</span></label>`).join("")}</div>` : ""}
      <div class="form-seccion">${abiertos.length ? "5" : "4"}. Compromisos nuevos<span class="sub"> · acuerdos que quedaron en este seguimiento (opcional)</span></div>
      <div class="comp-cab"><span>Compromiso</span><span>Responsable</span><span>Fecha límite</span><span></span></div>
      <div id="comp-nuevos">${filaComp()}</div>
      <button type="button" class="btn chico" id="b-add-comp">+ Agregar otro compromiso</button>`;
    const init = (m) => {
      const caja = $("#sug-sem");
      let sugerido = "";
      const sug = () => {
        const real = m.querySelector("[name=Avance_Real]").value, plan = m.querySelector("[name=Avance_Planeado]").value;
        let html = "";
        sugerido = "";
        if (real !== "" && plan !== "") {
          const s = R.semaforoSugerido(real, plan);
          sugerido = s.semaforo;
          const actual = m.querySelector("[name=Semaforo]").value;
          const texto = s.desviacion > 0 ? `${s.desviacion} pts por debajo del plan` : s.desviacion < 0 ? `${-s.desviacion} pts por encima del plan` : "igual al plan";
          html = `Semáforo sugerido: ${chipSemaforo(s.semaforo)} (${texto}).${actual !== s.semaforo ? ` <button type="button" class="btn enlace" data-usar-sug>Usar sugerido</button>` : " Coincide con el elegido."}`;
        }
        if (caja.innerHTML !== html) caja.innerHTML = html;   // no redibujar si no cambió (evita perder el clic)
      };
      // Delegado en la caja: el botón puede redibujarse sin perder el evento.
      caja.addEventListener("mousedown", (ev) => { if (ev.target.closest("[data-usar-sug]")) ev.preventDefault(); });
      caja.addEventListener("click", (ev) => {
        if (!ev.target.closest("[data-usar-sug]") || !sugerido) return;
        m.querySelector("[name=Semaforo]").value = sugerido;
        sug();
      });
      ["Avance_Real", "Avance_Planeado", "Semaforo"].forEach((k) => ["input", "change"].forEach((t) => m.querySelector(`[name=${k}]`).addEventListener(t, sug)));
      sug();
      const cont = $("#comp-nuevos");
      const enlazar = () => cont.querySelectorAll(".c-quitar").forEach((b) => { b.onclick = () => { if (cont.children.length > 1) b.parentElement.remove(); else b.parentElement.querySelectorAll("input").forEach((i) => { i.value = ""; }); }; });
      enlazar();
      $("#b-add-comp").addEventListener("click", () => { cont.insertAdjacentHTML("beforeend", filaComp()); enlazar(); cont.lastElementChild.querySelector("input").focus(); });
      $("#b-acta-pdf").addEventListener("click", () => elegirPDF(async (archivo) => {
        const est = $("#acta-estado");
        if (archivo.size > MAX_ACTA) { est.className = "acta-estado error"; est.textContent = "El acta pesa más de 5 MB. Guarda una versión más liviana del PDF."; return; }
        est.className = "acta-estado"; est.textContent = "Leyendo el acta…";
        try {
          const res = await ACTAS.leerActa(archivo);
          acta = { bytes: res.datos, nombre: archivo.name };
          if (res.fecha) {
            m.querySelector("[name=Fecha_Acta]").value = res.fecha;
            if (res.fecha <= R.hoyISO()) m.querySelector("[name=Fecha_Corte]").value = res.fecha;
          }
          [...cont.querySelectorAll(".comp-fila")].forEach((f) => { if (![...f.querySelectorAll("input")].some((i) => i.value)) f.remove(); });
          res.compromisos.forEach((c) => cont.insertAdjacentHTML("beforeend", filaComp(c)));
          if (!cont.children.length) cont.insertAdjacentHTML("beforeend", filaComp());
          enlazar();
          const sinFecha = res.compromisos.filter((c) => !c.Fecha_Compromiso).length;
          est.className = "acta-estado ok";
          est.innerHTML = `<b>${esc(archivo.name)}</b> se guardará con la sesión. ${res.fecha ? `Fecha del acta: ${fecha(res.fecha)}. ` : "No encontré la fecha del acta. "}` +
            (res.encontrada ? `Cargué ${res.compromisos.length} compromiso(s) abajo${sinFecha ? ` (${sinFecha} sin fecha en el acta)` : ""}; revísalos antes de guardar.` : "No encontré la tabla «Compromisos de la reunión»; escríbelos abajo si hay.");
        } catch (e) {
          acta = null;
          est.className = "acta-estado error"; est.textContent = "No se pudo leer el PDF: " + e.message;
        }
      }));
    };
    let acta = null;
    const recoger = (m, fd) => {
      const nuevos = [];
      for (const f of m.querySelectorAll(".comp-fila")) {
        const t = f.querySelector(".c-texto").value.trim(), r = f.querySelector(".c-resp").value.trim(), d = f.querySelector(".c-fecha").value;
        if (!t && !r && !d) continue;
        if (d && fd.Fecha_Corte && d < fd.Fecha_Corte) return { error: `La fecha del compromiso «${t || "sin descripción"}» no puede ser anterior a la fecha de corte.` };
        nuevos.push({ Compromiso: t || "(sin descripción)", Responsable: r, Fecha_Compromiso: d });
      }
      const cerrados = [...m.querySelectorAll(".comp-previos input:checked")].map((i) => i.value);
      return { datos: { nuevos, cerrados } };
    };
    modal(`Seguimiento · ${p.Nombre}`, campos, {}, (fd, extra) => {
      if (!fd.Fecha_Corte) fd.Fecha_Corte = R.hoyISO();
      const idSeg = R.siguienteIdHijo("SEG", S.datos.Seguimientos, "ID_Seguimiento", p.ID_Proyecto);
      const fila = {
        ID_Seguimiento: idSeg, ID_Proyecto: p.ID_Proyecto, Fecha_Corte: fd.Fecha_Corte, Semana: R.semanaISO(fd.Fecha_Corte),
        Avance_Real: fd.Avance_Real, Semaforo: fd.Semaforo, Ejecutado: fd.Ejecutado, Logros: fd.Logros, Proximos_Pasos: fd.Proximos_Pasos,
        Bloqueos: fd.Bloqueos, Reportado_Por: S.usuario.Correo, Fecha_Acta: fd.Fecha_Acta, URL_Acta: fd.URL_Acta,
        Acta_Archivo: acta ? acta.nombre : "",
      };
      const actaElegida = acta;
      const comps = extra.nuevos.map((c, i) => ({
        ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", p.ID_Proyecto, i),
        ID_Proyecto: p.ID_Proyecto, ID_Seguimiento: idSeg, ...c, Correo_Responsable: (recursoPor({ nombre: c.Responsable }) || {}).Correo || "", Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo,
      }));
      guardar(async () => {
        await S.api.agregarFila("Seguimientos", fila);
        await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}),
          ...(fd.Avance_Planeado !== "" ? { Avance_Planeado: fd.Avance_Planeado } : {}),
          ...(fd.Ejecutado !== "" ? { Ejecutado: fd.Ejecutado } : {}),
          ...(fd.Semaforo ? { Semaforo: fd.Semaforo } : {}),
          ...(fd.Bloqueos || fd.Logros ? { Comentario_Estado: fd.Bloqueos ? `Bloqueo: ${fd.Bloqueos}` : fd.Logros } : {}), ...sello(),
        });
        if (comps.length) await S.api.agregarFilas("Compromisos", comps);
        if (extra.cerrados.length) {
          await S.api.actualizarVarios("Compromisos", "ID_Compromiso", extra.cerrados.map((id) => ({ id, cambios: { Estado: "Cerrado", Fecha_Cierre: fd.Fecha_Corte } })));
          const cerradosC = S.datos.Compromisos.filter((c) => extra.cerrados.includes(c.ID_Compromiso));
          const nuevosCom = [];
          cerradosC.forEach((c) => { const nc = nuevoComentario(c, `Cerrado en la sesión de seguimiento del ${fecha(fd.Fecha_Corte)}.`); nuevosCom.push(nc); });
          if (nuevosCom.length) await S.api.agregarFilas("Comentarios", nuevosCom);
        }
        if (actaElegida) await subirActa(idSeg, actaElegida.bytes, actaElegida.nombre);
      }, `Seguimiento registrado${comps.length ? ` con ${comps.length} compromiso(s)` : ""}${actaElegida ? " y acta guardada" : ""}`);
    }, (fd) => (fd.Fecha_Corte && fd.Fecha_Corte > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""), { antes: cargaActa + contexto, despues, init, recoger });
  }

  function formHito(p, h) {
    const campos = [
      { k: "Hito", label: "Hito", placeholder: "Ej.: Salida a producción", ancho: true },
      { k: "Fecha_Plan", label: "Fecha planeada", tipo: "date" },
      { k: "Fecha_Real", label: "Fecha real", tipo: "date", ayuda: "Solo si ya se cumplió; el hito queda como Cumplido." },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Hito, def: "Pendiente" },
    ];
    modal(h ? "Editar hito" : `Nuevo hito · ${p.Nombre}`, campos, h || {}, (fd) => {
      if (fd.Fecha_Real) fd.Estado = "Cumplido";
      if (h) guardar(() => S.api.actualizarPorId("Hitos", "ID_Hito", h.ID_Hito, fd));
      else guardar(() => S.api.agregarFila("Hitos", { ID_Hito: R.siguienteIdHijo("HIT", S.datos.Hitos, "ID_Hito", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }), "Hito agregado");
    }, null, h ? { eliminar: { texto: "Eliminar hito", mensaje: `Se borrará el hito «${h.Hito}».`, accion: () => guardar(() => S.api.eliminarFilas("Hitos", "ID_Hito", [h.ID_Hito]), "Hito eliminado") } } : {});
  }

  function formRiesgo(p, r) {
    const probs = Object.entries(PROB_TXT).map(([k, t]) => [Number(k), t]);
    const imps = Object.entries(IMP_TXT).map(([k, t]) => [Number(k), t]);
    const campos = [
      { seccion: "1. Riesgo" },
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: S.cat.Tipo_Riesgo, def: "Amenaza" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Riesgo, def: "Abierto" },
      { k: "Responsable", label: "Responsable", def: S.usuario.Nombre || S.usuario.Correo, placeholder: "Elige del directorio de recursos", ayuda: " " },
      { k: "Descripcion", label: "Descripción del riesgo", tipo: "textarea", placeholder: "Si ocurre X, entonces Y" },
      { seccion: "2. ¿Dónde se identificó?" },
      { k: "ID_Seguimiento", label: "Sesión de seguimiento", tipo: "select", opciones: opcionesSesion(p.ID_Proyecto), textoVacio: "No salió de una sesión registrada" },
      { k: "Origen", label: "Instancia o fuente", placeholder: "Ej.: Comité Operativo, Kickoff, revisión técnica" },
      { k: "Fecha_Identificacion", label: "Fecha de identificación", tipo: "date", def: R.hoyISO() },
      { seccion: "3. Antes de mitigar (inherente)" },
      { k: "Probabilidad_Inherente", label: "Probabilidad (inherente)", tipo: "select", opciones: probs },
      { k: "Impacto_Inherente", label: "Impacto (inherente)", tipo: "select", opciones: imps },
      { html: `<div class="calif-viva" id="cv-inh"></div>` },
      { seccion: "4. Planes" },
      { k: "Plan_Mitigacion", label: "Plan de mitigación", tipo: "textarea", placeholder: "Qué hacemos para que no ocurra o afecte menos" },
      { k: "Plan_Contingencia", label: "Plan de contingencia", tipo: "textarea", placeholder: "Qué hacemos si ocurre" },
      { seccion: "5. Después de mitigar (residual)" },
      { k: "Probabilidad_Residual", label: "Probabilidad (residual)", tipo: "select", opciones: probs },
      { k: "Impacto_Residual", label: "Impacto (residual)", tipo: "select", opciones: imps },
      { html: `<div class="calif-viva" id="cv-res"></div>` },
    ];
    const init = (m) => {
      selectorPersona(m, '[name="Responsable"]', {}, m.querySelector('[name="Responsable"]').parentElement.querySelector(".sub"));
      const calc = () => {
        const v = (k) => Number(m.querySelector(`[name="${k}"]`).value) || 0;
        m.querySelector("#cv-inh").innerHTML = `Calificación inherente: <b>${califTxt(v("Probabilidad_Inherente"), v("Impacto_Inherente"))}</b>`;
        m.querySelector("#cv-res").innerHTML = `Calificación residual: <b>${califTxt(v("Probabilidad_Residual"), v("Impacto_Residual"))}</b>`;
      };
      m.querySelectorAll("select").forEach((x) => x.addEventListener("change", calc)); calc();
    };
    modal(r ? `Editar riesgo ${numRiesgo(r)}` : `Nuevo riesgo · ${p.Nombre}`, campos, r || {}, (fd) => {
      ["Probabilidad_Inherente", "Impacto_Inherente", "Probabilidad_Residual", "Impacto_Residual"].forEach((k) => { fd[k] = fd[k] === "" ? "" : Number(fd[k]); });
      fd.Calificacion_Inherente = R.calificacion(fd.Probabilidad_Inherente, fd.Impacto_Inherente) || "";
      fd.Calificacion_Residual = R.calificacion(fd.Probabilidad_Residual, fd.Impacto_Residual) || "";
      const resp = async () => { if (fd.Responsable) { const rec = await asegurarRecurso({ Nombre: fd.Responsable }); if (rec) fd.Responsable = rec.Nombre; } };
      if (r) guardar(async () => { await resp(); await S.api.actualizarPorId("Riesgos", "ID_Riesgo", r.ID_Riesgo, fd); });
      else guardar(async () => { await resp(); await S.api.agregarFila("Riesgos", { ID_Riesgo: R.siguienteIdHijo("RSG", S.datos.Riesgos, "ID_Riesgo", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }); }, "Riesgo agregado");
    }, null, r ? { init, eliminar: { texto: "Eliminar riesgo", mensaje: `Se borrará el riesgo «${r.Descripcion}».`, accion: () => guardar(() => S.api.eliminarFilas("Riesgos", "ID_Riesgo", [r.ID_Riesgo]), "Riesgo eliminado") } } : { init });
  }

  function formUsuario(uEdit) {
    const campos = [
      { k: "Correo", label: "Correo corporativo", tipo: "email", bloqueado: !!uEdit, placeholder: "nombre@empresa.com" },
      { k: "Nombre", label: "Nombre", placeholder: "Como aparecerá en la app" },
      { k: "Rol", label: "Roles", tipo: "checks", opciones: R.ROLES, ayuda: "Puedes marcar varios; los permisos se suman." },
      { k: "Proyectos", label: "Proyectos que puede consultar (rol Lector)", def: "Todos", ancho: true,
        ayuda: "IDs separados por coma (PRY-0001, PRY-0004) o Todos. Los proyectos de un PM salen de la tabla Proyectos." },
      { k: "Activo", label: "Activo", tipo: "select", opciones: ["Sí", "No"], def: "Sí" },
    ];
    modal(uEdit ? "Editar usuario" : "Nuevo usuario", campos, uEdit || {}, (fd) => {
      if (uEdit) { fd.Correo = uEdit.Correo; guardar(() => S.api.actualizarPorId("Usuarios", "Correo", uEdit.Correo, fd)); }
      else guardar(() => S.api.agregarFila("Usuarios", fd), "Usuario creado");
    }, (fd) => {
      if (!uEdit && !fd.Correo) return "Escribe el correo: es con lo que el usuario entra a la app.";
      if (!uEdit && S.datos.Usuarios.some((x) => lc(x.Correo) === lc(fd.Correo))) return "Ese correo ya existe.";
      if (uEdit && lc(uEdit.Correo) === lc(S.usuario.Correo) && (!String(fd.Rol).split(", ").includes("Admin") || fd.Activo !== "Sí")) return "No puedes quitarte el rol Admin ni desactivarte a ti mismo.";
      const ids = new Set(S.datos.Proyectos.map((p) => p.ID_Proyecto));
      const malos = String(fd.Proyectos || "").split(",").map((s) => s.trim()).filter((s) => s && s !== "Todos" && !ids.has(s));
      if (malos.length) return `Estos proyectos no existen: ${malos.join(", ")}.`;
      return "";
    }, uEdit && lc(uEdit.Correo) !== lc(S.usuario.Correo) ? { eliminar: { texto: "Eliminar usuario",
      mensaje: `${uEdit.Nombre || uEdit.Correo} ya no podrá entrar. Sus proyectos y registros se conservan. Si es temporal, mejor márcalo como inactivo.`,
      accion: () => guardar(() => S.api.eliminarFilas("Usuarios", "Correo", [uEdit.Correo]), "Usuario eliminado") } } : {});
  }

  // ---------- Editar y eliminar registros ----------
  async function eliminarProyecto(pid) {
    const segs = S.datos.Seguimientos.filter((s) => s.ID_Proyecto === pid);
    for (const s of segs.filter((x) => x.Acta_Archivo)) await S.api.borrarArchivo(s.ID_Seguimiento);
    const borrar = (tabla, col, filas) => (filas.length ? S.api.eliminarFilas(tabla, col, filas.map((f) => f[col])) : null);
    await borrarAdjuntos((S.datos.Comentarios || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Comentarios", "ID_Comentario", (S.datos.Comentarios || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Compromisos", "ID_Compromiso", S.datos.Compromisos.filter((c) => c.ID_Proyecto === pid));
    await borrar("Seguimientos", "ID_Seguimiento", segs);
    await borrar("Hitos", "ID_Hito", S.datos.Hitos.filter((h) => h.ID_Proyecto === pid));
    await borrar("Riesgos", "ID_Riesgo", S.datos.Riesgos.filter((r) => r.ID_Proyecto === pid));
    await borrar("Stakeholders", "ID_Stakeholder", (S.datos.Stakeholders || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Tickets", "ID_Ticket", (S.datos.Tickets || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Cambios", "ID_Cambio", (S.datos.Cambios || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Lecciones", "ID_Leccion", (S.datos.Lecciones || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("RACI", "ID_RACI", (S.datos.RACI || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Dependencias", "ID_Dependencia", (S.datos.Dependencias || []).filter((x) => x.ID_Proyecto === pid || x.Depende_De === pid));
    await S.api.eliminarFilas("Proyectos", "ID_Proyecto", [pid]);
  }

  function formCompromiso(c, pNuevo, segSugerida) {
    const nuevo = !c;
    const pid = nuevo ? pNuevo.ID_Proyecto : c.ID_Proyecto;
    const sesiones = R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse()
      .map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)}${s.Fecha_Acta && s.Fecha_Acta !== s.Fecha_Corte ? ` (acta ${fecha(s.Fecha_Acta)})` : ""}`]);
    const gente = [...recursos().filter((r) => r.Activo !== "No").map((r) => [r.Nombre, r.Correo || ""]), ...S.datos.Usuarios.filter((u) => u.Activo === "Sí").map((u) => [u.Nombre || u.Correo, u.Correo]),
      ...(S.datos.Stakeholders || []).filter((x) => x.ID_Proyecto === pid && x.Nombre).map((x) => [x.Nombre, x.Correo || ""])];
    const personas = [...new Set(gente.map((g) => g[0]))].sort((a, b) => a.localeCompare(b, "es"));
    const correoPorNombre = (n) => (gente.find((g) => lc(g[0]) === lc(n) && g[1]) || [])[1] || "";
    const autollenar = { init: (m) => {
      const r = m.querySelector('[name="Responsable"]'), co = m.querySelector('[name="Correo_Responsable"]');
      r.addEventListener("change", () => { const c2 = correoPorNombre(r.value); if (c2 && !co.value) co.value = c2; });
    } };
    const campos = [
      { k: "Compromiso", label: "Compromiso", tipo: "textarea", placeholder: "Qué se hará" },
      { k: "ID_Seguimiento", label: "¿De qué sesión salió?", tipo: "select", opciones: sesiones, ancho: true, textoVacio: "Sin sesión (solo del proyecto)",
        ayuda: sesiones.length ? "Elige la sesión de seguimiento donde se acordó, o déjalo sin sesión." : "Este proyecto aún no tiene sesiones; quedará atado solo al proyecto." },
      { k: "Responsable", label: "Responsable", recurso: true, placeholder: "Busca en Recursos o escribe un nombre nuevo" },
      { k: "Correo_Responsable", label: "Correo del responsable", tipo: "email", placeholder: "Se llena solo si la persona está registrada", ayuda: "Para enviarle recordatorios." },
      { k: "Fecha_Compromiso", label: "Fecha límite", tipo: "date" },
      { k: "Dias_Alerta", label: "Avisar con (días de anticipación)", tipo: "number", min: 0, max: 60, def: R.DIAS_COMPROMISO_POR_VENCER, ayuda: "Desde ese día aparece en «Por vencer» y en la campana 🔔." },
      ...(nuevo ? [{ k: "Comentario", label: "Comentario inicial (opcional)", tipo: "textarea", placeholder: "Contexto o primer avance" }] : [
        { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Compromiso, def: "Pendiente" },
        { k: "Fecha_Cierre", label: "Fecha de cierre", tipo: "date", ayuda: "Se llena sola al marcarlo cumplido; se borra si lo reabres." }]),
    ];
    if (nuevo) {
      modal(`Nuevo compromiso · ${pNuevo.Nombre}`, campos, { ID_Seguimiento: segSugerida || "" }, (fd) => {
        const fila = { ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", pNuevo.ID_Proyecto), ID_Proyecto: pNuevo.ID_Proyecto, ID_Seguimiento: fd.ID_Seguimiento || "",
          Compromiso: fd.Compromiso || "(sin descripción)", Responsable: fd.Responsable, Correo_Responsable: fd.Correo_Responsable || correoPorNombre(fd.Responsable), Dias_Alerta: fd.Dias_Alerta === "" ? "" : fd.Dias_Alerta, Fecha_Compromiso: fd.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo };
        guardar(async () => {
          if (fila.Correo_Responsable) await asegurarRecurso({ Nombre: fila.Responsable, Correo: fila.Correo_Responsable });
          await S.api.agregarFila("Compromisos", fila);
          if (fd.Comentario) await S.api.agregarFila("Comentarios", nuevoComentario(fila, fd.Comentario));
        }, "Compromiso agregado");
      }, null, autollenar);
      return;
    }
    const coms = comentariosDe(c.ID_Compromiso);
    modal("Editar compromiso", campos, { ...c, Estado: R.estadoBase(c) }, (fd) => {
      if (fd.Estado === "Cerrado" && !fd.Fecha_Cierre) fd.Fecha_Cierre = R.hoyISO();
      if (fd.Estado !== "Cerrado") fd.Fecha_Cierre = "";
      const cambioEstado = fd.Estado !== R.estadoBase(c);
      guardar(async () => {
        if (fd.Correo_Responsable && fd.Responsable) await asegurarRecurso({ Nombre: fd.Responsable, Correo: fd.Correo_Responsable });
        await S.api.actualizarPorId("Compromisos", "ID_Compromiso", c.ID_Compromiso, fd);
        if (cambioEstado) await S.api.agregarFila("Comentarios", nuevoComentario(c, TEXTO_ESTADO[fd.Estado] || `Estado: ${fd.Estado}.`));
      }, "Compromiso actualizado");
    }, null, { ...autollenar, eliminar: { texto: "Eliminar compromiso", mensaje: `Se borrará «${c.Compromiso}»${coms.length ? ` con sus ${coms.length} comentario(s)` : ""}.`,
      accion: () => guardar(async () => {
        await borrarAdjuntos(coms);
        if (coms.length) await S.api.eliminarFilas("Comentarios", "ID_Comentario", coms.map((x) => x.ID_Comentario));
        await S.api.eliminarFilas("Compromisos", "ID_Compromiso", [c.ID_Compromiso]);
      }, "Compromiso eliminado") } });
  }

  function formEditarSeguimiento(s) {
    const p = proyecto(s.ID_Proyecto);
    const segs = R.seguimientosDe(s.ID_Proyecto, S.datos.Seguimientos);
    const esUltimo = segs.length && segs[segs.length - 1].ID_Seguimiento === s.ID_Seguimiento;
    const comps = S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento);
    const campos = [
      { seccion: "1. Avance" },
      { k: "Fecha_Corte", label: "Fecha de corte", tipo: "date" },
      { k: "Avance_Real", label: "Avance real (%)", tipo: "number", min: 0, max: 100 },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo },
      { k: "Ejecutado", label: "Presupuesto ejecutado a la fecha (COP)", tipo: "number", min: 0 },
      { seccion: "2. Qué pasó en el periodo" },
      { k: "Logros", label: "Logros del periodo", tipo: "textarea" },
      { k: "Proximos_Pasos", label: "Próximos pasos", tipo: "textarea" },
      { k: "Bloqueos", label: "Bloqueos", tipo: "textarea" },
      { seccion: "3. Acta de la sesión" },
      { k: "Fecha_Acta", label: "Fecha del acta", tipo: "date" },
      { k: "URL_Acta", label: "Enlace al acta (opcional)", tipo: "url", placeholder: "https://…" },
    ];
    const bloqueActa = `<div class="carga-acta">
      <div><b>Archivo del acta</b><div class="sub" id="acta-actual">${s.Acta_Archivo ? `Guardada: ${esc(s.Acta_Archivo)}` : "Esta sesión no tiene acta guardada."}</div></div>
      <div class="acciones">${s.Acta_Archivo ? `<button type="button" class="btn" id="b-ver-acta">Ver</button><button type="button" class="btn" id="b-quitar-acta">Quitar</button>` : ""}
        <button type="button" class="btn primario" id="b-reemplazar-acta">${s.Acta_Archivo ? "Reemplazar PDF" : "Subir PDF"}</button></div>
      <div id="acta-estado" class="acta-estado" aria-live="polite"></div></div>`;
    const despues = `<p class="sub">${comps.length ? `Esta sesión tiene ${comps.length} compromiso(s); edítalos en la tabla «Compromisos» de la ficha.` : "Esta sesión no tiene compromisos."}${esUltimo ? " Es la última sesión: al guardar también se actualiza el avance y el semáforo del proyecto." : ""}</p>`;
    let accionActa = null;   // null = sin cambio · { tipo: "subir", bytes, nombre } · { tipo: "quitar" }
    const init = (m) => {
      const est = $("#acta-estado");
      if ($("#b-ver-acta")) $("#b-ver-acta").addEventListener("click", () => { const volver = () => formEditarSeguimiento(s); verActa(s.ID_Seguimiento, s.Acta_Archivo).then(() => { const c = $("#v-cerrar"); if (c) c.addEventListener("click", volver, { once: true }); }); });
      if ($("#b-quitar-acta")) $("#b-quitar-acta").addEventListener("click", () => { accionActa = { tipo: "quitar" }; est.className = "acta-estado error"; est.textContent = "El acta se quitará al guardar."; });
      $("#b-reemplazar-acta").addEventListener("click", () => elegirPDF(async (archivo) => {
        if (archivo.size > MAX_ACTA) { est.className = "acta-estado error"; est.textContent = "El acta pesa más de 5 MB."; return; }
        est.className = "acta-estado"; est.textContent = "Leyendo el acta…";
        try {
          const res = await ACTAS.leerActa(archivo);
          accionActa = { tipo: "subir", bytes: res.datos, nombre: archivo.name };
          if (res.fecha) m.querySelector("[name=Fecha_Acta]").value = res.fecha;
          est.className = "acta-estado ok";
          est.textContent = `${archivo.name} se guardará al dar Guardar.${res.compromisos.length && !comps.length ? ` Trae ${res.compromisos.length} compromiso(s); se agregarán a la sesión.` : ""}`;
          accionActa.compromisos = !comps.length ? res.compromisos : [];
        } catch (e) { est.className = "acta-estado error"; est.textContent = "No se pudo leer el PDF: " + e.message; }
      }));
    };
    modal(`Editar sesión del ${fecha(s.Fecha_Corte)} · ${p ? p.Nombre : ""}`, campos, s, (fd) => {
      if (!fd.Fecha_Corte) fd.Fecha_Corte = s.Fecha_Corte || R.hoyISO();
      fd.Semana = R.semanaISO(fd.Fecha_Corte);
      const accion = accionActa;
      guardar(async () => {
        const cambios = { ...fd };
        if (accion && accion.tipo === "quitar") { await S.api.borrarArchivo(s.ID_Seguimiento); cambios.Acta_Archivo = ""; }
        if (accion && accion.tipo === "subir") { await subirActa(s.ID_Seguimiento, accion.bytes, accion.nombre); cambios.Acta_Archivo = accion.nombre; }
        await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", s.ID_Seguimiento, cambios);
        if (accion && accion.compromisos && accion.compromisos.length) await S.api.agregarFilas("Compromisos", accion.compromisos.map((c, i) => ({
          ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", s.ID_Proyecto, i), ID_Proyecto: s.ID_Proyecto, ID_Seguimiento: s.ID_Seguimiento,
          Compromiso: c.Compromiso, Responsable: c.Responsable, Fecha_Compromiso: c.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo })));
        if (esUltimo && p) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}), ...(fd.Semaforo ? { Semaforo: fd.Semaforo } : {}), ...(fd.Ejecutado !== "" ? { Ejecutado: fd.Ejecutado } : {}), ...sello() });
      }, "Sesión actualizada");
    }, (fd) => (fd.Fecha_Corte && fd.Fecha_Corte > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""), {
      antes: bloqueActa, despues, init,
      eliminar: { texto: "Eliminar sesión", titulo: `Eliminar la sesión del ${fecha(s.Fecha_Corte)}`,
        mensaje: `Se borrará la sesión${comps.length ? `, sus ${comps.length} compromiso(s)` : ""}${s.Acta_Archivo ? " y el acta guardada" : ""}. No se puede deshacer.`,
        accion: () => guardar(async () => {
          if (s.Acta_Archivo) await S.api.borrarArchivo(s.ID_Seguimiento);
          const idsC = new Set(comps.map((c) => c.ID_Compromiso));
          const comsS = (S.datos.Comentarios || []).filter((x) => idsC.has(x.ID_Compromiso));
          await borrarAdjuntos(comsS);
          if (comsS.length) await S.api.eliminarFilas("Comentarios", "ID_Comentario", comsS.map((x) => x.ID_Comentario));
          if (comps.length) await S.api.eliminarFilas("Compromisos", "ID_Compromiso", comps.map((c) => c.ID_Compromiso));
          await S.api.eliminarFilas("Seguimientos", "ID_Seguimiento", [s.ID_Seguimiento]);
        }, "Sesión eliminada") },
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
      S.api = conAuditoria(await DATOS.crear(S.modo));
      await recargar();
      const correo = leerSesion();
      S.usuario = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(correo) && x.Activo === "Sí") || null;
      filtrosIniciales();
      render();
    } catch (e) {
      app().innerHTML = `<div class="login"><div class="login-caja"><h2>No se pudo leer el Excel</h2><p>${esc(e.message)}</p>
        <p class="sub">Verifica que el archivo abierto sea Portafolio.xlsx y que tenga las tablas Proyectos, Hitos, Seguimientos, Riesgos, Usuarios y Catalogos.</p></div></div>`;
    } finally {
      cargando(false);
    }
  }

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(() => iniciar());
  else document.addEventListener("DOMContentLoaded", iniciar);
})();
