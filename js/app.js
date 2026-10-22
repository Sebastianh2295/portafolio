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
    { id: "agenda", t: "Agenda semanal" },
    { id: "backlog", t: "Backlog de demanda" },
    { id: "proyectos", t: "Proyectos" },
    { id: "cronograma", t: "Cronograma" },
    { id: "riesgos", t: "Riesgos" },
    { id: "recursos", t: "Recursos" },
    { id: "capacidad", t: "Capacidad del equipo" },
    { id: "priorizacion", t: "Priorización" },
    { id: "valor", t: "Valor del portafolio" },
    { id: "lecciones", t: "Lecciones aprendidas" },
    { id: "sugerencias", t: "Sugerencias de mejora" },
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
    capacidad: ["Cuánto está dedicada cada persona sumando sus proyectos activos o en pausa y sus otras actividades (operación, comités, demanda, ausencias…).",
      "La dedicación del PM se pone en «Editar proyecto»; la de cada stakeholder, en su ficha o en la sección Equipo del formulario.",
      "Las otras actividades se registran abajo en «Otras actividades»; las ausencias bajan la capacidad mientras duran.",
      "Más de 100% = sobreasignado. Cambia la fecha de corte para ver la carga de otro momento."],
    backlog: ["Iniciativas que asigna Gestión de la Demanda en Almera, antes de ser proyectos. El código Almera es la llave.",
      "Flujo: Recibida → En análisis → Priorizada → En comité → (Aprobada / Aplazada / Rechazada) → Convertida en proyecto.",
      "Ranking: obligatorias primero y luego por puntaje. En el tablero arrastras las tarjetas entre estados; la decisión la registra el comité."],
    recursos: ["Directorio único de personas: internas y de proveedores, sin duplicados.",
      "Al agregar stakeholders, equipo o responsables de compromisos eliges a la persona de aquí y sus datos se llenan solos.",
      "Si aparecen «Posibles duplicados», usa «Fusionar» para dejar un solo registro."],
    agenda: ["Las sesiones de seguimiento de todos los proyectos en la semana: daily, seguimientos, comités y mesas técnicas, por día y hora.",
      "Se programan en la pestaña «Seguimientos» de cada proyecto (Agenda de seguimiento). Rojo = cruce de horario del mismo PM o sesión sin registro.",
      "Clic en una sesión pasada o de hoy para registrarla con la fecha puesta. Los daily no exigen registro."],
    valor: ["Qué valor le deja cada proyecto a la FSFB: no solo financiero, también seguridad del paciente, eficiencia operacional, experiencia, cumplimiento y lo estratégico.",
      "Cada proyecto tiene un nivel de valor (crítico, alto, medio, bajo) y beneficios con indicador, línea base y meta; el PM registra las mediciones en la pestaña «Valor» del proyecto.",
      "Los proyectos cerrados siguen aquí mientras sus beneficios se sigan midiendo."],
    priorizacion: ["Ordena el portafolio por un puntaje de 0 a 100 según valor, urgencia, complejidad y esfuerzo.",
      "La calificación de cada proyecto se hace en «Editar proyecto» → Priorización.",
      "La matriz valor vs. esfuerzo ayuda a ver ganancias rápidas y proyectos a reconsiderar."],
    sugerencias: ["Aquí quedan todas las mejoras, errores e ideas que se reportan sobre la app, con su estado.",
      "Desde cualquier pantalla usa «💡 Sugerir mejora» (arriba a la derecha): la app guarda sola en qué módulo y proyecto estabas.",
      "Admin y PMO cambian el estado y responden. Marca varias y usa «Copiar para Claude» para pedir los cambios de una vez."],
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
  // Fecha legible: acepta ISO, ISO con hora y el número de serie de Excel (por si una columna llegó sin convertir).
  const fecha = (iso) => {
    let s = typeof iso === "number" && iso > 20000 && iso < 90000 ? new Date(Math.round((iso - 25569) * 86400000)).toISOString().slice(0, 10) : String(iso ?? "");
    if (/^\d{5}$/.test(s) && Number(s) > 20000) s = new Date(Math.round((Number(s) - 25569) * 86400000)).toISOString().slice(0, 10);
    if (/^\d{4}-\d{2}-\d{2}[T ]/.test(s)) s = s.slice(0, 10);
    return /^\d{4}-\d{2}-\d{2}$/.test(s) ? `${s.slice(8, 10)}/${s.slice(5, 7)}/${s.slice(0, 4)}` : s || "—";
  };
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
  // Aviso de espera. Si Excel tarda más de lo normal, se ofrece cerrar el aviso para no quedar bloqueado.
  let relojCarga = null;
  function cargando(on, msg = "Cargando…") {
    let o = $("#cargando");
    clearTimeout(relojCarga);
    if (on) {
      if (!o) { o = document.createElement("div"); o.id = "cargando"; document.body.appendChild(o); }
      o.innerHTML = `<div class="spinner"></div><div>${esc(msg)}</div><div class="cargando-lento" hidden>Excel está tardando más de lo normal. Puedes seguir esperando o cerrar este aviso; si el cambio no aparece, usa «Actualizar datos del Excel».<div class="cargando-op"></div><button class="btn chico" id="cg-cerrar">Cerrar aviso</button> <button class="btn chico primario" id="cg-reconectar">Reconectar con Excel</button></div>`;
      const t0 = Date.now();
      relojCarga = setTimeout(function mostrarLento() {
        const l = o.querySelector(".cargando-lento");
        if (!l || !o.isConnected) return;
        if (l.hidden) { l.hidden = false; o.querySelector("#cg-cerrar").addEventListener("click", () => o.remove()); o.querySelector("#cg-reconectar").addEventListener("click", () => { o.remove(); reconectar(); }); }
        const det = o.querySelector(".cargando-op");
        if (det) det.textContent = `${opsEnCurso()} · ${Math.round((Date.now() - t0) / 1000)} s`;
        relojCarga = setTimeout(mostrarLento, 1000);
      }, 12000);
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
    Tipo_Documento: ["Acta", "Cronograma", "Contrato / orden de compra", "Especificación / requerimientos", "Diseño / arquitectura", "Presentación", "Informe", "Manual", "Carpeta", "Otro"],
    Categoria_Leccion: ["Planeación", "Técnica", "Proveedores", "Comunicación", "Equipo", "Calidad", "Gestión del cambio"],
    Tipo_Sesion_Agenda: ["Daily", "Seguimiento", "Comité", "Mesa técnica"],
    Dimension_Valor: ["Seguridad del paciente", "Eficiencia operacional", "Experiencia del paciente", "Cumplimiento / acreditación", "Estratégico", "Financiero"],
    Tipo_Actividad: ["Operación / soporte", "Comités y reuniones recurrentes", "Gestión de demanda", "Administrativo / informes", "Capacitación", "Ausencia (vacaciones, incapacidad, permiso)", "Otro"] };
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
    await vincularCompromisosRecursos();
    await completarFechaCreacion();
    await migrarLineasBase();
    await migrarCronogramas();
    const faltan = Object.keys(CATALOGOS_BASE).filter((k) => !S.datos.Catalogos.some((f) => f.Lista === k));
    if (sembrado || !faltan.length) return;
    sembrado = true;
    try { await S.api.agregarFilas("Catalogos", faltan.flatMap((k) => CATALOGOS_BASE[k].map((v) => ({ Lista: k, Valor: v })))); armarCatalogos(); }
    catch (e) { /* sin permiso de edición: se usan los valores base en memoria */ }
  }
  function refrescarUsuario() {
    if (S.usuario) S.usuario = S.datos.Usuarios.find((x) => lc(x.Correo) === lc(S.usuario.Correo) && x.Activo === "Sí") || null;
  }
  async function recargar() {
    S.datos = await S.api.leerTodo();
    S.cargadoEl = Date.now();
    armarCatalogos();
    await sembrarCatalogos();
    refrescarUsuario();
  }
  // ---------- Reconectar con Excel (sin recargar el complemento) ----------
  async function reconectar() {
    cerrarModal();
    cargando(true, "Reconectando con Excel…");
    try {
      if (S.api.reiniciar) S.api.reiniciar();   // suelta lo que haya quedado esperando
      EN_CURSO.clear();
      const r = S.api.ping ? await S.api.ping() : null;
      await recargar();
      S.conexion = "ok";
      render();
      toast(`Conectado con Excel${r && r.libro ? ` · ${r.libro}` : ""}. Datos actualizados.`);
    } catch (e) {
      S.conexion = "error";
      render();
      ayudaReconexion(e.message);
    } finally { cargando(false); }
  }
  function ayudaReconexion(motivo) {
    const viejo = /versi[oó]n anterior/i.test(motivo || "");
    const pasos = S.modo === "ventana"
      ? `<ol><li>En Excel, busca el panel lateral <b>Portafolio PMO</b> (si no lo ves: <b>Inicio → Complementos → Portafolio PMO</b>).</li>
          <li>${viejo ? "Ciérralo con la ✕ y ábrelo de nuevo para que cargue la versión nueva." : "Si el panel muestra «Conectando…» o se ve en blanco, ciérralo con la ✕ y ábrelo de nuevo."}</li>
          <li>En el panel, haz clic en <b>Abrir Portafolio</b>: se abre una ventana nueva ya conectada. Esta ventana la puedes cerrar.</li></ol>`
      : `<ol><li>Cierra este panel con la ✕ y ábrelo de nuevo desde <b>Inicio → Complementos → Portafolio PMO</b>.</li></ol>`;
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="Reconectar con Excel">
      <div class="titulo-fila"><h2>No pude conectarme con Excel</h2><button class="btn enlace" id="rx-cerrar">✕</button></div>
      <p>${viejo ? "El panel lateral de Excel tiene una versión anterior de la app." : "El panel lateral de Excel no respondió. Puede que se haya cerrado, recargado o que Excel esté ocupado."}</p>
      ${pasos}<p class="sub">Tus datos están guardados en el Excel; no se pierde nada. Detalle: ${esc(motivo || "")}</p>
      <div class="acciones derecha"><button class="btn" id="rx-no">Cerrar</button><button class="btn primario" id="rx-otra">Intentar de nuevo</button></div></div>`;
    document.body.appendChild(m);
    $("#rx-cerrar").addEventListener("click", cerrarModal); $("#rx-no").addEventListener("click", cerrarModal);
    $("#rx-otra").addEventListener("click", reconectar);
  }
  // Nota: no se relee el Excel en segundo plano (en la ventana grande competía con los guardados). Para ver cambios de otros: menú «Actualizar datos del Excel».
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
      // Los cambios ya quedaron en memoria (ver conAuditoria): no se relee todo el Excel.
      armarCatalogos();
      refrescarUsuario();
      cerrarModal();
      render();
      toast(exito);
    } catch (e) {
      toast("No se pudo guardar: " + e.message, true);
      if (/no respondi|conexi|panel lateral|tard/i.test(e.message)) { S.conexion = "error"; const b = $("#b-conexion"); if (b) b.classList.add("desconectado"); }
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
    ED.clear();
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
    const nSug = gestionaSug() ? (S.datos.Sugerencias || []).filter((x) => (x.Estado || "Nueva") === "Nueva").length : 0;
    const itemNav = (v) => `<button class="nav ${v.id === activa ? "activa" : ""}" data-vista="${v.id}" title="${esc(v.t)}">${ico(v.id)}<span>${esc(v.t)}</span>${v.id === "sugerencias" && nSug ? `<span class="nav-badge" title="${nSug} sugerencia(s) nueva(s)">${nSug}</span>` : ""}</button>`;
    app().innerHTML = `
      <div class="shell ${S.navMini ? "mini" : ""}">
        <header class="appbar">
          <button class="ab-btn" id="b-menu" aria-label="Mostrar u ocultar el menú">${ico("menu")}</button>
          <div class="ab-marca"><span class="logo">P</span><span class="ab-nombre">Portafolio PMO</span><span class="ab-org">FSFB</span></div>
          <button class="ab-buscar" id="b-buscar" title="Búsqueda global (Ctrl+K)">${ico("buscar")}<span>Buscar proyectos, personas, riesgos…</span><kbd>Ctrl K</kbd></button>
          <div class="ab-der">
            <button class="modo ${S.modo} ${S.conexion === "error" ? "desconectado" : ""}" id="b-conexion" title="${S.modo === "demo" ? esc(etiquetaModo) : S.conexion === "error" ? "Sin conexión con Excel · clic para reconectar" : "Conectado con Excel · clic para reconectar si algo se queda pegado"}"><span class="punto-con"></span>${S.modo === "demo" ? "Demo" : "Excel"}</button>
            <button class="ab-btn campana ${al.vencidos.length ? "roja" : nAl ? "ambar" : ""}" id="b-alertas" title="Compromisos vencidos o por vencer">${ico("campana")}${nAl ? `<span class="ab-badge">${nAl}</span>` : ""}</button>
            <button class="ab-btn" id="b-ayuda" title="Mostrar la ayuda de esta pantalla">${ico("ayuda")}</button>
            <div class="ab-usuario"><button class="ab-avatar" id="b-usuario" aria-haspopup="true" aria-expanded="false" title="${esc(u.Nombre || u.Correo)}">${avatar(u.Nombre || u.Correo)}</button>
              <div class="menu-usuario" id="menu-usuario" hidden>
                <div class="mu-cab">${avatar(u.Nombre || u.Correo, "grande")}<div><b>${esc(u.Nombre || u.Correo)}</b><div class="sub">${esc(u.Correo)}</div><div class="sub">${esc(R.roles(u).join(", "))}</div></div></div>
                <button class="mu-item" id="b-recargar">${ico("actualizar")} Actualizar datos del Excel</button>
                <button class="mu-item" id="b-reconectar">${ico("actualizar")} Reconectar con Excel</button>
                <button class="mu-item" id="b-diag">${ico("ayuda")} Diagnóstico de conexión</button>
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
        <main class="contenido"><div class="barra-migas">${migas()}<button class="btn chico btn-sugerir" id="b-sugerir" title="Proponer un cambio o reportar un error de esta pantalla">💡 <span>Sugerir mejora</span></button>${S.datos && S.vista !== "ficha" ? `<button class="btn chico btn-ppt" id="b-ppt" title="Informe de estado general o por PM en PowerPoint, con la marca FSFB">📊 <span>Informe PPT</span></button>` : ""}${EXPORTADORES[S.vista] && S.datos ? `<button class="btn chico btn-exportar" id="b-exportar" title="Descarga en Excel lo que ves, con los filtros aplicados">${ico("exportar")}<span>${S.vista === "ficha" ? (S.fichaTab === "riesgos" ? "Exportar riesgos (formato FSFB)" : "Exportar proyecto a Excel") : "Exportar a Excel"}</span></button>` : ""}</div><section id="vista"></section></main>
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
    $("#b-sugerir").addEventListener("click", () => formSugerencia());
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
    $("#b-conexion").addEventListener("click", reconectar);
    $("#b-reconectar").addEventListener("click", reconectar);
    if ($("#b-ppt")) $("#b-ppt").addEventListener("click", formInformePPT);
    $("#b-diag").addEventListener("click", () => { $("#menu-usuario").hidden = true; verDiagnostico(); });
    $("#b-recargar").addEventListener("click", async () => {
      cargando(true, "Leyendo Excel…");
      try { await recargar(); render(); toast("Datos actualizados"); } catch (e) { toast(e.message, true); } finally { cargando(false); }
    });
    const vistas = { dashboard: vDashboard, avances: vAvances, seguimiento: vSeguimiento, proyectos: vProyectos, ficha: vFicha, backlog: vBacklog, recursos: vRecursos, capacidad: vCapacidad, priorizacion: vPriorizacion, valor: vValor, agenda: vAgenda, lecciones: vLecciones, sugerencias: vSugerencias, auditoria: vAuditoria, cronograma: vCronograma, riesgos: vRiesgos, catalogos: vCatalogos, usuarios: vUsuarios };
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
    animarVista(el);
  }
  // Entrada suave de la pantalla y contadores que suben en los indicadores (se respeta «reducir movimiento»).
  const sinMovimiento = () => window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  let ultimaVista = "";
  function animarVista(el) {
    const clave = `${S.vista}|${S.pid || ""}|${S.fichaTab || ""}`;
    const cambio = clave !== ultimaVista;
    ultimaVista = clave;
    if (sinMovimiento() || !cambio) return;   // al guardar en la misma pantalla no se repite la animación
    el.classList.remove("entrando"); void el.offsetWidth; el.classList.add("entrando");
    el.querySelectorAll(".kpi-v, [data-contar]").forEach((n) => {
      const txt = n.textContent.trim(), m = txt.match(/^(-?\d+(?:[.,]\d+)?)(.*)$/);
      if (!m || n.children.length > 1) return;
      const fin = parseFloat(m[1].replace(",", ".")), suf = m[2], dec = /[.,]/.test(m[1]) ? 1 : 0;
      if (!Number.isFinite(fin) || fin === 0) return;
      const t0 = performance.now(), dur = 700;
      const paso = (t) => { const k = Math.min(1, (t - t0) / dur), v = fin * (1 - Math.pow(1 - k, 3)); n.textContent = (dec ? v.toFixed(dec) : Math.round(v)) + suf; if (k < 1) requestAnimationFrame(paso); else n.textContent = txt; };
      requestAnimationFrame(paso);
    });
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

  // Etiquetas de color para semáforo, estado del proyecto y fase.
  const SEM_CLASE = { Verde: "verde", Amarillo: "amarillo", Rojo: "rojo" };
  const chipSemaforo = (s) => (s ? `<span class="sem sem-${SEM_CLASE[s] || "otro"}"><span class="dot" style="background:${COLORES[s] || "#999"}" aria-hidden="true"></span>${esc(s)}</span>` : `<span class="sub">—</span>`);
  const ESTADO_P = { Activo: ["ep-activo", "●"], "En pausa": ["ep-pausa", "❚❚"], Cerrado: ["ep-cerrado", "✓"], Cancelado: ["ep-cancelado", "✕"] };
  const chipEstadoP = (e) => { if (!e) return `<span class="sub">—</span>`; const [c, i] = ESTADO_P[e] || ["ep-otro", "•"]; return `<span class="ep ${c}">${e === "Activo" ? `<span class="ep-dot" aria-hidden="true"></span>` : `<i aria-hidden="true">${i}</i>`}${esc(e)}</span>`; };
  const FASE_CLASE = { Inicio: "fa-inicio", "Planeación": "fa-plan", "Ejecución": "fa-ejec", Cierre: "fa-cierre" };
  const chipFase = (f) => {
    if (!f) return `<span class="sub">—</span>`;
    const fs = S.cat.Fase || [], i = fs.indexOf(f);
    const pasos = i >= 0 ? `<span class="fase-pasos" aria-hidden="true">${fs.map((_, k) => `<i class="${k <= i ? "on" : ""}"></i>`).join("")}</span>` : "";
    return `<span class="fase ${FASE_CLASE[f] || "fa-otra"}" title="${i >= 0 ? `Fase ${i + 1} de ${fs.length}` : ""}">${pasos}${esc(f)}</span>`;
  };
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
        const segs = R.avancesDe(p.ID_Proyecto, S.datos.Seguimientos).filter((s) => s.Fecha_Corte && s.Fecha_Corte <= corte && s.Avance_Real !== "");
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

    const hora = new Date().getHours();
    const saludo = hora < 12 ? "Buenos días" : hora < 19 ? "Buenas tardes" : "Buenas noches";
    const nombreCorto = String(S.usuario.Nombre || S.usuario.Correo || "").split(/[\s@]/)[0];
    const al = alertas();
    el.innerHTML = `
      <section class="hero">
        <div class="hero-txt"><div class="hero-fecha">${esc((() => { const f = new Date().toLocaleDateString("es-CO", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); return f.charAt(0).toUpperCase() + f.slice(1); })())}</div>
          <h1>${saludo}${nombreCorto ? `, ${esc(nombreCorto.charAt(0).toUpperCase() + nombreCorto.slice(1))}` : ""}</h1>
          <p>Este es el estado del portafolio: ${k.activos} proyecto(s) activos${k.rojos ? `, <b>${k.rojos} en rojo</b>` : ""}${al.vencidos.length ? ` y <b>${al.vencidos.length} compromiso(s) vencido(s)</b>` : ""}.</p></div>
        <div class="hero-anillo" style="--p:${Math.max(0, Math.min(100, Math.round(k.avanceReal)))}"><div><span class="hero-num"><b data-contar="${Math.round(k.avanceReal)}">${Math.round(k.avanceReal)}</b>%</span><small>avance promedio</small></div></div>
      </section>
      ${barraFiltros()}
      <div class="kpis">
        ${tarjeta("Proyectos activos", k.activos, `de ${k.total} en la selección`, false, "activos")}
        ${tarjeta("Avance real promedio", pct(k.avanceReal), `plan ${pct(k.avancePlan)} · ${desv > 0 ? "+" : ""}${desv} pts`, desv < -5, "activos")}
        ${tarjeta("Proyectos en rojo", k.rojos, `${Math.round((k.rojos / Math.max(1, k.activos)) * 100)}% de los activos`, k.rojos > 0, "rojos")}
        ${tarjeta("Presupuesto ejecutado", pct(k.pctEjecutado), `${cop(k.ejecutado)} de ${cop(k.presupuesto)}`, k.pctEjecutado > 100, "activos")}
        ${tarjeta("Seguimientos al día", pct(k.pctAlDia), "activos sin reporte vencido", k.pctAlDia < 80, "seguimiento")}
        ${(() => { const vs = activos.map(evmDe).filter(Boolean), sum = (f) => vs.reduce((a, v) => a + (v[f] || 0), 0);
          const pv = sum("PV"), ev = sum("EV"), ac = vs.filter((v) => v.AC > 0 && v.EV > 0).reduce((a, v) => a + v.AC, 0), evc = vs.filter((v) => v.AC > 0 && v.EV > 0).reduce((a, v) => a + v.EV, 0);
          const spi = pv > 0 ? ev / pv : null, cpi = ac > 0 ? evc / ac : null, f = (x) => (x === null ? "—" : x.toFixed(2).replace(".", ","));
          return tarjeta("Desempeño (valor ganado)", `${f(spi)} · ${f(cpi)}`, "SPI cronograma · CPI costo (1 = según plan)", (spi !== null && spi < 0.9) || (cpi !== null && cpi < 0.9), "activos"); })()}
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
    const segs = S.datos.Seguimientos.filter((s) => ids.has(s.ID_Proyecto) && !R.esSesion(s));
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
        <td>${chipSemaforo(s.Semaforo)}</td><td${edSeg(s, "Logros", "Logros")}>${esc(s.Logros || "—")}</td><td class="opc"${edSeg(s, "Proximos_Pasos", "Próximos pasos")}>${esc(s.Proximos_Pasos || "—")}</td>
        <td class="${s.Bloqueos ? "bloqueo" : ""}"${edSeg(s, "Bloqueos", "Bloqueos")}>${esc(s.Bloqueos || "—")}</td>
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

  // Logros, próximos pasos y bloqueos de una sesión, editables en la tabla.
  const edSeg = (s, campo, etiqueta) => ed({ puede: R.puedeEditar(S.usuario, proyecto(s.ID_Proyecto) || {}, "seguimiento"), tabla: "Seguimientos", idCol: "ID_Seguimiento", id: s.ID_Seguimiento, campo, tipo: "textarea", valor: s[campo] ?? "", etiqueta });
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
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(nombreUsuario(p.PM))}</div></td><td${ed({ puede: R.puedeEditar(S.usuario, p, "editarFrecuencia"), tabla: "Proyectos", idCol: "ID_Proyecto", id: p.ID_Proyecto, campo: "Frecuencia_Seguimiento", tipo: "select", opciones: Object.keys(R.DIAS_FRECUENCIA), valor: p.Frecuencia_Seguimiento, extra: () => sello(), etiqueta: "Frecuencia" })}>${esc(p.Frecuencia_Seguimiento)}</td>
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
    const ids = comentarios.flatMap((x) => adjuntosDe(x).map((a) => a.id));
    if (ids.length) { try { await S.api.borrarArchivos(ids); } catch (e) { /* ya no existen */ } }
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
        return `<tr>${conProyecto ? `<td><button class="btn enlace" data-pid="${esc(c.ID_Proyecto)}" title="Ir al proyecto"><b>${esc(p.Nombre)}</b></button></td>` : ""}
          <td${edCompTexto(c)}>${esc(c.Compromiso)}<div class="sub">${esc(origen(c))}</div></td><td>${celdaResp(c, R.puedeEditar(S.usuario, p, "compromisos"))}</td>
          <td class="nowrap"${edCompFecha(c)}>${c.Fecha_Compromiso ? fecha(c.Fecha_Compromiso) : `<span class="sub">—</span>`}</td>
          <td${edCompEstado(c)}>${pillComp(c)}${e === "Cerrado" && c.Fecha_Cierre ? `<div class="sub">${fecha(c.Fecha_Cierre)}</div>` : ""}</td>
          <td class="opc">${ult ? `<span class="ult-com">${esc(resumenCom(ult))}</span><div class="sub">${esc(nombreUsuario(ult.Autor))} · ${fechaHora(ult.Fecha_Hora)}</div>` : `<span class="sub">Sin comentarios</span>`}</td>
          <td class="derecha nowrap"><button class="btn chico ${coms.length ? "" : "primario-suave"}" data-abrir-comp="${esc(c.ID_Compromiso)}">${coms.length ? `Comentarios (${coms.length})` : "Abrir"}</button>
            ${R.puedeEditar(S.usuario, p, "compromisos") ? botonesEstado(c) : ""}</td></tr>`;
      }).join("")}</tbody></table></div>`;
  }
  // ---------- Alertas de compromisos ----------
  // A quién escribirle: el correo del compromiso, o el de un usuario/stakeholder con ese nombre, o el texto si ya es un correo.
  // Un compromiso puede tener uno o varios responsables: Responsable, Correo_Responsable e ID_Recurso guardan listas separadas por «; ».
  const partesLista = (v) => String(v ?? "").split(";").map((x) => x.trim());
  function correoSuelto(nombre) {
    const r = lc(nombre);
    if (!r) return "";
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r)) return String(nombre).trim();
    const u = S.datos.Usuarios.find((x) => lc(x.Nombre) === r || lc(x.Correo) === r);
    if (u) return u.Correo;
    const st = (S.datos.Stakeholders || []).find((x) => lc(x.Nombre) === r && x.Correo);
    return st ? st.Correo : "";
  }
  // «Grupo de proyecto»: un responsable especial que representa a todos los stakeholders configurados en el proyecto.
  const GRUPO_PROY = "Grupo de proyecto";
  const esGrupo = (n) => normTxt(n) === normTxt(GRUPO_PROY);
  const integrantesGrupo = (pid) => stakeholdersDe(pid).filter((x) => x.Nombre).map((x) => { const r = recursoPor({ id: x.ID_Recurso, correo: x.Correo, nombre: x.Nombre });
    return { Nombre: r ? r.Nombre : x.Nombre, Correo: x.Correo || (r && r.Correo) || "", ID_Recurso: r ? r.ID_Recurso : "", r, Rol: x.Rol, grupo: true }; });
  // expandir = true: el «Grupo de proyecto» se reemplaza por sus integrantes (para correos, alertas y «a tu cargo»).
  function responsablesDe(c, expandir = true) {
    const ns = partesLista(c.Responsable).filter(Boolean);
    if (!ns.length) return [];
    const ids = partesLista(c.ID_Recurso), ms = partesLista(c.Correo_Responsable).filter((x) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
    const lista = ns.map((n, i) => {
      if (esGrupo(n)) return { Nombre: GRUPO_PROY, Correo: "", ID_Recurso: "", r: null, esGrupo: true };
      const r = recursoPor({ id: ids.length === ns.length ? ids[i] : "", nombre: n });
      return { Nombre: n, Correo: (ms.length === ns.length ? ms[i] : "") || (r && r.Correo) || correoSuelto(n), ID_Recurso: r ? r.ID_Recurso : "", r };
    });
    if (!expandir || !lista.some((x) => x.esGrupo)) return lista;
    const out = [];
    lista.forEach((x) => (x.esGrupo ? integrantesGrupo(c.ID_Proyecto) : [x]).forEach((y) => { if (!out.some((z) => normTxt(z.Nombre) === normTxt(y.Nombre))) out.push(y); }));
    return out;
  }
  const serializarResp = (rs) => ({ Responsable: rs.map((x) => x.Nombre).join("; "),
    Correo_Responsable: rs.every((x) => validoCorreo(x.Correo)) ? rs.map((x) => x.Correo.trim()).join("; ") : rs.map((x) => x.Correo).filter(validoCorreo).join("; "),
    ID_Recurso: rs.some((x) => x.ID_Recurso) ? rs.map((x) => x.ID_Recurso || "").join("; ") : "" });
  const correosDe = (c) => [...new Set(responsablesDe(c).map((x) => String(x.Correo || "").trim()).filter(validoCorreo))];
  const correoDe = (c) => correosDe(c).join("; ");
  const esMioComp = (c) => { const u = S.usuario; return responsablesDe(c).some((x) => (x.Correo && lc(x.Correo) === lc(u.Correo)) || (u.Nombre && lc(x.Nombre) === lc(u.Nombre))); };
  function alertas() {
    const ids = new Set(visibles().filter((p) => p.Estado !== "Cancelado").map((p) => p.ID_Proyecto));   // lo cancelado no genera alertas
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
    const para = [...new Set(cs.flatMap(correosDe))];
    const nombres = [...new Set(cs.flatMap((c) => responsablesDe(c).map((x) => (x.r || {}).Nombre || x.Nombre)))];
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
      `${responsablesDe(c).length > 1 ? "Responsables" : "Responsable"}: ${responsablesDe(c).map((x) => x.Nombre).join(", ") || "—"}`,
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
      const para = [...new Set(items.flatMap((x) => correosDe(x.c)))];
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
    const V = S.compVista = S.compVista || { ver: "abiertos", orden: leerLocal("pmo_orden_comp", "fecha"), agrupar: false };
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
      recientes: (a, b) => String(fechaCreacion(b.c)).localeCompare(String(fechaCreacion(a.c))) || String(b.c.ID_Compromiso).localeCompare(String(a.c.ID_Compromiso), "es", { numeric: true }),
      antiguos: (a, b) => String(fechaCreacion(a.c) || "9999").localeCompare(String(fechaCreacion(b.c) || "9999")) || String(a.c.ID_Compromiso).localeCompare(String(b.c.ID_Compromiso), "es", { numeric: true }),
    };
    lista = lista.slice().sort(ORDEN[V.orden] || porFecha);
    const seg = (id, t) => `<button class="btn chico ${V.ver === id ? "primario" : ""}" data-cver="${id}">${t} (${cuenta[id]})</button>`;
    const fila = ({ c, e }) => {
      const coms = comentariosDe(c.ID_Compromiso);
      const ult = coms[coms.length - 1];
      return `<tr><td${edCompTexto(c)}>${esc(c.Compromiso)}${ult ? `<div class="sub ult-com">💬 ${esc(resumenCom(ult))}</div>` : ""}</td>
        <td class="nowrap"${edCompCreado(c)}>${fechaCreacion(c) ? fecha(fechaCreacion(c)) : `<span class="sub">—</span>`}</td>
        <td${edCompSesion(c)}>${c.ID_Seguimiento ? `Sesión ${fecha(fSes(c))}` : `<span class="sub">Sin sesión</span>`}</td>
        <td>${celdaResp(c, puede)}</td><td class="nowrap"${edCompFecha(c)}>${c.Fecha_Compromiso ? fecha(c.Fecha_Compromiso) : `<span class="sub">—</span>`}</td>
        <td${edCompEstado(c)}>${pillComp(c)}${e === "Cerrado" && c.Fecha_Cierre ? `<div class="sub">${fecha(c.Fecha_Cierre)}</div>` : ""}</td>
        <td class="derecha nowrap"><button class="btn chico ${coms.length ? "" : "primario-suave"}" data-abrir-comp="${esc(c.ID_Compromiso)}">${coms.length ? `Comentarios (${coms.length})` : "Abrir"}</button>
          ${puede ? botonesEstado(c) : ""}</td></tr>`;
    };
    let filas;
    if (V.agrupar) {
      const grupos = [...new Set(lista.map((x) => x.c.ID_Seguimiento || ""))]
        .sort((a, b) => (a === "" ? 1 : b === "" ? -1 : String(fSes({ ID_Seguimiento: b })).localeCompare(String(fSes({ ID_Seguimiento: a })))));
      filas = grupos.map((g) => {
        const del = lista.filter((x) => (x.c.ID_Seguimiento || "") === g);
        return `<tr class="grupo"><td colspan="7">${g ? `Sesión del ${fecha(fSes({ ID_Seguimiento: g }))}` : "Registrados en el proyecto (sin sesión)"} <span class="contador">${del.length}</span></td></tr>${del.map(fila).join("")}`;
      }).join("");
    } else filas = lista.map(fila).join("");
    return `<div class="card"><div class="titulo-fila"><h2>Compromisos del proyecto</h2>${puede ? `<button class="btn primario" id="b-comp">+ Compromiso</button>` : ""}</div>
      <p class="sub">Un compromiso puede quedar atado a una sesión de seguimiento o solo al proyecto. Ábrelo para comentar su avance hasta cerrarlo.</p>
      <div class="barra-comp">
        <div class="seg" role="group" aria-label="Qué compromisos ver">${seg("abiertos", "Abiertos")}${seg("pendientes", "Pendientes")}${seg("encurso", "En curso")}${seg("vencidos", "Vencidos")}${seg("cerrados", "Cerrados")}${seg("todos", "Todos")}</div>
        <label class="filtro"><span>Ordenar por</span><select id="c-orden">
          ${[["fecha", "Fecha límite"], ["recientes", "Creación (más recientes primero)"], ["antiguos", "Creación (más antiguos primero)"], ["estado", "Estado (vencidos primero)"], ["responsable", "Responsable"]].map(([v, t]) => `<option value="${v}" ${V.orden === v ? "selected" : ""}>${t}</option>`).join("")}</select></label>
        <label class="check"><input type="checkbox" id="c-agrupar" ${V.agrupar ? "checked" : ""}> Agrupar por sesión</label>
      </div>
      ${lista.length ? `<div class="tabla-scroll"><table class="tabla-comp"><thead><tr><th>Compromiso</th><th>Creado</th><th>Sesión</th><th>Responsable</th><th>Fecha límite</th><th>Estado</th><th></th></tr></thead>
        <tbody>${filas}</tbody></table></div>`
        : vacio(comps.length ? "No hay compromisos en esta vista. Prueba «Todos»." : `Aún no hay compromisos.${puede ? " Crea el primero con «+ Compromiso»." : ""}`)}</div>`;
  }
  function enlazarSeccionCompromisos(el) {
    el.querySelectorAll("[data-cver]").forEach((b) => b.addEventListener("click", () => { S.compVista.ver = b.dataset.cver; render(); }));
    const o = el.querySelector("#c-orden"); if (o) o.addEventListener("change", () => { S.compVista.orden = o.value; guardarLocal("pmo_orden_comp", o.value); render(); });
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
  // Celdas editables de compromisos (texto, fecha, estado, sesión, creación) con constancia en el hilo.
  const edComp = (c, campo, o = {}) => ed({ puede: R.puedeEditar(S.usuario, proyecto(c.ID_Proyecto) || {}, "compromisos"), tabla: "Compromisos", idCol: "ID_Compromiso", id: c.ID_Compromiso, campo, valor: c[campo] ?? "", ...o });
  const edCompTexto = (c) => edComp(c, "Compromiso", { tipo: "textarea", requerido: true, etiqueta: "Compromiso" });
  const edCompFecha = (c) => edComp(c, "Fecha_Compromiso", { tipo: "date", etiqueta: "Fecha límite", despues: (v) => S.api.agregarFila("Comentarios", nuevoComentario(c, `Fecha límite: ${c.Fecha_Compromiso ? fecha(c.Fecha_Compromiso) : "sin fecha"} → ${v ? fecha(v) : "sin fecha"}.`)) });
  const edCompEstado = (c) => edComp(c, "Estado", { tipo: "select", opciones: R.ESTADOS_COMPROMISO, valor: R.estadoBase(c), etiqueta: "Estado del compromiso",
    extra: (v) => ({ Fecha_Cierre: v === "Cerrado" ? R.hoyISO() : "" }), despues: (v) => S.api.agregarFila("Comentarios", nuevoComentario(c, TEXTO_ESTADO[v])) });
  const edCompSesion = (c) => edComp(c, "ID_Seguimiento", { tipo: "select", opciones: opcionesSesion(c.ID_Proyecto), vacio: "Sin sesión (registrado en el proyecto)", etiqueta: "Sesión" });
  const edCompCreado = (c) => edComp(c, "Fecha_Creacion", { tipo: "date", valor: c.Fecha_Creacion || "", validar: (v) => (v && v > R.hoyISO() ? "La fecha de creación no puede ser futura." : ""), etiqueta: "Fecha de creación" });
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
        <div><span class="sub">${responsablesDe(c).length > 1 ? "Responsables" : "Responsable"}${puede ? ` <button class="btn enlace chico" data-cambiar-resp="${esc(c.ID_Compromiso)}">✎ Cambiar</button>` : ""}</span>${responsablesDe(c, false).some((x) => x.esGrupo) ? `<div class="sub">👥 ${GRUPO_PROY}:</div>` : ""}${responsablesDe(c).map((x) => `<b>${esc(x.Nombre)}</b>${(() => { const rol = rolEnProyecto(c.ID_Proyecto, x); const d = [rol, x.r && x.r.Cargo, x.r && empresaDe(x.r), x.Correo].filter(Boolean); return d.length ? `<div class="sub">${esc(d.join(" · "))}</div>` : ""; })()}`).join("") || "<b>—</b>"}</div>
        <div><span class="sub">Creado</span><b>${fechaCreacion(c) ? fecha(fechaCreacion(c)) : "—"}</b></div>
        <div><span class="sub">Fecha límite</span><b>${fecha(c.Fecha_Compromiso)}</b></div>
        ${c.Fecha_Cierre ? `<div><span class="sub">Cerrado el</span><b>${fecha(c.Fecha_Cierre)}</b></div>` : ""}
      </div>
      <div class="form-seccion">Seguimiento del compromiso <span class="contador">${coms.length}</span></div>
      <div class="hilo">${coms.map((x) => `<div class="com"><div class="com-cab"><b>${esc(nombreUsuario(x.Autor))}</b><span class="sub">${fechaHora(x.Fecha_Hora)}${x.Editado_El ? ` · editado` : ""}${lc(x.Autor) === lc(S.usuario.Correo) || R.tieneRol(S.usuario, "Admin") ? ` <button class="btn enlace com-borrar" data-editar-com="${esc(x.ID_Comentario)}" title="Editar este comentario">Editar</button> <button class="btn enlace com-borrar" data-borrar-com="${esc(x.ID_Comentario)}" title="Eliminar este comentario">Eliminar</button>` : ""}</span></div><div class="com-texto" data-texto-com="${esc(x.ID_Comentario)}" ${x.Texto ? "" : "hidden"}>${esc(x.Texto || "")}</div>${htmlAdjuntos(x)}</div>`).join("") || `<div class="vacio">Aún no hay comentarios. Escribe el primero: avances, dudas o bloqueos de este compromiso.</div>`}</div>
      ${puede ? `<textarea id="d-texto" rows="3" placeholder="Escribe un comentario: qué se avanzó, qué falta, quién debe actuar… (puedes pegar aquí una imagen con Ctrl+V)"></textarea>
        <div class="adj-barra"><button type="button" class="btn chico" id="d-adjuntar">📎 Adjuntar archivo</button><span class="sub">o pega una imagen del correo con Ctrl+V · máx. 5 MB por archivo</span></div>
        <div id="d-pend" class="adj-pend"></div>
        <div class="error-campo" id="d-error" role="alert"></div>
        <div class="acciones derecha">
          <button class="btn" id="d-editar">Editar compromiso</button>
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
    // Editar un comentario propio (o cualquiera, si es Admin) en el mismo hilo.
    m.querySelectorAll("[data-editar-com]").forEach((b) => b.addEventListener("click", () => {
      const x = coms.find((y) => y.ID_Comentario === b.dataset.editarCom);
      const div = m.querySelector(`[data-texto-com="${CSS.escape(x.ID_Comentario)}"]`);
      if (!div || div.querySelector("textarea")) return;
      div.hidden = false;
      div.innerHTML = `<textarea rows="3" class="com-edit">${esc(x.Texto || "")}</textarea><div class="acciones"><button class="btn primario chico" data-ce-ok>Guardar</button><button class="btn chico" data-ce-no>Cancelar</button></div>`;
      const ta = div.querySelector("textarea"); ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length);
      const cancelar = () => detalleCompromiso(id);
      ta.addEventListener("keydown", (ev) => { if (ev.key === "Escape") { ev.stopPropagation(); cancelar(); } });
      div.querySelector("[data-ce-no]").addEventListener("click", cancelar);
      div.querySelector("[data-ce-ok]").addEventListener("click", async () => {
        const nuevo = ta.value.trim();
        if (!nuevo && !adjuntosDe(x).length) { ta.focus(); return; }
        if (nuevo === String(x.Texto || "").trim()) return cancelar();
        await guardar(() => S.api.actualizarPorId("Comentarios", "ID_Comentario", x.ID_Comentario, { Texto: nuevo, Editado_El: ahora() }), "Comentario editado");
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
    // Celdas editables con clic (según permisos). Cerrar/cancelar y mover la fecha comprometida pasan por su flujo.
    const pmsEd = usuariosConRol("PM").sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es")).map((x) => [x.Correo, x.Nombre || x.Correo]);
    const EP = (p, campo, o = {}) => ed({ puede: R.puedeEditar(S.usuario, p, "editarProyecto") && (o.puede ?? true), tabla: "Proyectos", idCol: "ID_Proyecto", id: p.ID_Proyecto, campo, valor: p[campo] ?? "", extra: () => sello(), ...o, ...(o.puede !== undefined ? { puede: R.puedeEditar(S.usuario, p, "editarProyecto") && o.puede } : {}) });
    const edEstado = (p) => (v) => {
      if (!["Cerrado", "Cancelado"].includes(v) || p.Estado === v) return true;
      formProyecto(p);
      const sel = document.querySelector('#m-form [name="Estado"]'); if (sel) { sel.value = v; sel.dispatchEvent(new Event("change", { bubbles: true })); }
      toast(`Para ${v === "Cerrado" ? "cerrar" : "cancelar"} el proyecto revisa y guarda el formulario.`);
      return false;
    };
    const edFin = (p) => () => {
      if (!lineaBase(p).tiene) return true;
      confirmar("Fecha comprometida en línea base", `«${p.Nombre}» tiene línea base: mover el fin planeado se registra como control de cambios (tiempo) para no perder la trazabilidad.`, "Registrar control de cambios", () => { cerrarModal(); formCambio(p); }, "primario");
      return false;
    };
    el.innerHTML = `
      <div class="titulo-fila"><h1>Proyectos</h1>${R.puede(S.usuario, "crearProyecto") ? `<button class="btn primario" id="b-nuevo">+ Nueva iniciativa</button>` : ""}</div>
      ${barraFiltros()}
      <div class="filtros">
        <label class="filtro crece"><span>Buscar</span><input id="f-texto" type="search" placeholder="ID, nombre, área o código Almera" value="${esc(FL.texto)}"></label>
        <label class="filtro"><span>Semáforo</span><select id="f-sem"><option value="">Todos</option>${(S.cat.Semaforo || []).map((v) => `<option ${v === FL.semaforo ? "selected" : ""}>${esc(v)}</option>`).join("")}</select></label>
        <span class="sub">${lista.length} resultado(s)</span>
      </div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th class="opc">Área</th><th>PM</th><th class="opc">Metodología</th><th class="opc">Fase</th><th>Estado</th><th>Semáforo</th><th>Avance real / plan</th><th class="opc">Fin plan</th><th class="opc">Valor FSFB</th><th>Seguimiento</th></tr></thead>
        <tbody>${lista.map((p) => `<tr class="clic ${["Cerrado", "Cancelado"].includes(p.Estado) ? "fila-inactiva" : ""}" data-pid="${esc(p.ID_Proyecto)}" tabindex="0">
          <td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.ID_Proyecto)} · prioridad ${esc(String(p.Prioridad).toLowerCase())}</div></td><td class="opc"${EP(p, "Cliente_Area", { tipo: "select", opciones: S.cat.Cliente_Area, etiqueta: "Área" })}>${esc(p.Cliente_Area)}</td>
          <td${EP(p, "PM", { tipo: "select", opciones: pmsEd, puede: R.puede(S.usuario, "verTodo"), etiqueta: "PM" })}>${persona(nombreUsuario(p.PM))}</td><td class="opc"${EP(p, "Metodologia", { tipo: "select", opciones: S.cat.Metodologia, etiqueta: "Metodología" })}>${esc(p.Metodologia)}</td>
          <td class="opc"${EP(p, "Fase", { tipo: "select", opciones: S.cat.Fase, etiqueta: "Fase" })}>${chipFase(p.Fase)}</td><td${EP(p, "Estado", { tipo: "select", opciones: S.cat.Estado, antes: edEstado(p), etiqueta: "Estado" })}>${chipEstadoP(p.Estado)}</td>
          <td${EP(p, "Semaforo", { tipo: "select", opciones: S.cat.Semaforo, vacio: "—", etiqueta: "Semáforo" })}>${p.Estado === "Activo" || p.Semaforo ? chipSemaforo(p.Semaforo) : `<span class="sub">—</span>`}</td>
          <td class="celda-avance"${EP(p, "Avance_Real", { tipo: "number", min: 0, max: 100, requerido: true, etiqueta: "Avance real" })}>${barraAvance(p.Avance_Real, p.Avance_Planeado)}</td><td class="opc"${EP(p, "Fecha_Fin_Plan", { tipo: "date", requerido: true, antes: edFin(p), validar: (v) => (p.Fecha_Inicio && v < p.Fecha_Inicio ? "El fin no puede ser anterior al inicio." : ""), etiqueta: "Fin planeado" })}>${fecha(p.Fecha_Fin_Plan)}</td>
          <td class="opc"${EP(p, "Nivel_Valor", { tipo: "select", opciones: NIVELES_VALOR, vacio: `Automático (${valorProyecto(p).sugerido || "sin datos"})`, etiqueta: "Nivel de valor" })}>${(() => { const V = valorProyecto(p); return V.nivel ? `${chipNivel(V.nivel)}${V.indice !== null ? `<div class="sub">logro ${V.indice}%</div>` : ""}` : `<span class="sub">—</span>`; })()}</td>
          <td>${pill(R.estadoSeguimiento(p, S.datos.Seguimientos).estado)}</td></tr>`).join("") || `<tr><td colspan="11">${vacio(visibles().length ? "No hay proyectos con estos filtros. Prueba «Limpiar filtros»." : "Aún no tienes proyectos. Registra el primero con «+ Nueva iniciativa».")}</td></tr>`}</tbody>
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
    const VP = valorProyecto(p);
    const TABS = [["resumen", "Resumen"], ["valor", `Valor${VP.nivel ? ` <span class="contador ${VP.riesgo || VP.pendientes ? "rojo" : ""}">${esc(VP.nivel)}</span>` : ""}`], ["compromisos", `Compromisos <span class="contador ${vencidosN ? "rojo" : ""}">${abiertosN}</span>`],
      ["seguimientos", `Seguimientos <span class="contador">${segs.length}</span>`], ["hitos", `Hitos <span class="contador">${hitos.length}</span>`],
      ["gantt", (() => { const ts = tareasDe(p.ID_Proyecto); const atr = ts.length ? calcCronograma(ts).atrasadas.length : 0; return `Cronograma <span class="contador ${atr ? "rojo" : ""}">${atr || ts.length}</span>`; })()], ["riesgos", `Riesgos <span class="contador">${riesgos.filter(riesgoActivo).length}</span>`],
      ["tickets", `Tickets <span class="contador">${ticketsDe(p.ID_Proyecto).filter(ticketAbierto).length}</span>`],
      ["documentos", `Documentos <span class="contador">${documentosDe(p.ID_Proyecto).length}</span>`],
      ["cambios", `Cambios${cambiosDe(p.ID_Proyecto).some((c) => c.Estado === "Solicitado") ? ` <span class="contador rojo">${cambiosDe(p.ID_Proyecto).filter((c) => c.Estado === "Solicitado").length}</span>` : ""}`],
      ["raci", "RACI"], ["lecciones", `Lecciones <span class="contador">${leccionesDe(p.ID_Proyecto).length}</span>`], ["auditoria", "Historial de cambios"]];
    const tab = TABS.some((t) => t[0] === S.fichaTab) ? S.fichaTab : "resumen";
    const pestanas = `<div class="pestanas" role="tablist">${TABS.map(([id, t]) => `<button class="pestana ${id === tab ? "activa" : ""}" role="tab" aria-selected="${id === tab}" data-tab="${id}">${t}</button>`).join("")}</div>`;
    const cuerpo = {
      resumen: () => `      <div class="card datos">
        ${dato("Código Almera", p.Codigo_Almera ? esc(p.Codigo_Almera) : `<span class="sub">Sin registrar</span>`)}${dato("Estado", chipEstadoP(p.Estado))}${dato("Fase", chipFase(p.Fase))}${dato("Semáforo", chipSemaforo(p.Semaforo))}
        ${dato("Avance real / plan", barraAvance(p.Avance_Real, p.Avance_Planeado))}
        ${dato("Inicio", fecha(p.Fecha_Inicio))}${dato("Fin planeado", fecha(p.Fecha_Fin_Plan))}
        ${dato("Presupuesto", cop(p.Presupuesto))}${dato("Ejecutado", cop(p.Ejecutado))}
        ${dato("Metodología", esc(p.Metodologia))}${dato("Prioridad", esc(p.Prioridad))}
        ${dato("Frecuencia", agendaDe(p.ID_Proyecto).filter((a) => a.Activa !== "No").length ? agendaDe(p.ID_Proyecto).filter((a) => a.Activa !== "No").map((a) => `<div>${esc(resumenAgenda(a))}</div>`).join("") : esc(p.Frecuencia_Seguimiento))}${dato("Seguimiento", `${pill(e.estado)} <span class="sub">próximo ${fecha(e.proximo)}</span>`)}
        ${dato("Repositorio", enlace(p.URL_Repositorio, "Abrir") || `<span class="sub">Sin registrar</span>`)}${dato("Documentos", enlace(p.URL_Documentos, "Abrir") || `<span class="sub">Sin registrar</span>`)}
        <div class="dato ancho"><div class="dato-k">Comentario de estado</div><div class="dato-v">${esc(p.Comentario_Estado || "—")}</div></div>
      </div>
      ${cardDesempeno(p)}
      ${(() => { const cc = cambiosDe(p.ID_Proyecto); const ap = cc.filter((c) => c.Estado === "Aprobado"), pe = cc.filter((c) => c.Estado === "Solicitado");
        return `<div class="card"><div class="titulo-fila"><h2>Control de cambios <span class="contador">${cc.length}</span></h2><div class="acciones">${puedeP("editarProyecto") ? `<button class="btn chico primario" id="b-cambio-res">+ Control de cambios</button>` : ""}<button class="btn chico" data-ir-tab="cambios">Ver todos</button></div></div>
          ${cc.length ? `<div class="cc-resumen">${ap.length ? `<span class="pill aprobado">${ap.length} aprobado(s)</span>` : ""}${pe.length ? `<span class="pill solicitado">${pe.length} por decidir</span>` : ""}${IMPACTOS_CC.map((i) => { const n = ap.filter((c) => impactosDe(c).includes(i)).length; return n ? `<span class="pill cc-imp">${esc(i)}: ${n}</span>` : ""; }).join("")}</div>
            <ul class="lista cc-ult">${cc.slice(0, 3).map((c) => `<li>${c.Numero_CC ? `<b>${esc(c.Numero_CC)}</b> · ` : ""}${fecha(c.Fecha)} · ${impactosDe(c).map((i) => `<span class="pill cc-imp">${esc(i)}</span>`).join(" ")} ${esc(c.Descripcion)} ${pill(c.Estado || "Solicitado")}${c.URL_Documento && esURL(c.URL_Documento) ? ` <a class="cc-doc" href="${esc(c.URL_Documento)}" target="_blank" rel="noopener">📄 Documento ↗</a>` : ""}</li>`).join("")}</ul>`
          : `<p class="sub">Sin controles de cambio registrados. Registra aquí los cambios de tiempo, costo o alcance con el enlace a su documento.</p>`}</div>`; })()}
      ${cardStakeholders(p, puedeP("editarProyecto"))}
      ${cardDependencias(p, puedeP("editarProyecto"))}
      <div class="card"><h2>Curva S del presupuesto</h2>${(() => { S._curvaP = datosCurvaS([p]); return S._curvaP ? `<div class="grafico-alto"><canvas id="g-curvas-p"></canvas></div><p class="sub">Planeado: ${num0(p.Presupuesto_Base) !== null ? "presupuesto de la línea base" : "presupuesto"} repartido entre inicio y fin. Real: ejecutado reportado en cada seguimiento.</p>` : vacio("Registra presupuesto y fechas para ver la curva S."); })()}</div>
      <div class="card"><h2>Curva de avance</h2>${segs.some((x) => !R.esSesion(x)) ? `<div class="grafico-alto"><canvas id="g-curva"></canvas></div>` : vacio("Aún no hay seguimientos. La curva aparece con el primer reporte.")}</div>
`,
      compromisos: () => seccionCompromisos(p, comps, segs, puedeP("compromisos")),
      tickets: () => seccionTickets(p, puedeP("compromisos")),
      documentos: () => seccionDocumentos(p, puedeP("compromisos")),
      gantt: () => seccionGantt(p, puedeP("hitos")),
      cambios: () => seccionCambios(p, puedeP("editarProyecto"), R.puede(u, "verTodo")),
      raci: () => seccionRACI(p, puedeP("editarProyecto")),
      valor: () => seccionValor(p, puedeP("seguimiento")),
      lecciones: () => seccionLecciones(p, puedeP("editarProyecto")),
      auditoria: () => seccionAuditoria(p),
      seguimientos: () => `${seccionAgenda(p, puedeP("seguimiento"))}      <div class="card"><div class="titulo-fila"><h2>Historial de seguimientos y sesiones</h2><div class="acciones"><button class="btn chico" id="b-expandir">Expandir compromisos</button><button class="btn chico" id="b-contraer">Contraer</button>${puedeP("seguimiento") && p.Estado !== "Cancelado" ? `<button class="btn" id="b-ses2">+ Registrar sesión</button>` : ""}${puedeP("seguimiento") && p.Estado === "Activo" ? `<button class="btn" id="b-seg2">+ Registrar seguimiento</button>` : ""}</div></div>
        <p class="sub">Cada fila es una sesión: su avance, lo que pasó, los compromisos que se acordaron y el acta.</p><div class="tabla-scroll"><table class="historial">
        <thead><tr><th>Sesión</th><th>Avance</th><th>Logros y próximos pasos</th><th>Bloqueos</th><th>Compromisos acordados</th><th>Acta</th></tr></thead>
        <tbody>${segs.slice().reverse().map((s) => {
          const acordados = S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento).sort(porCreacion);
          if (R.esSesion(s)) return `<tr class="fila-sesion"><td><b>${fecha(s.Fecha_Corte)}</b> <span class="pill sesion">Sesión</span>${s.Nombre_Sesion ? `<div><b>${esc(s.Nombre_Sesion)}</b></div>` : ""}<div class="sub">${esc(nombreUsuario(s.Reportado_Por))}</div></td>
          <td><span class="sub">No reporta avance</span></td>
          <td${edSeg(s, "Logros", "Logros")}>${esc(s.Logros || "—")}${s.Proximos_Pasos ? `<div class="sub">Sigue: ${esc(s.Proximos_Pasos)}</div>` : ""}</td><td${edSeg(s, "Bloqueos", "Bloqueos")}>${esc(s.Bloqueos || "—")}</td>
          <td>${acordados.length ? `<details class="acordados" data-seg-det="${esc(s.ID_Seguimiento)}" ${(S.segAbiertos || new Set()).has(s.ID_Seguimiento) ? "open" : ""}><summary><b>${acordados.length} compromiso(s)</b> <span class="sub">· ${acordados.filter((c) => !R.cerrado(c)).length} abierto(s)</span></summary>
            <ul class="mini">${acordados.map((c) => `<li><button class="btn enlace" data-abrir-comp="${esc(c.ID_Compromiso)}">${esc(c.Compromiso)}</button> <span class="sub">${esc(c.Responsable)} · ${fecha(c.Fecha_Compromiso)}</span> ${pillComp(c)}</li>`).join("")}</ul></details>` : `<span class="sub">Sin compromisos</span>`}
            ${puedeP("compromisos") ? `<div><button class="btn chico" data-comp-seg="${esc(s.ID_Seguimiento)}">+ Compromiso de esta sesión</button></div>` : ""}</td>
          <td>${s.Fecha_Acta ? fecha(s.Fecha_Acta) : `<span class="sub">Sin fecha</span>`}<div class="celda-acta">${celdaActa(s, puedeP("seguimiento"))}</div>
            ${puedeP("seguimiento") ? `<button class="btn chico enlace-edicion" data-edit-seg="${esc(s.ID_Seguimiento)}">Editar sesión</button>` : ""}</td></tr>`;
          return `<tr><td><b>${fecha(s.Fecha_Corte)}</b> <span class="pill encurso">Seguimiento</span><div class="sub">${esc(s.Semana)} · ${esc(nombreUsuario(s.Reportado_Por))}</div></td>
          <td>${pct(s.Avance_Real)}${s.Avance_Planeado !== "" && s.Avance_Planeado != null ? ` <span class="sub">/ ${pct(s.Avance_Planeado)}</span>` : ""}<div>${chipSemaforo(s.Semaforo)}</div>${(() => { const v = (S.datos.Cronogramas || []).find((x) => x.ID_Seguimiento === s.ID_Seguimiento); return v ? `<button class="btn enlace" data-ir-crv="${esc(v.ID_Version)}" title="Ver el cronograma de esta sesión">📅 V${esc(v.Numero)}</button>` : ""; })()}</td>
          <td${edSeg(s, "Logros", "Logros")}>${esc(s.Logros || "—")}<div class="sub">Sigue: ${esc(s.Proximos_Pasos || "—")}</div></td><td${edSeg(s, "Bloqueos", "Bloqueos")}>${esc(s.Bloqueos || "—")}</td>
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
        <tbody>${hitos.map((h) => { const est = h.Estado !== "Cumplido" && h.Fecha_Plan < hoy ? "Atrasado" : h.Estado; const EH = (campo, o) => ed({ puede: puedeP("hitos"), tabla: "Hitos", idCol: "ID_Hito", id: h.ID_Hito, campo, valor: h[campo] ?? "", ...o });
          return `<tr><td${EH("Hito", { tipo: "text", requerido: true, etiqueta: "Hito" })}>${esc(h.Hito)}</td><td${EH("Fecha_Plan", { tipo: "date", requerido: true, etiqueta: "Fecha plan" })}>${fecha(h.Fecha_Plan)}</td>
          <td${EH("Fecha_Real", { tipo: "date", validar: (v) => (v && v > hoy ? "La fecha real no puede ser futura." : ""), extra: (v) => ({ Estado: v ? "Cumplido" : h.Estado === "Cumplido" ? "Pendiente" : h.Estado }), etiqueta: "Fecha real" })}>${h.Fecha_Real ? fecha(h.Fecha_Real) : "—"}</td>
          <td${EH("Estado", { tipo: "select", opciones: S.cat.Estado_Hito, extra: (v) => (v === "Cumplido" ? { Fecha_Real: h.Fecha_Real || hoy } : { Fecha_Real: "" }), etiqueta: "Estado del hito" })}>${pill(est)}</td>
          <td class="derecha">${puedeP("hitos") ? `<button class="btn chico" data-hito="${esc(h.ID_Hito)}">Editar</button>` : ""}</td></tr>`; }).join("") || `<tr><td colspan="5">${vacio("Sin hitos. Agrega los hitos clave del proyecto con «+ Hito».")}</td></tr>`}</tbody></table></div></div>
`,
      riesgos: () => seccionRiesgos(p, puedeP("riesgos")),
    }[tab]();
    el.innerHTML = `
      <button class="btn enlace" id="b-volver">← Volver a proyectos</button>
      <div class="titulo-fila"><div><h1>${esc(p.Nombre)}</h1><div class="ficha-chips">${chipEstadoP(p.Estado)}${chipFase(p.Fase)}${p.Semaforo ? chipSemaforo(p.Semaforo) : ""}</div><div class="sub">${esc(p.ID_Proyecto)}${p.Codigo_Almera ? ` · Almera ${esc(p.Codigo_Almera)}` : ""} · ${esc(p.Cliente_Area)} · PM ${esc(nombreUsuario(p.PM))}</div></div>
        <div class="acciones">
          ${enlace(p.URL_Repositorio, "Repositorio")}${enlace(p.URL_Documentos, "Documentos")}
          ${puedeP("seguimiento") && p.Estado === "Activo" ? `<button class="btn primario" id="b-seg">Registrar seguimiento</button>` : ""}
          <button class="btn" id="b-hv" title="Documento con todo el historial del proyecto, listo para guardar como PDF">📄 Informe de proyecto</button>
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
    if ($("#b-ses2")) $("#b-ses2").addEventListener("click", () => formSeguimiento(p, "Sesión"));
    if ($("#b-tk")) $("#b-tk").addEventListener("click", () => formTicket(p));
    enlazarDocumentos(el, p);
    enlazarGantt(el, p);
    el.querySelectorAll("[data-ir-crv]").forEach((b) => b.addEventListener("click", (ev) => { ev.stopPropagation(); const F = ((S.ganttF = S.ganttF || {})[p.ID_Proyecto] = S.ganttF[p.ID_Proyecto] || { nivel: "", vista: "todo", q: "", cerr: [], escala: "ajustar", ver: "", comp: "" }); F.ver = b.dataset.irCrv; delete F.scroll; S.fichaTab = "gantt"; render(); }));
    $("#b-hv").addEventListener("click", () => hojaDeVida(p));
    if ($("#g-curvas-p") && S._curvaP) graficoCurvaS($("#g-curvas-p"), S._curvaP);
    if ($("#b-cambio")) $("#b-cambio").addEventListener("click", () => formCambio(p));
    if ($("#b-cambio-res")) $("#b-cambio-res").addEventListener("click", () => formCambio(p));
    el.querySelectorAll("[data-ir-tab]").forEach((b) => b.addEventListener("click", () => { S.fichaTab = b.dataset.irTab; render(); }));
    if ($("#b-lb")) $("#b-lb").addEventListener("click", () => fijarLineaBase(p));
    el.querySelectorAll("[data-lb]").forEach((b) => b.addEventListener("click", () => fijarLineaBase(p, (S.datos.LineasBase || []).find((x) => x.ID_LB === b.dataset.lb))));
    el.querySelectorAll("[data-cambio]").forEach((b) => b.addEventListener("click", () => formCambio(p, cambiosDe(p.ID_Proyecto).find((c) => c.ID_Cambio === b.dataset.cambio))));
    el.querySelectorAll("[data-decidir]").forEach((b) => b.addEventListener("click", () => decidirCambio(p, cambiosDe(p.ID_Proyecto).find((c) => c.ID_Cambio === b.dataset.decidir))));
    if ($("#b-leccion")) $("#b-leccion").addEventListener("click", () => formLeccion(p));
    el.querySelectorAll("[data-leccion]").forEach((b) => b.addEventListener("click", () => formLeccion(p, leccionesDe(p.ID_Proyecto).find((l) => l.ID_Leccion === b.dataset.leccion))));
    if ($("#b-dep")) $("#b-dep").addEventListener("click", () => formDependencia(p));
    el.querySelectorAll("[data-dep]").forEach((b) => b.addEventListener("click", () => formDependencia(p, (S.datos.Dependencias || []).find((d) => d.ID_Dependencia === b.dataset.dep))));
    if (tab === "raci") enlazarRACI(el, p);
    if (tab === "valor") enlazarValor(el, p);
    if (tab === "seguimientos") enlazarAgenda(el, p);
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
    const segsAv = segs.filter((x) => !R.esSesion(x));
    if (segsAv.length && $("#g-curva")) grafico($("#g-curva"), { type: "line", data: { labels: segsAv.map((s) => fecha(s.Fecha_Corte)), datasets: [
      { label: "Avance real", data: segsAv.map((s) => Number(s.Avance_Real) || 0), borderColor: COLORES.azul, backgroundColor: COLORES.azul, tension: 0.2 },
      ...(segsAv.some((s) => s.Avance_Planeado !== "" && s.Avance_Planeado !== undefined && s.Avance_Planeado !== null) ? [{ label: "Avance planeado", data: segsAv.map((s) => (s.Avance_Planeado === "" || s.Avance_Planeado == null ? null : Number(s.Avance_Planeado))), borderColor: "#9aa5b1", backgroundColor: "#9aa5b1", borderDash: [6, 4], spanGaps: true, tension: 0.2 }] : [])] },
      options: { maintainAspectRatio: false, plugins: { legend: { display: segsAv.some((s) => s.Avance_Planeado !== "" && s.Avance_Planeado != null) } }, scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + "%" } } } } });
  }

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
    const agrupar = S.cronoAgrupar ?? "pm";
    const selRango = `<div class="filtros"><label class="filtro"><span>Agrupar</span><select id="crono-agrupar"><option value="pm" ${agrupar === "pm" ? "selected" : ""}>Por PM</option><option value="" ${agrupar ? "" : "selected"}>Sin agrupar</option></select></label><label class="filtro"><span>Periodo</span><select id="crono-rango">${Object.entries(RANGOS).map(([k, t]) => `<option value="${k}" ${k === rango ? "selected" : ""}>${t}</option>`).join("")}</select></label></div>`;
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
        ${(() => { const filaP = (p) => {
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
        };
        if (!agrupar) return ps.map(filaP).join("");
        // Agrupado por PM: encabezado con sus proyectos, avance promedio y el rango que cubren.
        const cerr = S.cronoCerr = S.cronoCerr || new Set();
        const grupos = [...new Set(ps.map((p) => lc(p.PM)))].map((k) => ({ k, nombre: nombreUsuario(ps.find((p) => lc(p.PM) === k).PM) || "Sin PM", lista: ps.filter((p) => lc(p.PM) === k) }))
          .sort((a, b) => String(a.nombre).localeCompare(String(b.nombre), "es"));
        return grupos.map((g) => {
          const a = g.lista.reduce((m, p) => (p.Fecha_Inicio < m ? p.Fecha_Inicio : m), g.lista[0].Fecha_Inicio), b = g.lista.reduce((m, p) => (p.Fecha_Fin_Plan > m ? p.Fecha_Fin_Plan : m), g.lista[0].Fecha_Fin_Plan);
          const prom = Math.round(g.lista.reduce((t, p) => t + (Number(p.Avance_Real) || 0), 0) / g.lista.length);
          const rojos = g.lista.filter((p) => p.Semaforo === "Rojo").length, abierto = !cerr.has(g.k);
          return `<div class="g-fila g-grupo" data-gpm="${esc(g.k)}" role="button" tabindex="0" aria-expanded="${abierto}"><div class="g-nombre"><span class="g-tog">${abierto ? "▾" : "▸"}</span>${persona(g.nombre)}
              <div class="sub">${g.lista.length} proyecto(s) · avance prom. ${prom}%${rojos ? ` <span class="baja" style="white-space:nowrap">· ${rojos} en rojo</span>` : ""}</div></div>
            <div class="g-pista">${meses.map((m) => `<span class="g-rejilla ${m.m === 0 ? "enero" : ""}" style="left:${x(m.iso)}%"></span>`).join("")}
              <div class="g-rango" style="left:${x(a)}%;width:${Math.max(1, x(b) - x(a))}%" title="${esc(g.nombre)}: ${fecha(a)} → ${fecha(b)}"></div>
              ${hoy >= ini && hoy <= fin ? `<span class="g-hoy" style="left:${x(hoy)}%"></span>` : ""}</div></div>
            ${abierto ? g.lista.map(filaP).join("") : ""}`;
        }).join("");
        })()}
      </div>
      ${rango !== "todo" && ps.some((p) => recorte(p)) ? `<p class="sub">Las barras con borde punteado siguen antes o después del periodo mostrado. Elige «Todo el portafolio» para verlas completas.</p>` : ""}`;
    enlazarFiltros();
    $("#crono-rango").addEventListener("change", (e) => { S.cronoRango = e.target.value; render(); });
    $("#crono-agrupar").addEventListener("change", (e) => { S.cronoAgrupar = e.target.value; render(); });
    el.querySelectorAll("[data-gpm]").forEach((g) => { const t = () => { const c = S.cronoCerr = S.cronoCerr || new Set(); c.has(g.dataset.gpm) ? c.delete(g.dataset.gpm) : c.add(g.dataset.gpm); render(); }; g.addEventListener("click", t); g.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); t(); } }); });
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
    { lista: "Tipo_Actividad", t: "Tipo de otra actividad (capacidad)", campo: "Tipo", tabla: "Actividades", idCol: "ID_Actividad" },
    { lista: "Tipo_Sesion_Agenda", t: "Tipo de sesión (agenda de seguimiento)", campo: "Tipo", tabla: "Agenda", idCol: "ID_Sesion" },
    { lista: "Dimension_Valor", t: "Dimensión de valor (beneficios)", campo: "Dimension", tabla: "Beneficios", idCol: "ID_Beneficio" },
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
          <tbody>${proveedores().map((x) => { const EV = (campo, o) => ed({ puede: true, tabla: "Proveedores", idCol: "ID_Proveedor", id: x.ID_Proveedor, campo, valor: x[campo] ?? "", ...o }); return `<tr><td${EV("Nombre", { tipo: "text", requerido: true, etiqueta: "Proveedor" })}><b>${esc(x.Nombre)}</b>${x.NIT ? `<div class="sub">NIT ${esc(x.NIT)}</div>` : ""}</td><td${EV("Servicio", { tipo: "text", etiqueta: "Servicio" })}>${esc(x.Servicio || "—")}</td>
            <td${EV("Contacto", { tipo: "text", etiqueta: "Contacto" })}>${esc(x.Contacto || "")}${x.Contacto ? "<br>" : ""}${contacto(x.Correo, x.Telefono)}</td><td>${usosProv(x.ID_Proveedor)}</td><td${EV("Activo", { tipo: "select", opciones: ["Sí", "No"], valor: x.Activo || "Sí", etiqueta: "Activo" })}>${esc(x.Activo || "Sí")}</td>
            <td class="derecha"><button class="btn chico" data-prov="${esc(x.ID_Proveedor)}">Editar</button></td></tr>`; }).join("") || `<tr><td colspan="6">${vacio("Aún no hay proveedores. Crea el primero con «+ Proveedor».")}</td></tr>`}</tbody></table></div>
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
          const EU = (campo, o) => ed({ puede: R.puede(S.usuario, "usuarios"), tabla: "Usuarios", idCol: "Correo", id: u.Correo, campo, valor: u[campo] ?? "", ...o });
          return `<tr><td${EU("Nombre", { tipo: "text", requerido: true, etiqueta: "Nombre" })}><b>${esc(u.Nombre)}</b><div class="sub">${esc(u.Correo)}</div></td>
          <td${EU("Rol", { tipo: "multi", opciones: R.ROLES, valor: R.roles(u).join(", "), etiqueta: "Roles", validar: (v) => (lc(u.Correo) === lc(S.usuario.Correo) && !/Admin/.test(v) && R.tieneRol(S.usuario, "Admin") ? "No puedes quitarte el rol Admin a ti mismo." : "") })}>${R.roles(u).map((r) => `<span class="pill rol">${esc(r)}</span>`).join(" ") || `<span class="pill vencido">Sin rol</span>`}</td>
          <td>${R.tieneRol(u, "PM") ? (pm.length ? esc(pm.join(", ")) : `<span class="sub">Sin proyectos</span>`) : "—"}</td>
          <td>${R.tieneRol(u, "Lector") ? esc(u.Proyectos || "—") : "—"}</td><td${EU("Activo", { tipo: "select", opciones: ["Sí", "No"], etiqueta: "Activo", validar: (v) => (v === "No" && lc(u.Correo) === lc(S.usuario.Correo) ? "No puedes desactivarte a ti mismo." : "") })}>${esc(u.Activo)}</td>
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
  // Comprime con gzip cuando vale la pena (XML, PDF de texto…). Un .xlsx ya viene comprimido y se deja igual.
  async function gzip(bytes) {
    if (typeof CompressionStream === "undefined") return null;
    try { const out = new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"))).arrayBuffer()); return out.length < bytes.length * 0.85 ? out : null; } catch (e) { return null; }
  }
  async function gunzip(bytes) { return new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).arrayBuffer()); }
  async function guardarArchivo(id, bytes, nombre, tipo, etiqueta = "el archivo") {
    const comp = await gzip(bytes);
    if (comp) { bytes = comp; tipo = `gzip;${tipo || "application/octet-stream"}`; }
    const b64 = ACTAS.aBase64(bytes);
    const total = Math.max(1, Math.ceil(b64.length / TAM_PARTE));
    // Se sube, se verifica que estén todas las partes (una vez cada una) y, si no, se repite una vez.
    for (let intento = 1; intento <= 2; intento++) {
      await S.api.borrarArchivo(id);
      if (total === 1) await S.api.guardarParte({ ID_Archivo: id, Parte: 0, Total: 1, Nombre: nombre, Tipo: tipo || "application/octet-stream", Datos: b64 });
      else {
        // Las partes se pasan al panel (rápido) y se escriben en Excel de una sola vez, por lotes.
        for (let i = 0; i < total; i++) {
          if (i % 5 === 0) cargando(true, `Preparando ${etiqueta}… ${Math.round(((i + 1) / total) * 100)}%`);
          await S.api.bufferParte({ ID_Archivo: id, Parte: i, Total: total, Nombre: nombre, Tipo: tipo || "application/octet-stream", Datos: b64.slice(i * TAM_PARTE, (i + 1) * TAM_PARTE) });
        }
        cargando(true, `Guardando ${etiqueta} en Excel${intento > 1 ? " (reintento)" : ""}…`);
        await S.api.escribirPartes(id);
      }
      cargando(true, `Verificando ${etiqueta}…`);
      if (archivoCompleto(await S.api.partesArchivo(id), total)) { delete cacheArchivos[id]; return; }
    }
    throw new Error(`${etiqueta} no quedó completo en Excel. Intente subirlo de nuevo.`);
  }
  // Completo = las partes 0..total-1, cada una una sola vez y con el mismo total.
  const archivoCompleto = (partes, total) => partes.length === total && partes.every((x) => x.total === total) && new Set(partes.map((x) => x.parte)).size === total;
  const cacheArchivos = {};
  async function leerArchivo(id, avance) {
    if (cacheArchivos[id]) return cacheArchivos[id];
    const b64 = await leerBase64(id, avance);
    const primera = b64.primera;
    let bytes = ACTAS.deBase64(b64.datos), tipo = String(primera.Tipo || "");
    if (tipo.startsWith("gzip;")) { bytes = await gunzip(bytes); tipo = tipo.slice(5); }
    return (cacheArchivos[id] = { bytes, nombre: primera.Nombre, tipo });
  }
  // Lee todas las partes de un archivo; si faltan o están mezcladas, avisa con un error claro.
  async function leerBase64(id, avance) {
    const partes = S.api.partesArchivo ? await S.api.partesArchivo(id) : null;
    if (partes && !partes.length) throw new Error("no está guardado en el Excel.");
    if (partes && !archivoCompleto(partes, partes[0].total)) throw new Error("quedó incompleto cuando se guardó (faltan partes o hay partes repetidas). Vuelve a subirlo.");
    const primera = await S.api.leerParte(id, 0);
    let datos = primera.Datos;
    for (let i = 1; i < primera.Total; i++) { if (avance) avance(i + 1, primera.Total); const p = await S.api.leerParte(id, i); if (Number(p.Total) !== Number(primera.Total)) throw new Error("tiene partes de dos versiones distintas. Vuelve a subirlo."); datos += p.Datos; }
    return { datos, primera };
  }
  const esActaWord = (nombre) => /\.docx$/i.test(String(nombre || ""));
  async function subirActa(idSeg, bytes, nombre) { await guardarArchivo(idSeg, bytes, nombre, esActaWord(nombre) ? ACTAS.TIPO_DOCX : "application/pdf", "el acta"); }
  // Compromisos del acta que ya existen en el proyecto (mismo texto): no se vuelven a crear.
  const claveComp = (t) => normTxt(t).replace(/[^a-z0-9 ]/g, "").slice(0, 70);
  const yaExisteComp = (pid, c) => S.datos.Compromisos.some((x) => x.ID_Proyecto === pid && claveComp(x.Compromiso) === claveComp(c.Compromiso));
  const msgActa = (res) => (res.compromisos.length
    ? `Encontré ${res.compromisos.length} compromiso(s)${res.formato && res.formato !== "FSFB" ? ` (${res.formato === "lista" ? "en lista" : res.formato === "Word" ? "en el Word" : "en tabla"})` : ""}`
    : res.titulo ? "Encontré la sección de compromisos, pero no pude separar sus filas" : "No encontré compromisos en el acta");
  // Revisión de los compromisos leídos de un acta antes de agregarlos a la sesión.
  function revisarCompromisosActa(s, lista) {
    const filas = lista.map((c) => ({ ...c, dup: yaExisteComp(s.ID_Proyecto, c) }));
    const html = `<p class="sub">Revisa y ajusta antes de agregar. Los que ya existen en el proyecto vienen desmarcados.</p>
      <div class="ra-lista">${filas.map((c, i) => `<div class="ra-fila ${c.dup ? "dup" : ""}"><label class="check"><input type="checkbox" class="ra-ok" data-i="${i}" ${c.dup ? "" : "checked"}></label>
        <textarea class="ra-t" rows="2" data-i="${i}">${esc(c.Compromiso)}</textarea>
        <input class="ra-r" data-i="${i}" value="${esc(separarNombres(c.Responsable).join("; "))}" placeholder="Responsable(s)" data-recurso data-multi autocomplete="off">
        <input class="ra-f" type="date" data-i="${i}" value="${esc(c.Fecha_Compromiso || "")}" title="${esc(c.Fecha_Texto && !c.Fecha_Compromiso ? `En el acta: ${c.Fecha_Texto}` : "")}">
        ${c.dup ? `<span class="pill noaplica">Ya existe</span>` : ""}</div>`).join("")}</div>`;
    modal(`Compromisos del acta (${filas.length})`, [{ html }], {}, (fd, sel) => guardar(async () => {
      if (!sel.length) return;
      const nuevas = [];
      for (const [k, c] of sel.entries()) {
        const resp = await resolverResponsables(separarNombres(c.Responsable).map((n) => ({ Nombre: n })), c._nuevos);
        nuevas.push({ ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", s.ID_Proyecto, k), ID_Proyecto: s.ID_Proyecto, ID_Seguimiento: s.ID_Seguimiento,
          Compromiso: c.Compromiso, ...resp, Fecha_Compromiso: c.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo, Fecha_Creacion: R.hoyISO() });
      }
      await S.api.agregarFilas("Compromisos", nuevas);
    }, `${sel.length} compromiso(s) agregados`), null, {
      textoBoton: "Agregar compromisos",
      init: (m) => { if (typeof enlazarRecursos === "function") enlazarRecursos(m); },
      recoger: (m) => {
        const sel = [...m.querySelectorAll(".ra-ok:checked")].map((x) => { const i = x.dataset.i, q = (c) => m.querySelector(`.${c}[data-i="${i}"]`); const inp = q("ra-r");
          return { Compromiso: q("ra-t").value.trim(), Responsable: inp.value.trim(), Fecha_Compromiso: q("ra-f").value, _nuevos: inp._nuevos || new Set() }; }).filter((c) => c.Compromiso);
        return sel.length ? { datos: sel } : { error: "Marca al menos un compromiso (o usa Cancelar)." };
      } });
  }
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
      const { datos, primera } = await leerBase64(idSeg, (i, n) => { const v = $("#v-paginas"); if (v) v.innerHTML = `<div class="vacio">Abriendo el acta… ${i} de ${n}</div>`; });
      let bytes = ACTAS.deBase64(datos);
      if (primera && String(primera.Tipo || "").startsWith("gzip;")) bytes = await gunzip(bytes);
      const word = esActaWord(nombre);
      const d = $("#v-descargar");
      if (d) { d.href = URL.createObjectURL(new Blob([bytes], { type: word ? ACTAS.TIPO_DOCX : "application/pdf" })); d.download = nombre || (word ? "acta.docx" : "acta.pdf"); d.hidden = false; }
      if (word) { const v = $("#v-paginas"); if (v) v.innerHTML = `<div class="vacio">Esta acta es un documento de Word. Usa «Descargar» para abrirla en Word.</div>`; }
      else await ACTAS.mostrar(bytes, $("#v-paginas"));
    } catch (e) {
      const c = $("#v-paginas");
      const s = S.datos.Seguimientos.find((x) => x.ID_Seguimiento === idSeg);
      const puedeSubir = s && R.puedeEditar(S.usuario, proyecto(s.ID_Proyecto) || {}, "seguimiento");
      if (c) {
        c.innerHTML = `${vacio("No se pudo abrir el acta: el archivo " + e.message)}${puedeSubir ? `<div class="acciones" style="justify-content:center"><button class="btn primario" id="v-resubir">Volver a subir el acta (PDF)</button></div>` : ""}`;
        const b = $("#v-resubir"); if (b) b.addEventListener("click", () => adjuntarActa(idSeg, true));
      }
    }
  }
  function elegirPDF(alElegir) {
    const i = document.createElement("input");
    i.type = "file"; i.accept = "application/pdf,.pdf,.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    i.addEventListener("change", () => { if (i.files[0]) alElegir(i.files[0]); });
    i.click();
  }
  // Adjuntar el acta a una sesión ya registrada; si trae compromisos y la sesión no tiene, ofrece agregarlos.
  function adjuntarActa(idSeg, reemplazo) {
    const s = S.datos.Seguimientos.find((x) => x.ID_Seguimiento === idSeg);
    if (!s) return;
    elegirPDF(async (archivo) => {
      if (archivo.size > MAX_ACTA) { toast("El acta pesa más de 5 MB. Guarda una versión más liviana del PDF.", true); return; }
      cargando(true, "Leyendo el acta…");
      let res;
      try { res = await ACTAS.leerActa(archivo); } catch (e) { cargando(false); toast("No se pudo leer el acta: " + e.message, true); return; }
      cargando(false);
      await guardar(async () => {
        await subirActa(idSeg, res.datos, archivo.name);
        await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", idSeg, { Acta_Archivo: archivo.name, ...(s.Fecha_Acta ? {} : { Fecha_Acta: res.fecha }) });
      }, "Acta adjuntada");
      if (res.compromisos.length) revisarCompromisosActa(s, res.compromisos);
      else toast(`Acta adjuntada. ${msgActa(res)}.`);
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
      <div class="acciones derecha">${extra.eliminar ? `<button type="button" class="btn peligro-suave" id="m-eliminar">${esc(extra.eliminar.texto)}</button>` : ""}<button type="button" class="btn" id="m-cancelar">Cancelar</button><button type="submit" class="btn primario">${esc(extra.textoBoton || "Guardar")}</button></div></form></div>`;
    document.body.appendChild(m);
    $("#m-cerrar").addEventListener("click", cerrarModal);
    $("#m-cancelar").addEventListener("click", cerrarModal);
    if (extra.eliminar) $("#m-eliminar").addEventListener("click", () => confirmar(extra.eliminar.titulo || extra.eliminar.texto, extra.eliminar.mensaje, "Eliminar", extra.eliminar.accion));
    m.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarModal(); });
    // Enter en un campo auxiliar (crear rol, persona, proveedor…) ejecuta su propio botón, no guarda todo el formulario.
    m.addEventListener("keydown", (e) => {
      if (e.key !== "Enter" || e.target.tagName !== "INPUT" || e.target.name || e.defaultPrevented) return;
      e.preventDefault();
      const caja = e.target.closest(".eq-form, #pn-form, .prov-nuevo");
      const btn = caja && caja.querySelector(".btn.primario");
      if (btn && !e.target.matches("[data-recurso]")) btn.click();
    });
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

  // ---------- Cronograma del proyecto: importado de Planner (Excel) o MS Project (XML) ----------
  const esSi = (v) => v === true || /^(s[ií]|si|yes|true|1)$/i.test(String(v ?? "").trim());
  const cabNorm = (s) => normTxt(s).replace(/[^a-z0-9%]/g, "");
  // Fecha en cualquier formato habitual (Date, serial de Excel, dd/mm/aaaa, aaaa-mm-dd…) → aaaa-mm-dd.
  function aFechaISO(v, X) {
    if (v === null || v === undefined || v === "") return "";
    const d2 = (n) => String(n).padStart(2, "0");
    if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 12 * 3600e3); return `${d.getFullYear()}-${d2(d.getMonth() + 1)}-${d2(d.getDate())}`; }
    if (typeof v === "number" && X && X.SSF) { const o = X.SSF.parse_date_code(v); return o ? `${o.y}-${d2(o.m)}-${d2(o.d)}` : ""; }
    const t = String(v).trim();
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${d2(m[2])}-${d2(m[3])}`;
    m = t.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/);
    if (m) {
      let dd = Number(m[1]), mm = Number(m[2]); const yy = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
      if (mm > 12 && dd <= 12) [dd, mm] = [mm, dd];          // venía mes/día
      return `${yy}-${d2(mm)}-${d2(dd)}`;
    }
    return "";
  }
  const numPct = (v) => { const n = parseFloat(String(v ?? "").replace(",", ".").replace("%", "")); return isNaN(n) ? null : Math.max(0, Math.min(100, n <= 1 && /\./.test(String(v)) && !/%/.test(String(v)) ? n * 100 : n)); };

  // Planner (o Planner Premium / Project guardado como Excel): encabezados en cualquiera de las primeras 20 filas.
  const SINONIMOS_PLAN = {
    Nombre: ["nombredetarea", "nombredelatarea", "taskname", "nombre", "tarea", "name", "titulo", "title"],
    Grupo: ["nombredeldeposito", "deposito", "bucketname", "bucket", "cubo", "fase"],
    Progreso: ["progreso", "progress", "estado", "status"],
    Pct: ["%completado", "porcentajecompletado", "%complete", "percentcomplete", "completado", "avance", "%"],
    Responsable: ["asignadoa", "assignedto", "asignado", "responsable", "nombresdelosrecursos", "resourcenames", "recursos"],
    Inicio: ["fechadeinicio", "startdate", "inicio", "start", "comienzo"],
    Fin: ["fechadevencimiento", "duedate", "vencimiento", "fin", "finish", "finalizar", "finalizacion", "fechafin", "fechadefin"],
    Duracion: ["duracion", "duration"],
    Checklist: ["elementoscompletadosdelalistadecomprobacion", "completedchecklistitems"],
    Etiquetas: ["etiquetas", "labels"],
    EDT: ["edt", "wbs", "numerodeesquema", "outlinenumber"],
    Nivel: ["niveldeesquema", "outlinelevel", "nivel"],
    Hito: ["hito", "milestone"],
    Id: ["iddetarea", "iddelatarea", "taskid", "numerodetarea", "tasknumber", "id", "uid", "identificadorunico", "uniqueid"],
  };
  async function leerPlannerExcel(file) {
    const X = await libXlsx();
    // Fechas como número de serie (sin zona horaria): así una tarea que termina a las 5 p. m. no salta al día siguiente.
    const wb = X.read(new Uint8Array(await file.arrayBuffer()), { type: "array", cellDates: false });
    let hoja = null, filas = [], iCab = -1, mapa = null;
    for (const n of wb.SheetNames) {
      const f = X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" });
      const i = f.slice(0, 20).findIndex((r) => { const c = r.map(cabNorm); return SINONIMOS_PLAN.Nombre.some((s) => c.includes(s)) && ["Inicio", "Fin", "Progreso", "Pct"].some((k) => SINONIMOS_PLAN[k].some((s) => c.includes(s))); });
      if (i >= 0) { hoja = n; filas = f; iCab = i; break; }
    }
    if (iCab < 0) throw new Error("No encontré las columnas del plan (nombre de la tarea y fechas o progreso). Usa «Exportar plan a Excel» de Planner.");
    const cab = filas[iCab].map(cabNorm);
    mapa = Object.fromEntries(Object.entries(SINONIMOS_PLAN).map(([k, sin]) => [k, sin.map((s) => cab.indexOf(s)).find((i) => i >= 0) ?? -1]));
    const val = (r, k) => (mapa[k] >= 0 ? r[mapa[k]] : "");
    // Nombre del plan: Planner lo deja en las primeras filas («Nombre del plan»).
    const filaPlan = filas.slice(0, iCab).find((r) => /nombre del (plan|proyecto)|plan name|project name/i.test(String(r[0]))) || filas.slice(0, iCab).find((r) => /plan/i.test(String(r[0])));
    const crudas = filas.slice(iCab + 1).filter((r) => String(val(r, "Nombre")).trim());
    // % completado como fracción (0,29 = 29 %) cuando todos los valores numéricos están entre 0 y 1.
    const nums = crudas.map((r) => val(r, "Pct")).filter((v) => typeof v === "number");
    const fraccion = nums.length > 0 && nums.every((v) => v >= 0 && v <= 1);
    const durCero = (v) => { const t = String(v ?? "").trim(); return t !== "" && /^0+([.,]0+)?\s*(d|h|dí|di|hor|day|hour|sem|w|$)/i.test(t); };
    const tareas = crudas.map((r) => {
      const pv = val(r, "Pct");
      let pct = typeof pv === "number" && fraccion ? pv * 100 : numPct(pv);
      if (pct === null) {
        const pr = normTxt(val(r, "Progreso"));
        const ch = String(val(r, "Checklist")).match(/(\d+)\s*\/\s*(\d+)/);
        pct = /complet|finaliz|termin|done/.test(pr) ? 100 : /curso|progres|iniciad[ao]$/.test(pr) && !/^no/.test(pr) ? (ch && Number(ch[2]) ? Math.max(10, Math.round((Number(ch[1]) / Number(ch[2])) * 100)) : 50) : 0;
      }
      const nombre = String(val(r, "Nombre")).trim();
      const etiquetas = String(val(r, "Etiquetas"));
      const inicio = aFechaISO(val(r, "Inicio"), X), fin = aFechaISO(val(r, "Fin"), X);
      return { ID_Externo: String(val(r, "Id") || "").trim(), EDT: String(val(r, "EDT") || ""), Nivel: Number(val(r, "Nivel")) || 0, Nombre: nombre, Grupo: String(val(r, "Grupo") || "").trim(),
        Inicio: inicio || fin, Fin: fin || inicio, Pct_Completado: Math.round(pct),
        Hito: esSi(val(r, "Hito")) || durCero(val(r, "Duracion")) || /\bhito\b|milestone/i.test(etiquetas) || /^hito\b/i.test(nombre) ? "Sí" : "No",
        Resumen: "No", Responsable: String(val(r, "Responsable") || "").replace(/\s*;\s*/g, "; ").trim(), Inicio_Base: "", Fin_Base: "", Predecesoras: "" };
    });
    // Jerarquía: columna de nivel o, si no, la profundidad del número de esquema/EDT (1.2.3 = nivel 3), como en Planner Premium.
    // Solo si el archivo no trae ninguna de las dos se agrupa por depósito (Planner básico).
    if (!tareas.some((t) => t.Nivel) && tareas.filter((t) => /^\d+(\.\d+)*$/.test(t.EDT.trim())).length >= tareas.length * 0.8) tareas.forEach((t) => { t.Nivel = t.EDT.trim() ? t.EDT.trim().split(".").length : 1; });
    if (tareas.some((t) => t.Nivel)) {
      tareas.forEach((t, i) => { if (!t.Nivel) t.Nivel = 1; const sig = tareas[i + 1]; t.Resumen = sig && sig.Nivel > t.Nivel ? "Sí" : "No"; if (t.Resumen === "Sí") t.Hito = "No"; });
      return { origen: mapa.Grupo >= 0 || filaPlan ? "Planner" : "MS Project (Excel)", nombre: filaPlan ? String(filaPlan[1] || filaPlan[0]).trim() : hoja, tareas };
    }
    const grupos = [...new Set(tareas.map((t) => t.Grupo || "Sin depósito"))];
    const salida = [];
    grupos.forEach((g, gi) => {
      const hijas = tareas.filter((t) => (t.Grupo || "Sin depósito") === g).sort((a, b) => String(a.Inicio || "9").localeCompare(String(b.Inicio || "9")));
      salida.push({ EDT: String(gi + 1), Nivel: 1, Nombre: g, Grupo: g, Inicio: "", Fin: "", Pct_Completado: 0, Hito: "No", Resumen: "Sí", Responsable: "", Inicio_Base: "", Fin_Base: "", Predecesoras: "" });
      hijas.forEach((t, ti) => salida.push({ ...t, EDT: `${gi + 1}.${ti + 1}`, Nivel: 2 }));
    });
    resumirGrupos(salida);
    return { origen: "Planner", nombre: filaPlan ? String(filaPlan[1] || filaPlan[0]).trim() : hoja, tareas: salida };
  }
  // Fechas y % de las filas resumen a partir de sus hijas (ponderado por duración).
  function resumirGrupos(ts) {
    for (let i = ts.length - 1; i >= 0; i--) {
      const t = ts[i];
      if (t.Resumen !== "Sí") continue;
      const hijas = [];
      for (let j = i + 1; j < ts.length && Number(ts[j].Nivel) > Number(t.Nivel); j++) if (Number(ts[j].Nivel) === Number(t.Nivel) + 1) hijas.push(ts[j]);
      const fi = hijas.map((h) => h.Inicio).filter(Boolean).sort(), ff = hijas.map((h) => h.Fin).filter(Boolean).sort();
      if (!t.Inicio && fi.length) t.Inicio = fi[0];
      if (!t.Fin && ff.length) t.Fin = ff[ff.length - 1];
      const w = (h) => (h.Inicio && h.Fin ? Math.max(1, R.difDias(h.Fin, h.Inicio) + 1) : 1);
      const tw = hijas.reduce((a, h) => a + w(h), 0);
      if (tw) t.Pct_Completado = Math.round(hijas.reduce((a, h) => a + w(h) * (Number(h.Pct_Completado) || 0), 0) / tw);
    }
  }

  // MS Project guardado como XML (Archivo → Guardar como → XML).
  async function leerProjectXML(file) {
    const doc = new DOMParser().parseFromString(await file.text(), "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("El archivo XML está dañado o incompleto.");
    const raiz = doc.documentElement;
    if (!raiz || raiz.localName !== "Project") throw new Error("Este XML no es de MS Project. En Project usa Archivo → Guardar como → «Formato XML».");
    const hijos = (el, n) => (el ? [...el.children].filter((c) => c.localName === n) : []);
    const uno = (el, n) => hijos(el, n)[0];
    const v = (el, n) => { const c = uno(el, n); return c ? c.textContent.trim() : ""; };
    const recursos = Object.fromEntries(hijos(uno(raiz, "Resources"), "Resource").map((r) => [v(r, "UID"), v(r, "Name")]).filter(([, n]) => n));
    const asign = {};
    hijos(uno(raiz, "Assignments"), "Assignment").forEach((a) => { const n = recursos[v(a, "ResourceUID")]; if (n) (asign[v(a, "TaskUID")] = asign[v(a, "TaskUID")] || []).push(n); });
    const tareasXml = hijos(uno(raiz, "Tasks"), "Task").filter((t) => v(t, "IsNull") !== "1" && Number(v(t, "OutlineLevel")) > 0 && v(t, "Name"));
    const idPorUid = Object.fromEntries(tareasXml.map((t) => [v(t, "UID"), v(t, "ID")]));
    const tareas = tareasXml.map((t) => {
      const base = hijos(t, "Baseline").find((b) => v(b, "Number") === "0");
      return { ID_Externo: v(t, "UID"), EDT: v(t, "WBS") || v(t, "OutlineNumber"), Nivel: Number(v(t, "OutlineLevel")) || 1, Nombre: v(t, "Name"), Grupo: "",
        Inicio: v(t, "Start").slice(0, 10), Fin: v(t, "Finish").slice(0, 10), Pct_Completado: Number(v(t, "PercentComplete")) || 0,
        Hito: v(t, "Milestone") === "1" ? "Sí" : "No", Resumen: v(t, "Summary") === "1" ? "Sí" : "No",
        Responsable: (asign[v(t, "UID")] || []).join("; "),
        Inicio_Base: base ? v(base, "Start").slice(0, 10) : "", Fin_Base: base ? v(base, "Finish").slice(0, 10) : "",
        Predecesoras: hijos(t, "PredecessorLink").map((l) => idPorUid[v(l, "PredecessorUID")]).filter(Boolean).join(", ") };
    });
    // El nivel 1 de Project suele ser la tarea resumen del proyecto: se conserva (da el total).
    return { origen: "MS Project", nombre: v(raiz, "Title") || v(raiz, "Name") || file.name, tareas };
  }

  // Avance real y planeado ponderados por duración, sobre las tareas que no son resumen.
  function calcCronograma(tareas, hoy = R.hoyISO()) {
    const hojas = tareas.filter((t) => t.Resumen !== "Sí" && t.Inicio && t.Fin);
    const conBase = hojas.some((t) => t.Inicio_Base && t.Fin_Base);
    const dur = (a, b) => Math.max(1, R.difDias(b, a) + 1);
    const peso = (t) => (t.Hito === "Sí" ? 0 : dur(t.Inicio, t.Fin));
    const planDe = (t) => {
      const ini = (conBase && t.Inicio_Base) || t.Inicio, fin = (conBase && t.Fin_Base) || t.Fin;
      if (hoy >= fin) return 100;
      if (hoy < ini) return 0;
      return Math.min(100, (dur(ini, hoy) / dur(ini, fin)) * 100);
    };
    let tw = hojas.reduce((a, t) => a + peso(t), 0);
    const w = tw ? peso : () => 1;
    if (!tw) tw = hojas.length;
    const real = tw ? Math.round(hojas.reduce((a, t) => a + w(t) * (Number(t.Pct_Completado) || 0), 0) / tw) : 0;
    const plan = tw ? Math.round(hojas.reduce((a, t) => a + w(t) * planDe(t), 0) / tw) : 0;
    const atrasadas = hojas.filter((t) => (Number(t.Pct_Completado) || 0) < 100 && t.Fin < hoy);
    const fechas = tareas.filter((t) => t.Inicio && t.Fin);
    const inicio = fechas.map((t) => t.Inicio).sort()[0] || "", fin = fechas.map((t) => t.Fin).sort().pop() || "";
    return { real, plan, atrasadas, inicio, fin, conBase, n: tareas.length, hojas: hojas.length, hitos: tareas.filter((t) => t.Hito === "Sí").length };
  }
  const estadoTarea = (t, hoy = R.hoyISO()) => ((Number(t.Pct_Completado) || 0) >= 100 ? "ok" : t.Fin && t.Fin < hoy ? "atr" : t.Inicio && t.Inicio <= hoy ? "curso" : "pend");
  const TXT_EST_TAREA = { ok: "Completada", atr: "Atrasada", curso: "En curso", pend: "Por iniciar" };

  // Hitos del cronograma → Hitos del proyecto (por nombre): crea los nuevos y actualiza fechas/estado.
  function planHitos(p, tareas) {
    const hs = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto);
    const nuevos = [], cambios = [];
    tareas.filter((t) => t.Hito === "Sí").forEach((t) => {
      const fPlan = t.Fin_Base || t.Fin || t.Inicio;
      const listo = (Number(t.Pct_Completado) || 0) >= 100;
      const datos = { Fecha_Plan: fPlan, Fecha_Real: listo ? t.Fin : "", Estado: listo ? "Cumplido" : "Pendiente" };
      const h = hs.find((x) => normTxt(x.Hito) === normTxt(t.Nombre));
      if (!h) nuevos.push({ Hito: t.Nombre, ...datos });
      else {
        const dif = Object.fromEntries(Object.entries(datos).filter(([k, val]) => String(h[k] || "") !== String(val || "") && !(k === "Estado" && val === "Pendiente" && h.Estado === "Atrasado")));
        if (Object.keys(dif).length) cambios.push({ id: h.ID_Hito, cambios: dif });
      }
    });
    return { nuevos, cambios };
  }

  // ---------- Versiones del cronograma ----------
  // Cada importación es una versión (V1, V2…), con su fecha de corte y, si se quiere, la sesión de seguimiento a la que corresponde.
  // Cada versión guarda todas sus tareas y, desde la v81, también el archivo original (Planner o Project) para descargarlo.
  const vidLegado = (pid) => `CRV-${pid}-1`;
  const vidDe = (t) => t.ID_Version || vidLegado(t.ID_Proyecto);
  const versionesDe = (pid) => (S.datos.Cronogramas || []).filter((v) => v.ID_Proyecto === pid).sort((a, b) => Number(a.Numero) - Number(b.Numero));
  const versionVigente = (pid) => { const l = versionesDe(pid); return l[l.length - 1] || null; };
  function tareasDe(pid, vid) {
    const todas = (S.datos.Tareas || []).filter((t) => t.ID_Proyecto === pid);
    if (!todas.length) return [];
    const v = vid || (versionVigente(pid) || {}).ID_Version || vidLegado(pid);
    return todas.filter((t) => vidDe(t) === v).sort((a, b) => Number(a.Orden) - Number(b.Orden));
  }
  // Proyectos con cronograma importado antes de existir las versiones: se registra como V1.
  async function migrarCronogramas() {
    const pids = [...new Set((S.datos.Tareas || []).filter((t) => !t.ID_Version).map((t) => t.ID_Proyecto))].filter((pid) => !versionesDe(pid).length);
    if (!pids.length) return;
    const filas = pids.map((pid) => {
      const p = proyecto(pid) || {}, ts = (S.datos.Tareas || []).filter((t) => t.ID_Proyecto === pid), f = p.Cronograma_Fecha || R.hoyISO(), c = calcCronograma(ts, f);
      return { ID_Version: vidLegado(pid), ID_Proyecto: pid, Numero: 1, Fecha_Corte: f, Fecha_Importacion: f, Origen: p.Cronograma_Origen || (ts[0] || {}).Origen || "", Archivo: p.Cronograma_Archivo || "",
        ID_Seguimiento: "", Tareas: c.hojas, Avance_Real: c.real, Avance_Plan: c.plan, Inicio: c.inicio, Fin: c.fin, Atrasadas: c.atrasadas.length, Hitos: c.hitos, Nota: "Primera versión importada", Registrado_Por: p.Cronograma_Por || "" };
    });
    try { await S.api.agregarFilas("Cronogramas", filas); } catch (e) { /* se reintenta en la próxima carga */ }
  }

  // Lee el archivo (XML de Project o Excel de Planner) y devuelve { origen, nombre, tareas }.
  function elegirCronograma(alLeer) {
    const i = document.createElement("input");
    i.type = "file"; i.accept = ".xml,.xlsx,.xls,.csv,application/xml,text/xml";
    i.addEventListener("change", async () => {
      const f = i.files[0];
      if (!f) return;
      cargando(true, "Leyendo el cronograma…");
      let res;
      try { res = /\.xml$/i.test(f.name) ? await leerProjectXML(f) : await leerPlannerExcel(f); res.bytes = new Uint8Array(await f.arrayBuffer()); res.tipoArchivo = f.type || (/\.xml$/i.test(f.name) ? "application/xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); }
      catch (e) { cargando(false); toast("No se pudo leer el cronograma: " + e.message, true); return; }
      cargando(false);
      if (!res.tareas.length) { toast("El archivo no trae tareas.", true); return; }
      alLeer(res, f.name);
    });
    i.click();
  }
  const importarCronograma = (p) => elegirCronograma((res, archivo) => previsualizarCronograma(p, res, archivo));

  // Línea base del cronograma cuando el archivo no la trae (Planner / Excel): cada tarea conserva las fechas con las que
  // apareció por primera vez (se arrastran de versión en versión). modo: "archivo" | "mantener" | "nueva".
  function conLineaBase(pid, tareas, modo) {
    if (modo === "archivo") return tareas;
    if (modo === "nueva") return tareas.map((t) => ({ ...t, Inicio_Base: t.Inicio, Fin_Base: t.Fin }));
    const prev = tareasDe(pid);
    const porId = new Map(), porEdtNom = new Map(), porNom = new Map();
    prev.forEach((t) => { if (t.ID_Externo) porId.set(String(t.ID_Externo), t); porEdtNom.set(`${t.EDT}|${normTxt(t.Nombre)}`, t); if (!porNom.has(normTxt(t.Nombre))) porNom.set(normTxt(t.Nombre), t); });
    return tareas.map((t) => {
      const a = (t.ID_Externo && porId.get(String(t.ID_Externo))) || porEdtNom.get(`${t.EDT}|${normTxt(t.Nombre)}`) || porNom.get(normTxt(t.Nombre));
      const ib = a ? a.Inicio_Base || a.Inicio : t.Inicio, fb = a ? a.Fin_Base || a.Fin : t.Fin;
      return { ...t, Inicio_Base: ib || t.Inicio, Fin_Base: fb || t.Fin };
    });
  }
  const traeBase = (tareas) => tareas.some((t) => t.Inicio_Base && t.Fin_Base);

  // Guarda una versión nueva: agrega tareas y la fila de versión. Devuelve el cálculo.
  async function guardarVersionCronograma(p, res, archivo, op) {
    const pid = p.ID_Proyecto, hoy = R.hoyISO(), corte = op.corte || hoy;
    const lista = op.nivel ? res.tareas.filter((t) => (Number(t.Nivel) || 1) <= Number(op.nivel)) : res.tareas;
    const c = calcCronograma(res.tareas, corte);
    const vs = versionesDe(pid), num = vs.length ? Math.max(...vs.map((v) => Number(v.Numero) || 0)) + 1 : 1;
    let vid = `CRV-${pid}-${num}`;
    while ((S.datos.Cronogramas || []).some((v) => v.ID_Version === vid)) vid += "b";
    // Cada versión conserva TODAS sus tareas (antes las anteriores se compactaban a fases e hitos).
    const filas = lista.map((t, k) => ({ ID_Tarea: `${vid}-${k + 1}`, ID_Proyecto: pid, ID_Version: vid, ID_Externo: t.ID_Externo || "", Orden: k + 1, EDT: t.EDT, Nivel: Number(t.Nivel) || 1, Nombre: t.Nombre, Grupo: t.Grupo,
      Inicio: t.Inicio, Fin: t.Fin, Duracion_Dias: t.Inicio && t.Fin ? R.difDias(t.Fin, t.Inicio) + 1 : "", Pct_Completado: Number(t.Pct_Completado) || 0, Hito: t.Hito, Resumen: t.Resumen,
      Responsable: t.Responsable, Inicio_Base: t.Inicio_Base, Fin_Base: t.Fin_Base, Predecesoras: t.Predecesoras, Origen: res.origen, Fecha_Importacion: hoy }));
    for (let k = 0; k < filas.length; k += 100) await S.api.agregarFilas("Tareas", filas.slice(k, k + 100));
    // El archivo original (Planner .xlsx o Project .xml) se guarda en el libro para poder descargarlo después.
    let guardado = "";
    if (res.bytes && res.bytes.length) {
      try { await guardarArchivo(vid, res.bytes, archivo, res.tipoArchivo, "el archivo del cronograma"); guardado = "Sí"; }
      catch (e) { toast(`La versión se guardó, pero no el archivo original: ${e.message}`, true); }
    }
    await S.api.agregarFila("Cronogramas", { Archivo_Guardado: guardado, ID_Version: vid, ID_Proyecto: pid, Numero: num, Fecha_Corte: corte, Fecha_Importacion: hoy, Origen: res.origen, Archivo: archivo,
      ID_Seguimiento: op.idSeg || "", Tareas: c.hojas, Avance_Real: c.real, Avance_Plan: c.plan, Inicio: c.inicio, Fin: c.fin, Atrasadas: c.atrasadas.length, Hitos: c.hitos, Nota: op.nota || "", Registrado_Por: S.usuario.Correo });
    await S.api.actualizarPorId("Proyectos", "ID_Proyecto", pid, { Cronograma_Origen: res.origen, Cronograma_Archivo: archivo, Cronograma_Fecha: hoy, Cronograma_Por: S.usuario.Correo,
      ...(op.avance ? { Avance_Real: c.real, Avance_Planeado: c.plan } : {}), ...sello() });
    if (op.idSeg && op.avanceSesion) await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", op.idSeg, { Avance_Real: c.real, Avance_Planeado: c.plan });
    if (op.hitos) {
      const h = planHitos(p, res.tareas);
      if (h.cambios.length) await S.api.actualizarVarios("Hitos", "ID_Hito", h.cambios);
      if (h.nuevos.length) await S.api.agregarFilas("Hitos", h.nuevos.map((x, k) => ({ ID_Hito: R.siguienteIdHijo("HIT", S.datos.Hitos, "ID_Hito", pid, k), ID_Proyecto: pid, ...x })));
    }
    return { c, num, vid };
  }

  function previsualizarCronograma(p, res, archivo) {
    const pid = p.ID_Proyecto, hoy = R.hoyISO();
    const segs = R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse();
    const usadas = new Set(versionesDe(pid).map((v) => v.ID_Seguimiento).filter(Boolean));
    const cercana = segs.find((s) => Math.abs(R.difDias(hoy, s.Fecha_Corte)) <= 7 && !usadas.has(s.ID_Seguimiento));
    // Contra qué se calcula el planeado: la línea base del archivo (Project) o las fechas originales de cada tarea (Planner / Excel).
    const prevTs = tareasDe(pid), tieneArchivo = traeBase(res.tareas);
    const modoDef = tieneArchivo ? "archivo" : prevTs.length ? "mantener" : "nueva";
    const calcCon = (modo) => calcCronograma(conLineaBase(pid, res.tareas, modo), hoy);
    const c = calcCon(modoDef);
    const h = planHitos(p, res.tareas);
    const prev = versionVigente(pid);
    const num = prev ? Math.max(...versionesDe(pid).map((v) => Number(v.Numero) || 0)) + 1 : 1;
    const niveles = [...new Set(res.tareas.map((t) => Number(t.Nivel) || 1))].sort((a, b) => a - b);
    const porNivel = (n) => res.tareas.filter((t) => (Number(t.Nivel) || 1) <= n).length;
    const grande = res.tareas.length > 600 && niveles.includes(3);
    const finTarde = c.fin && p.Fecha_Fin_Plan && c.fin > p.Fecha_Fin_Plan;
    const corrimiento = prev && prev.Fin && c.fin ? R.difDias(c.fin, prev.Fin) : null;
    const muestra = res.tareas.slice(0, 10).map((t) => `<tr><td class="nowrap">${esc(t.EDT)}</td><td style="padding-left:${((Number(t.Nivel) || 1) - 1) * 12 + 6}px">${t.Resumen === "Sí" ? `<b>${esc(t.Nombre)}</b>` : `${t.Hito === "Sí" ? "◆ " : ""}${esc(t.Nombre)}`}</td>
      <td class="nowrap">${fecha(t.Inicio)}</td><td class="nowrap">${fecha(t.Fin)}</td><td>${Number(t.Pct_Completado) || 0}%</td></tr>`).join("");
    const antes = `<div class="crono-prev">
      <div class="crono-prev-cab"><span class="pill vnum">V${num}</span><span class="pill ${res.origen === "Planner" ? "planner" : "msproject"}">${esc(res.origen)}</span> <b>${esc(res.nombre || archivo)}</b> <span class="sub">${esc(archivo)}</span></div>
      <div class="crono-kpis"><div><span>Tareas</span><b>${c.hojas}</b></div><div><span>Desde · hasta</span><b>${fecha(c.inicio)} → ${fecha(c.fin)}</b></div>
        <div><span>Avance real</span><b>${c.real}%</b></div><div><span>Planeado a hoy</span><b id="cp-plan">${c.plan}%</b></div><div><span>Desviación</span><b id="cp-desv" class="${c.real - c.plan < -5 ? "texto-rojo" : ""}">${c.real - c.plan > 0 ? "+" : ""}${c.real - c.plan} pts</b></div><div><span>Atrasadas</span><b class="${c.atrasadas.length ? "texto-rojo" : ""}">${c.atrasadas.length}</b></div>
        ${prev ? `<div><span>Fin vs V${esc(prev.Numero)}</span><b class="${corrimiento > 0 ? "texto-rojo" : ""}">${corrimiento === null ? "—" : corrimiento === 0 ? "igual" : `${corrimiento > 0 ? "+" : ""}${corrimiento} días`}</b></div>` : ""}</div>
      <p class="sub" id="cp-base">${tieneArchivo ? "El planeado se calcula contra la línea base guardada en el archivo." : prevTs.length ? "El archivo no trae línea base: el planeado se calcula contra las fechas originales de cada tarea (las de la primera versión en que apareció), así los corrimientos se notan." : "Primera versión: sus fechas quedan como línea base; desde la próxima carga el planeado se mide contra ellas."} Proyecto hoy: real ${pct(p.Avance_Real)} · planeado ${pct(p.Avance_Planeado)}.${prev ? ` La V${esc(prev.Numero)} queda en el historial (con sus fases e hitos) para comparar.` : ""}</p>
      ${finTarde ? `<div class="aviso-seg">⚠ El cronograma termina el ${fecha(c.fin)}, después del fin planeado del proyecto (${fecha(p.Fecha_Fin_Plan)}). Si es un cambio aprobado, registra un control de cambios.</div>` : ""}
      <details class="crono-det"><summary>Ver las primeras tareas</summary><div class="tabla-scroll"><table class="no-orden no-recortar crono-muestra"><thead><tr><th>EDT</th><th>Tarea</th><th>Inicio</th><th>Fin</th><th>%</th></tr></thead><tbody>${muestra}</tbody></table></div></details></div>`;
    const campos = [
      { k: "Fecha_Corte", label: "Fecha de corte de esta versión", tipo: "date", def: cercana ? cercana.Fecha_Corte : hoy, ayuda: "El % planeado se calcula a esta fecha." },
      { k: "ID_Seguimiento", label: "Sesión de seguimiento", tipo: "select", opciones: segs.map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)} · ${pct(s.Avance_Real)}${usadas.has(s.ID_Seguimiento) ? " (ya tiene versión)" : ""}`]), textoVacio: "No asociar a una sesión", def: cercana ? cercana.ID_Seguimiento : "" },
      { k: "Base", label: "Calcular el planeado contra", tipo: "select", ancho: true, def: modoDef, opciones: [
        ...(tieneArchivo ? [["archivo", "La línea base guardada en el archivo"]] : []),
        ...(prevTs.length ? [["mantener", "Las fechas originales de cada tarea (versiones anteriores)"]] : []),
        ["nueva", prevTs.length ? "Re-planear: las fechas de esta versión son la nueva línea base" : "Las fechas de esta versión (primera línea base)"]] },
      { k: "Avance", label: "Avance del proyecto", tipo: "select", opciones: [["Sí", "Actualizar con el cronograma"], ["No", "No cambiar"]], def: "Sí" },
      { k: "Hitos", label: "Hitos", tipo: "select", opciones: [["Sí", c.hitos ? `Sincronizar (${h.nuevos.length} nuevo(s) · ${h.cambios.length} actualizado(s))` : "Sincronizar"], ["No", "No tocar los hitos"]], def: c.hitos ? "Sí" : "No" },
      ...(niveles.length > 1 ? [{ k: "Nivel", label: "Niveles a guardar", tipo: "select", textoVacio: `Todos (${res.tareas.length} tareas)`, opciones: niveles.slice(0, -1).map((n) => [String(n), `Hasta el nivel ${n} (${porNivel(n)} tareas)`]), def: grande ? "3" : "" }] : []),
      { k: "Nota", label: "Nota de la versión (opcional)", placeholder: "Ej.: Re-planeación aprobada en CC-003 · ajuste de fechas de pruebas", ancho: true },
    ];
    modal(`Nueva versión del cronograma · ${p.Nombre}`, campos, {}, (fd) => {
      let r;
      guardar(async () => {
        r = await guardarVersionCronograma(p, { ...res, tareas: conLineaBase(pid, res.tareas, fd.Base || modoDef) }, archivo, { corte: fd.Fecha_Corte, idSeg: fd.ID_Seguimiento, avance: fd.Avance === "Sí", avanceSesion: !!fd.ID_Seguimiento, hitos: fd.Hitos === "Sí", nivel: fd.Nivel, nota: fd.Nota });
        S.fichaTab = "gantt";
        const F = (S.ganttF || {})[pid]; if (F) { F.ver = ""; F.comp = ""; delete F.scroll; }
      }, `Cronograma V${num} guardado`);
    }, (fd) => (fd.Fecha_Corte && fd.Fecha_Corte > hoy ? "La fecha de corte no puede ser futura." : ""), { antes, init: (m) => {
      const sel = m.querySelector("[name=Base]");
      if (sel) sel.addEventListener("change", () => { const k = calcCon(sel.value), d = k.real - k.plan; m.querySelector("#cp-plan").textContent = k.plan + "%"; const e = m.querySelector("#cp-desv"); e.textContent = `${d > 0 ? "+" : ""}${d} pts`; e.className = d < -5 ? "texto-rojo" : ""; });
    } });
  }

  // ---------- Vista «Avance por EDT»: árbol de fases con avance real vs planeado al corte ----------
  // Verde oscuro = completado · verde = en plazo (≤ 5 pp) · amarillo = atraso 6-15 pp · rojo = atraso ≥ 16 pp · gris = no iniciado.
  function claseAvance(real, plan) {
    if (real >= 100) return "ed-comp";
    if (real === 0 && plan === 0) return "ed-noini";
    const d = real - plan;
    return d >= -5 ? "ed-ok" : d >= -15 ? "ed-warn" : "ed-bad";
  }
  const partirItems = (txt) => String(txt || "").split(/\n|•|·|;|(?:^|\s)-\s/).map((x) => x.trim()).filter(Boolean);
  function vistaEDT(p, ts, corte, conLado = true) {
    if (!ts.length) return vacio("Sin tareas.");
    // Avance de un nodo: sus tareas hoja (o él mismo si es hoja), ponderadas por duración.
    const hijosDe = (i) => { const n = Number(ts[i].Nivel) || 1, out = []; for (let j = i + 1; j < ts.length && (Number(ts[j].Nivel) || 1) > n; j++) out.push(j); return out; };
    const avance = (i) => {
      const desc = hijosDe(i).map((j) => ts[j]);
      const hojas = desc.length ? desc : [ts[i]];
      const c = calcCronograma(hojas.map((t) => ({ ...t, Resumen: desc.length && t.Resumen === "Sí" ? "Sí" : "No" })), corte);
      return { real: c.real, plan: c.plan };
    };
    const niv = (i) => Number(ts[i].Nivel) || 1;
    const minN = Math.min(...ts.map((t, i) => niv(i)));
    const raices = ts.map((t, i) => i).filter((i) => niv(i) === minN);
    // Una sola raíz (Project): raíz = esa tarea, fases = su nivel siguiente. Varias (Planner): raíz = el proyecto.
    let raizTxt, raizAv, fases, edtRaiz = "";
    if (raices.length === 1 && hijosDe(raices[0]).length) {
      const r = raices[0]; edtRaiz = String(ts[r].EDT || "");
      raizTxt = ts[r].Nombre; raizAv = avance(r);
      fases = hijosDe(r).filter((j) => niv(j) === minN + 1);
    } else {
      raizTxt = p.Nombre; raizAv = calcCronograma(ts, corte); fases = raices;
    }
    // Numeración relativa (1, 1.1, 1.1.1) como en el comité; el EDT real queda en el título.
    const num = (t) => { const e = String(t.EDT || ""); return edtRaiz && e.startsWith(edtRaiz) ? "1" + e.slice(edtRaiz.length) : e; };
    const dif = (a) => { const d = a.real - a.plan; return d === 0 ? "" : `<b class="ed-pp ${d > 0 ? "mas" : d < -5 ? "menos" : ""}">${d > 0 ? "+" : ""}${d} pp</b>`; };
    const caja = (t, a, cls = "") => `<div class="ed-caja ${claseAvance(a.real, a.plan)} ${cls}" title="${esc(`${t.EDT || ""} · ${t.Nombre}${t.Inicio ? ` · ${fecha(t.Inicio)} → ${fecha(t.Fin)}` : ""}${t.Responsable ? ` · ${t.Responsable}` : ""}`)}">
        <div class="ed-tit"><span class="ed-num">${esc(num(t))}</span> ${t.Hito === "Sí" ? "◆ " : ""}${esc(t.Nombre)}</div>
        <div class="ed-av">Real <b>${a.real}%</b> <span>/ ${a.plan}%</span> ${dif(a)}</div></div>`;
    const columnas = fases.map((i) => {
      const subs = hijosDe(i).filter((j) => niv(j) === niv(i) + 1);
      return `<div class="ed-col">${caja(ts[i], avance(i), "ed-fase")}${subs.length ? `<div class="ed-subs">${subs.map((j) => caja(ts[j], avance(j))).join("")}</div>` : ""}</div>`;
    }).join("");
    const ult = R.avancesDe(p.ID_Proyecto, S.datos.Seguimientos).slice(-1)[0];
    // Paneles de la última sesión: se puede agregar o quitar ítems aquí mismo (quedan en esa sesión de seguimiento).
    const puedeSeg = R.puedeEditar(S.usuario, p, "seguimiento");
    const itemsDe = (txt) => partirItems(txt);
    const TIT_SEG = { Logros: "Logros del periodo", Proximos_Pasos: "Próximos pasos", Bloqueos: "Bloqueos" };
    const panelSeg = (k) => {
      if (k === "Bloqueos" && (!ult || (!ult.Bloqueos && !puedeSeg))) return "";
      const items = ult ? itemsDe(ult[k]) : [];
      return `<div class="ed-panel ${k === "Bloqueos" ? "ed-bloq" : ""}"><div class="ed-panel-cab"><h3>${TIT_SEG[k]}</h3>${ult && puedeSeg ? `<button type="button" class="btn enlace" data-ed-editseg="${esc(ult.ID_Seguimiento)}">Editar sesión</button>` : ""}</div>
        ${items.length ? `<ul class="ed-items">${items.map((x, i) => `<li>${esc(x)}${ult && puedeSeg ? ` <button type="button" class="ed-quitar" data-ed-quitar="${k}" data-i="${i}" title="Quitar">✕</button>` : ""}</li>`).join("")}</ul>` : `<p class="sub">${ult ? "Sin registrar." : "Aún no hay sesiones de seguimiento."}</p>`}
        ${ult && puedeSeg ? `<form class="ed-agregar" data-ed-add="${k}"><input placeholder="${k === "Logros" ? "Agregar un logro…" : k === "Proximos_Pasos" ? "Agregar un próximo paso…" : "Agregar un bloqueo…"}" aria-label="Agregar"><button class="btn chico">+</button></form>` : ""}
        ${k === "Logros" ? `<div class="sub">${ult ? `Se guardan en la sesión del ${fecha(ult.Fecha_Corte)}.` : "Registra la primera sesión para empezar."}</div>${!ult && p.Estado === "Activo" && puedeSeg ? `<button type="button" class="btn chico primario" data-ed-seg>Registrar seguimiento</button>` : ""}` : ""}</div>`;
    };
    const lista = (txt) => { const l = String(txt || "").split(/\n|•|·|;|(?:^|\s)-\s/).map((x) => x.trim()).filter(Boolean); return l.length ? `<ul>${l.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : `<p class="sub">Sin registrar.</p>`; };
    const equipo = [["Gerente de proyecto", nombreUsuario(p.PM)], ...stakeholdersDe(p.ID_Proyecto).map((x) => [x.Rol, x.Nombre])].filter(([, n]) => n);
    return `<div class="ed-wrap ${conLado ? "" : "sin-lado"} ${fases.length > 5 ? "muchas" : ""}">
      <div class="ed-arbol">
        <div class="ed-raiz">${caja({ EDT: edtRaiz, Nombre: raizTxt }, raizAv, "ed-root")}</div>
        <div class="ed-fases" style="grid-template-columns:repeat(${Math.max(1, fases.length)},minmax(128px,1fr));--ed-l:${(100 / (2 * Math.max(1, fases.length))).toFixed(2)}%">${columnas}</div>
        <div class="ed-leyenda"><span><i class="ed-comp"></i>Completado</span><span><i class="ed-ok"></i>En plazo (≤ 5 pp)</span><span><i class="ed-warn"></i>Atraso 6-15 pp</span><span><i class="ed-bad"></i>Atraso ≥ 16 pp</span><span><i class="ed-noini"></i>No iniciado</span><span class="sub">Real / plan al corte ${fecha(corte)}</span></div>
      </div>
      ${conLado ? `<aside class="ed-lado">
        ${["Logros", "Proximos_Pasos", "Bloqueos"].map((k) => panelSeg(k)).join("")}
        <div class="ed-panel"><h3>Equipo</h3>${equipo.map(([r, n]) => `<div><b>${esc(r)}:</b> ${esc(n)}</div>`).join("")}</div>
      </aside>` : ""}</div>`;
  }

  // ---------- Diagrama de Gantt ----------
  const ESCALAS_GT = { ajustar: { t: "Ajustar" }, semanas: { px: 7, t: "Semanas" }, meses: { px: 2.4, t: "Meses" }, trimestres: { px: 0.9, t: "Trimestres" } };
  const VISTAS_GT = { todo: "Todo", activas: "En curso y próximas", atr: "Atrasadas", hitos: "Hitos" };
  const MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  // Tarea equivalente en otra versión (por EDT y, si no, por nombre).
  const indiceVersion = (ts) => { const m = new Map(); ts.forEach((t) => { if (t.EDT) m.set("e:" + t.EDT, t); m.set("n:" + normTxt(t.Nombre), t); }); return (t) => m.get("e:" + t.EDT) || m.get("n:" + normTxt(t.Nombre)); };

  function seccionGantt(p, puede) {
    const pid = p.ID_Proyecto;
    const vs = versionesDe(pid), vig = versionVigente(pid);
    const ayuda = `<div class="crono-ayuda"><div><b>Planner</b><span>En el plan: «…» (más opciones) → <b>Exportar plan a Excel</b>.</span></div>
      <div><b>MS Project</b><span>Archivo → Guardar como → tipo <b>Formato XML (*.xml)</b>.</span></div>
      <div><b>Versiones</b><span>Cada carga queda como una versión (V1, V2…) asociada a su sesión de seguimiento. También se carga desde «Registrar seguimiento».</span></div></div>`;
    if (!tareasDe(pid).length) return `<div class="card"><div class="titulo-fila"><h2>Cronograma</h2>${puede ? `<button class="btn primario" id="b-crono-imp">⇪ Cargar cronograma</button>` : ""}</div>
      ${vacio("Este proyecto aún no tiene cronograma. Cárgalo desde Planner o MS Project: el avance real y planeado se calcularán solos.")}${ayuda}</div>`;
    const hoy = R.hoyISO();
    const F = (S.ganttF = S.ganttF || {})[pid] = (S.ganttF[pid] || { nivel: "", vista: "todo", q: "", cerr: [], escala: "ajustar", ver: "", comp: "", modo: "edt" });
    const verV = (F.ver && vs.find((v) => v.ID_Version === F.ver)) || vig;
    const esVig = !verV || !vig || verV.ID_Version === vig.ID_Version;
    const ts = tareasDe(pid, verV && verV.ID_Version);
    // Fecha de estado: por defecto hoy (versión vigente) o el corte de la versión; se puede mover para ver cómo cambia el planeado.
    const corteBase = esVig ? hoy : verV.Fecha_Corte || hoy;
    const corte = F.estado || corteBase, simulado = !!F.estado && F.estado !== corteBase;
    const etqCorte = simulado ? "Estado" : esVig ? "Hoy" : "Corte";
    const c = calcCronograma(ts, corte);
    const compV = F.comp && F.comp !== "lb" ? vs.find((v) => v.ID_Version === F.comp) : null;
    const buscarComp = compV ? indiceVersion(tareasDe(pid, compV.ID_Version)) : null;
    const niveles = [...new Set(ts.map((t) => Number(t.Nivel) || 1))].sort((a, b) => a - b);
    const finTarde = esVig && c.fin && p.Fecha_Fin_Plan && c.fin > p.Fecha_Fin_Plan;
    const viejo = esVig && p.Cronograma_Fecha && R.difDias(hoy, p.Cronograma_Fecha) > 15 && p.Estado === "Activo";
    const difAv = esVig && !simulado && (Math.abs((Number(p.Avance_Real) || 0) - c.real) >= 3 || Math.abs((Number(p.Avance_Planeado) || 0) - c.plan) >= 3);
    const conResp = ts.some((t) => t.Responsable);
    const filtrando = F.vista !== "todo" || F.q;
    const en28 = R.sumarDias(corte, 28);
    const pasaVista = (t) => {
      const e = estadoTarea(t, corte);
      if (F.vista === "atr") return e === "atr";
      if (F.vista === "hitos") return t.Hito === "Sí";
      if (F.vista === "activas") return t.Resumen !== "Sí" && e !== "ok" && t.Inicio && t.Inicio <= en28;
      return true;
    };
    const cerr = new Set(F.cerr.map(String));
    const lista = [];
    let ocultarHasta = null;
    ts.forEach((t) => {
      const n = Number(t.Nivel) || 1;
      if (ocultarHasta !== null && !filtrando) { if (n > ocultarHasta) return; ocultarHasta = null; }
      if (F.nivel && n > Number(F.nivel) && !filtrando) return;
      if (!pasaVista(t)) return;
      if (F.q && !normTxt(`${t.Nombre} ${t.Responsable} ${t.EDT}`).includes(normTxt(F.q))) return;
      lista.push(t);
      if (t.Resumen === "Sí" && cerr.has(String(t.Orden))) ocultarHasta = n;
    });
    // Rango de la línea de tiempo: el de las filas que se ven (y hoy, si el proyecto está activo).
    const fechasVis = lista.flatMap((t) => [t.Inicio, t.Fin, F.comp === "lb" ? t.Inicio_Base : "", F.comp === "lb" ? t.Fin_Base : ""]).filter(Boolean).sort();
    let r0 = fechasVis[0] || c.inicio || hoy, r1 = fechasVis[fechasVis.length - 1] || c.fin || hoy;
    if (p.Estado === "Activo" && esVig) { if (hoy < r0) r0 = hoy; if (hoy > r1) r1 = hoy; }
    const spanDias = Math.max(14, R.difDias(r1, r0) + 1);
    const pad = Math.max(3, Math.round(spanDias * 0.03));
    const ini = R.sumarDias(r0, -pad), fin = R.sumarDias(r1, pad + 1);
    const dias = R.difDias(fin, ini) + 1;
    const anchoDisp = Math.max(320, S.ganttAncho || 640);
    const pxDia = F.escala === "ajustar" ? anchoDisp / dias : ESCALAS_GT[F.escala].px;
    const W = Math.max(anchoDisp, Math.round(dias * pxDia));
    const x = (f) => Math.round(R.difDias(f, ini) * pxDia);
    const unDia = Math.max(1, Math.round(pxDia));
    // Encabezado de dos niveles según la densidad.
    const meses = mesesEntre(ini.slice(0, 7), fin.slice(0, 7));
    const celda = (a, b, txt, cls = "") => { a = Math.max(0, a); b = Math.min(W, b); const w = b - a; return w > 0 ? `<span class="${cls}" style="left:${a}px;width:${w}px">${w >= 24 ? `<em>${txt}</em>` : ""}</span>` : ""; };
    const anchoMes = 30.4 * pxDia;
    const lineas = [];
    let arriba = "", abajo = "";
    const anios = [...new Set(meses.map((m) => m.slice(0, 4)))];
    if (anchoMes >= 160) {
      arriba = meses.map((ym) => celda(x(`${ym}-01`), x(finDeMes(ym)) + unDia, `${MES_CORTO[Number(ym.slice(5, 7)) - 1]} ${ym.slice(0, 4)}`)).join("");
      let d = ini; while (new Date(d + "T12:00").getDay() !== 1) d = R.sumarDias(d, 1);
      for (; d <= fin; d = R.sumarDias(d, 7)) { abajo += celda(x(d), x(d) + Math.round(7 * pxDia), String(Number(d.slice(8, 10)))); lineas.push(x(d)); }
    } else if (anchoMes >= 34) {
      arriba = anios.map((y) => celda(x(`${y}-01-01`), x(`${y}-12-31`) + unDia, y)).join("");
      abajo = meses.map((ym) => { lineas.push(x(`${ym}-01`)); return celda(x(`${ym}-01`), x(finDeMes(ym)) + unDia, anchoMes >= 60 ? MES_CORTO[Number(ym.slice(5, 7)) - 1] : MES_CORTO[Number(ym.slice(5, 7)) - 1].charAt(0).toUpperCase()); }).join("");
    } else {
      arriba = anios.map((y) => celda(x(`${y}-01-01`), x(`${y}-12-31`) + unDia, y)).join("");
      const trims = [...new Set(meses.map((ym) => `${ym.slice(0, 4)}-${Math.floor((Number(ym.slice(5, 7)) - 1) / 3)}`))];
      abajo = trims.map((k) => { const [y, q] = k.split("-").map(Number), a = x(`${y}-${String(q * 3 + 1).padStart(2, "0")}-01`); lineas.push(a); return celda(a, x(finDeMes(`${y}-${String(q * 3 + 3).padStart(2, "0")}`)) + unDia, `T${q + 1}`); }).join("");
    }
    const xh = x(corte);
    const verHoy = xh >= 0 && xh <= W;
    const grid = lineas.filter((l) => l > 0 && l < W).map((l) => `transparent ${l}px, #e6ebf1 ${l}px, #e6ebf1 ${l + 1}px, transparent ${l + 1}px`).join(", ");
    const fondo = [verHoy ? `linear-gradient(90deg, transparent ${xh - 1}px, #e08a00 ${xh - 1}px, #e08a00 ${xh + 1}px, transparent ${xh + 1}px)` : "", grid ? `linear-gradient(90deg, transparent 0, ${grid}, transparent ${W}px)` : ""].filter(Boolean).join(", ") || "none";
    const der = [];
    const filas = lista.map((t) => {
      const est = estadoTarea(t, corte), n = Number(t.Nivel) || 1, pc = Number(t.Pct_Completado) || 0, res = t.Resumen === "Sí";
      let barra = "", difTxt = "";
      // Barra de comparación (línea base del archivo u otra versión), debajo.
      const ref = F.comp === "lb" ? (t.Inicio_Base && t.Fin_Base ? { Inicio: t.Inicio_Base, Fin: t.Fin_Base } : null) : buscarComp ? buscarComp(t) : null;
      if (ref && ref.Inicio && ref.Fin) {
        if (t.Hito === "Sí") barra += `<span class="gt-hito gt-hito-ref" style="left:${x(ref.Fin)}px" title="${F.comp === "lb" ? "Línea base" : `V${compV.Numero}`}: ${fecha(ref.Fin)}"></span>`;
        else { const a = x(ref.Inicio), w = Math.max(3, x(ref.Fin) + unDia - a); barra += `<span class="gt-base" style="left:${a}px;width:${w}px" title="${F.comp === "lb" ? "Línea base" : `V${compV.Numero}`}: ${fecha(ref.Inicio)} → ${fecha(ref.Fin)}"></span>`; }
        const dd = t.Fin && ref.Fin ? R.difDias(t.Fin, ref.Fin) : 0;
        if (dd) difTxt = `<span class="gt-dif ${dd > 0 ? "mal" : "bien"}" title="Fin frente a ${F.comp === "lb" ? "la línea base" : `la V${compV.Numero}`}">${dd > 0 ? "+" : ""}${dd} d</span>`;
      }
      if (t.Inicio && t.Fin) {
        const tip = `${t.Nombre} · ${fecha(t.Inicio)} → ${fecha(t.Fin)} · ${pc}%${t.Responsable ? ` · ${t.Responsable}` : ""}`;
        if (t.Hito === "Sí") barra += `<span class="gt-hito gt-${est}" style="left:${x(t.Fin)}px" title="${esc(tip)}"></span>`;
        else {
          const a = x(t.Inicio), w = Math.max(5, x(t.Fin) + unDia - a);
          barra += `<span class="gt-barra gt-${est} ${res ? "gt-resumen" : ""}" style="left:${a}px;width:${w}px" title="${esc(tip)}"><i style="width:${pc}%"></i>${!res && w >= 44 && pc > 0 ? `<b>${pc}%</b>` : ""}</span>`;
        }
      }
      const abierto = !cerr.has(String(t.Orden));
      const tog = res && !filtrando ? `<button type="button" class="gt-tog" data-gt-tog="${esc(t.Orden)}" aria-label="${abierto ? "Contraer" : "Expandir"}">${abierto ? "▾" : "▸"}</button>` : `<span class="gt-tog-v"></span>`;
      const cls = `gt-n${Math.min(n, 4)} ${res ? "gt-res" : ""} ${est === "atr" && !res ? "gt-fila-atr" : ""}`;
      der.push(`<div class="gt-r ${cls}">${barra}</div>`);
      return `<tr class="gt-fila ${cls}">
        <td class="gt-c1" title="${esc(`${t.EDT} · ${t.Nombre}${t.Responsable ? ` · ${t.Responsable}` : ""}`)}"><div class="gt-nom" style="padding-left:${filtrando ? 0 : (n - 1) * 14}px">${tog}<span class="gt-edt">${esc(t.EDT)}</span>${t.Hito === "Sí" ? `<span class="gt-rombo">◆</span>` : ""}<span class="gt-txt">${esc(t.Nombre)}</span></div></td>
        ${conResp ? `<td class="gt-c2"><span class="gt-resp" title="${esc(t.Responsable || "")}">${esc(t.Responsable || "")}</span></td>` : ""}
        <td class="gt-c3">${fecha(t.Inicio)}</td><td class="gt-c4">${fecha(t.Fin)}</td>${F.comp ? `<td class="gt-c6">${difTxt}</td>` : ""}
        <td class="gt-c5"><span class="gt-est gt-est-${est}">${t.Hito === "Sí" ? (pc >= 100 ? "✓" : "◆") : `${pc}%`}</span></td></tr>`;
    }).join("");
    const resumenes = ts.filter((t) => t.Resumen === "Sí").length;
    const nVista = (k) => (k === "todo" ? ts.length : ts.filter((t) => { const e = estadoTarea(t, corte); return k === "atr" ? e === "atr" : k === "hitos" ? t.Hito === "Sí" : t.Resumen !== "Sí" && e !== "ok" && t.Inicio && t.Inicio <= en28; }).length);
    const optVer = vs.slice().reverse().map((v) => `<option value="${esc(v.ID_Version)}" ${verV && v.ID_Version === verV.ID_Version ? "selected" : ""}>V${esc(v.Numero)} · ${fecha(v.Fecha_Corte)}${v === vig ? " (vigente)" : ""}</option>`).join("");
    const optComp = `<option value="">Sin comparar</option>${ts.some((t) => t.Inicio_Base) ? `<option value="lb" ${F.comp === "lb" ? "selected" : ""}>Línea base del cronograma</option>` : ""}${vs.filter((v) => !verV || v.ID_Version !== verV.ID_Version).slice().reverse().map((v) => `<option value="${esc(v.ID_Version)}" ${F.comp === v.ID_Version ? "selected" : ""}>V${esc(v.Numero)} · ${fecha(v.Fecha_Corte)}</option>`).join("")}`;
    return `<div class="card"><div class="titulo-fila"><h2>Cronograma ${verV ? `<span class="pill vnum">V${esc(verV.Numero)}${esVig ? " vigente" : ""}</span>` : ""} <span class="pill ${(verV ? verV.Origen : p.Cronograma_Origen) === "Planner" ? "planner" : "msproject"}">${esc((verV && verV.Origen) || p.Cronograma_Origen || "Importado")}</span></h2>
        <div class="acciones">${puede ? `<button class="btn primario" id="b-crono-imp">⇪ Cargar nueva versión</button>` : ""}</div></div>
      <p class="sub">${verV ? `Corte ${fecha(verV.Fecha_Corte)}${verV.ID_Seguimiento ? ` · sesión del ${fecha((S.datos.Seguimientos.find((s) => s.ID_Seguimiento === verV.ID_Seguimiento) || {}).Fecha_Corte)}` : ""} · cargado el ${fecha(verV.Fecha_Importacion)}${verV.Registrado_Por ? ` por ${esc(nombreUsuario(verV.Registrado_Por))}` : ""}${verV.Archivo ? ` · ${esc(verV.Archivo)}` : ""}${verV.Nota ? ` · <i>${esc(verV.Nota)}</i>` : ""}` : ""}${!esVig ? ` · <b>Versión anterior: solo conserva fases e hitos.</b>` : ""}</p>
      <div class="crono-kpis"><div><span>Avance real</span><b>${c.real}%</b></div><div class="${simulado ? "crono-sim" : ""}"><span>Planeado al ${fecha(corte)}</span><b>${c.plan}%</b></div>
        <div><span>Desviación</span><b class="${c.real - c.plan < -5 ? "texto-rojo" : ""}">${c.real - c.plan > 0 ? "+" : ""}${c.real - c.plan} pts</b></div>
        <div><span>Fin del cronograma</span><b class="${finTarde ? "texto-rojo" : ""}">${fecha(c.fin)}</b></div>
        <div><span>Atrasadas</span><b class="${c.atrasadas.length ? "texto-rojo" : ""}">${c.atrasadas.length}</b></div><div><span>Versiones</span><b>${vs.length || 1}</b></div></div>
      ${viejo ? `<div class="aviso-seg">⚠ El cronograma tiene ${R.difDias(hoy, p.Cronograma_Fecha)} días sin actualizar. Carga una nueva versión con el próximo seguimiento.</div>` : ""}
      ${finTarde ? `<div class="aviso-seg">⚠ El cronograma termina el ${fecha(c.fin)}, después del fin planeado (${fecha(p.Fecha_Fin_Plan)}).${puede ? ` <button class="btn enlace" id="b-crono-cc">Registrar control de cambios</button>` : ""}</div>` : ""}
      ${difAv && puede ? `<div class="aviso-seg info">El avance del proyecto (${pct(p.Avance_Real)} / ${pct(p.Avance_Planeado)}) no coincide con el cronograma (${c.real}% / ${c.plan}%). <button class="btn enlace" id="b-crono-av">Actualizar avance</button></div>` : ""}
      <div class="gt-modos"><div class="gt-seg gt-seg-grande" role="tablist">${[["edt", "📊 Avance por EDT"], ["gantt", "▤ Gantt"]].map(([k, t]) => `<button type="button" class="${(F.modo || "edt") === k ? "activo" : ""}" data-gt-modo="${k}">${t}</button>`).join("")}</div>
        ${vs.length > 1 ? `<label>Versión <select id="gt-ver">${optVer}</select></label>` : ""}
        <label title="Cambia la fecha para ver cuánto debería llevar el proyecto en ese momento. No se guarda.">Fecha de estado <input type="date" id="gt-estado" value="${esc(corte)}"></label>
        ${simulado ? `<button type="button" class="btn chico" id="gt-estado-x">↺ Volver a ${esVig ? "hoy" : "el corte"}</button><span class="pill crono-sim-pill">Simulación: el planeado y los atrasos se calculan al ${fecha(corte)}</span>` : ""}</div>
      ${(F.modo || "edt") === "edt" ? vistaEDT(p, ts, corte) : `
      <div class="gt-barra-h">
        <div class="gt-chips" role="group" aria-label="Qué mostrar">${Object.entries(VISTAS_GT).map(([k, v]) => `<button type="button" class="${F.vista === k ? "activo" : ""}" data-gt-vista="${k}">${v} <span>${nVista(k)}</span></button>`).join("")}</div>
        <input id="gt-q" type="search" placeholder="Buscar tarea, responsable o EDT" value="${esc(F.q)}" aria-label="Buscar">
      </div>
      <div class="gt-barra-h gt-barra-2">
        <label>Escala <span class="gt-seg">${Object.entries(ESCALAS_GT).map(([k, v]) => `<button type="button" class="${F.escala === k ? "activo" : ""}" data-gt-esc="${k}">${v.t}</button>`).join("")}</span></label>
        ${niveles.length > 1 ? `<label>Niveles <select id="gt-nivel"><option value="">Todos</option>${niveles.slice(0, -1).map((n) => `<option value="${n}" ${String(F.nivel) === String(n) ? "selected" : ""}>Hasta ${n}</option>`).join("")}</select></label>` : ""}
        <label>Comparar con <select id="gt-comp">${optComp}</select></label>
        <span class="gt-acc">${resumenes && !filtrando ? `<button type="button" class="btn chico" id="gt-exp" title="Expandir todo">⊞</button><button type="button" class="btn chico" id="gt-con" title="Contraer todo">⊟</button>` : ""}${verHoy && F.escala !== "ajustar" ? `<button type="button" class="btn chico" id="gt-hoy">◎ Hoy</button>` : ""}</span>
      </div>
      ${filas ? `<div class="gt-cont" id="gt-cont">
        <div class="gt-grid">
          <table class="no-orden no-recortar gt-izq" style="width:${270 + (conResp ? 110 : 0) + 82 + 82 + 48 + (F.comp ? 58 : 0)}px"><thead><tr><th class="gt-c1">Tarea</th>${conResp ? `<th class="gt-c2">Responsable</th>` : ""}<th class="gt-c3">Inicio</th><th class="gt-c4">Fin</th>${F.comp ? `<th class="gt-c6" title="Diferencia de la fecha fin frente a la referencia">Δ fin</th>` : ""}<th class="gt-c5">%</th></tr></thead><tbody>${filas}</tbody></table>
          <div class="gt-der" id="gt-der">
            <div class="gt-cabw" id="gt-cabw"><div class="gt-tl" style="width:${W}px"><div class="gt-cab gt-cab1"><div class="gt-escala">${arriba}</div></div><div class="gt-cab gt-cab2"><div class="gt-escala">${abajo}${verHoy ? `<span class="gt-hoy-cab" style="left:${xh}px">${etqCorte}</span>` : ""}</div></div></div></div>
            <div class="gt-cuerpo" id="gt-cuerpo" data-hoy="${verHoy ? xh : -1}"><div class="gt-tl" style="width:${W}px;--gt-w:${W}px;--gt-fondo:${fondo}">${der.join("")}</div></div>
          </div>
        </div>
        <div class="gt-hbar" id="gt-hbar" ${W <= anchoDisp + 2 ? "hidden" : ""}><div style="width:${W}px;height:1px"></div></div></div>
        <div class="gt-leyenda"><span><i class="gt-l-ok"></i>Completada</span><span><i class="gt-l-curso"></i>En curso</span><span><i class="gt-l-atr"></i>Atrasada</span><span><i class="gt-l-pend"></i>Por iniciar</span><span><i class="gt-l-res"></i>Fase / resumen</span><span><i class="gt-l-hito"></i>Hito</span>${F.comp ? `<span><i class="gt-l-base"></i>${F.comp === "lb" ? "Línea base" : `V${esc(compV ? compV.Numero : "")}`}</span>` : ""}<span><i class="gt-l-hoy"></i>${etqCorte}</span><span class="sub">Arrastra la línea de tiempo para moverte.</span></div>`
      : vacio("Ninguna tarea coincide con el filtro.")}
      `}
    </div>
    ${seccionVersiones(p, puede)}`;
  }

  // Historial de versiones: tabla + evolución del avance y del fin proyectado.
  function seccionVersiones(p, puede) {
    const vs = versionesDe(p.ID_Proyecto);
    if (!vs.length) return "";
    const filas = vs.slice().reverse().map((v, i, arr) => {
      const ant = arr[i + 1];
      const corr = ant && ant.Fin && v.Fin ? R.difDias(v.Fin, ant.Fin) : null;
      const seg = v.ID_Seguimiento && S.datos.Seguimientos.find((s) => s.ID_Seguimiento === v.ID_Seguimiento);
      const desv = (Number(v.Avance_Real) || 0) - (Number(v.Avance_Plan) || 0);
      return `<tr><td><span class="pill vnum">V${esc(v.Numero)}</span>${i === 0 ? ` <span class="pill aprobado">vigente</span>` : ""}</td><td class="nowrap"${ed({ puede, tabla: "Cronogramas", idCol: "ID_Version", id: v.ID_Version, campo: "Fecha_Corte", tipo: "date", requerido: true, valor: v.Fecha_Corte, validar: (x) => (x > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""), etiqueta: "Fecha de corte" })}>${fecha(v.Fecha_Corte)}</td>
        <td>${seg ? `Sesión del ${fecha(seg.Fecha_Corte)}` : `<span class="sub">—</span>`}</td>
        <td class="nowrap"><b>${pct(v.Avance_Real)}</b> / ${pct(v.Avance_Plan)} <span class="${desv < -5 ? "texto-rojo" : "sub"}">(${desv > 0 ? "+" : ""}${desv})</span></td>
        <td class="nowrap">${fecha(v.Fin)}${corr ? ` <span class="gt-dif ${corr > 0 ? "mal" : "bien"}">${corr > 0 ? "+" : ""}${corr} d</span>` : ""}</td>
        <td>${Number(v.Atrasadas) ? `<span class="texto-rojo">${esc(v.Atrasadas)}</span>` : "0"}</td>
        <td><span class="sub">${esc(v.Origen || "")}${v.Archivo ? ` · ${v.Archivo_Guardado === "Sí" ? `<button type="button" class="btn enlace crv-arch" data-crv-arch="${esc(v.ID_Version)}" title="Descargar el archivo original de esta versión">⬇ ${esc(v.Archivo)}</button>` : `${esc(v.Archivo)} <span class="sub" title="Esta versión se cargó antes de que la app guardara el archivo original">(archivo no guardado)</span>`}` : ""}</span>${tareasDe(p.ID_Proyecto, v.ID_Version).some((t) => t.Resumen !== "Sí" && t.Hito !== "Sí") ? "" : `<div class="sub">Solo fases e hitos (versión compactada antes de la v81)</div>`}${v.Nota ? `<div>${esc(v.Nota)}</div>` : ""}</td>
        <td class="derecha nowrap"><button class="btn chico" data-crv-ver="${esc(v.ID_Version)}">Ver</button>${puede ? ` <button class="btn chico" data-crv-edit="${esc(v.ID_Version)}">Editar</button>` : ""}</td></tr>`;
    }).join("");
    return `<div class="card"><div class="titulo-fila"><h2>Versiones del cronograma <span class="contador">${vs.length}</span></h2></div>
      <p class="sub">Cada versión guarda su fecha de corte y el avance real y planeado de ese momento. El corrimiento compara la fecha fin con la versión anterior.</p>
      ${vs.length > 1 ? `<div class="grid2 crv-graf"><div><h3>Avance real vs planeado</h3><div class="grafico-alto"><canvas id="g-crv-av"></canvas></div></div><div><h3>Fin proyectado por versión</h3><div class="grafico-alto"><canvas id="g-crv-fin"></canvas></div></div></div>` : ""}
      <div class="tabla-scroll"><table class="no-orden no-recortar"><thead><tr><th>Versión</th><th>Corte</th><th>Seguimiento</th><th>Real / plan</th><th>Fin</th><th>Atrasadas</th><th>Archivo / nota</th><th></th></tr></thead><tbody>${filas}</tbody></table></div></div>`;
  }
  function editarVersion(p, v) {
    const vs = versionesDe(p.ID_Proyecto), esVig = versionVigente(p.ID_Proyecto) === v;
    const segs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos).slice().reverse();
    modal(`Versión V${v.Numero} del cronograma`, [
      { k: "Fecha_Corte", label: "Fecha de corte", tipo: "date" },
      { k: "ID_Seguimiento", label: "Sesión de seguimiento", tipo: "select", opciones: segs.map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)}`]), textoVacio: "No asociar a una sesión" },
      { k: "Nota", label: "Nota", ancho: true },
    ], v, (fd) => {
      const ts = tareasDe(p.ID_Proyecto, v.ID_Version), c = calcCronograma(ts, fd.Fecha_Corte || v.Fecha_Corte);
      // Las versiones anteriores ya no tienen el detalle: su % se conserva.
      const cambios = { ...fd, ...(esVig ? { Avance_Real: c.real, Avance_Plan: c.plan, Atrasadas: c.atrasadas.length } : {}) };
      guardar(() => S.api.actualizarPorId("Cronogramas", "ID_Version", v.ID_Version, cambios), `V${v.Numero} actualizada`);
    }, (fd) => (fd.Fecha_Corte && fd.Fecha_Corte > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""), {
      eliminar: { texto: "Eliminar versión", titulo: `Eliminar V${v.Numero}`, mensaje: esVig && vs.length > 1 ? `Se borrará la V${v.Numero} y su archivo. La versión anterior quedará como vigente.` : `Se borrará la V${v.Numero} con sus tareas.`,
        accion: () => guardar(async () => {
          const ids = (S.datos.Tareas || []).filter((t) => t.ID_Proyecto === p.ID_Proyecto && vidDe(t) === v.ID_Version).map((t) => t.ID_Tarea);
          if (ids.length) await S.api.eliminarFilas("Tareas", "ID_Tarea", ids);
          if (v.Archivo_Guardado === "Sí") await S.api.borrarArchivo(v.ID_Version);
          await S.api.eliminarFilas("Cronogramas", "ID_Version", [v.ID_Version]);
          const F = (S.ganttF || {})[p.ID_Proyecto]; if (F) { F.ver = ""; F.comp = ""; }
          if (!versionesDe(p.ID_Proyecto).length) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { Cronograma_Origen: "", Cronograma_Archivo: "", Cronograma_Fecha: "", Cronograma_Por: "" });
        }, `V${v.Numero} eliminada`) },
    });
  }

  let arrGT = null;
  window.addEventListener("mousemove", (ev) => { if (arrGT) arrGT.mover(ev.clientX - arrGT.x); });
  window.addEventListener("mouseup", () => { if (arrGT) { arrGT.cuerpo.classList.remove("arrastrando"); arrGT = null; } });
  function enlazarGantt(el, p) {
    const b = el.querySelector("#b-crono-imp");
    if (!b) return;
    b.addEventListener("click", () => importarCronograma(p));
    const pid = p.ID_Proyecto, F = (S.ganttF || {})[pid];
    const cc = el.querySelector("#b-crono-cc"); if (cc) cc.addEventListener("click", () => formCambio(p));
    const av = el.querySelector("#b-crono-av");
    if (av) av.addEventListener("click", () => { const c = calcCronograma(tareasDe(pid)); guardar(() => S.api.actualizarPorId("Proyectos", "ID_Proyecto", pid, { Avance_Real: c.real, Avance_Planeado: c.plan, ...sello() }), `Avance actualizado: ${c.real}% / ${c.plan}%`); });
    el.querySelectorAll("[data-crv-arch]").forEach((x) => x.addEventListener("click", async () => {
      const txt = x.innerHTML; x.innerHTML = "Descargando…";
      try { const f = await leerArchivo(x.dataset.crvArch); descargar(f.bytes, f.nombre, f.tipo); } catch (e) { toast("No se pudo descargar el archivo: " + e.message, true); }
      x.innerHTML = txt;
    }));
    el.querySelectorAll("[data-crv-ver]").forEach((x) => x.addEventListener("click", () => { F.ver = x.dataset.crvVer; if (F.comp === F.ver) F.comp = ""; delete F.scroll; render(); const g = $("#gt-cont"); if (g) g.scrollIntoView({ block: "start", behavior: "smooth" }); }));
    el.querySelectorAll("[data-crv-edit]").forEach((x) => x.addEventListener("click", () => { const v = versionesDe(pid).find((y) => y.ID_Version === x.dataset.crvEdit); if (v) editarVersion(p, v); }));
    // Gráficos de evolución entre versiones.
    const vs = versionesDe(pid);
    if (vs.length > 1 && el.querySelector("#g-crv-av")) {
      const et = vs.map((v) => `V${v.Numero} · ${fecha(v.Fecha_Corte).slice(0, 5)}`);
      grafico(el.querySelector("#g-crv-av"), { type: "line", data: { labels: et, datasets: [
        { label: "Real", data: vs.map((v) => Number(v.Avance_Real) || 0), borderColor: COLORES.azul, backgroundColor: COLORES.azul, tension: 0.2 },
        { label: "Planeado", data: vs.map((v) => Number(v.Avance_Plan) || 0), borderColor: "#9aa5b1", backgroundColor: "#9aa5b1", borderDash: [6, 4], tension: 0.2 }] },
        options: { maintainAspectRatio: false, scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + "%" } } } } });
      const base = vs[0].Fin;
      grafico(el.querySelector("#g-crv-fin"), { type: "bar", data: { labels: et, datasets: [{ label: "Días de corrimiento frente a la V1", data: vs.map((v) => (v.Fin && base ? R.difDias(v.Fin, base) : 0)), backgroundColor: vs.map((v) => (v.Fin && base && v.Fin > base ? "#d64535" : "#2e9e5b")) }] },
        options: { maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => `${ctx.raw > 0 ? "+" : ""}${ctx.raw} días · fin ${fecha(vs[ctx.dataIndex].Fin)}` } } }, scales: { y: { ticks: { callback: (v) => `${v} d` } } } } });
    }
    if (!F) return;
    const der = el.querySelector("#gt-der"), cab = el.querySelector("#gt-cabw"), cuerpo = el.querySelector("#gt-cuerpo"), hbar = el.querySelector("#gt-hbar");
    if (der) {
      // «Ajustar»: el ancho real disponible se mide al pintar; si cambió, se vuelve a pintar una vez.
      const real = der.clientWidth;
      if (real && Math.abs(real - (S.ganttAncho || 0)) > 16) { S.ganttAncho = real; if (F.escala === "ajustar" && !F._reajuste) { F._reajuste = true; render(); return; } }
      F._reajuste = false;
      const izqT = el.querySelector(".gt-izq"); if (izqT) hbar.style.marginLeft = izqT.offsetWidth + "px";
      const mover = (sl) => { cab.scrollLeft = sl; cuerpo.scrollLeft = sl; if (hbar.scrollLeft !== sl) hbar.scrollLeft = sl; F.scroll = sl; };
      hbar.addEventListener("scroll", () => mover(hbar.scrollLeft), { passive: true });
      const irHoy = () => { const h = Number(cuerpo.dataset.hoy); if (h >= 0) mover(Math.max(0, h - der.clientWidth / 3)); };
      if (F.scroll !== undefined) mover(F.scroll); else irHoy();
      const bh = el.querySelector("#gt-hoy"); if (bh) bh.addEventListener("click", irHoy);
      // Rueda horizontal (o Shift + rueda) y arrastrar para desplazar la línea de tiempo.
      cuerpo.addEventListener("wheel", (ev) => { const d = ev.deltaX || (ev.shiftKey ? ev.deltaY : 0); if (d) { ev.preventDefault(); mover(Math.max(0, Math.min(hbar.scrollWidth - hbar.clientWidth, hbar.scrollLeft + d))); } }, { passive: false });
      cuerpo.addEventListener("mousedown", (ev) => { if (ev.button === 0 && hbar.scrollWidth > hbar.clientWidth) { arrGT = { x: ev.clientX, sl: hbar.scrollLeft, cuerpo, mover: (dx) => mover(Math.max(0, Math.min(hbar.scrollWidth - hbar.clientWidth, arrGT.sl - dx))) }; cuerpo.classList.add("arrastrando"); ev.preventDefault(); } });
      // Resaltar la misma fila a ambos lados.
      const izq = [...el.querySelectorAll(".gt-izq tbody tr")], dr = [...el.querySelectorAll(".gt-r")];
      const marcar = (i, on) => { if (izq[i]) izq[i].classList.toggle("gt-sobre", on); if (dr[i]) dr[i].classList.toggle("gt-sobre", on); };
      [izq, dr].forEach((arrF) => arrF.forEach((r, i) => { r.addEventListener("mouseenter", () => marcar(i, true)); r.addEventListener("mouseleave", () => marcar(i, false)); }));
    }
    const ultSeg = () => R.avancesDe(pid, S.datos.Seguimientos).slice(-1)[0];
    el.querySelectorAll("[data-ed-add]").forEach((f) => f.addEventListener("submit", (ev) => {
      ev.preventDefault();
      const k = f.dataset.edAdd, txt = f.querySelector("input").value.trim(), sg = ultSeg();
      if (!txt || !sg) return;
      const nuevo = [...partirItems(sg[k]), txt].join("\n");
      guardar(() => S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", sg.ID_Seguimiento, { [k]: nuevo }), "Agregado a la sesión del " + fecha(sg.Fecha_Corte));
    }));
    el.querySelectorAll("[data-ed-quitar]").forEach((x) => x.addEventListener("click", () => {
      const k = x.dataset.edQuitar, i = Number(x.dataset.i), sg = ultSeg();
      if (!sg) return;
      const l = partirItems(sg[k]); l.splice(i, 1);
      guardar(() => S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", sg.ID_Seguimiento, { [k]: l.join("\n") }), "Quitado");
    }));
    el.querySelectorAll("[data-ed-seg]").forEach((x) => x.addEventListener("click", () => formSeguimiento(p)));
    el.querySelectorAll("[data-ed-editseg]").forEach((x) => x.addEventListener("click", () => { const sg = S.datos.Seguimientos.find((y) => y.ID_Seguimiento === x.dataset.edEditseg); if (sg) formEditarSeguimiento(sg); }));
    el.querySelectorAll("[data-gt-modo]").forEach((x) => x.addEventListener("click", () => { F.modo = x.dataset.gtModo; delete F.scroll; render(); }));
    el.querySelectorAll("[data-gt-esc]").forEach((x) => x.addEventListener("click", () => { F.escala = x.dataset.gtEsc; delete F.scroll; render(); }));
    el.querySelectorAll("[data-gt-vista]").forEach((x) => x.addEventListener("click", () => { F.vista = x.dataset.gtVista; delete F.scroll; render(); }));
    el.querySelectorAll("[data-gt-tog]").forEach((x) => x.addEventListener("click", (ev) => { ev.stopPropagation(); const k = x.dataset.gtTog; F.cerr = F.cerr.includes(k) ? F.cerr.filter((y) => y !== k) : [...F.cerr, k]; render(); }));
    const ex = el.querySelector("#gt-exp"); if (ex) ex.addEventListener("click", () => { F.cerr = []; render(); });
    const co = el.querySelector("#gt-con"); if (co) co.addEventListener("click", () => { const ts = tareasDe(pid, F.ver || undefined); const min = Math.min(...ts.map((t) => Number(t.Nivel) || 1)); F.cerr = ts.filter((t) => t.Resumen === "Sí" && (Number(t.Nivel) || 1) > min).map((t) => String(t.Orden)); render(); });
    const n = el.querySelector("#gt-nivel"); if (n) n.addEventListener("change", () => { F.nivel = n.value; render(); });
    const fe = el.querySelector("#gt-estado"); if (fe) fe.addEventListener("change", () => { F.estado = fe.value || ""; delete F.scroll; render(); });
    const fx = el.querySelector("#gt-estado-x"); if (fx) fx.addEventListener("click", () => { F.estado = ""; delete F.scroll; render(); });
    const vv = el.querySelector("#gt-ver"); if (vv) vv.addEventListener("change", () => { F.estado = ""; F.ver = vv.value; if (F.comp === F.ver) F.comp = ""; F.cerr = []; delete F.scroll; render(); });
    const cp = el.querySelector("#gt-comp"); if (cp) cp.addEventListener("change", () => { F.comp = cp.value; render(); });
    const t = el.querySelector("#gt-q"); if (t) t.addEventListener("change", () => { F.q = t.value; render(); });
  }

  // ---------- Informe de proyecto (antes «hoja de vida»; PDF con marca FSFB) ----------
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
    // Palabras muy largas (URLs, códigos) se pueden partir sin cortar las palabras normales.
    const wbr = (v) => String(v).split(/(\s+)/).map((w) => (w.length < 28 ? esc(w) : (w.match(/.{1,20}/g) || []).map(esc).join("<wbr>"))).join("");
    const t = (v) => (v === 0 ? "0" : v ? wbr(v) : "—");
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
    const vencidosC = comps.filter((c) => R.estadoCompromiso(c) === "Vencido").length;
    const cambiosP = cambiosDe(pid), depsP = depsDe(pid).length + bloqueaA(pid).length, lecc = leccionesDe(pid);
    const vEvm = evmDe(p), f2 = (x) => (x === null ? "—" : x.toFixed(2).replace(".", ","));
    const plural = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;
    const enlaceHV = (u, txt) => (esURL(u) ? `<a class="hv-link" href="${esc(u)}" target="_blank" rel="noopener noreferrer">${esc(txt || u)}</a>` : t(u));
    const docs = documentosDe(pid);
    const tsHV = tareasDe(pid), cHV = calcCronograma(tsHV);
    const secciones = [
      { t: "Datos generales", d: `PM ${nombreUsuario(p.PM)} · ${fecha(p.Fecha_Inicio)} → ${fecha(p.Fecha_Fin_Plan)}`, h: `      <table class="hv-ficha">${[
        fila("PM responsable", t(nombreUsuario(p.PM))), fila("Metodología", t(p.Metodologia)), fila("Prioridad", t(p.Prioridad)),
        fila("Fecha de inicio", fecha(p.Fecha_Inicio)), fila("Fin planeado", fecha(p.Fecha_Fin_Plan)), fila("Frecuencia de seguimiento", agendaDe(pid).filter((a) => a.Activa !== "No").length ? agendaDe(pid).filter((a) => a.Activa !== "No").map((a) => esc(resumenAgenda(a))).join("<br>") : t(p.Frecuencia_Seguimiento)),
        fila("Estado del seguimiento", `${t(e.estado)} · próximo ${fecha(e.proximo)}`), fila("Presupuesto", cop(p.Presupuesto)), fila("Ejecutado", cop(p.Ejecutado)),
        fila("Repositorio", enlaceHV(p.URL_Repositorio)), fila("Documentos", enlaceHV(p.URL_Documentos)), fila("Comentario de estado", t(p.Comentario_Estado)),
      ].join("")}</table>` },
      { t: "Desempeño y línea base", d: vEvm ? `SPI ${f2(vEvm.SPI)} · CPI ${f2(vEvm.CPI)}${lineaBase(p).tiene ? ` · LB${lineaBase(p).num}` : ""}` : "sin presupuesto", h: `${(() => { const v = evmDe(p); if (!v) return `<p class="hv-vacio">Sin presupuesto registrado: no se calcula el valor ganado.</p>`;
        const f = (x) => (x === null ? "—" : x.toFixed(2).replace(".", ","));
        return `<table class="hv-ficha">${[fila("SPI · cronograma", `${f(v.SPI)} <small>(valor ganado / valor planeado; 1 = según plan)</small>`), fila("CPI · costo", `${f(v.CPI)} <small>(valor ganado / costo real; 1 = según presupuesto)</small>`),
          fila("Valor planeado (PV)", cop(Math.round(v.PV))), fila("Valor ganado (EV)", cop(Math.round(v.EV))), fila("Costo real (AC)", cop(v.AC)), fila("Estimado al terminar (EAC)", v.EAC !== null ? cop(Math.round(v.EAC)) : "—"),
          fila("Presupuesto base (BAC)", `${cop(v.BAC)} <small>${lineaBase(p).tiene ? `línea base LB${lineaBase(p).num}` : "presupuesto actual"}</small>`)].join("")}</table>`; })()}
      ${(() => { const lb = lineaBase(p); return lb.tiene ? `<p class="hv-p"><b>Línea base vigente (LB${lb.num}):</b> ${lb.ib ? `inicio ${fecha(lb.ib)} → actual ${fecha(p.Fecha_Inicio)}${lb.diasInicio ? ` (${lb.diasInicio > 0 ? "+" : ""}${lb.diasInicio} días)` : ""} · ` : ""}fin ${fecha(lb.fb)} → actual ${fecha(p.Fecha_Fin_Plan)}${lb.dias !== null ? ` (${lb.dias > 0 ? "+" : ""}${lb.dias} días)` : ""} · presupuesto ${lb.pb !== null ? cop(lb.pb) : "—"} → actual ${cop(p.Presupuesto)}${lb.pctCosto !== null ? ` (${lb.pctCosto > 0 ? "+" : ""}${lb.pctCosto}%)` : ""}</p>` : ""; })()}
      ${(() => { const l = lineasBaseDe(p.ID_Proyecto); return l.length > 1 ? `<table class="hv-tabla"><thead><tr><th>LB</th><th>Fijada</th><th>Inicio</th><th>Fin</th><th>Presupuesto</th><th>Motivo</th></tr></thead><tbody>${l.map((x) => `<tr><td>LB${esc(x.Numero)}</td><td>${fecha(x.Fecha)}</td><td>${fecha(x.Fecha_Inicio)}</td><td>${fecha(x.Fecha_Fin)}</td><td>${x.Presupuesto !== "" && x.Presupuesto != null ? cop(x.Presupuesto) : "—"}</td><td>${esc(x.Motivo || "")}</td></tr>`).join("")}</tbody></table>` : ""; })()}` },
      (() => { const V = valorProyecto(p); return { t: "Valor y beneficios para la FSFB", d: V.nivel ? `nivel ${V.nivel.toLowerCase()}${V.indice !== null ? ` · logro ${V.indice}%` : ""} · ${plural(V.bs.length, "beneficio", "beneficios")}` : "sin propuesta de valor",
        h: `<table class="hv-ficha">${[fila("Beneficio esperado", t(p.Beneficio_Esperado)), fila("Valor para la FSFB", t(p.Valor_FSFB)), fila("Nivel de valor", `${t(V.nivel)}${V.nivel && !V.manual ? " <small>(sugerido)</small>" : ""}`),
          fila("Impacto por dimensión", Object.entries(V.imp).map(([d, i]) => `${esc(d)}: <b>${esc(i)}</b>`).join(" · ") || "—"), fila("Índice de valor", V.indice === null ? "sin mediciones" : `${V.indice}% <small>(logro ponderado por importancia)</small>`),
          ...(V.valorFin ? [fila("Valor financiero (meta)", `${cop(V.valorFin)}${V.roi ? ` · ROI ${V.roi.toFixed(1).replace(".", ",")}x · payback ${V.payback} meses` : ""}`)] : [])].join("")}</table>
        ${tabla(["Importancia", "Dimensión", "Beneficio esperado / valor para la FSFB", "Indicador", "Base → meta", "Actual", "Logro", "Estado"], V.bs.map(({ b, e }) => `<tr><td>${t(b.Importancia)}</td><td>${t(b.Dimension)}</td><td>${t(b.Beneficio_Esperado)}${b.Valor_FSFB ? `<br><small>${t(b.Valor_FSFB)}</small>` : ""}</td><td>${t(b.Indicador)}</td>
          <td class="hv-nw">${b.Tipo_Medida === "Cualitativo" ? "Cualitativo" : `${esc(fmtValor(b, b.Linea_Base))} → ${esc(fmtValor(b, b.Meta))}`}</td><td class="hv-nw">${e.ult ? `${esc(fmtValor(b, e.ult.Valor))}<br><small>${fecha(e.ult.Fecha)}</small>` : "—"}</td><td>${e.logro === null ? "—" : `${e.logro}%`}</td><td>${t(e.estado)}</td></tr>`), "Sin beneficios registrados.")}` }; })(),
      { t: "Equipo: stakeholders y proveedores", d: `${plural(stks.length, "persona", "personas")} · ${plural(provs.length, "proveedor", "proveedores")}`, h: `      ${tabla(["Rol", "Nombre", "Cargo / área", "Empresa", "Correo"], stks.map((x) => `<tr><td>${t(x.Rol)}</td><td>${t(x.Nombre)}</td><td>${t([x.Cargo, x.Area].filter(Boolean).join(" · "))}</td><td>${x.ID_Proveedor ? t((proveedor(x.ID_Proveedor) || {}).Nombre) : "Interno"}</td><td>${t(x.Correo)}</td></tr>`), "Sin stakeholders registrados.")}
      <p class="hv-p"><b>Proveedores:</b> ${provs.length ? provs.map((x) => `${esc(x.Nombre)}${x.Servicio ? ` (${esc(x.Servicio)})` : ""}`).join(" · ") : "ninguno"}</p>` },
      { t: "Hitos", d: `${hitos.length} · ${hitos.filter((h) => h.Estado === "Cumplido").length} cumplidos`, h: `      ${tabla(["Hito", "Fecha plan", "Fecha real", "Estado"], hitos.map((h) => `<tr><td>${t(h.Hito)}</td><td>${fecha(h.Fecha_Plan)}</td><td>${h.Fecha_Real ? fecha(h.Fecha_Real) : "—"}</td><td>${t(h.Estado)}</td></tr>`), "Sin hitos.")}` },
      { t: "Cronograma", d: tsHV.length ? `${p.Cronograma_Origen || "importado"} · ${cHV.real}% real / ${cHV.plan}% plan${cHV.atrasadas.length ? ` · ${cHV.atrasadas.length} atrasadas` : ""}` : "sin importar",
        h: !tsHV.length ? `<p class="hv-vacio">Sin cronograma importado. Se importa desde Planner (Excel) o MS Project (XML) en la pestaña Cronograma.</p>` : `<p class="hv-p">Importado de <b>${esc(p.Cronograma_Origen || "")}</b> el ${fecha(p.Cronograma_Fecha)} · ${cHV.hojas} tareas y ${cHV.hitos} hitos · del ${fecha(cHV.inicio)} al ${fecha(cHV.fin)}${p.Fecha_Fin_Plan && cHV.fin > p.Fecha_Fin_Plan ? ` <b>(termina después del fin planeado ${fecha(p.Fecha_Fin_Plan)})</b>` : ""} · avance real ${cHV.real}% · planeado ${cHV.plan}%.</p>
          <div class="hv-edt"><p class="hv-p"><b>Avance por EDT</b> (real / planeado al ${fecha(R.hoyISO())})</p>${vistaEDT(p, tsHV, R.hoyISO(), false)}</div>
          <p class="hv-p"><b>Detalle de tareas</b></p>
          ${tabla(["EDT", "Tarea", "Inicio", "Fin", "%", "Estado"], (() => { const nmax = tsHV.filter((x) => Number(x.Nivel) <= 2).length <= 60 ? 2 : 1; return tsHV.filter((x) => Number(x.Nivel) <= nmax || estadoTarea(x) === "atr").slice(0, 80); })().map((x) => `<tr><td class="hv-nw">${t(x.EDT)}</td><td style="padding-left:${(Number(x.Nivel) - 1) * 10 + 6}px">${x.Resumen === "Sí" ? `<b>${t(x.Nombre)}</b>` : `${x.Hito === "Sí" ? "◆ " : ""}${t(x.Nombre)}`}</td><td class="hv-nw">${fecha(x.Inicio)}</td><td class="hv-nw">${fecha(x.Fin)}</td><td>${Number(x.Pct_Completado) || 0}%</td><td>${TXT_EST_TAREA[estadoTarea(x)]}</td></tr>`), "Sin tareas.")}
          <p class="hv-nota">Se muestran los niveles principales y las tareas atrasadas; el detalle completo está en la pestaña Cronograma de la app.</p>
          ${versionesDe(pid).length > 1 ? `<p class="hv-p"><b>Versiones del cronograma</b></p>${tabla(["Versión", "Corte", "Sesión", "Real / plan", "Fin", "Nota"], versionesDe(pid).slice().reverse().map((v) => { const sg = v.ID_Seguimiento && segs.find((x) => x.ID_Seguimiento === v.ID_Seguimiento); return `<tr><td>V${t(v.Numero)}</td><td class="hv-nw">${fecha(v.Fecha_Corte)}</td><td class="hv-nw">${sg ? fecha(sg.Fecha_Corte) : "—"}</td><td class="hv-nw">${pct(v.Avance_Real)} / ${pct(v.Avance_Plan)}</td><td class="hv-nw">${fecha(v.Fin)}</td><td>${t(v.Nota)}</td></tr>`; }), "")}` : ""}` },
      { t: "Seguimientos y sesiones", d: segs.length ? `${plural(segs.length, "sesión", "sesiones")} · última ${fecha(segs[segs.length - 1].Fecha_Corte)}` : "sin sesiones", h: `      ${tabla(["Fecha", "Avance", "Semáforo", "Logros", "Próximos pasos", "Bloqueos", "Acta"], segs.slice().reverse().map((x) => `<tr><td class="hv-nw">${fecha(x.Fecha_Corte)}${R.esSesion(x) ? `<br><small>Sesión${x.Nombre_Sesion ? `: ${esc(x.Nombre_Sesion)}` : ""}</small>` : ""}</td><td>${R.esSesion(x) ? "—" : pct(x.Avance_Real)}</td><td>${R.esSesion(x) ? "—" : t(x.Semaforo)}</td><td>${t(x.Logros)}</td><td>${t(x.Proximos_Pasos)}</td><td>${t(x.Bloqueos)}</td><td>${x.Fecha_Acta ? fecha(x.Fecha_Acta) : "—"}${x.Acta_Archivo ? " · PDF" : ""}${esURL(x.URL_Acta) ? `<br>${enlaceHV(x.URL_Acta, "Ver acta ↗")}` : ""}</td></tr>`), "Sin seguimientos.")}` },
      { t: "Compromisos", d: `${comps.length} · ${abiertos} abiertos${vencidosC ? ` · ${vencidosC} vencidos` : ""}`, h: `      ${tabla(["Compromiso", "Origen", "Responsable", "Fecha límite", "Estado", "Seguimiento"], comps.map((c) => {
        const hilo = comentariosDe(c.ID_Compromiso);
        return `<tr><td>${t(c.Compromiso)}</td><td>${esc(origen(c))}</td><td>${t(c.Responsable)}</td><td class="hv-nw">${fecha(c.Fecha_Compromiso)}</td>
          <td>${esc(R.estadoCompromiso(c))}${R.cerrado(c) && c.Fecha_Cierre ? `<br><small>cerrado ${fecha(c.Fecha_Cierre)}</small>` : ""}</td>
          <td>${hilo.length ? hilo.map((x) => `<div class="hv-com"><small>${fechaEv(String(x.Fecha_Hora))} · ${esc(nombreUsuario(x.Autor))}</small><br>${wbr(resumenCom(x))}</div>`).join("") : "—"}</td></tr>`;
      }), "Sin compromisos.")}` },
      { t: "Riesgos", d: `${riesgos.length} · ${riesgos.filter((r) => ["Abierto", "Materializado"].includes(r.Estado)).length} abiertos`, h: `      ${tabla(["No.", "Riesgo", "Responsable", "Inherente", "Planes", "Residual", "Origen"], riesgos.slice().sort((a, b) => numRiesgo(a) - numRiesgo(b)).map((r) => `<tr><td>${numRiesgo(r)}</td><td><small>${t(r.Tipo)} · ${t(r.Estado)}</small><br>${t(r.Descripcion)}</td><td>${t(r.Responsable)}</td>
        <td class="hv-nw">${califTxt(r.Probabilidad_Inherente, r.Impacto_Inherente)}</td><td>${r.Plan_Mitigacion ? `<b>Mitigación:</b> ${t(r.Plan_Mitigacion)}` : ""}${r.Plan_Contingencia ? `<br><b>Contingencia:</b> ${t(r.Plan_Contingencia)}` : ""}</td>
        <td class="hv-nw">${califTxt(r.Probabilidad_Residual, r.Impacto_Residual)}</td><td>${esc(origenRiesgo(r))}</td></tr>`), "Sin riesgos.")}` },
      { t: "Control de cambios", d: cambiosP.length ? `${cambiosP.length} · ${cambiosP.filter((c) => c.Estado === "Aprobado").length} aprobados${cambiosP.some((c) => (c.Estado || "Solicitado") === "Solicitado") ? ` · ${cambiosP.filter((c) => (c.Estado || "Solicitado") === "Solicitado").length} por decidir` : ""}` : "sin cambios", h: `      ${(() => { const ap = cambiosDe(pid).filter((c) => c.Estado === "Aprobado"); return ap.length ? `<p class="hv-p"><b>Cambios aprobados:</b> ${ap.length} · ${IMPACTOS_CC.map((i) => { const n = ap.filter((c) => impactosDe(c).includes(i)).length; return n ? `${i}: ${n}` : ""; }).filter(Boolean).join(" · ")}</p>` : ""; })()}
      ${tabla(["N.º / Fecha", "Impacta en", "Cambio", "Efecto", "Estado", "Documento"], cambiosDe(pid).map((c) => `<tr><td class="hv-nw">${c.Numero_CC ? `<b>${t(c.Numero_CC)}</b><br>` : ""}${fecha(c.Fecha)}</td><td>${t(impactosDe(c).join(", "))}</td><td>${t(c.Descripcion)}${c.Justificacion ? `<br><small>Por qué: ${t(c.Justificacion)}</small>` : ""}</td><td>${t(c.Impacto)}${c.Nueva_Fecha_Fin ? `<br><small>Nueva fecha fin ${fecha(c.Nueva_Fecha_Fin)}</small>` : ""}${num0(c.Nuevo_Presupuesto) !== null ? `<br><small>Nuevo presupuesto ${cop(c.Nuevo_Presupuesto)}</small>` : ""}</td><td>${t(c.Estado)}${c.Fecha_Decision ? `<br><small>${fecha(c.Fecha_Decision)}</small>` : ""}</td><td>${c.URL_Documento && esURL(c.URL_Documento) ? `<a href="${esc(c.URL_Documento)}" target="_blank" rel="noopener">Ver documento</a>` : "—"}</td></tr>`), "Sin cambios registrados.")}` },
      { t: "Documentos del proyecto", d: docs.length ? plural(docs.length, "enlace", "enlaces") : "sin documentos", h: tabla(["Documento", "Tipo", "Descripción", "Enlace"], docs.map((d) => `<tr><td>${t(d.Nombre)}</td><td>${t(d.Tipo)}</td><td>${t(d.Descripcion)}</td><td>${esURL(d.URL) ? enlaceHV(d.URL, "Abrir documento ↗") : t(d.URL)}</td></tr>`), "Sin documentos registrados. Agrégalos en la pestaña Documentos.") },
      { t: "Dependencias", d: depsP ? String(depsP) : "ninguna", h: `      ${tabla(["Relación", "Proyecto", "Tipo", "Alerta"], [...depsDe(pid).map((d) => `<tr><td>Depende de</td><td>${esc((proyecto(d.Depende_De) || {}).Nombre || d.Depende_De)}</td><td>${t(d.Tipo)}</td><td>${esc(conflictoDep(d) || "—")}</td></tr>`),
        ...bloqueaA(pid).map((d) => `<tr><td>Bloquea a</td><td>${esc((proyecto(d.ID_Proyecto) || {}).Nombre || d.ID_Proyecto)}</td><td>${t(d.Tipo)}</td><td>${esc(conflictoDep(d) || "—")}</td></tr>`)], "Sin dependencias.")}` },
      { t: "Tickets del helpdesk", d: tks.length ? `${tks.length} · ${tks.filter((k) => !["Resuelto", "Cerrado"].includes(k.Estado)).length} abiertos` : "ninguno", h: `      ${tabla(["N.º", "Título", "Estado", "Prioridad", "Registrado"], tks.map((k) => `<tr><td>${t(k.Numero)}</td><td>${t(k.Titulo)}</td><td>${t(k.Estado)}</td><td>${t(k.Prioridad)}</td><td>${fecha(k.Fecha_Registro)}</td></tr>`), "Sin tickets.")}` },
      { t: "Lecciones aprendidas", d: lecc.length ? String(lecc.length) : "sin registrar", h: `      ${tabla(["Tipo", "Categoría", "Qué pasó", "Lección", "Recomendación"], leccionesDe(pid).map((l) => `<tr><td>${t(l.Tipo)}</td><td>${t(l.Categoria)}</td><td>${t(l.Situacion)}</td><td>${t(l.Leccion)}</td><td>${t(l.Recomendacion)}</td></tr>`), "Sin lecciones registradas.")}` },
      { t: "Bitácora del proyecto", d: `${plural(eventos.length, "evento", "eventos")} en orden cronológico`, h: `${(() => {
        if (!eventos.length) return `<p class="hv-vacio">Sin eventos registrados.</p>`;
        const MES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
        const recorta = (x, n) => (x.length > n ? x.slice(0, n).replace(/\s+\S*$/, "") + "…" : x);
        let mes = ""; const filas = [];
        eventos.forEach((x) => {
          const m = x.f.slice(0, 7);
          if (m !== mes) { mes = m; filas.push(`<tr class="hv-mes"><td colspan="3">${MES[Number(m.slice(5, 7)) - 1] || ""} ${m.slice(0, 4)}</td></tr>`); }
          filas.push(`<tr><td class="hv-nw">${fechaEv(x.f)}</td><td><span class="hv-tag hv-tag-${x.tipo.toLowerCase()}">${x.tipo}</span></td><td>${wbr(recorta(x.txt, 280))}</td></tr>`);
        });
        return `<table class="hv-t hv-bitacora"><colgroup><col style="width:18%"><col style="width:15%"><col></colgroup><thead><tr><th>Fecha</th><th>Tipo</th><th>Qué pasó</th></tr></thead><tbody>${filas.join("")}</tbody></table>
          <p class="hv-nota">Los textos largos se resumen aquí; el detalle completo de cada comentario está en la sección «Compromisos».</p>`;
      })()}` },
    ];
    const indice = `<nav class="hv-indice" id="hv-indice" aria-label="Contenido"><div class="hv-indice-t">Contenido</div>
      <ol style="grid-template-rows:repeat(${Math.ceil(secciones.length / 2)},auto)">${secciones.map((x, i) => `<li><a href="#hv-s${i + 1}" data-hv-ir="hv-s${i + 1}"><span class="hv-ind-n">${i + 1}</span><span class="hv-ind-t">${esc(x.t)}</span><span class="hv-ind-d">${esc(x.d)}</span></a></li>`).join("")}</ol></nav>`;
    const doc = `
      ${cabDocumento("Informe de proyecto", p.Nombre, `${esc(pid)}${p.Codigo_Almera ? ` · Almera ${esc(p.Codigo_Almera)}` : ""} · ${esc(p.Cliente_Area || "")}`)}
      <section class="hv-kpis">
        <div><span>Estado</span><b>${chipEstadoP(p.Estado)}</b></div><div><span>Fase</span><b>${chipFase(p.Fase)}</b></div>
        <div><span>Semáforo</span><b>${chipSemaforo(p.Semaforo)}</b></div><div><span>Avance real / plan</span><b>${pct(p.Avance_Real)} / ${pct(p.Avance_Planeado)}</b></div>
        <div><span>Presupuesto ejecutado</span><b>${Number(p.Presupuesto) ? Math.round((Number(p.Ejecutado) || 0) / Number(p.Presupuesto) * 100) + "%" : "—"}</b></div>
        <div><span>Compromisos abiertos</span><b>${abiertos} de ${comps.length}</b></div>
        <div><span>Riesgos abiertos</span><b>${riesgos.filter((r) => ["Abierto", "Materializado"].includes(r.Estado)).length} de ${riesgos.length}</b></div>
        ${(() => { const v = evmDe(p); const f = (x) => (x === null ? "—" : x.toFixed(2).replace(".", ",")); return v ? `<div><span>SPI · CPI</span><b>${f(v.SPI)} · ${f(v.CPI)}</b></div>` : ""; })()}
      </section>
      ${indice}
      ${secciones.map((x, i) => `<section class="hv-sec" id="hv-s${i + 1}"><h2><span>${i + 1}. ${esc(x.t)}</span><a class="hv-arriba" href="#hv-indice" data-hv-ir="hv-indice">↑ Contenido</a></h2>
${x.h}</section>`).join("")}
      <footer class="hv-pie">Fundación Santa Fe de Bogotá · Oficina de Proyectos · Informe de proyecto ${esc(pid)} · Documento generado desde el Portafolio PMO</footer>`;
    mostrarDocumento(`Informe de proyecto · ${p.Nombre}`, doc, `Informe de proyecto ${pid} - ${p.Nombre}`);
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
    // Índice: salta a la sección dentro del documento (en el PDF quedan como vínculos internos).
    o.addEventListener("click", (ev) => {
      const a = ev.target.closest("[data-hv-ir]");
      if (!a) return;
      ev.preventDefault();
      const destino = o.querySelector("#" + a.dataset.hvIr);
      if (destino) destino.scrollIntoView({ behavior: "smooth", block: "start" });
    });
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
  // ---------- Líneas base (varias por proyecto: LB0 original, LB1, LB2… re-planeaciones) ----------
  // ---------- Desempeño (valor ganado) ----------
  const bacDe = (p) => { const lb = lineaBase(p); return lb.pb !== null && lb.pb > 0 ? lb.pb : Number(p.Presupuesto) || 0; };
  const evmDe = (p) => R.valorGanado(bacDe(p), p.Avance_Real, p.Avance_Planeado, p.Ejecutado);
  const semIndice = (x) => (x === null ? "" : x >= 0.95 ? "Verde" : x >= 0.85 ? "Amarillo" : "Rojo");
  const fmtIndice = (x) => (x === null ? `<span class="sub">—</span>` : `<span class="dot" style="background:${COLORES[semIndice(x)]}" aria-hidden="true"></span><b>${x.toFixed(2).replace(".", ",")}</b>`);
  function cardDesempeno(p) {
    const v = evmDe(p);
    if (!v) return "";
    const dato = (k, val, ayuda) => `<div class="dato" title="${esc(ayuda)}"><div class="dato-k">${k}</div><div class="dato-v">${val}</div></div>`;
    const lect = [
      v.SPI !== null ? `Cronograma: va al <b>${Math.round(v.SPI * 100)}%</b> del ritmo planeado${v.SPI >= 1 ? " (al día o adelantado)" : ""}.` : "Cronograma: registra el avance planeado para calcular el SPI.",
      v.CPI !== null ? `Costo: por cada $1 ejecutado se ha logrado <b>$${v.CPI.toFixed(2).replace(".", ",")}</b> de avance${v.CPI >= 1 ? " (dentro del presupuesto)" : ""}.` : "Costo: registra el ejecutado en los seguimientos para calcular el CPI.",
      v.EAC !== null ? `Si sigue así, terminaría costando <b>${cop(Math.round(v.EAC))}</b> (${v.VAC >= 0 ? "ahorro" : "sobrecosto"} de ${cop(Math.abs(Math.round(v.VAC)))}).` : "",
    ].filter(Boolean);
    return `<div class="card"><h2>Desempeño · valor ganado</h2>
      <div class="datos">${dato("SPI · cronograma", fmtIndice(v.SPI), "Valor ganado / valor planeado. 1 = según plan; menor a 1 = atrasado.")}${dato("CPI · costo", fmtIndice(v.CPI), "Valor ganado / costo real. 1 = según presupuesto; menor a 1 = sobrecosto.")}
        ${dato("Valor planeado (PV)", cop(Math.round(v.PV)), "Avance planeado × presupuesto de la línea base")}${dato("Valor ganado (EV)", cop(Math.round(v.EV)), "Avance real × presupuesto de la línea base")}
        ${dato("Costo real (AC)", cop(v.AC), "Presupuesto ejecutado a la fecha")}${dato("Estimado al terminar (EAC)", v.EAC !== null ? cop(Math.round(v.EAC)) : "—", "Presupuesto / CPI")}</div>
      <p class="sub">${lect.join(" ")} Base: ${cop(v.BAC)} (${lineaBase(p).tiene ? `línea base LB${lineaBase(p).num}` : "presupuesto actual"}). Verde ≥ 0,95 · amarillo ≥ 0,85 · rojo &lt; 0,85.</p></div>`;
  }

  const lineasBaseDe = (pid) => (S.datos.LineasBase || []).filter((x) => x.ID_Proyecto === pid).sort((a, b) => Number(a.Numero) - Number(b.Numero));
  // La vigente es la última; la original, la LB0. Si el proyecto es de antes de esta versión, se usan sus columnas *_Base.
  function lineaBase(p) {
    const l = lineasBaseDe(p.ID_Proyecto), vig = l[l.length - 1];
    const ib = vig ? vig.Fecha_Inicio || "" : p.Fecha_Inicio_Base || "", fb = vig ? vig.Fecha_Fin || "" : p.Fecha_Fin_Base || "", pb = vig ? num0(vig.Presupuesto) : num0(p.Presupuesto_Base);
    const tiene = !!(fb || ib || pb !== null);
    return {
      tiene, fb, pb, ib, lista: l, vig, num: vig ? Number(vig.Numero) : tiene ? 0 : null, orig: l[0] || null,
      diasInicio: ib && p.Fecha_Inicio ? R.difDias(p.Fecha_Inicio, ib) : null,
      dias: fb && p.Fecha_Fin_Plan ? R.difDias(p.Fecha_Fin_Plan, fb) : null,
      pctCosto: pb ? Math.round(((Number(p.Presupuesto) || 0) - pb) / pb * 100) : null,
    };
  }
  // Crea una nueva línea base (la siguiente en número) y la deja como vigente en el proyecto.
  async function crearLineaBase(p, v, motivo, idCambio) {
    const l = lineasBaseDe(p.ID_Proyecto);
    const num = l.length ? Math.max(...l.map((x) => Number(x.Numero) || 0)) + 1 : 0;
    await S.api.agregarFila("LineasBase", { ID_LB: R.siguienteIdHijo("LB", S.datos.LineasBase || [], "ID_LB", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, Numero: num, Fecha: R.hoyISO(),
      Fecha_Inicio: v.Fecha_Inicio || "", Fecha_Fin: v.Fecha_Fin || "", Presupuesto: v.Presupuesto === "" || v.Presupuesto === undefined ? "" : v.Presupuesto, Motivo: motivo || "", ID_Cambio: idCambio || "", Registrado_Por: S.usuario.Correo });
    await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { Fecha_Inicio_Base: v.Fecha_Inicio || "", Fecha_Fin_Base: v.Fecha_Fin || "", Presupuesto_Base: v.Presupuesto ?? "", ...sello() });
    return num;
  }
  // Una vez: los proyectos que ya tenían línea base (columnas *_Base) quedan con su LB0 registrada.
  let lbMigradas = false;
  async function migrarLineasBase() {
    if (lbMigradas || !S.datos.LineasBase) return;
    lbMigradas = true;
    const filas = S.datos.Proyectos.filter((p) => (p.Fecha_Fin_Base || p.Fecha_Inicio_Base || num0(p.Presupuesto_Base) !== null) && !lineasBaseDe(p.ID_Proyecto).length)
      .map((p, i) => ({ ID_LB: R.siguienteIdHijo("LB", S.datos.LineasBase || [], "ID_LB", p.ID_Proyecto, i), ID_Proyecto: p.ID_Proyecto, Numero: 0, Fecha: String(p.Actualizado_El || R.hoyISO()).slice(0, 10),
        Fecha_Inicio: p.Fecha_Inicio_Base || p.Fecha_Inicio || "", Fecha_Fin: p.Fecha_Fin_Base || "", Presupuesto: p.Presupuesto_Base ?? "", Motivo: "Línea base original (registrada antes)", ID_Cambio: "", Registrado_Por: "" }));
    if (!filas.length) return;
    try { await S.api.agregarFilas("LineasBase", filas); } catch (e) { /* sin permiso de edición */ }
  }
  function seccionCambios(p, puedeSolicitar, puedeDecidir) {
    const lb = lineaBase(p);
    const lista = cambiosDe(p.ID_Proyecto);
    const pend = lista.filter((c) => c.Estado === "Solicitado").length;
    const desvio = (v, u) => (v === null ? "—" : `<span class="${v > 0 ? "baja" : v < 0 ? "sube" : ""}">${v > 0 ? "+" : ""}${v}${u}</span>`);
    const o = lb.orig, dO = o && lb.lista.length > 1 && o.Fecha_Fin && p.Fecha_Fin_Plan ? R.difDias(p.Fecha_Fin_Plan, o.Fecha_Fin) : null, pO = o && lb.lista.length > 1 && num0(o.Presupuesto) ? Math.round(((Number(p.Presupuesto) || 0) - num0(o.Presupuesto)) / num0(o.Presupuesto) * 100) : null;
    return `<div class="card"><div class="titulo-fila"><h2>Líneas base ${lb.tiene ? `<span class="pill cc-imp">Vigente: LB${lb.num}</span>` : ""}</h2>${puedeDecidir ? `<button class="btn chico ${lb.tiene ? "" : "primario"}" id="b-lb">${lb.tiene ? "+ Nueva línea base" : "Fijar línea base (LB0)"}</button>` : ""}</div>
        ${lb.tiene ? `<p class="sub">Comparación del proyecto hoy contra la línea base vigente (LB${lb.num}${lb.vig && lb.vig.Fecha ? `, fijada el ${fecha(lb.vig.Fecha)}` : ""}).${dO !== null || pO !== null ? ` Contra la original (LB0): fin ${desvio(dO, " días")} · presupuesto ${desvio(pO, "%")}.` : ""}</p>
          <div class="datos">
          <div class="dato"><div class="dato-k">Inicio LB${lb.num}</div><div class="dato-v">${lb.ib ? fecha(lb.ib) : "—"}</div></div>
          <div class="dato"><div class="dato-k">Inicio actual</div><div class="dato-v">${fecha(p.Fecha_Inicio)}${lb.ib ? ` · ${desvio(lb.diasInicio, " días")}` : ""}</div></div>
          <div class="dato"><div class="dato-k">Fin LB${lb.num}</div><div class="dato-v">${fecha(lb.fb)}</div></div>
          <div class="dato"><div class="dato-k">Fin planeado actual</div><div class="dato-v">${fecha(p.Fecha_Fin_Plan)} · ${desvio(lb.dias, " días")}</div></div>
          <div class="dato"><div class="dato-k">Presupuesto LB${lb.num}</div><div class="dato-v">${lb.pb !== null ? cop(lb.pb) : "—"}</div></div>
          <div class="dato"><div class="dato-k">Presupuesto actual</div><div class="dato-v">${cop(p.Presupuesto)} · ${desvio(lb.pctCosto, "%")}</div></div></div>
          ${lb.lista.length ? `<div class="tabla-scroll"><table class="no-orden tabla-lb"><thead><tr><th>Línea base</th><th>Fijada</th><th>Inicio</th><th>Fin</th><th>Presupuesto</th><th>Motivo</th><th></th></tr></thead><tbody>
            ${lb.lista.slice().reverse().map((x, i, arr) => { const prev = arr[i + 1]; const dd = prev && prev.Fecha_Fin && x.Fecha_Fin ? R.difDias(x.Fecha_Fin, prev.Fecha_Fin) : null;
              const cc = x.ID_Cambio && (S.datos.Cambios || []).find((c) => c.ID_Cambio === x.ID_Cambio);
              const ELB = (y, campo, tipo, etiqueta) => ed({ puede: puedeDecidir, tabla: "LineasBase", idCol: "ID_LB", id: y.ID_LB, campo, tipo, valor: y[campo] ?? "", etiqueta, min: tipo === "number" ? 0 : undefined,
                despues: (v) => { const mapa = { Fecha_Inicio: "Fecha_Inicio_Base", Fecha_Fin: "Fecha_Fin_Base", Presupuesto: "Presupuesto_Base" }; return y === lb.vig && mapa[campo] ? S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { [mapa[campo]]: v }) : null; } });
              return `<tr class="${x === lb.vig ? "lb-vigente" : ""}"><td><b>LB${esc(x.Numero)}</b>${x === lb.vig ? ` <span class="pill aprobado">vigente</span>` : Number(x.Numero) === 0 ? ` <span class="pill">original</span>` : ""}</td><td class="nowrap"${ELB(x, "Fecha", "date", "Fecha en que se fijó")}>${fecha(x.Fecha)}</td>
                <td class="nowrap"${ELB(x, "Fecha_Inicio", "date", "Inicio de la línea base")}>${fecha(x.Fecha_Inicio)}</td><td class="nowrap"${ELB(x, "Fecha_Fin", "date", "Fin de la línea base")}>${fecha(x.Fecha_Fin)}${dd ? `<div class="sub">${dd > 0 ? "+" : ""}${dd} días vs LB${esc(prev.Numero)}</div>` : ""}</td><td class="nowrap"${ELB(x, "Presupuesto", "number", "Presupuesto de la línea base")}>${num0(x.Presupuesto) !== null ? cop(x.Presupuesto) : "—"}</td>
                <td${ELB(x, "Motivo", "textarea", "Motivo")}>${esc(x.Motivo || "—")}${cc ? `<div class="sub">Control de cambios ${esc(cc.Numero_CC || cc.ID_Cambio)}</div>` : ""}</td>
                <td class="derecha">${puedeDecidir ? `<button class="btn chico" data-lb="${esc(x.ID_LB)}">Editar</button>` : ""}</td></tr>`; }).join("")}</tbody></table></div>` : ""}`
        : `<p class="sub">Aún no hay línea base. Se fija sola con el primer cambio aprobado (toma el inicio, el fin y el presupuesto de ese momento)${puedeDecidir ? ", o puedes fijarla y editarla ahora con «Fijar línea base»" : ""}.</p>`}
      </div>
      <div class="card"><div class="titulo-fila"><h2>Control de cambios ${pend ? `<span class="contador rojo">${pend} por decidir</span>` : ""}</h2>${puedeSolicitar ? `<button class="btn primario" id="b-cambio">+ Control de cambios</button>` : ""}</div>
        <p class="sub">Cambios que impactan tiempo, costo, alcance, recursos o calidad. El PM los registra y la PMO o el Admin los aprueban; al aprobar, la fecha fin y el presupuesto nuevos se aplican al proyecto. Adjunta el enlace del documento de control de cambios firmado.</p>
        ${lista.length ? `<div class="cc-resumen">${IMPACTOS_CC.map((i) => { const n = lista.filter((c) => c.Estado === "Aprobado" && impactosDe(c).includes(i)).length; return n ? `<span class="pill cc-imp">${esc(i)}: ${n} aprobado${n > 1 ? "s" : ""}</span>` : ""; }).join(" ")}</div>
          <div class="tabla-scroll"><table><thead><tr><th>Fecha</th><th>Impacta en</th><th>Cambio</th><th>Efecto</th><th>Estado</th><th></th></tr></thead><tbody>
          ${lista.map((c) => `<tr><td class="nowrap">${c.Numero_CC ? `<b>${esc(c.Numero_CC)}</b><br>` : ""}${fecha(c.Fecha)}<div class="sub">${esc(nombreUsuario(c.Solicitado_Por))}</div></td><td>${impactosDe(c).map((i) => `<span class="pill cc-imp">${esc(i)}</span>`).join(" ") || "—"}</td>
            <td${ed({ puede: c.Estado === "Solicitado" && (puedeSolicitar || puedeDecidir), tabla: "Cambios", idCol: "ID_Cambio", id: c.ID_Cambio, campo: "Descripcion", tipo: "textarea", requerido: true, valor: c.Descripcion, etiqueta: "Cambio" })}><b>${esc(c.Descripcion)}</b>${c.Justificacion ? `<div class="sub">Por qué: ${esc(c.Justificacion)}</div>` : ""}${c.URL_Documento && esURL(c.URL_Documento) ? `<div><a class="cc-doc" href="${esc(c.URL_Documento)}" target="_blank" rel="noopener">📄 Documento del control de cambios ↗</a></div>` : ""}</td>
            <td>${esc(c.Impacto || "")}${c.Nueva_Fecha_Fin ? `<div class="sub">Nueva fecha fin: ${fecha(c.Nueva_Fecha_Fin)}</div>` : ""}${num0(c.Nuevo_Presupuesto) !== null ? `<div class="sub">Nuevo presupuesto: ${cop(c.Nuevo_Presupuesto)}</div>` : ""}</td>
            <td>${pill(c.Estado || "Solicitado")}${c.Fecha_Decision ? `<div class="sub">${fecha(c.Fecha_Decision)} · ${esc(nombreUsuario(c.Decidido_Por))}</div>` : ""}${c.Comentario_Decision ? `<div class="sub">${esc(c.Comentario_Decision)}</div>` : ""}</td>
            <td class="derecha nowrap">${c.Estado === "Solicitado" && puedeDecidir ? `<button class="btn chico primario" data-decidir="${esc(c.ID_Cambio)}">Decidir</button> ` : ""}${puedeSolicitar || puedeDecidir ? `<button class="btn chico" data-cambio="${esc(c.ID_Cambio)}">${c.Estado === "Solicitado" ? "Editar" : "Editar / documento"}</button>` : ""}</td></tr>`).join("")}
        </tbody></table></div>` : vacio("Sin cambios registrados.")}</div>`;
  }
  // Un control de cambios puede impactar varias cosas a la vez (se guardan en la columna Tipo, separadas por coma).
  const IMPACTOS_CC = ["Tiempo", "Costo", "Alcance", "Recursos", "Calidad"];
  const impactosDe = (c) => String(c.Tipo || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  function formCambio(p, c) {
    const nuevo = !c;
    const decidido = !nuevo && c.Estado !== "Solicitado";
    const puedeDecidir = R.puede(S.usuario, "verTodo");
    const opcionesImp = [...new Set([...IMPACTOS_CC, ...(S.cat.Tipo_Cambio || []), ...(nuevo ? [] : impactosDe(c))])];
    const campos = [
      { k: "Tipo", label: "¿En qué impacta?", tipo: "checks", opciones: opcionesImp, vacio: "Sin opciones" },
      { k: "Numero_CC", label: "N.º o código del control de cambios (opcional)", placeholder: "Ej.: CC-003 o FM-GPRO-19 #3" },
      { k: "Fecha", label: "Fecha del cambio", tipo: "date", def: R.hoyISO() },
      { k: "URL_Documento", label: "Enlace del documento (opcional)", tipo: "url", placeholder: "https://…sharepoint.com/… (formato firmado, acta del comité)", ancho: true },
      { k: "Descripcion", label: "¿Qué cambia?", tipo: "textarea", placeholder: "Describe el cambio" },
      { k: "Justificacion", label: "¿Por qué?", tipo: "textarea", placeholder: "Motivo o necesidad" },
      { k: "Impacto", label: "Impacto", tipo: "textarea", placeholder: "Efecto en alcance, calidad, equipo, riesgos…" },
      { seccion: "Si cambia la fecha o el costo", ayuda: "se aplican al proyecto cuando se aprueba" },
      { k: "Nueva_Fecha_Fin", label: "Nueva fecha fin", tipo: "date", ayuda: decidido ? "Ya aplicado al proyecto" : `Hoy: ${fecha(p.Fecha_Fin_Plan)}`, bloqueado: decidido },
      { k: "Nuevo_Presupuesto", label: "Nuevo presupuesto (COP)", tipo: "number", min: 0, ayuda: decidido ? "Ya aplicado al proyecto" : `Hoy: ${cop(p.Presupuesto)}`, bloqueado: decidido },
      ...(nuevo && puedeDecidir ? [{ k: "Estado_Inicial", label: "Estado", tipo: "select", opciones: [["Solicitado", "Solicitado (queda por decidir)"], ["Aprobado", "Ya aprobado (p. ej. en comité): se aplica de una vez"]], def: "Solicitado" },
        { k: "Rebase", label: "Si ya está aprobado y cambia fecha o presupuesto…", tipo: "select", opciones: [["Sí", "Generar nueva línea base"], ["No", "Mantener la línea base actual"]], def: "Sí" }] : []),
    ];
    const validar = (fd) => (!fd.Tipo ? "Marca en qué impacta el cambio (tiempo, costo, alcance…)." : !fd.Descripcion ? "Describe qué cambia." : !String(fd.Justificacion || "").trim() ? "Explica por qué se necesita el cambio (justificación)." :
      fd.Nueva_Fecha_Fin && p.Fecha_Inicio && fd.Nueva_Fecha_Fin < p.Fecha_Inicio ? "La nueva fecha fin no puede ser anterior al inicio del proyecto." : "");
    modal(nuevo ? `Control de cambios · ${p.Nombre}` : `Control de cambios${c.Numero_CC ? ` ${c.Numero_CC}` : ""}`, campos, nuevo ? {} : c, (fd) => {
      if (decidido) { delete fd.Nueva_Fecha_Fin; delete fd.Nuevo_Presupuesto; }
      if (nuevo) {
        const id = R.siguienteIdHijo("CAM", S.datos.Cambios || [], "ID_Cambio", p.ID_Proyecto);
        const aprobar = fd.Estado_Inicial === "Aprobado", rebase = fd.Rebase === "Sí"; delete fd.Estado_Inicial; delete fd.Rebase;
        const fila = { ID_Cambio: id, ID_Proyecto: p.ID_Proyecto, ...fd, Estado: "Solicitado", Solicitado_Por: S.usuario.Correo };
        guardar(async () => {
          await S.api.agregarFila("Cambios", fila);
          if (aprobar) await aplicarCambio(p, fila, "Aprobado", "Registrado como ya aprobado.", rebase);
        }, aprobar ? "Control de cambios registrado y aplicado" : "Cambio solicitado");
      } else guardar(() => S.api.actualizarPorId("Cambios", "ID_Cambio", c.ID_Cambio, fd), "Control de cambios actualizado");
    }, validar, nuevo ? {} : { eliminar: { texto: "Eliminar", mensaje: decidido ? "Se borrará este registro de cambio (lo que ya se aplicó al proyecto no se revierte)." : "Se borrará esta solicitud de cambio.",
      accion: () => guardar(() => S.api.eliminarFilas("Cambios", "ID_Cambio", [c.ID_Cambio]), "Registro eliminado") } });
  }
  // Aprobar o rechazar: si se aprueba, la fecha fin y el presupuesto nuevos pasan al proyecto (y se fija la línea base si no existía).
  // rebase: además de aplicar el cambio, crea una nueva línea base con los valores aprobados.
  async function aplicarCambio(p, c, estado, comentario, rebase) {
    if (estado === "Aprobado") {
      // Si aún no hay línea base, la original (LB0) se toma de los valores de antes del cambio.
      if (!lineaBase(p).tiene) await crearLineaBase(p, { Fecha_Inicio: p.Fecha_Inicio, Fecha_Fin: p.Fecha_Fin_Plan, Presupuesto: p.Presupuesto }, "Línea base original (antes del primer cambio aprobado)");
      const cambiosP = {};
      if (c.Nueva_Fecha_Fin) cambiosP.Fecha_Fin_Plan = c.Nueva_Fecha_Fin;
      if (num0(c.Nuevo_Presupuesto) !== null) cambiosP.Presupuesto = Number(c.Nuevo_Presupuesto);
      if (Object.keys(cambiosP).length) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...cambiosP, ...sello() });
      if (rebase && Object.keys(cambiosP).length) await crearLineaBase(p, { Fecha_Inicio: p.Fecha_Inicio, Fecha_Fin: p.Fecha_Fin_Plan, Presupuesto: p.Presupuesto },
        `Re-planeación por control de cambios${c.Numero_CC ? ` ${c.Numero_CC}` : ""}`, c.ID_Cambio);
    }
    await S.api.actualizarPorId("Cambios", "ID_Cambio", c.ID_Cambio, { Estado: estado, Decidido_Por: S.usuario.Correo, Fecha_Decision: R.hoyISO(), Comentario_Decision: comentario || "" });
  }
  function decidirCambio(p, c) {
    const campos = [
      { k: "Decision", label: "Decisión", tipo: "select", opciones: ["Aprobado", "Rechazado"], def: "Aprobado" },
      { k: "Comentario_Decision", label: "Comentario (opcional)", tipo: "textarea", placeholder: "Condiciones, instancia que aprobó…" },
    ];
    const efectos = [c.Nueva_Fecha_Fin ? `fecha fin ${fecha(p.Fecha_Fin_Plan)} → ${fecha(c.Nueva_Fecha_Fin)}` : "", num0(c.Nuevo_Presupuesto) !== null ? `presupuesto ${cop(p.Presupuesto)} → ${cop(c.Nuevo_Presupuesto)}` : ""].filter(Boolean);
    if (efectos.length) campos.push({ k: "Rebase", label: "¿Generar una nueva línea base con los valores aprobados?", tipo: "select", opciones: [["Sí", "Sí, re-planear (nueva LB)"], ["No", "No, mantener la línea base actual"]], def: "Sí" });
    modal(`Decidir cambio · ${impactosDe(c).join(", ") || c.Tipo}`, campos, {}, (fd) => guardar(() => aplicarCambio(p, c, fd.Decision === "Rechazado" ? "Rechazado" : "Aprobado", fd.Comentario_Decision, fd.Rebase === "Sí"),
      fd.Decision === "Rechazado" ? "Cambio rechazado" : "Cambio aprobado y aplicado"), (fd) => (fd.Decision === "Rechazado" && !String(fd.Comentario_Decision || "").trim() ? "Escribe el motivo del rechazo." : ""),
    { antes: `<div class="resumen-filtro"><b>${esc(c.Descripcion)}</b>${efectos.length ? `<br>Si se aprueba: ${esc(efectos.join(" · "))}` : "<br>No cambia fechas ni presupuesto."}</div>` });
  }
  // Nueva línea base (o editar una existente): inicio, fin, presupuesto, motivo y el control de cambios que la originó.
  function fijarLineaBase(p, x) {
    const nueva = !x, lb = lineaBase(p);
    const ccs = cambiosDe(p.ID_Proyecto).filter((c) => c.Estado === "Aprobado").map((c) => [c.ID_Cambio, `${c.Numero_CC ? c.Numero_CC + " · " : ""}${fecha(c.Fecha)} · ${String(c.Descripcion || "").slice(0, 60)}`]);
    const num = nueva ? (lb.lista.length ? Math.max(...lb.lista.map((y) => Number(y.Numero) || 0)) + 1 : 0) : Number(x.Numero);
    const campos = [
      { html: `<div class="resumen-filtro">Valores actuales del proyecto: inicio <b>${fecha(p.Fecha_Inicio)}</b> · fin <b>${fecha(p.Fecha_Fin_Plan)}</b> · presupuesto <b>${cop(p.Presupuesto)}</b>
        <div style="margin-top:6px"><button type="button" class="btn chico" id="lb-actuales">Usar los valores actuales</button></div></div>` },
      { k: "Fecha_Inicio", label: "Inicio", tipo: "date" },
      { k: "Fecha_Fin", label: "Fin", tipo: "date" },
      { k: "Presupuesto", label: "Presupuesto (COP)", tipo: "number", min: 0 },
      { k: "Motivo", label: "Motivo", tipo: "textarea", placeholder: num === 0 ? "Ej.: Línea base aprobada en el acta de inicio" : "Ej.: Re-planeación aprobada en comité por ampliación de alcance" },
      { k: "ID_Cambio", label: "Control de cambios que la origina (opcional)", tipo: "select", opciones: ccs, textoVacio: "Ninguno", ancho: true },
    ];
    const valores = nueva ? { Fecha_Inicio: p.Fecha_Inicio, Fecha_Fin: p.Fecha_Fin_Plan, Presupuesto: p.Presupuesto, Motivo: num === 0 ? "Línea base original" : "" } : x;
    const esVigente = !nueva && lb.vig && lb.vig.ID_LB === x.ID_LB;
    modal(nueva ? `Nueva línea base · LB${num}` : `Editar LB${x.Numero}`, campos, valores, (fd) => {
      if (nueva) guardar(() => crearLineaBase(p, fd, fd.Motivo, fd.ID_Cambio), `Línea base LB${num} creada`);
      else guardar(async () => {
        await S.api.actualizarPorId("LineasBase", "ID_LB", x.ID_LB, fd);
        if (esVigente) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { Fecha_Inicio_Base: fd.Fecha_Inicio, Fecha_Fin_Base: fd.Fecha_Fin, Presupuesto_Base: fd.Presupuesto, ...sello() });
      }, `LB${x.Numero} actualizada`);
    }, (fd) => (fd.Fecha_Inicio && fd.Fecha_Fin && fd.Fecha_Fin < fd.Fecha_Inicio ? "El fin no puede ser anterior al inicio." : ""), {
      init: (m) => m.querySelector("#lb-actuales").addEventListener("click", () => {
        m.querySelector("[name=Fecha_Inicio]").value = p.Fecha_Inicio || ""; m.querySelector("[name=Fecha_Fin]").value = p.Fecha_Fin_Plan || ""; m.querySelector("[name=Presupuesto]").value = p.Presupuesto || "";
      }),
      ...(nueva ? {} : { eliminar: { texto: "Eliminar línea base", mensaje: `Se borrará LB${x.Numero}.${esVigente ? " La anterior quedará como vigente." : ""}`,
        accion: () => guardar(async () => {
          await S.api.eliminarFilas("LineasBase", "ID_LB", [x.ID_LB]);
          if (esVigente) { const l = lineasBaseDe(p.ID_Proyecto), v = l[l.length - 1];
            await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, v ? { Fecha_Inicio_Base: v.Fecha_Inicio, Fecha_Fin_Base: v.Fecha_Fin, Presupuesto_Base: v.Presupuesto } : { Fecha_Inicio_Base: "", Fecha_Fin_Base: "", Presupuesto_Base: "" }); }
        }, "Línea base eliminada") } }),
    });
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
  // Un solo modelo de priorización para demanda y portafolio (los pesos se ajustan en Backlog → pesos).
  const CRITERIOS = [
    { k: "Valor", t: "Valor / beneficio para la institución", peso: 25 },
    { k: "Alineacion", t: "Alineación estratégica", peso: 20 },
    { k: "Impacto_Paciente", t: "Impacto en paciente / seguridad", peso: 15 },
    { k: "Urgencia", t: "Urgencia", peso: 15 },
    { k: "Riesgo_Prio", t: "Complejidad / riesgo de ejecución", peso: 10, inv: true },
    { k: "Esfuerzo", t: "Esfuerzo / costo", peso: 15, inv: true },
  ];
  const BASE_PRIO = ["Valor", "Urgencia", "Riesgo_Prio", "Esfuerzo"];
  // Puntaje 0-100: más valor, alineación, impacto y urgencia suben; más riesgo y esfuerzo bajan.
  // Proyectos anteriores sin alineación/impacto se califican con los cuatro criterios base (pesos re-normalizados).
  function puntaje(p) {
    if (BASE_PRIO.some((k) => !num0(p[k]))) return null;
    const w = pesosDemanda(), cs = CRITERIOS.filter((c) => num0(p[c.k]));
    const total = cs.reduce((a, c) => a + (w[c.k] || 0), 0) || 1;
    const v = cs.reduce((a, c) => a + (w[c.k] || 0) * (c.inv ? 6 - Number(p[c.k]) : Number(p[c.k])), 0) / total;
    return Math.round(((v - 1) / 4) * 100);
  }
  const cuadrante = (p) => { const v = Number(p.Valor), e = Number(p.Esfuerzo); return v >= 3 ? (e <= 3 ? "Ganancia rápida" : "Proyecto mayor") : (e <= 3 ? "Relleno" : "Reconsiderar"); };
  function vPriorizacion(el) {
    const ps = filtrados();
    const cal = ps.filter((p) => puntaje(p) !== null).sort((a, b) => puntaje(b) - puntaje(a));
    const sin = ps.filter((p) => puntaje(p) === null);
    const puede = (p) => R.puedeEditar(S.usuario, p, "editarProyecto");
    el.innerHTML = `<h1>Priorización del portafolio</h1>${barraFiltros()}
      <div class="card"><h2>Cómo se calcula</h2><p class="sub">Cada proyecto se califica de 1 a 5 en «Editar proyecto» → sección Priorización. Puntaje (0 a 100) = ${(() => { const w = pesosDemanda(), t = Object.values(w).reduce((a, b) => a + b, 0) || 1; return CRITERIOS.map((c) => `${c.t} ${Math.round((w[c.k] / t) * 100)}%${c.inv ? " (a menor, mejor)" : ""}`).join(" · "); })()}. Es el mismo modelo con el que se prioriza la demanda en el Backlog.</p></div>
      <div class="grid2">
        <div class="card"><h2>Ranking</h2>${cal.length ? `<div class="tabla-scroll"><table><thead><tr><th>#</th><th>Proyecto</th><th>Puntaje</th><th class="opc">V · U · R · E</th><th>Cuadrante</th></tr></thead><tbody>
          ${cal.map((p, i) => `<tr class="clic" data-pid="${esc(p.ID_Proyecto)}"><td>${i + 1}</td><td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.Estado)} · ${esc(nombreUsuario(p.PM))}</div></td>
            <td class="nowrap"><span class="cap-barra"><span class="cap-lleno" style="width:${puntaje(p)}%;background:var(--azul)"></span></span> <b>${puntaje(p)}</b></td>
            <td class="opc">${BASE_PRIO.map((k) => esc(p[k])).join(" · ")}</td><td>${pill(cuadrante(p))}</td></tr>`).join("")}</tbody></table></div>` : vacio("Aún no hay proyectos calificados.")}</div>
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
    const fin = p.Fecha_Fin_Base || p.Fecha_Fin_Plan, ini = p.Fecha_Inicio_Base || p.Fecha_Inicio;
    if (!ini || !fin || !pres) return 0;
    if (hasta < ini) return 0;
    if (hasta >= fin) return pres;
    return pres * (R.difDias(hasta, ini) + 1) / (R.difDias(fin, ini) + 1);
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
    Comentarios: "ID_Comentario", Stakeholders: "ID_Stakeholder", Proveedores: "ID_Proveedor", Recursos: "ID_Recurso", Demandas: "ID_Demanda", Sugerencias: "ID_Sugerencia", Documentos: "ID_Documento", LineasBase: "ID_LB", Tareas: "ID_Tarea", Cronogramas: "ID_Version", Tickets: "ID_Ticket", Cambios: "ID_Cambio", Lecciones: "ID_Leccion", RACI: "ID_RACI", Dependencias: "ID_Dependencia" };
  const CAMPO_DESC = ["Nombre", "Compromiso", "Titulo", "Hito", "Descripcion", "Entregable", "Leccion", "Texto", "Valor", "Logros"];
  const NO_AUDITAR = new Set(["Auditoria", "Archivos", "Filtros", "Tareas"]);   // el cronograma se audita como un cambio del proyecto
  const IGNORAR_CAMPOS = new Set(["Actualizado_Por", "Actualizado_El"]);
  // Qué está haciendo Excel ahora (para el aviso de espera).
  const EN_CURSO = new Map();
  let seqOp = 0;
  const ETQ_OP = { leerTodo: "Leyendo todo el Excel", leerTabla: "Leyendo la tabla", agregarFila: "Agregando en", agregarFilas: "Agregando en", actualizarPorId: "Actualizando", actualizarVarios: "Actualizando",
    eliminarFila: "Eliminando en", eliminarFilas: "Eliminando en", guardarParte: "Subiendo archivo", borrarArchivo: "Borrando archivo", borrarArchivos: "Borrando archivos", partesArchivo: "Verificando archivo", leerParte: "Leyendo archivo" };
  function opsEnCurso() {
    const l = [...EN_CURSO.values()];
    if (!l.length) return "Esperando respuesta de Excel";
    return "Ahora: " + l.map((x) => `${x.etq} (${Math.round((Date.now() - x.t0) / 1000)} s)`).join(", ");
  }
  // Registro de errores de Excel (últimos 40) para diagnosticar fallas intermitentes.
  function registrarDiag(operacion, e, ms) {
    const d = { t: new Date().toLocaleString("es-CO"), op: operacion, cod: (e && e.code) || "", msg: String((e && e.message) || e || "").slice(0, 300), ms: Math.round(ms || 0), modo: S.modo || "" };
    try { const l = JSON.parse(localStorage.getItem("pmo_diag") || "[]"); l.unshift(d); localStorage.setItem("pmo_diag", JSON.stringify(l.slice(0, 40))); } catch (x) { /* sin almacenamiento */ }
    (S.diag = S.diag || []).unshift(d);
  }
  function verDiagnostico() {
    let l = [];
    try { l = JSON.parse(localStorage.getItem("pmo_diag") || "[]"); } catch (x) { l = S.diag || []; }
    const texto = l.map((d) => `${d.t} · ${d.op} · ${d.cod} · ${d.msg} · ${d.ms} ms · ${d.modo}`).join("\n");
    const m = document.createElement("div");
    m.id = "modal"; m.className = "modal";
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true"><div class="titulo-fila"><h2>Diagnóstico de conexión con Excel</h2><button class="btn enlace" id="m-cerrar">✕</button></div>
      <p class="sub">Errores recientes al leer o guardar en Excel en este equipo (versión ${esc(window.APP_VERSION || "")}). Si Excel muestra «Sorry, something went wrong», copia esta lista y envíala.</p>
      ${l.length ? `<div class="tabla-scroll"><table class="no-orden no-recortar"><thead><tr><th>Cuándo</th><th>Operación</th><th>Código</th><th>Mensaje</th><th>Duró</th></tr></thead><tbody>${l.map((d) => `<tr><td class="nowrap">${esc(d.t)}</td><td>${esc(d.op)}</td><td>${esc(d.cod)}</td><td>${esc(d.msg)}</td><td class="nowrap">${d.ms} ms</td></tr>`).join("")}</tbody></table></div>` : vacio("No se han registrado errores en este equipo.")}
      <div class="acciones derecha">${l.length ? `<button class="btn" id="dg-limpiar">Limpiar</button><button class="btn primario" id="dg-copiar">Copiar</button>` : ""}<button class="btn" id="dg-cerrar">Cerrar</button></div></div>`;
    cerrarModal(); document.body.appendChild(m);
    const cerrar = () => m.remove();
    m.querySelector("#m-cerrar").addEventListener("click", cerrar); m.querySelector("#dg-cerrar").addEventListener("click", cerrar);
    const cp = m.querySelector("#dg-copiar"); if (cp) cp.addEventListener("click", async () => { try { await navigator.clipboard.writeText(texto); toast("Copiado"); } catch (x) { toast("No se pudo copiar; selecciona el texto de la tabla.", true); } });
    const lp = m.querySelector("#dg-limpiar"); if (lp) lp.addEventListener("click", () => { try { localStorage.removeItem("pmo_diag"); } catch (x) { /* nada */ } S.diag = []; cerrar(); toast("Diagnóstico limpio"); });
  }
  function conMedicion(api) {
    const w = {};
    Object.keys(api).forEach((op) => {
      if (typeof api[op] !== "function") { w[op] = api[op]; return; }
      w[op] = async (...a) => {
        const id = ++seqOp;
        EN_CURSO.set(id, { etq: `${ETQ_OP[op] || op}${typeof a[0] === "string" && /^(agregar|actualizar|eliminar|leerTabla)/.test(op) ? " " + a[0] : ""}`, t0: Date.now() });
        try { return await api[op](...a); }
        catch (e) { registrarDiag(EN_CURSO.get(id) ? EN_CURSO.get(id).etq : op, e, Date.now() - (EN_CURSO.get(id) || { t0: Date.now() }).t0); throw e; }
        finally { EN_CURSO.delete(id); }
      };
    });
    return w;
  }
  function conAuditoria(api) {
    const pend = [];
    let timer = null;
    const buscar = (tabla, id) => (S.datos && S.datos[tabla] || []).find((r) => String(r[ID_COLS[tabla]]) === String(id));
    const desc = (o) => { const k = CAMPO_DESC.find((c) => o && o[c]); return k ? String(o[k]).slice(0, 140) : ""; };
    const corto = (v) => { const t = String(v ?? ""); return t.length > 80 ? t.slice(0, 80) + "…" : t; };
    const reg = (tabla, id, pid, accion, detalle) => {
      if (NO_AUDITAR.has(tabla) || !S.usuario) return;
      pend.push({ Fecha_Hora: ahora(), Usuario: S.usuario.Correo, Tabla: tabla, ID_Registro: String(id || ""), ID_Proyecto: pid || "", Accion: accion, Detalle: String(detalle || "").slice(0, 1500) });
      clearTimeout(timer); timer = setTimeout(vaciar, 600);
    };
    async function vaciar() {
      clearTimeout(timer);
      if (!pend.length) return;
      const lote = pend.splice(0);
      try { await api.agregarFilas("Auditoria", lote); S.auditoria = null; } catch (e) { /* la auditoría nunca bloquea el trabajo */ }
    }
    // Tras escribir en Excel, el mismo cambio se aplica a S.datos: así no hay que releer todo el libro en cada guardado.
    const val = (v) => (v === null || v === undefined ? "" : v);
    const local = {
      agregar(tabla, objs) {
        if (!S.datos) return;
        S.datos[tabla] = S.datos[tabla] || [];
        const modelo = S.datos[tabla][0] || {};
        objs.forEach((o) => {
          const fila = Object.fromEntries(Object.keys(modelo).map((k) => [k, ""]));
          Object.entries(o).forEach(([k, v]) => { fila[k] = val(v); });
          S.datos[tabla].push(fila);
        });
      },
      actualizar(tabla, colId, lista) {
        if (!S.datos || !S.datos[tabla]) return;
        lista.forEach(({ id, cambios }) => {
          const f = S.datos[tabla].find((r) => String(r[colId]) === String(id));
          if (f) Object.entries(cambios).forEach(([k, v]) => { f[k] = val(v); });
        });
      },
      eliminar(tabla, colId, ids) {
        if (!S.datos || !S.datos[tabla]) return;
        const set = new Set(ids.map(String));
        S.datos[tabla] = S.datos[tabla].filter((r) => !set.has(String(r[colId])));
      },
      eliminarUna(tabla, criterios) {
        if (!S.datos || !S.datos[tabla]) return;
        const i = S.datos[tabla].findIndex((r) => Object.entries(criterios).every(([k, v]) => String(r[k]) === String(v)));
        if (i >= 0) S.datos[tabla].splice(i, 1);
      },
    };
    const w = { ...api, vaciar };
    w.agregarFilas = async (tabla, objs) => {
      if (tabla === "Compromisos") objs.forEach((o) => { if (!o.Fecha_Creacion) o.Fecha_Creacion = R.hoyISO(); });   // fecha en que se creó, para ordenar
      const r = await api.agregarFilas(tabla, objs);
      local.agregar(tabla, objs);
      objs.forEach((o) => reg(tabla, o[ID_COLS[tabla]] || o.Valor || "", tabla === "Proyectos" ? o.ID_Proyecto : o.ID_Proyecto, "Creó", desc(o) || (tabla === "Catalogos" ? `${o.Lista}: ${o.Valor}` : "")));
      return r;
    };
    w.agregarFila = (tabla, obj) => w.agregarFilas(tabla, [obj]);
    w.actualizarVarios = async (tabla, colId, lista) => {
      const antes = lista.map(({ id }) => ({ ...(buscar(tabla, id) || {}) }));
      const r = await api.actualizarVarios(tabla, colId, lista);
      local.actualizar(tabla, colId, lista);
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
      local.eliminar(tabla, colId, ids);
      ids.forEach((id, i) => reg(tabla, id, tabla === "Proyectos" ? id : antes[i].ID_Proyecto, "Eliminó", desc(antes[i])));
      return r;
    };
    w.eliminarFila = async (tabla, criterios) => {
      const r = await api.eliminarFila(tabla, criterios);
      local.eliminarUna(tabla, criterios);
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
    Stakeholders: "Stakeholder", Sugerencias: "Sugerencia", Documentos: "Documento", LineasBase: "Línea base", Tareas: "Tarea", Cronogramas: "Versión de cronograma", Proveedores: "Proveedor", Recursos: "Recurso", Tickets: "Ticket", Cambios: "Cambio", Lecciones: "Lección", RACI: "RACI", Dependencias: "Dependencia", Catalogos: "Catálogo" };
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
      const hist = R.avancesDe(p.ID_Proyecto, S.datos.Seguimientos).filter((s) => s.Fecha_Corte <= fin);
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
        return `<tr><td><b>${t(p.Nombre)}</b>${p.Codigo_Almera ? `<br><small>Almera ${esc(p.Codigo_Almera)}</small>` : ""}</td><td>${t(nombreUsuario(p.PM))}</td><td>${chipSemaforo(p.Semaforo)}</td>
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
    return S.datos.Recursos.find((x) => x.ID_Recurso === r.ID_Recurso) || r;
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
  // Una vez: los compromisos que ya existían quedan enlazados a la persona del directorio (por correo o nombre).
  let compromisosEnlazados = false;
  async function vincularCompromisosRecursos() {
    if (compromisosEnlazados || !(S.datos.Recursos || []).length) return;
    compromisosEnlazados = true;
    const enlaces = S.datos.Compromisos.filter((c) => !c.ID_Recurso && String(c.Responsable || "").trim()).map((c) => {
      const rs = responsablesDe(c, false);
      if (rs.length === 1 && !rs[0].r && c.Correo_Responsable) { const r = recursoPor({ correo: c.Correo_Responsable }); if (r) rs[0] = { ...rs[0], r, ID_Recurso: r.ID_Recurso }; }
      if (!rs.some((x) => x.r)) return null;
      return { id: c.ID_Compromiso, cambios: serializarResp(rs.map((x) => (x.r ? { Nombre: x.r.Nombre, Correo: x.Correo || x.r.Correo || "", ID_Recurso: x.r.ID_Recurso } : x))) };
    }).filter(Boolean);
    if (!enlaces.length) return;
    try { await S.api.actualizarVarios("Compromisos", "ID_Compromiso", enlaces); } catch (e) { /* sin permiso de edición: se resuelve por nombre/correo */ }
  }
  const porCreacion = (a, b) => String(fechaCreacion(a) || "9999").localeCompare(String(fechaCreacion(b) || "9999")) || String(a.ID_Compromiso).localeCompare(String(b.ID_Compromiso), "es", { numeric: true });
  // Fecha de creación de un compromiso: la guardada; si es antiguo, la de su sesión o la de su primer comentario.
  function fechaCreacion(c) {
    if (c.Fecha_Creacion) return String(c.Fecha_Creacion).slice(0, 10);
    const s = c.ID_Seguimiento && S.datos.Seguimientos.find((x) => x.ID_Seguimiento === c.ID_Seguimiento);
    if (s && s.Fecha_Corte) return s.Fecha_Corte;
    const primero = (S.datos.Comentarios || []).filter((x) => x.ID_Compromiso === c.ID_Compromiso).map((x) => String(x.Fecha_Hora).slice(0, 10)).sort()[0];
    return primero || "";
  }
  // Una vez: los compromisos que ya existían quedan con su fecha de creación en el Excel.
  let fechasCompletadas = false;
  async function completarFechaCreacion() {
    if (fechasCompletadas) return;
    fechasCompletadas = true;
    const cambios = S.datos.Compromisos.filter((c) => !c.Fecha_Creacion && fechaCreacion(c)).map((c) => ({ id: c.ID_Compromiso, cambios: { Fecha_Creacion: fechaCreacion(c) } }));
    if (!cambios.length) return;
    try { await S.api.actualizarVarios("Compromisos", "ID_Compromiso", cambios); } catch (e) { /* sin permiso de edición: se calcula al vuelo */ }
  }
  // Responsables de un compromiso → personas del directorio (crea las nuevas).
  // Solo se crea en Recursos a quien se eligió con «+ Agregar como recurso nuevo» (crearNuevos); el texto libre queda como nombre, sin crear basura en el directorio.
  async function resolverResponsables(lista, crearNuevos) {
    const out = [];
    for (const x of lista) {
      const nombre = String(x.Nombre || "").trim();
      if (!nombre) continue;
      if (esGrupo(nombre)) { out.push({ Nombre: GRUPO_PROY, Correo: "", ID_Recurso: "" }); continue; }
      const ya = recursoPor({ id: x.ID_Recurso, correo: x.Correo, nombre });
      const crear = !ya && (crearNuevos === true || (crearNuevos instanceof Set && crearNuevos.has(normTxt(nombre))));
      const r = ya ? await asegurarRecurso({ ...x, Nombre: ya.Nombre, ID_Recurso: ya.ID_Recurso }) : crear ? await asegurarRecurso({ Nombre: nombre, Correo: x.Correo }) : null;
      out.push(r ? { Nombre: r.Nombre, Correo: x.Correo || r.Correo || "", ID_Recurso: r.ID_Recurso } : { Nombre: nombre, Correo: x.Correo || correoSuelto(nombre), ID_Recurso: "" });
    }
    return serializarResp(out);
  }
  // «Ana Pérez / Luis Gómez», «Ana; Luis», «Ana y Luis» → varios nombres.
  const separarNombres = (t) => String(t || "").split(/\s*(?:;|\/|\s+y\s+|\s+&\s+)\s*/i).map((x) => x.trim()).filter(Boolean);
  // Rol de la persona en el proyecto (si es stakeholder).
  const stakeholderDe = (pid, x) => (S.datos.Stakeholders || []).find((s) => s.ID_Proyecto === pid && ((x.ID_Recurso && s.ID_Recurso === x.ID_Recurso) || normTxt(s.Nombre) === normTxt(x.Nombre)));
  const rolEnProyecto = (pid, x) => (stakeholderDe(pid, x) || {}).Rol || "";
  // Cambios en un recurso se reflejan donde aparece (stakeholders y responsables de compromisos).
  async function propagarRecurso(antes, r) {
    const stk = (S.datos.Stakeholders || []).filter((x) => x.ID_Recurso === r.ID_Recurso || (!x.ID_Recurso && ((antes.Correo && lc(x.Correo) === lc(antes.Correo)) || normTxt(x.Nombre) === normTxt(antes.Nombre))));
    if (stk.length) await S.api.actualizarVarios("Stakeholders", "ID_Stakeholder", stk.map((x) => ({ id: x.ID_Stakeholder,
      cambios: { ID_Recurso: r.ID_Recurso, Nombre: r.Nombre, Correo: r.Correo, Cargo: r.Cargo, Area: r.Area, Telefono: r.Telefono, ID_Proveedor: r.ID_Proveedor } })));
    if (antes.Nombre === r.Nombre && antes.Correo === r.Correo) return;
    const es = (x) => x.ID_Recurso === r.ID_Recurso || (!x.ID_Recurso && ((antes.Correo && lc(x.Correo) === lc(antes.Correo)) || (antes.Nombre && normTxt(x.Nombre) === normTxt(antes.Nombre))));
    const cmp = S.datos.Compromisos.map((c) => ({ c, rs: responsablesDe(c, false) })).filter(({ rs }) => rs.some(es));
    if (cmp.length) await S.api.actualizarVarios("Compromisos", "ID_Compromiso", cmp.map(({ c, rs }) => ({ id: c.ID_Compromiso,
      cambios: serializarResp(rs.map((x) => (es(x) ? { Nombre: r.Nombre, Correo: r.Correo || "", ID_Recurso: r.ID_Recurso } : x))) })));
  }
  const dupRecurso = (fd, excluir) => (S.datos.Recursos || []).find((r) => r.ID_Recurso !== excluir && ((fd.Correo && lc(r.Correo) === lc(fd.Correo)) || normTxt(r.Nombre) === normTxt(fd.Nombre)));
  function formRecurso(r, alCrear) {
    const nuevo = !r;
    const empresas = [["", "Interno (FSFB)"], ...proveedoresActivos(r && r.ID_Proveedor).map((v) => [v.ID_Proveedor, v.Nombre])];
    const usos = nuevo ? { stk: [], cmp: [] } : {
      stk: (S.datos.Stakeholders || []).filter((x) => x.ID_Recurso === r.ID_Recurso || (r.Correo && lc(x.Correo) === lc(r.Correo))),
      cmp: S.datos.Compromisos.filter((c) => responsablesDe(c).some((x) => x.ID_Recurso === r.ID_Recurso)),
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
    }, nuevo ? {} : { antes: proyectosRecursoHTML(r), init: (m) => m.querySelectorAll("[data-rp-ir]").forEach((b) => b.addEventListener("click", () => { cerrarModal(); ir("ficha", b.dataset.rpIr); })), eliminar: { texto: "Eliminar recurso",
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
    const fp = S.recP || "";
    const lista = recursos().filter((r) => (!q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${r.Area} ${empresaDe(r)}`).includes(normTxt(q))) && (!fp || proyectosDeRecurso(r).some((x) => x.p.ID_Proyecto === fp)));
    const etiquetasProy = (r) => proyectosDeRecurso(r).map(({ p, roles }) => `<button class="chip-proy-rec ${/cerrad|finaliz|cancel/i.test(p.Estado || "") ? "inactivo" : ""}" data-pid="${esc(p.ID_Proyecto)}" title="${esc(`${p.Nombre}: ${roles.join(", ")}`)}">${esc(p.Nombre)}</button>`).join("") || `<span class="sub">Sin proyectos</span>`;
    const dups = gruposDuplicados();
    const dudosos = recursosDudosos();
    el.innerHTML = `<div class="titulo-fila"><h1>Recursos</h1>${puede ? `<button class="btn primario" id="b-rec">+ Recurso</button>` : ""}</div>
      ${dudosos.length && puede ? `<div class="card aviso-dup"><h2>Registros que parecen varias personas en uno <span class="contador rojo">${dudosos.length}</span></h2>
        <p class="sub">Se crearon con texto como «Ana / Luis». Al eliminarlos, los compromisos que los tenían quedan con cada persona por separado (si ya existe en el directorio).</p>
        <ul class="lista">${dudosos.map((r) => `<li><label class="check"><input type="checkbox" class="rec-dud" value="${esc(r.ID_Recurso)}" ${separarNombres(r.Nombre).length > 1 ? "checked" : ""}> <b>${esc(r.Nombre)}</b> <span class="sub">${esc([r.Correo, separarNombres(r.Nombre).length > 1 ? `→ ${separarNombres(r.Nombre).join(" · ")}` : ""].filter(Boolean).join(" "))}</span></label></li>`).join("")}</ul>
        <button class="btn" id="rec-dud-borrar">Eliminar los marcados</button></div>` : ""}
      <div class="filtros"><label class="filtro crece"><span>Buscar</span><input id="rec-q" type="search" value="${esc(q)}" placeholder="Nombre, correo, cargo, área o empresa"></label>
        <label class="filtro"><span>Proyecto</span><select id="rec-p"><option value="">Todos</option>${visibles().slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es")).map((p) => `<option value="${esc(p.ID_Proyecto)}" ${p.ID_Proyecto === fp ? "selected" : ""}>${esc(p.Nombre)}</option>`).join("")}</select></label></div>
      ${dups.length && puede ? `<div class="card aviso-dup"><h2>Posibles duplicados <span class="contador rojo">${dups.length}</span></h2>
        <ul class="lista">${dups.map((g, i) => `<li>${g.map((r) => `<b>${esc(r.Nombre)}</b>${r.Correo ? ` <span class="sub">${esc(r.Correo)}</span>` : ""}`).join(" · ")} <button class="btn chico" data-fusionar="${i}">Fusionar</button></li>`).join("")}</ul></div>` : ""}
      <div class="card"><p class="sub">Directorio único de personas del portafolio. En stakeholders, equipo y responsables de compromisos se eligen de aquí y sus datos se llenan solos; a alguien nuevo lo agregas con «+ Agregar como recurso nuevo» o con «+ Recurso».</p>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Persona</th><th>Proyectos</th><th>Empresa</th><th>Área</th><th>Contacto</th><th>Asignado</th><th></th></tr></thead><tbody>
          ${lista.map((r) => { const a = asig(r), c = num0(r.Capacidad) ?? 100; return `<tr class="${r.Activo === "No" ? "inactivo" : ""}"><td${ed({ puede, tabla: "Recursos", idCol: "ID_Recurso", id: r.ID_Recurso, campo: "Nombre", tipo: "text", requerido: true, valor: r.Nombre, etiqueta: "Nombre", despues: (v) => propagarRecurso(r, { ...r, Nombre: v }) })}>${persona(r.Nombre, r.Cargo)}${r.Activo === "No" ? ` <span class="pill noaplica">Inactivo</span>` : ""}</td>
            <td class="rec-proys">${etiquetasProy(r)}</td><td${ed({ puede, tabla: "Recursos", idCol: "ID_Recurso", id: r.ID_Recurso, campo: "ID_Proveedor", tipo: "select", opciones: [["", "Interno (FSFB)"], ...proveedoresActivos(r.ID_Proveedor).map((v) => [v.ID_Proveedor, v.Nombre])], valor: r.ID_Proveedor || "", etiqueta: "Empresa", despues: (v) => propagarRecurso(r, { ...r, ID_Proveedor: v }) })}>${esc(empresaDe(r))}</td><td${ed({ puede, tabla: "Recursos", idCol: "ID_Recurso", id: r.ID_Recurso, campo: "Area", tipo: "select", opciones: S.cat.Cliente_Area, vacio: "—", valor: r.Area, etiqueta: "Área" })}>${esc(r.Area || "—")}</td><td>${contacto(r.Correo, r.Telefono)}</td>
            <td class="nowrap"${ed({ puede, tabla: "Recursos", idCol: "ID_Recurso", id: r.ID_Recurso, campo: "Capacidad", tipo: "number", min: 0, max: 100, requerido: true, valor: r.Capacidad ?? "", etiqueta: "Capacidad" })}>${a ? `${a.total}% de ${c}%${a.total > c ? ` <span class="baja">⚠</span>` : ""}<div class="sub">${a.asign.length} asignación(es)</div>` : `<span class="sub">— de ${c}%</span>`}</td>
            <td class="derecha">${puede ? `<button class="btn chico" data-rec="${esc(r.ID_Recurso)}">Editar</button>` : ""}</td></tr>`; }).join("")}</tbody></table></div>`
        : vacio(q ? "Nadie coincide con la búsqueda." : "Aún no hay recursos. Se crean solos con las personas que ya existen o con «+ Recurso».")}</div>`;
    const qi = $("#rec-q"); qi.addEventListener("change", () => { S.recQ = qi.value; render(); });
    $("#rec-p").addEventListener("change", (e) => { S.recP = e.target.value; render(); });
    if ($("#b-rec")) $("#b-rec").addEventListener("click", () => formRecurso());
    el.querySelectorAll("[data-rec]").forEach((b) => b.addEventListener("click", () => formRecurso(recursoId(b.dataset.rec))));
    el.querySelectorAll("[data-fusionar]").forEach((b) => b.addEventListener("click", () => fusionar(dups[Number(b.dataset.fusionar)])));
    if ($("#rec-dud-borrar")) $("#rec-dud-borrar").addEventListener("click", () => {
      const ids = [...el.querySelectorAll(".rec-dud:checked")].map((i) => i.value);
      if (!ids.length) return toast("Marca al menos uno.", true);
      confirmar("Eliminar registros", `Se eliminarán ${ids.length} registro(s) del directorio de Recursos.`, "Eliminar", () => guardar(async () => {
        const set = new Set(ids);
        // Compromisos que apuntaban a esos registros: cada nombre se vuelve a enlazar por separado.
        const cambios = S.datos.Compromisos.map((c) => ({ c, rs: responsablesDe(c, false) })).filter(({ rs }) => rs.some((x) => set.has(x.ID_Recurso)))
          .map(({ c, rs }) => ({ id: c.ID_Compromiso, cambios: serializarResp(rs.flatMap((x) => (set.has(x.ID_Recurso) ? separarNombres(x.Nombre) : [x.Nombre]).map((n) => {
            const r = recursoPor({ nombre: n }); return r && !set.has(r.ID_Recurso) ? { Nombre: r.Nombre, Correo: r.Correo || "", ID_Recurso: r.ID_Recurso } : { Nombre: n, Correo: correoSuelto(n), ID_Recurso: "" }; }))) }));
        await S.api.eliminarFilas("Recursos", "ID_Recurso", ids);
        if (cambios.length) await S.api.actualizarVarios("Compromisos", "ID_Compromiso", cambios);
      }, `${ids.length} registro(s) eliminados`));
    });
  }
  // En qué proyectos (visibles para el usuario) está cada persona y con qué papel.
  // Proyectos del recurso dentro de su ficha: rol, dedicación, estado, avance y compromisos abiertos.
  function proyectosRecursoHTML(r) {
    const l = proyectosDeRecurso(r);
    const ded = (pid) => { const x = (S.datos.Stakeholders || []).find((y) => y.ID_Proyecto === pid && ((y.ID_Recurso && y.ID_Recurso === r.ID_Recurso) || (r.Correo && y.Correo && lc(y.Correo) === lc(r.Correo)))); const pr = proyecto(pid); return x && num0(x.Dedicacion) !== null ? `${num0(x.Dedicacion)}%` : pr && r.Correo && lc(pr.PM) === lc(r.Correo) && num0(pr.Dedicacion_PM) !== null ? `${num0(pr.Dedicacion_PM)}%` : "—"; };
    const activos = l.filter((x) => x.p.Estado === "Activo").length;
    return `<div class="rp-caja"><div class="rp-cab"><b>Proyectos asociados</b> <span class="sub">${l.length ? `${l.length} proyecto(s) · ${activos} activo(s)` : "Aún no está en ningún proyecto"}</span></div>
      ${l.length ? `<div class="tabla-scroll"><table class="no-orden no-recortar rp-tabla"><thead><tr><th>Proyecto</th><th>Rol</th><th>Dedicación</th><th>Estado</th><th>Avance</th></tr></thead><tbody>
        ${l.map(({ p, roles }) => `<tr><td><button type="button" class="btn enlace" data-rp-ir="${esc(p.ID_Proyecto)}" title="Abrir el proyecto">${esc(p.Nombre)}</button><div class="sub">${esc(p.ID_Proyecto)}</div></td>
          <td>${roles.map((x) => `<span class="pill rol">${esc(x)}</span>`).join(" ")}</td><td>${ded(p.ID_Proyecto)}</td><td>${chipEstadoP(p.Estado)}</td><td>${p.Estado === "Activo" ? barraAvance(p.Avance_Real, p.Avance_Planeado) : pct(p.Avance_Real)}</td></tr>`).join("")}</tbody></table></div>` : ""}</div>
      ${actividadesRecursoHTML(r)}`;
  }
  // Otras actividades del recurso (vigentes y programadas) con su carga total a hoy.
  function actividadesRecursoHTML(r) {
    const hoy = R.hoyISO();
    const l = actividades().filter((a) => (a.ID_Recurso && a.ID_Recurso === r.ID_Recurso) || (r.Correo && a.Correo && lc(a.Correo) === lc(r.Correo)) || (!a.Correo && normTxt(a.Nombre) === normTxt(r.Nombre)))
      .filter((a) => estadoAct(a, hoy) !== "Finalizada").sort((a, b) => String(a.Fecha_Inicio).localeCompare(String(b.Fecha_Inicio)));
    const c = capacidad(hoy).find((x) => x.k === clavePersona(r.Correo, r.Nombre));
    return `<div class="rp-caja"><div class="rp-cab"><b>Otras actividades</b> <span class="sub">${l.length ? `${l.length} registrada(s)` : "Ninguna registrada"}${c ? ` · carga hoy ${c.tProy}% proyectos + ${c.tOtras}% otras = <b>${c.total}%</b> de ${c.cap}%` : ""}</span></div>
      ${l.length ? `<div class="tabla-scroll"><table class="no-orden no-recortar rp-tabla"><thead><tr><th>Tipo</th><th>Descripción</th><th>Dedicación</th><th>Vigencia</th><th>Estado</th></tr></thead><tbody>
        ${l.map((a) => `<tr><td><span class="chip-act ${esAusencia(a) ? "aus" : ""}">${esc(a.Tipo || "—")}</span></td><td>${esc(a.Descripcion || "—")}</td><td>${num0(a.Dedicacion) ?? "—"}%</td><td class="nowrap">${fecha(a.Fecha_Inicio)} → ${a.Fecha_Fin ? fecha(a.Fecha_Fin) : "Permanente"}</td><td>${pill(estadoAct(a, hoy))}</td></tr>`).join("")}</tbody></table></div>` : ""}
      <p class="sub">Se registran en «Capacidad del equipo» → Otras actividades.</p></div>`;
  }
  function proyectosDeRecurso(r) {
    const mapa = new Map();
    const add = (pid, rol) => { const p = proyecto(pid); if (!p || !R.esMio(S.usuario, p)) return; const x = mapa.get(pid) || { p, roles: [] }; if (!x.roles.includes(rol)) x.roles.push(rol); mapa.set(pid, x); };
    const mismo = (id, correo, nombre) => (id && id === r.ID_Recurso) || (r.Correo && correo && lc(correo) === lc(r.Correo)) || (!id && nombre && normTxt(nombre) === normTxt(r.Nombre));
    S.datos.Proyectos.forEach((p) => { if (r.Correo && lc(p.PM) === lc(r.Correo)) add(p.ID_Proyecto, "PM"); });
    (S.datos.Stakeholders || []).forEach((x) => { if (mismo(x.ID_Recurso, x.Correo, x.Nombre)) add(x.ID_Proyecto, x.Rol || "Equipo"); });
    const nComp = {};
    S.datos.Compromisos.forEach((c) => { if (!R.cerrado(c) && responsablesDe(c).some((x) => mismo(x.ID_Recurso, x.Correo, x.Nombre))) nComp[c.ID_Proyecto] = (nComp[c.ID_Proyecto] || 0) + 1; });
    Object.entries(nComp).forEach(([pid, n]) => add(pid, `${n} compromiso${n > 1 ? "s" : ""} abierto${n > 1 ? "s" : ""}`));
    return [...mapa.values()].sort((a, b) => String(a.p.Nombre).localeCompare(String(b.p.Nombre), "es"));
  }
  // Recursos creados a partir de texto con varias personas («Ana / Luis», «Ana; Luis», «Ana y Luis»).
  function recursosDudosos() {
    const L = S.datos.Recursos || [];
    return L.filter((r) => {
      const n = String(r.Nombre || "");
      if (/[\/;]|\s&\s/.test(n) || separarNombres(n).length > 1) return true;
      // Nombre que contiene completo a otra persona del directorio y algo más (p. ej. «Ana Pérez Luis Gómez»).
      return n.split(/\s+/).length >= 5 && L.some((o) => o !== r && String(o.Nombre).split(/\s+/).length >= 2 && normTxt(n).includes(normTxt(o.Nombre)) && normTxt(n) !== normTxt(o.Nombre));
    });
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
  const opcionesSesion = (pid) => R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse().map((s) => [s.ID_Seguimiento, `${R.esSesion(s) ? (s.Nombre_Sesion ? `${s.Nombre_Sesion} · ` : "Sesión del ") : "Seguimiento del "}${fecha(s.Fecha_Corte)}${s.Fecha_Acta && s.Fecha_Acta !== s.Fecha_Corte ? ` (acta ${fecha(s.Fecha_Acta)})` : ""}`]);

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
    const multi = input.hasAttribute("data-multi");
    // En modo varios («A; B»), se busca y se reemplaza solo el último nombre.
    const segmentos = () => input.value.split(";");
    const actualTxt = () => (multi ? segmentos().pop() : input.value).trim();
    const elegir = (r, nuevo, esG) => {
      if (multi) {
        const previos = segmentos().slice(0, -1).map((x) => x.trim()).filter(Boolean);
        const nombre = r ? r.Nombre : nuevo;
        if (!r && !esG) (input._nuevos = input._nuevos || new Set()).add(normTxt(nombre));
        input.value = [...previos.filter((x) => normTxt(x) !== normTxt(nombre)), nombre].join("; ") + "; ";
        lista.hidden = true;
        input.dispatchEvent(new Event("change", { bubbles: true }));
        input.focus();
        return;
      }
      if (!r && !esG) (input._nuevos = input._nuevos || new Set()).add(normTxt(nuevo));
      input.value = r ? r.Nombre : nuevo;
      input.dataset.idRecurso = r ? r.ID_Recurso : "";
      lista.hidden = true;
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.dispatchEvent(new CustomEvent("elegido", { detail: { r, nuevo, grupo: !!esG } }));
    };
    const pintar = () => {
      const q = normTxt(actualTxt());
      const ya = multi ? new Set(segmentos().slice(0, -1).map((x) => normTxt(x.trim()))) : input._excluir ? input._excluir() : new Set();
      const rs = recursos().filter((r) => r.Activo !== "No" && !ya.has(normTxt(r.Nombre)) && (!q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${empresaDe(r)}`).includes(q))).slice(0, 8);
      const exacto = rs.some((r) => normTxt(r.Nombre) === q);
      const integ = input._pid ? integrantesGrupo(input._pid) : [];
      const conGrupo = integ.length && !ya.has(normTxt(GRUPO_PROY)) && (!q || normTxt("grupo de proyecto equipo del proyecto todos stakeholders").includes(q));
      items = [...(conGrupo ? [{ grupo: true, n: integ.length }] : []), ...rs.map((r) => ({ r })), ...(q && !exacto && !/[;\/]/.test(actualTxt()) ? [{ nuevo: actualTxt() }] : [])];
      activo = items.length ? 0 : -1;
      lista.innerHTML = items.map((it, i) => it.grupo
        ? `<li role="option" data-i="${i}" class="grupo-opt ${i === activo ? "activo" : ""}"><b>👥 ${GRUPO_PROY}</b><span class="sub">Los ${it.n} stakeholders configurados en el proyecto</span></li>`
        : it.r
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
      else if (e.key === "Enter" && activo >= 0) { e.preventDefault(); const it = items[activo]; elegir(it.r, it.grupo ? GRUPO_PROY : it.nuevo, it.grupo); }
      else if (e.key === "Escape") { e.stopPropagation(); lista.hidden = true; }
    });
    lista.addEventListener("mousedown", (e) => {
      const li = e.target.closest("li[data-i]");
      if (!li) return;
      e.preventDefault();
      const it = items[Number(li.dataset.i)];
      elegir(it.r, it.grupo ? GRUPO_PROY : it.nuevo, it.grupo);
    });
    input.addEventListener("blur", () => setTimeout(() => { lista.hidden = true; }, 150));
    // Al envolver el campo se pierde el foco: se recupera y se muestra la lista de una vez.
    if (document.activeElement !== input) input.focus(); else pintar();
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
    agenda: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01"/>',
    seguimiento: '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="m9 14 2 2 4-4"/>',
    proyectos: '<path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
    cronograma: '<path d="M3 3v18h18"/><rect x="7" y="6" width="9" height="3" rx="1"/><rect x="10" y="11" width="8" height="3" rx="1"/><rect x="6" y="16" width="6" height="3" rx="1"/>',
    riesgos: '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    recursos: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
    capacidad: '<path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>',
    priorizacion: '<path d="M3 3v18h18"/><path d="M7 16h4"/><path d="M7 11h8"/><path d="M7 6h12"/>',
    valor: '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
    lecciones: '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
    sugerencias: '<path d="M9 18h6"/><path d="M10 22h4"/><path d="M12 2a7 7 0 0 0-4 12.7c.6.5 1 1.3 1 2.1V17h6v-.2c0-.8.4-1.6 1-2.1A7 7 0 0 0 12 2z"/>',
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
    documento: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/><path d="M8 13h8M8 17h5"/>',
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
    ["Inicio", ["dashboard", "avances", "seguimiento", "agenda"]],
    ["Portafolio", ["backlog", "proyectos", "cronograma", "riesgos", "priorizacion", "valor"]],
    ["Personas", ["recursos", "capacidad"]],
    ["Conocimiento", ["lecciones", "sugerencias"]],
    ["Administración", ["catalogos", "usuarios", "auditoria"]],
  ];
  const NOMBRE_TAB = { resumen: "Resumen", valor: "Valor", compromisos: "Compromisos", seguimientos: "Seguimientos", hitos: "Hitos", gantt: "Cronograma", riesgos: "Riesgos", tickets: "Tickets", documentos: "Documentos", cambios: "Cambios", raci: "RACI", lecciones: "Lecciones", auditoria: "Historial de cambios" };
  function migas() {
    const v = VISTAS.find((x) => x.id === (S.vista === "ficha" ? "proyectos" : S.vista));
    const grupo = (GRUPOS_NAV.find((g) => g[1].includes(v ? v.id : "")) || ["Inicio"])[0];
    const partes = [`<span>${esc(grupo)}</span>`, S.vista === "ficha" ? `<button class="btn enlace" data-miga="proyectos">Proyectos</button>` : `<span>${esc(v ? v.t : "")}</span>`];
    if (S.vista === "ficha") { const p = proyecto(S.pid); if (p) partes.push(`<span>${esc(p.Nombre)}</span>`, `<span>${esc(NOMBRE_TAB[S.fichaTab] || "Resumen")}</span>`); }
    return `<nav class="migas" aria-label="Ubicación">${partes.join('<span class="sep">›</span>')}</nav>`;
  }

  // Abrir un registro haciendo clic en cualquier parte de su fila o tarjeta (los botones siguen igual).
  // Se usa el botón principal de la fila (Abrir, Ver o Editar). Las filas de proyectos ya abren la ficha.
  const ABRIR_SEL = ["[data-edit-seg]", "a[data-doc-ir]", "[data-abrir-comp]", "[data-al-abrir]", "[data-sg-ver]", "[data-tk]", "[data-riesgo]", "[data-hito]", "[data-stk]", "[data-prov]", "[data-rec]", "[data-u]",
    "[data-leccion]", "[data-dep]", "[data-cambio]"];
  const ITEM_SEL = "#vista tbody tr, #vista .rg-card, #modal .alerta-item, #vista .alerta-item";
  document.addEventListener("click", (e) => {
    if (e.button !== 0 || !e.target.closest) return;
    if (e.target.closest("button, a, input, select, textarea, label, summary, details, [contenteditable], .combo, [data-pid]")) return;
    const item = e.target.closest(ITEM_SEL);
    if (!item) return;
    if (String(window.getSelection ? window.getSelection() : "").trim()) return;   // estaba seleccionando texto para copiar
    const b = ABRIR_SEL.map((sel) => item.querySelector(sel)).find(Boolean);   // en orden de prioridad (la sesión antes que sus compromisos)
    if (b && !b.disabled) b.click();
  });

  // ---------- Listas largas: recortar, minimizar tarjetas y secciones ----------
  // Se aplica solo a todo lo que aparece en pantalla (vistas, formularios y fichas), sin tocar cada módulo.
  const LIM_LISTA = 6, LIM_TABLA = 12;
  const plegadas = () => leerLocal("pmo_plegadas", []);
  const clavePliegue = (titulo) => `${S.vista}|${titulo}`;
  // Campo de varios responsables como fichas: se ve cada persona y se agregan desde una lista del directorio.
  // El campo original queda oculto y guarda «A; B», así el resto del código sigue igual.
  function selectorMultiple(input) {
    input.dataset.chips = "1";
    input.removeAttribute("data-recurso");
    input.hidden = true;
    const caja = document.createElement("div");
    caja.className = "multi-resp";
    caja.innerHTML = `<div class="mr-chips"></div><input class="mr-buscar" data-recurso autocomplete="off" placeholder="+ Agregar" aria-label="Agregar responsable">`;
    input.after(caja);
    const chips = caja.querySelector(".mr-chips"), buscar = caja.querySelector(".mr-buscar");
    buscar._pid = input.dataset.pid || S.pidForm || "";
    const nombres = () => separarNombres(input.value);
    const fijar = (l) => { input.value = l.join("; "); input.dispatchEvent(new Event("change", { bubbles: true })); pintar(); };
    function pintar() {
      chips.innerHTML = nombres().map((n, i) => { if (esGrupo(n)) { const g = integrantesGrupo(buscar._pid);
          return `<span class="mr-chip grupo" title="${esc(g.map((x) => `${x.Nombre}${x.Rol ? ` (${x.Rol})` : ""}`).join(", ") || "Sin stakeholders en el proyecto")}">👥 ${GRUPO_PROY} (${g.length})<button type="button" data-mr-quitar="${i}" aria-label="Quitar el grupo">✕</button></span>`; }
        const r = recursoPor({ nombre: n });
        return `<span class="mr-chip ${r ? "" : "libre"}" title="${esc(r ? [r.Cargo, r.Correo].filter(Boolean).join(" · ") || n : "No está en el directorio de Recursos")}">${esc(n)}<button type="button" data-mr-quitar="${i}" aria-label="Quitar a ${esc(n)}">✕</button></span>`; }).join("");
      chips.querySelectorAll("[data-mr-quitar]").forEach((b) => b.addEventListener("click", (e) => { e.stopPropagation(); const l = nombres(); l.splice(Number(b.dataset.mrQuitar), 1); fijar(l); }));
    }
    buscar.addEventListener("elegido", (e) => {
      const { r, nuevo, grupo } = e.detail;
      const n = r ? r.Nombre : nuevo;
      if (!n) return;
      if (!r && !grupo) (input._nuevos = input._nuevos || new Set()).add(normTxt(n));
      const l = nombres();
      if (!l.some((x) => normTxt(x) === normTxt(n))) l.push(n);
      buscar.value = "";
      fijar(l);
      setTimeout(() => { buscar.focus(); buscar.dispatchEvent(new Event("input")); }, 0);   // la lista se vuelve a abrir para elegir otra persona
    });
    buscar._excluir = () => new Set(nombres().map(normTxt));   // en la lista no salen los ya elegidos
    buscar.addEventListener("keydown", (e) => { if (e.key === "Backspace" && !buscar.value) { const l = nombres(); if (l.length) { l.pop(); fijar(l); } } });
    caja.addEventListener("click", (e) => { if (e.target === caja || e.target === chips) buscar.focus(); });
    pintar();
  }
  function compactar(raiz) {
    raiz.querySelectorAll("input[data-multi]:not([data-chips])").forEach(selectorMultiple);
    raiz.querySelectorAll("textarea.c-texto:not([data-alto])").forEach((t) => { if (t.offsetParent) { t.dataset.alto = "1"; autoAlto(t); } });
    // 1) Listas con muchos elementos: se ven los primeros y un botón «Ver todos».
    raiz.querySelectorAll("ul.lista, ul.mini, ul.eq-lista, ul.alertas, ul.cat-lista, ul.deps").forEach((ul) => {
      const n = [...ul.children].filter((li) => li.tagName === "LI").length;
      if (!ul._verTodos) {
        if (n <= LIM_LISTA + 1) return;
        ul.classList.add("recortada");
        const b = document.createElement("button");
        b.type = "button"; b.className = "btn enlace ver-todos";
        b.addEventListener("click", (e) => { e.stopPropagation(); ul.classList.toggle("recortada"); ul._pintarVer(); });
        ul._verTodos = b;
        ul.after(b);
      }
      // La lista puede cambiar (p. ej. al agregar una persona): se recalcula el botón.
      ul._pintarVer = () => { const k = [...ul.children].filter((li) => li.tagName === "LI").length; if (k <= LIM_LISTA + 1) ul.classList.remove("recortada"); ul._verTodos.hidden = k <= LIM_LISTA + 1; ul._verTodos.textContent = ul.classList.contains("recortada") ? `Ver todos (${k}) ▾` : "Ver menos ▴"; };
      ul._pintarVer();
    });
    // 2) Tablas largas en las vistas: se ven las primeras filas (sigue funcionando el orden por columna).
    raiz.querySelectorAll("#vista table:not(.raci):not(.no-recortar) > tbody").forEach((tb) => {
      const tabla = tb.closest("table");
      if (tabla.dataset.compacto || tb.querySelector("tr.grupo")) return;
      const n = tb.rows.length;
      if (n <= LIM_TABLA + 3) return;
      tabla.dataset.compacto = "1";
      tabla.classList.add("recortada");
      const b = document.createElement("button");
      b.type = "button"; b.className = "btn chico ver-todos-tabla";
      const txt = () => (tabla.classList.contains("recortada") ? `Ver las ${n} filas ▾` : "Ver menos ▴");
      b.textContent = txt();
      b.addEventListener("click", () => { tabla.classList.toggle("recortada"); b.textContent = txt(); });
      (tabla.closest(".tabla-scroll") || tabla).after(b);
    });
    // 3) Tarjetas con título: botón para minimizarlas (se recuerda en este equipo).
    raiz.querySelectorAll("#vista .card").forEach((card) => {
      if (card.dataset.plegable) return;
      const h = card.querySelector(":scope > .titulo-fila h2, :scope > h2");
      if (!h) return;
      card.dataset.plegable = "1";
      const titulo = h.textContent.trim().replace(/\s+\d+$/, "");
      const b = document.createElement("button");
      b.type = "button"; b.className = "btn-plegar"; b.title = "Minimizar o expandir"; b.setAttribute("aria-label", `Minimizar ${titulo}`);
      const pintar = () => { const p = card.classList.contains("plegada"); b.textContent = p ? "▸" : "▾"; b.setAttribute("aria-expanded", String(!p)); };
      if (plegadas().includes(clavePliegue(titulo))) card.classList.add("plegada");
      pintar();
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        card.classList.toggle("plegada"); pintar();
        const k = clavePliegue(titulo), l = plegadas().filter((x) => x !== k);
        guardarLocal("pmo_plegadas", card.classList.contains("plegada") ? [...l, k] : l);
      });
      h.prepend(b);
    });
    // 4) Secciones de los formularios: clic en el título para minimizar la sección.
    raiz.querySelectorAll("#modal .form-grid > .form-seccion").forEach((sec) => {
      if (sec.dataset.plegable) return;
      sec.dataset.plegable = "1";
      sec.classList.add("plegable");
      sec.setAttribute("role", "button"); sec.tabIndex = 0; sec.title = "Clic para minimizar o expandir esta sección";
      const alternar = () => {
        const cerrar = !sec.classList.contains("plegada");
        sec.classList.toggle("plegada", cerrar);
        let el = sec.nextElementSibling;
        while (el && !el.classList.contains("form-seccion")) { el.classList.toggle("oculto-seccion", cerrar); el = el.nextElementSibling; }
      };
      sec.addEventListener("click", (e) => { if (!e.target.closest("button, a, input, select")) alternar(); });
      sec.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); alternar(); } });
    });
  }
  // Campos de compromiso: el texto largo baja de renglón y la caja crece sola.
  const autoAlto = (t) => { t.style.height = "auto"; t.style.height = `${t.scrollHeight + 2}px`; };
  document.addEventListener("input", (e) => { if (e.target.matches && e.target.matches("textarea.c-texto")) autoAlto(e.target); });
  document.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target.matches && e.target.matches("textarea.c-texto")) e.preventDefault(); });   // un compromiso es un solo párrafo
  let compactarPend = false;
  new MutationObserver(() => {
    if (compactarPend) return;
    compactarPend = true;
    requestAnimationFrame(() => { compactarPend = false; compactar(document); });
  }).observe(document.body || document.documentElement, { childList: true, subtree: true });

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
    (S.datos.Documentos || []).filter((d) => ids.has(d.ID_Proyecto)).forEach((d) => out.push({ tipo: "Documentos", ico: "documento", t: d.Nombre || d.URL, sub: `${nom(d.ID_Proyecto)} · ${d.Tipo || "Documento"}${d.Descripcion ? ` · ${d.Descripcion}` : ""}`, ir: () => window.open(d.URL, "_blank", "noopener") }));
    (S.datos.Lecciones || []).filter((l) => ids.has(l.ID_Proyecto)).forEach((l) => out.push({ tipo: "Lecciones", ico: "lecciones", t: l.Leccion, sub: `${nom(l.ID_Proyecto)} · ${l.Categoria || ""}`, ir: () => { S.fichaTab = "lecciones"; ir("ficha", l.ID_Proyecto); } }));
    return out;
  }
  function abrirBusqueda() {
    if (!S.usuario || $("#paleta")) return;
    const indice = indiceBusqueda();
    const o = document.createElement("div");
    o.id = "paleta";
    o.innerHTML = `<div class="paleta-caja" role="dialog" aria-modal="true" aria-label="Búsqueda global">
      <div class="paleta-in">${ico("buscar")}<input id="paleta-q" placeholder="Buscar proyectos, documentos, personas, compromisos, riesgos o pantallas…" autocomplete="off"><kbd>Esc</kbd></div>
      <div id="paleta-res" class="paleta-res" role="listbox"></div>
      <div class="paleta-pie"><span><kbd>↑</kbd><kbd>↓</kbd> moverse</span><span><kbd>Enter</kbd> abrir</span><span><kbd>Esc</kbd> cerrar</span></div></div>`;
    document.body.appendChild(o);
    const q = $("#paleta-q"), res = $("#paleta-res");
    let vis = [], act = 0;
    const cerrar = () => o.remove();
    const pintar = () => {
      const terms = normTxt(q.value).split(" ").filter(Boolean);
      const orden = ["Ir a", "Proyectos", "Documentos", "Demandas", "Personas", "Compromisos", "Riesgos", "Tickets", "Lecciones"];
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

  // ---------- Edición directa en las tablas (clic en la celda, como en Excel) ----------
  // ed(cfg) devuelve los atributos de una celda editable. cfg: { puede, tabla, idCol, id, campo, tipo, opciones, valor, min, max, vacio, antes(v), extra(v), despues(v) }.
  // Antes de guardar se relee la fila en Excel: si otra persona la cambió, se pregunta qué valor dejar.
  const ED = new Map();
  let edN = 0;
  function ed(cfg) {
    if (!cfg || !cfg.puede || !S.usuario) return "";
    const k = "e" + (++edN); ED.set(k, cfg);
    return ` data-ed="${k}" tabindex="0" title="Clic para editar"`;
  }
  const txtOpcion = (ops, v) => { const o = (ops || []).find((x) => String(Array.isArray(x) ? x[0] : x) === String(v)); return o ? (Array.isArray(o) ? o[1] : o) : v; };
  function abrirEditor(td) {
    const cfg = ED.get(td.dataset.ed);
    if (!cfg || td.classList.contains("editando")) return;
    const previo = td.innerHTML, v0 = cfg.valor ?? "";
    td.classList.add("editando");
    let ctl;
    if (cfg.tipo === "select") {
      const ops = cfg.opciones || [];
      const todas = v0 !== "" && !ops.some((o) => String(Array.isArray(o) ? o[0] : o) === String(v0)) ? [String(v0), ...ops] : ops;
      ctl = `<select class="ed-ctl">${cfg.vacio !== undefined ? `<option value="">${esc(cfg.vacio)}</option>` : ""}${todas.map((o) => { const [val, t] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}" ${String(val) === String(v0) ? "selected" : ""}>${esc(t)}</option>`; }).join("")}</select>`;
    } else if (cfg.tipo === "multi") {
      const sel = String(v0).split(/[,;]\s*/).map((x) => x.trim()).filter(Boolean);
      ctl = `<div class="ed-ctl ed-multi" tabindex="-1">${(cfg.opciones || []).map((o) => { const [val, t] = Array.isArray(o) ? o : [o, o]; return `<label class="check"><input type="checkbox" value="${esc(val)}" ${sel.includes(String(val)) ? "checked" : ""}> ${esc(t)}</label>`; }).join("")}<button type="button" class="btn chico primario ed-listo">Listo</button></div>`;
    } else if (cfg.tipo === "textarea") ctl = `<textarea class="ed-ctl" rows="3">${esc(v0)}</textarea>`;
    else ctl = `<input class="ed-ctl" type="${cfg.tipo === "number" ? "number" : cfg.tipo === "date" ? "date" : "text"}" value="${esc(v0)}" ${cfg.min !== undefined ? `min="${cfg.min}"` : ""} ${cfg.max !== undefined ? `max="${cfg.max}"` : ""}>`;
    td.innerHTML = `<div class="ed-caja">${ctl}<span class="ed-ayuda">${cfg.tipo === "multi" ? "Marca y pulsa Listo · Esc cancela" : cfg.tipo === "textarea" ? "Ctrl+Enter guarda · Esc cancela" : "Enter guarda · Esc cancela"}</span></div>`;
    td.querySelector(".ed-caja").addEventListener("click", (e) => e.stopPropagation());
    const i = td.querySelector(".ed-ctl");
    let hecho = false;
    const valorCtl = () => (cfg.tipo === "multi" ? [...i.querySelectorAll("input:checked")].map((x) => x.value).join(", ") : i.value);
    const cerrar = () => { td.classList.remove("editando"); td.innerHTML = previo; };
    const confirmarValor = () => {
      if (hecho) return; hecho = true;
      let v = String(valorCtl()).trim();
      if (String(v) === String(v0)) return cerrar();
      if (cfg.tipo === "number") {
        if (v === "" && cfg.requerido) { toast("Este campo no puede quedar vacío.", true); return cerrar(); }
        if (v !== "") { const n = Number(v); if (Number.isNaN(n) || (cfg.min !== undefined && n < cfg.min) || (cfg.max !== undefined && n > cfg.max)) { toast(`Valor no válido${cfg.max !== undefined ? ` (entre ${cfg.min} y ${cfg.max})` : ""}.`, true); return cerrar(); } v = n; }
      }
      if (cfg.requerido && v === "") { toast("Este campo no puede quedar vacío.", true); return cerrar(); }
      if (cfg.validar) { const msg = cfg.validar(v); if (msg) { toast(msg, true); return cerrar(); } }
      if (cfg.antes && cfg.antes(v) === false) return cerrar();
      guardarCelda(td, cfg, v, previo);
    };
    i.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); hecho = true; cerrar(); td.focus(); }
      else if (e.key === "Enter" && cfg.tipo !== "textarea") { e.preventDefault(); confirmarValor(); }
      else if (e.key === "Enter" && cfg.tipo === "textarea" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); confirmarValor(); }
    });
    if (cfg.tipo === "select") i.addEventListener("change", confirmarValor);
    if (cfg.tipo === "multi") {
      i.querySelector(".ed-listo").addEventListener("click", (e) => { e.stopPropagation(); confirmarValor(); });
      const fuera = (e) => { if (!td.contains(e.target)) { document.removeEventListener("mousedown", fuera, true); if (!hecho) confirmarValor(); } };
      document.addEventListener("mousedown", fuera, true);
    } else i.addEventListener("blur", () => setTimeout(() => { if (!hecho && document.activeElement !== i) confirmarValor(); }, 120));
    i.focus();
    if (i.select && cfg.tipo !== "select" && cfg.tipo !== "date" && cfg.tipo !== "multi") i.select();
  }
  async function guardarCelda(td, cfg, v, previo) {
    td.innerHTML = `<span class="ed-guardando">Guardando…</span>`;
    const y = window.scrollY;
    const fin = () => { render(); window.scrollTo(0, y); };
    const escribir = async () => {
      cerrarModal();
      try {
        await S.api.actualizarPorId(cfg.tabla, cfg.idCol, cfg.id, { [cfg.campo]: v, ...(cfg.extra ? cfg.extra(v) : {}) });
        if (cfg.despues) await cfg.despues(v);
        armarCatalogos(); fin(); toast(`${cfg.etiqueta || "Dato"} actualizado`);
      } catch (e) { toast("No se pudo guardar: " + e.message, true); td.innerHTML = previo; td.classList.remove("editando"); }
    };
    // ¿Alguien más lo cambió desde que se cargó?
    try {
      const filas = await S.api.leerTabla(cfg.tabla);
      const f = filas.find((r) => String(r[cfg.idCol]) === String(cfg.id));
      const enExcel = f ? f[cfg.campo] ?? "" : null;
      const norm = (x) => (typeof x === "number" && x > 20000 && x < 90000 && cfg.tipo === "date" ? new Date(Math.round((x - 25569) * 86400000)).toISOString().slice(0, 10) : String(x ?? ""));
      if (f && norm(enExcel) !== String(cfg.valor ?? "")) {
        const fila = (S.datos[cfg.tabla] || []).find((r) => String(r[cfg.idCol]) === String(cfg.id));
        if (fila) fila[cfg.campo] = norm(enExcel);
        const mostrar = (x) => (cfg.tipo === "date" ? fecha(x) : txtOpcion(cfg.opciones, x)) || "vacío";
        fin();
        confirmar("Otra persona cambió este dato", `Mientras lo tenías abierto, en Excel cambió de «${mostrar(cfg.valor)}» a «${mostrar(norm(enExcel))}». ¿Lo reemplazo por «${mostrar(v)}»?`, "Reemplazar", escribir, "primario");
        return;
      }
    } catch (e) { /* si no se puede releer, se guarda igual */ }
    return escribir();
  }
  function enlazarEdicionTablas() {
    document.addEventListener("click", (e) => {
      const td = e.target.closest && e.target.closest("[data-ed]");
      if (!td) return;
      if (td.classList.contains("editando")) return;   // clics dentro del editor: el editor los atiende (y no dejan pasar a la fila)
      // Un enlace o botón dentro de la celda sigue funcionando (abrir documento, comentarios…); el resto de la celda edita.
      const accion = e.target.closest("a, button, input, select, textarea, summary");
      if (accion && td.contains(accion) && accion !== td) return;
      e.stopPropagation();
      abrirEditor(td);
    }, true);
    document.addEventListener("keydown", (e) => {
      const td = e.target.closest && e.target.closest("[data-ed]");
      if (td && !td.classList.contains("editando") && (e.key === "Enter" || e.key === "F2")) { e.preventDefault(); e.stopPropagation(); abrirEditor(td); }
    }, true);
  }

  enlazarEdicionTablas();

  // ---------- Agenda de seguimiento (sesiones recurrentes por proyecto) ----------
  // Cada proyecto puede tener sesiones recurrentes (daily, seguimiento, comité…) con días, hora y duración.
  // Con agenda, «al día / vencido» se mide contra las sesiones programadas (los daily no cuentan para eso).
  const DIAS_AG = [["L", "Lun", 1, "MO"], ["M", "Mar", 2, "TU"], ["X", "Mié", 3, "WE"], ["J", "Jue", 4, "TH"], ["V", "Vie", 5, "FR"]];
  const FREC_AG = ["Semanal", "Quincenal", "Mensual"];
  const VECES_SEMANA = { Semanal: 1, Quincenal: 0.5, Mensual: 12 / 52 };
  const esDaily = (a) => /daily/i.test(String(a.Tipo || ""));
  const diasDe = (a) => String(a.Dias || "").split(/[,;\s]+/).map((x) => x.trim().toUpperCase()).filter((x) => DIAS_AG.some((d) => d[0] === x));
  const agendaDe = (pid) => (S.datos.Agenda || []).filter((a) => a.ID_Proyecto === pid).sort((a, b) => String(a.Hora).localeCompare(String(b.Hora)));
  const agendaActiva = (a, d) => a.Activa !== "No" && (!a.Desde || String(a.Desde) <= d) && (!a.Hasta || String(a.Hasta) >= d);
  const lunesDe = (iso) => { const d = R.aFecha ? R.aFecha(iso) : new Date(iso + "T00:00:00"); const w = (d.getDay() + 6) % 7; return R.sumarDias(iso, -w); };
  const diaSemana = (iso) => { const [y, m, dd] = iso.split("-").map(Number); return new Date(y, m - 1, dd).getDay(); };
  const hhmm = (h) => { const m = String(h || "").match(/^(\d{1,2}):(\d{2})/); return m ? `${m[1].padStart(2, "0")}:${m[2]}` : ""; };
  const horaTxt = (h) => { const v = hhmm(h); if (!v) return "—"; let [H, M] = v.split(":").map(Number); const pmx = H >= 12; H = H % 12 || 12; return `${H}:${String(M).padStart(2, "0")} ${pmx ? "p. m." : "a. m."}`; };
  const minDe = (h) => { const v = hhmm(h); if (!v) return 0; const [H, M] = v.split(":").map(Number); return H * 60 + M; };
  // ¿Hay sesión de esta recurrencia en la fecha d?
  function tocaEn(a, d) {
    if (!agendaActiva(a, d)) return false;
    const dw = diaSemana(d);
    if (!diasDe(a).some((x) => DIAS_AG.find((z) => z[0] === x)[2] === dw)) return false;
    const base = a.Desde || a.Fecha_Registro || d;
    if (a.Frecuencia === "Quincenal") return Math.round(R.difDias(lunesDe(d), lunesDe(base)) / 7) % 2 === 0;
    if (a.Frecuencia === "Mensual") { const nth = Math.ceil(Number(String(base).slice(8, 10)) / 7); return Math.ceil(Number(d.slice(8, 10)) / 7) === Math.min(nth, 4); }
    return true;
  }
  const ocurrencias = (a, ini, fin) => { const out = []; for (let d = ini; d <= fin; d = R.sumarDias(d, 1)) if (tocaEn(a, d)) out.push(d); return out; };
  const resumenAgenda = (a) => `${a.Tipo || "Sesión"}: ${diasDe(a).length === 5 ? "L a V" : diasDe(a).join(", ")}${a.Frecuencia && a.Frecuencia !== "Semanal" ? ` (${a.Frecuencia.toLowerCase()})` : ""} ${horaTxt(a.Hora)}`;
  const minutosSemana = (a) => (Number(a.Duracion_Min) || 0) * diasDe(a).length * (VECES_SEMANA[a.Frecuencia] || 1);
  const durTxt = (min) => { const h = Math.floor(min / 60), m = Math.round(min % 60); return h ? `${h} h${m ? ` ${m} min` : ""}` : `${m} min`; };

  // Estado de seguimiento con agenda: vencido si una sesión programada (no daily) pasó sin registro.
  const estadoSegBase = R.estadoSeguimiento;
  R.estadoSeguimiento = (p, segs, hoy = R.hoyISO()) => {
    const base = estadoSegBase(p, segs, hoy);
    const ag = agendaDe(p.ID_Proyecto).filter((a) => !esDaily(a) && diasDe(a).length && a.Activa !== "No");
    if (!ag.length || base.estado === "No aplica") return base;
    const regs = R.seguimientosDe(p.ID_Proyecto, segs);
    const pasadas = ag.flatMap((a) => ocurrencias(a, R.sumarDias(hoy, -45), R.sumarDias(hoy, -1))).sort();
    const sinRegistro = pasadas.filter((d) => !regs.some((s) => s.Fecha_Corte >= R.sumarDias(d, -1)));
    const proximas = ag.flatMap((a) => ocurrencias(a, hoy, R.sumarDias(hoy, 70))).sort();
    if (sinRegistro.length) { const d = sinRegistro[sinRegistro.length - 1]; return { ...base, proximo: d, faltan: R.difDias(d, hoy), estado: "Vencido", conAgenda: true }; }
    const prox = proximas[0];
    if (!prox) return { ...base, conAgenda: true };
    const faltan = R.difDias(prox, hoy);
    return { ...base, proximo: prox, faltan, estado: faltan <= 2 ? "Por vencer" : "Al día", conAgenda: true };
  };
  const kpisBase = R.kpis;
  R.kpis = (proyectos, segs) => { const k = kpisBase(proyectos, segs); const act = proyectos.filter((p) => p.Estado === "Activo"); k.pctAlDia = act.length ? (act.filter((p) => R.estadoSeguimiento(p, segs).estado !== "Vencido").length / act.length) * 100 : 100; return k; };

  function seccionAgenda(p, puede) {
    const ag = agendaDe(p.ID_Proyecto), hoy = R.hoyISO();
    const regs = R.seguimientosDe(p.ID_Proyecto, S.datos.Seguimientos);
    const minT = ag.filter((a) => agendaActiva(a, hoy)).reduce((t, a) => t + minutosSemana(a), 0);
    const fila = (a) => {
      const prox = ocurrencias(a, hoy, R.sumarDias(hoy, 70))[0];
      const ult = ocurrencias(a, R.sumarDias(hoy, -45), R.sumarDias(hoy, -1)).pop();
      const falta = ult && !esDaily(a) && !regs.some((s) => s.Fecha_Corte >= R.sumarDias(ult, -1));
      return `<div class="ag-fila ${a.Activa === "No" ? "inactiva" : ""}">
        <div><span class="pill ag-${normTxt(a.Tipo || "sesion").replace(/[^a-z]/g, "")}">${esc(a.Tipo || "Sesión")}</span>${a.Participantes ? `<div class="sub">${esc(a.Participantes)}</div>` : ""}</div>
        <div class="ag-dias">${DIAS_AG.map(([k]) => `<span class="ag-dia ${diasDe(a).includes(k) ? "on" : ""}">${k}</span>`).join("")}${a.Frecuencia && a.Frecuencia !== "Semanal" ? ` <span class="sub">${esc(a.Frecuencia.toLowerCase())}</span>` : ""}</div>
        <div><b>${horaTxt(a.Hora)}</b><div class="sub">${a.Duracion_Min ? `${a.Duracion_Min} min` : ""}${a.Lugar ? ` · ${esc(a.Lugar)}` : ""}</div></div>
        <div>${a.Activa === "No" ? pill("Inactiva") : esDaily(a) ? `<span class="sub">No exige registro</span>` : falta ? `<span class="pill vencido">Sin registro del ${fecha(ult)}</span>` : `<span class="pill aldía">Al día</span>`}<div class="sub">${prox ? `Próxima: ${fecha(prox)}` : ""}</div></div>
        <div class="derecha nowrap"><button class="btn chico" data-ag-ics="${esc(a.ID_Sesion)}" title="Descargar para Outlook (.ics)">📅 Outlook</button>${puede ? ` <button class="btn chico" data-ag-edit="${esc(a.ID_Sesion)}">Editar</button>` : ""}</div></div>`;
    };
    return `<div class="card ag-card"><div class="titulo-fila"><h2>Agenda de seguimiento</h2>${puede ? `<button class="btn chico primario" id="b-ag">+ Sesión recurrente</button>` : ""}</div>
      ${ag.length ? ag.map(fila).join("") : `<p class="sub">Sin agenda: el seguimiento se controla por la frecuencia «${esc(p.Frecuencia_Seguimiento || "Semanal")}». Agrega las sesiones recurrentes (daily, seguimiento, comité) con sus días y hora.</p>`}
      ${minT ? `<p class="sub">Tiempo en sesiones de este proyecto: <b>${durTxt(minT)} por semana</b> para el PM (se suma a su carga en Capacidad del equipo). Los daily no exigen registro; los demás sí.</p>` : ""}</div>`;
  }
  function enlazarAgenda(el, p) {
    if ($("#b-ag")) $("#b-ag").addEventListener("click", () => formAgenda(p));
    el.querySelectorAll("[data-ag-edit]").forEach((b) => b.addEventListener("click", () => formAgenda(p, agendaDe(p.ID_Proyecto).find((a) => a.ID_Sesion === b.dataset.agEdit))));
    el.querySelectorAll("[data-ag-ics]").forEach((b) => b.addEventListener("click", () => descargarICS(p, agendaDe(p.ID_Proyecto).find((a) => a.ID_Sesion === b.dataset.agIcs))));
  }
  function formAgenda(p, a) {
    const nuevo = !a;
    const campos = [
      { k: "Tipo", label: "Tipo de sesión", tipo: "select", opciones: S.cat.Tipo_Sesion_Agenda || CATALOGOS_BASE.Tipo_Sesion_Agenda, def: "Seguimiento" },
      { k: "Frecuencia", label: "Frecuencia", tipo: "select", opciones: FREC_AG, def: "Semanal", ayuda: "Quincenal: semana de por medio desde «Desde». Mensual: la misma semana del mes que «Desde»." },
      { k: "Dias", label: "Días", tipo: "checks", opciones: DIAS_AG.map(([k, t]) => [k, t]), def: "" },
      { k: "Hora", label: "Hora", tipo: "time", def: "08:00" },
      { k: "Duracion_Min", label: "Duración (minutos)", tipo: "number", min: 5, max: 480, def: 30 },
      { k: "Lugar", label: "Lugar o enlace", placeholder: "Teams, Sala 3…" },
      { k: "Participantes", label: "Participantes", ancho: true, placeholder: "Equipo técnico, proveedor, sponsor…" },
      { k: "Desde", label: "Desde", tipo: "date", def: R.hoyISO() },
      { k: "Hasta", label: "Hasta (opcional)", tipo: "date", ayuda: "Vacío = mientras dure el proyecto." },
      { k: "Activa", label: "¿Activa?", tipo: "select", opciones: ["Sí", "No"], def: "Sí" },
    ];
    modal(nuevo ? "Nueva sesión recurrente" : "Editar sesión recurrente", campos, a || {}, (fd) => {
      fd.Dias = String(fd.Dias).split(/,\s*/).filter(Boolean).join(",");
      if (nuevo) guardar(() => S.api.agregarFila("Agenda", { ID_Sesion: R.siguienteIdHijo("AGD", S.datos.Agenda || [], "ID_Sesion", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo }), "Sesión agendada");
      else guardar(() => S.api.actualizarPorId("Agenda", "ID_Sesion", a.ID_Sesion, fd), "Agenda actualizada");
    }, (fd) => (!fd.Dias ? "Marca al menos un día." : !hhmm(fd.Hora) ? "Escribe la hora." : fd.Hasta && fd.Desde && fd.Hasta < fd.Desde ? "«Hasta» no puede ser anterior a «Desde»." : ""),
    { init: (m) => {
        const t = m.querySelector('[name="Tipo"]');
        t.addEventListener("change", () => { if (/daily/i.test(t.value)) { m.querySelectorAll('[data-checks="Dias"] input').forEach((i) => { i.checked = true; }); const d = m.querySelector('[name="Duracion_Min"]'); if (d && (!d.value || d.value === "30")) d.value = 15; } });
      },
      ...(nuevo ? {} : { eliminar: { texto: "Eliminar sesión", mensaje: "Se quitará esta sesión recurrente de la agenda. Los seguimientos ya registrados no se tocan.", accion: () => guardar(() => S.api.eliminarFilas("Agenda", "ID_Sesion", [a.ID_Sesion]), "Sesión eliminada") } }) });
  }
  // Archivo de calendario (.ics) con la recurrencia, para agregarla a Outlook.
  function descargarICS(p, a) {
    if (!a) return;
    const d0 = ocurrencias(a, a.Desde && a.Desde > R.hoyISO() ? a.Desde : R.hoyISO(), R.sumarDias(R.hoyISO(), 70))[0] || a.Desde || R.hoyISO();
    const hm = hhmm(a.Hora) || "08:00";
    const ini = `${d0.replace(/-/g, "")}T${hm.replace(":", "")}00`;
    const finMin = minDe(hm) + (Number(a.Duracion_Min) || 30);
    const fin = `${d0.replace(/-/g, "")}T${String(Math.floor(finMin / 60) % 24).padStart(2, "0")}${String(finMin % 60).padStart(2, "0")}00`;
    const byday = diasDe(a).map((k) => DIAS_AG.find((z) => z[0] === k)[3]);
    const nth = Math.min(4, Math.ceil(Number(String(a.Desde || d0).slice(8, 10)) / 7));
    const rrule = a.Frecuencia === "Mensual" ? `FREQ=MONTHLY;BYDAY=${byday.map((x) => nth + x).join(",")}` : `FREQ=WEEKLY;INTERVAL=${a.Frecuencia === "Quincenal" ? 2 : 1};BYDAY=${byday.join(",")}`;
    const hasta = a.Hasta || p.Fecha_Fin_Plan;
    const escI = (t) => String(t || "").replace(/\\/g, "\\\\").replace(/[,;]/g, (c) => "\\" + c).replace(/\n/g, "\\n");
    const ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//FSFB//Portafolio PMO//ES", "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
      "BEGIN:VTIMEZONE", "TZID:America/Bogota", "BEGIN:STANDARD", "DTSTART:19700101T000000", "TZOFFSETFROM:-0500", "TZOFFSETTO:-0500", "TZNAME:COT", "END:STANDARD", "END:VTIMEZONE",
      "BEGIN:VEVENT", `UID:${a.ID_Sesion}@portafolio-fsfb`, `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, "").slice(0, 15)}Z`,
      `DTSTART;TZID=America/Bogota:${ini}`, `DTEND;TZID=America/Bogota:${fin}`, `RRULE:${rrule}${hasta ? `;UNTIL=${String(hasta).replace(/-/g, "")}T235959Z` : ""}`,
      `SUMMARY:${escI(`${a.Tipo || "Sesión"} · ${p.Nombre}`)}`, a.Lugar ? `LOCATION:${escI(a.Lugar)}` : "",
      `DESCRIPTION:${escI(`Proyecto ${p.ID_Proyecto} · ${p.Nombre}\nPM: ${nombreUsuario(p.PM)}${a.Participantes ? `\nParticipantes: ${a.Participantes}` : ""}`)}`,
      "END:VEVENT", "END:VCALENDAR"].filter(Boolean).join("\r\n");
    descargar(new TextEncoder().encode(ics), `${normTxt(`${a.Tipo || "sesion"} ${p.Nombre}`).replace(/[^a-z0-9]+/g, "_").slice(0, 50)}.ics`, "text/calendar");
    toast("Archivo de calendario descargado: ábrelo para agregarlo a Outlook.");
  }

  // Vista «Agenda semanal»: sesiones de todos los proyectos filtrados, por día y hora.
  function vAgenda(el) {
    const hoy = R.hoyISO();
    const lun = S.agLunes || lunesDe(hoy);
    const dias = DIAS_AG.map((d, i) => ({ k: d[0], t: d[1], f: R.sumarDias(lun, i) }));
    const ps = filtrados().filter(CUENTAN);
    const ev = [];
    ps.forEach((p) => agendaDe(p.ID_Proyecto).forEach((a) => dias.forEach((d) => { if (tocaEn(a, d.f)) ev.push({ p, a, f: d.f, ini: minDe(a.Hora), fin: minDe(a.Hora) + (Number(a.Duracion_Min) || 30) }); })));
    const regs = (pid) => R.seguimientosDe(pid, S.datos.Seguimientos);
    ev.forEach((e) => {
      e.reg = regs(e.p.ID_Proyecto).some((s) => s.Fecha_Corte === e.f || (s.Fecha_Corte >= R.sumarDias(e.f, -1) && s.Fecha_Corte <= R.sumarDias(e.f, 1)));
      e.falta = !esDaily(e.a) && e.f < hoy && !e.reg;
      e.choque = ev.some((o) => o !== e && o.f === e.f && lc(o.p.PM) === lc(e.p.PM) && o.ini < e.fin && e.ini < o.fin);
    });
    const horas = [...new Set(ev.map((e) => hhmm(e.a.Hora)))].sort();
    const porPM = {};
    ev.forEach((e) => { const k = lc(e.p.PM); porPM[k] = (porPM[k] || 0) + (e.fin - e.ini); });
    const nCh = ev.filter((e) => e.choque).length, nFal = ev.filter((e) => e.falta).length;
    const evHTML = (e) => `<button type="button" class="ag-ev ag-${normTxt(e.a.Tipo || "sesion").replace(/[^a-z]/g, "")} ${e.choque ? "choque" : ""} ${e.falta ? "falta" : ""} ${e.reg ? "hecho" : ""}" data-ag-ev="${esc(e.p.ID_Proyecto)}|${esc(e.a.ID_Sesion)}|${e.f}"
        title="${esc(`${e.p.Nombre} · ${e.a.Tipo} · ${horaTxt(e.a.Hora)} · ${e.a.Duracion_Min || 30} min · PM ${nombreUsuario(e.p.PM)}${e.choque ? " · cruce de horario del PM" : ""}${e.falta ? " · sin registro" : ""}`)}">
        <b>${esc(corta(e.p.Nombre, 26))}</b><span>${esc(e.a.Tipo || "Sesión")}${e.reg ? " ✓" : ""}${e.choque ? " · cruce" : ""}${e.falta ? " · sin registro" : ""}</span></button>`;
    el.innerHTML = `<h1>Agenda semanal</h1>${barraFiltros()}
      <div class="filtros ag-nav"><button class="btn chico" id="ag-ant">← Semana anterior</button><b>Semana del ${fecha(dias[0].f)} al ${fecha(dias[4].f)}</b><button class="btn chico" id="ag-sig">Semana siguiente →</button>${lun !== lunesDe(hoy) ? `<button class="btn chico" id="ag-hoy">Esta semana</button>` : ""}</div>
      <div class="kpis kpis-3">
        <div class="kpi estatico"><span class="kpi-t">Sesiones esta semana</span><span class="kpi-v">${ev.length}</span><span class="sub">${ev.filter((e) => esDaily(e.a)).length} daily</span></div>
        <div class="kpi estatico ${nFal ? "alerta" : ""}"><span class="kpi-t">Sin registro</span><span class="kpi-v">${nFal}</span><span class="sub">seguimientos y comités ya pasados</span></div>
        <div class="kpi estatico ${nCh ? "alerta" : ""}"><span class="kpi-t">Cruces de horario</span><span class="kpi-v">${nCh}</span><span class="sub">mismo PM, misma hora</span></div>
        <div class="kpi estatico"><span class="kpi-t">Horas en sesiones por PM</span><span class="kpi-v kpi-chico">${Object.entries(porPM).sort((a, b) => b[1] - a[1]).slice(0, 4).map(([k, m]) => `${esc(corta(nombreUsuario(k), 22))} ${durTxt(m)}`).join(" · ") || "—"}</span></div>
      </div>
      <div class="card">${ev.length ? `<div class="ag-grid"><div class="ag-h"></div>${dias.map((d) => `<div class="ag-h ${d.f === hoy ? "hoy" : ""}">${d.t} ${d.f.slice(8, 10)}</div>`).join("")}
        ${horas.map((h) => `<div class="ag-hora">${horaTxt(h)}</div>${dias.map((d) => `<div class="ag-celda ${d.f === hoy ? "hoy" : ""}">${ev.filter((e) => e.f === d.f && hhmm(e.a.Hora) === h).sort((a, b) => String(a.p.Nombre).localeCompare(String(b.p.Nombre), "es")).map(evHTML).join("")}</div>`).join("")}`).join("")}</div>
        <div class="ag-leyenda"><span class="ag-ev ag-daily">Daily</span><span class="ag-ev ag-seguimiento">Seguimiento</span><span class="ag-ev ag-comite">Comité</span><span class="ag-ev ag-mesatecnica">Mesa técnica</span><span class="ag-ev choque">Cruce de horario</span><span class="ag-ev falta">Sin registro</span><span class="sub">✓ = registrado. Clic en una sesión para registrarla.</span></div>`
        : vacio("No hay sesiones programadas esta semana. Agrégalas en la pestaña «Seguimientos» de cada proyecto (Agenda de seguimiento).")}</div>`;
    enlazarFiltros();
    $("#ag-ant").addEventListener("click", () => { S.agLunes = R.sumarDias(lun, -7); render(); });
    $("#ag-sig").addEventListener("click", () => { S.agLunes = R.sumarDias(lun, 7); render(); });
    if ($("#ag-hoy")) $("#ag-hoy").addEventListener("click", () => { S.agLunes = ""; render(); });
    el.querySelectorAll("[data-ag-ev]").forEach((b) => b.addEventListener("click", () => {
      const [pid, , f] = b.dataset.agEv.split("|"), p = proyecto(pid), a = agendaDe(pid).find((x) => x.ID_Sesion === b.dataset.agEv.split("|")[1]);
      if (!p) return;
      if (f > hoy || !R.puedeEditar(S.usuario, p, "seguimiento")) { S.fichaTab = "seguimientos"; S.fichaPid = pid; return ir("ficha", pid); }
      formSeguimiento(p, a && /seguimiento/i.test(a.Tipo) && p.Estado === "Activo" ? "Seguimiento" : "Sesión");
      const fc = document.querySelector('#m-form [name="Fecha_Corte"]'); if (fc) fc.value = f;
      const nm = document.querySelector('#m-form [name="Nombre_Sesion"]'); if (nm && a && !nm.value) nm.value = a.Tipo;
    }));
  }

  // ---------- Valor y beneficios para la FSFB ----------
  // Cada proyecto declara su propuesta de valor (beneficio esperado, valor para la FSFB, impacto por dimensión)
  // y sus beneficios medibles (indicador, línea base → meta). El PM registra mediciones, también después del cierre.
  const NIVELES_VALOR = ["Crítico", "Alto", "Medio", "Bajo"];
  const PESO_IMP = { "Crítico": 4, Alto: 3, Medio: 2, Bajo: 1 };
  const IMPACTOS_DIM = ["Alto", "Medio", "Bajo"];
  const FREC_MED = { Mensual: 30, Trimestral: 91, Semestral: 182, Anual: 365 };
  const ICO_DIM = { "Seguridad del paciente": "❤", "Eficiencia operacional": "⚙", "Experiencia del paciente": "☺", "Cumplimiento / acreditación": "✓", "Estratégico": "◎", "Financiero": "$" };
  const dimensiones = () => S.cat.Dimension_Valor || CATALOGOS_BASE.Dimension_Valor;
  const beneficiosDe = (pid) => (S.datos.Beneficios || []).filter((b) => b.ID_Proyecto === pid).sort((a, b) => (PESO_IMP[b.Importancia] || 0) - (PESO_IMP[a.Importancia] || 0) || String(a.ID_Beneficio).localeCompare(String(b.ID_Beneficio)));
  const medicionesDe = (bid) => (S.datos.Mediciones || []).filter((m) => m.ID_Beneficio === bid).sort((a, b) => String(a.Fecha).localeCompare(String(b.Fecha)));
  const impactoValorDe = (p) => { try { const o = JSON.parse(p.Impacto_Valor || "{}"); return o && typeof o === "object" ? o : {}; } catch (e) { return {}; } };
  const esMonetario = (b) => /\$|cop|peso|millon/i.test(String(b.Unidad || "")) || b.Dimension === "Financiero";
  const fmtValor = (b, v) => (v === "" || v === null || v === undefined ? "—" : b.Tipo_Medida === "Cualitativo" ? String(v) : esMonetario(b) ? cop(v) : `${Number(v).toLocaleString("es-CO", { maximumFractionDigits: 2 })}${b.Unidad ? ` ${b.Unidad}` : ""}`);
  const num1 = (v) => (v === "" || v === null || v === undefined || Number.isNaN(Number(v)) ? null : Number(v));

  // Logro (0-100) frente a la meta, estado y próxima medición de un beneficio.
  function evalBeneficio(b, hoy = R.hoyISO()) {
    const ms = medicionesDe(b.ID_Beneficio), ult = ms[ms.length - 1] || null;
    let logro = null;
    if (ult) {
      if (b.Tipo_Medida === "Cualitativo") logro = ult.Valor === "Cumplido" ? 100 : ult.Valor === "Parcial" ? 50 : 0;
      else {
        const v = num1(ult.Valor), base = num1(b.Linea_Base), meta = num1(b.Meta);
        if (v !== null && base !== null && meta !== null) logro = meta === base ? (v === meta ? 100 : 0) : Math.round(Math.max(0, Math.min(100, ((v - base) / (meta - base)) * 100)));
      }
    }
    const desde = b.Medir_Desde || b.Fecha_Registro || hoy;
    const iniciado = hoy >= desde;
    const dias = FREC_MED[b.Frecuencia_Medicion] || 30;
    const proxima = ult ? R.sumarDias(ult.Fecha, dias) : desde;
    const pendiente = iniciado && proxima <= hoy && logro !== 100;
    const esperado = b.Fecha_Meta && b.Fecha_Meta > desde ? Math.max(0, Math.min(100, Math.round((R.difDias(hoy, desde) / R.difDias(b.Fecha_Meta, desde)) * 100))) : null;
    let estado;
    if (logro === 100) estado = "Logrado";
    else if (!iniciado && logro === null) estado = "Por iniciar";
    else if (logro === null) estado = "Sin medición";
    else if (b.Fecha_Meta && hoy > b.Fecha_Meta) estado = "En riesgo";
    else if (esperado !== null && logro < esperado - 20) estado = "En riesgo";
    else estado = "En camino";
    return { ult, ms, logro, estado, proxima, pendiente, esperado, iniciado };
  }
  // Valor del proyecto: índice ponderado por importancia, nivel sugerido y cifras financieras (si hay beneficios en $).
  function valorProyecto(p, hoy = R.hoyISO()) {
    const bs = beneficiosDe(p.ID_Proyecto).map((b) => ({ b, e: evalBeneficio(b, hoy) }));
    const conLogro = bs.filter((x) => x.e.logro !== null);
    const pesoT = conLogro.reduce((t, x) => t + (PESO_IMP[x.b.Importancia] || 1), 0);
    const indice = pesoT ? Math.round(conLogro.reduce((t, x) => t + (PESO_IMP[x.b.Importancia] || 1) * x.e.logro, 0) / pesoT) : null;
    const imp = impactoValorDe(p);
    let sugerido = "";
    const vals = Object.values(imp);
    if (imp["Seguridad del paciente"] === "Alto" || bs.some((x) => x.b.Importancia === "Crítico")) sugerido = "Crítico";
    else if (vals.includes("Alto") || bs.some((x) => x.b.Importancia === "Alto")) sugerido = "Alto";
    else if (vals.includes("Medio") || bs.some((x) => x.b.Importancia === "Medio")) sugerido = "Medio";
    else if (vals.length || bs.length) sugerido = "Bajo";
    const fin = bs.filter((x) => x.b.Tipo_Medida !== "Cualitativo" && esMonetario(x.b));
    const valorFin = fin.reduce((t, x) => t + Math.abs((num1(x.b.Meta) || 0) - (num1(x.b.Linea_Base) || 0)), 0);
    const inversion = Number(p.Ejecutado) > Number(p.Presupuesto) ? Number(p.Ejecutado) : Number(p.Presupuesto) || 0;
    return { bs, indice, sugerido, nivel: p.Nivel_Valor || sugerido, manual: !!p.Nivel_Valor, imp,
      riesgo: bs.filter((x) => x.e.estado === "En riesgo").length, criticosRiesgo: bs.filter((x) => x.e.estado === "En riesgo" && x.b.Importancia === "Crítico").length,
      pendientes: bs.filter((x) => x.e.pendiente).length, valorFin, roi: valorFin && inversion ? valorFin / inversion : null, payback: valorFin && inversion ? Math.round(inversion / (valorFin / 12)) : null,
      tieneAlgo: !!(p.Beneficio_Esperado || p.Valor_FSFB || bs.length || vals.length) };
  }
  const clsNivel = (n) => ({ "Crítico": "nv-critico", Alto: "nv-alto", Medio: "nv-medio", Bajo: "nv-bajo" }[n] || "nv-na");
  const chipNivel = (n, sug) => (n ? `<span class="pill nv ${clsNivel(n)}" title="${sug ? "Sugerido por la app según el impacto y los beneficios" : "Nivel de valor para la FSFB"}">${esc(n)}${sug ? " · sugerido" : ""}</span>` : `<span class="sub">Sin definir</span>`);
  const chipImp = (n) => (n ? `<span class="pill nv ${clsNivel(n === "Alto" ? "Crítico" : n === "Medio" ? "Alto" : "Medio")}">${esc(n)}</span>` : `<span class="pill nv nv-na">No aplica</span>`);
  const chipEstB = (e) => `<span class="pill bn-${normTxt(e).replace(/\s+/g, "")}">${esc(e)}</span>`;
  const barraLogro = (l) => (l === null ? `<span class="sub">sin medir</span>` : `<span class="bn-barra" title="${l}% de logro"><i class="${l >= 100 ? "ok" : l >= 60 ? "ok" : l >= 30 ? "warn" : "bad"}" style="width:${Math.max(2, l)}%"></i></span> <b>${l}%</b>`);

  function seccionValor(p, puede) {
    const V = valorProyecto(p), hoy = R.hoyISO();
    const dims = [...new Set([...dimensiones(), ...Object.keys(V.imp), ...V.bs.map((x) => x.b.Dimension).filter(Boolean)])];
    const filasDim = dims.map((d) => {
      const bs = V.bs.filter((x) => x.b.Dimension === d), im = V.imp[d] || "";
      if (!im && !bs.length) return "";
      return `<tr><td class="nowrap"><span class="dim-ico">${ICO_DIM[d] || "•"}</span> ${esc(d)}</td><td>${chipImp(im)}</td>
        <td>${bs.length ? bs.map((x) => `<div class="dim-ind">${esc(x.b.Indicador || x.b.Beneficio_Esperado || "Beneficio")}: ${x.b.Tipo_Medida === "Cualitativo" ? esc(x.e.ult ? x.e.ult.Valor : "sin medir") : `${fmtValor(x.b, x.b.Linea_Base)} → ${fmtValor(x.b, x.b.Meta)} · actual ${x.e.ult ? fmtValor(x.b, x.e.ult.Valor) : "—"}`} ${barraLogro(x.e.logro)}</div>`).join("") : `<span class="sub">Sin indicador</span>`}</td></tr>`;
    }).join("");
    const nImp = (n) => V.bs.filter((x) => x.b.Importancia === n).length;
    const tarjeta = ({ b, e }) => `<div class="bn-card ${e.estado === "En riesgo" ? "riesgo" : ""}">
        <div class="bn-cab">${chipNivel(b.Importancia)}<b>${esc(b.Dimension || "Beneficio")}</b>${b.Indicador ? `<span class="sub">· ${esc(b.Indicador)}</span>` : ""}<span class="bn-der">${chipEstB(e.estado)}${e.pendiente ? ` <span class="pill vencido" title="Debía medirse el ${fecha(e.proxima)}">Medición pendiente</span>` : ""}</span></div>
        <div class="bn-grid">
          <div><div class="bn-k">Beneficio esperado</div>${esc(b.Beneficio_Esperado || "—")}</div>
          <div><div class="bn-k">Valor para la FSFB</div>${esc(b.Valor_FSFB || "—")}</div>
          <div><div class="bn-k">Indicador</div>${b.Tipo_Medida === "Cualitativo" ? `Cualitativo · último: <b>${esc(e.ult ? e.ult.Valor : "sin medir")}</b>` : `${fmtValor(b, b.Linea_Base)} → <b>${fmtValor(b, b.Meta)}</b>${b.Fecha_Meta ? ` <span class="sub">a ${fecha(b.Fecha_Meta)}</span>` : ""}<br>Actual: <b>${e.ult ? fmtValor(b, e.ult.Valor) : "—"}</b>`}${b.Tipo_Medida === "Cualitativo" && e.logro === null ? "" : `<div>${barraLogro(e.logro)}${e.esperado !== null && e.logro !== 100 && e.logro !== null ? ` <span class="sub">esperado hoy ≈ ${e.esperado}%</span>` : ""}</div>`}</div>
        </div>
        <div class="bn-pie"><span class="sub">${b.Dueno ? `Dueño: ${esc(b.Dueno)} · ` : ""}Medición ${esc(String(b.Frecuencia_Medicion || "Mensual").toLowerCase())} desde ${fecha(b.Medir_Desde || b.Fecha_Registro)} · próxima ${fecha(e.proxima)}</span>
          ${puede ? `<span class="acciones"><button class="btn chico primario" data-bn-medir="${esc(b.ID_Beneficio)}">Registrar medición</button><button class="btn chico" data-bn-edit="${esc(b.ID_Beneficio)}">Editar</button></span>` : ""}</div>
        ${e.ms.length ? `<details class="bn-hist"><summary>Historial de mediciones (${e.ms.length})</summary><table class="no-orden"><thead><tr><th>Fecha</th><th>Valor</th><th>Comentario</th><th>Registró</th><th></th></tr></thead><tbody>
          ${e.ms.slice().reverse().map((m) => { const EM = (campo, o) => ed({ puede, tabla: "Mediciones", idCol: "ID_Medicion", id: m.ID_Medicion, campo, valor: m[campo] ?? "", ...o }); return `<tr><td${EM("Fecha", { tipo: "date", requerido: true, etiqueta: "Fecha de la medición" })}>${fecha(m.Fecha)}</td><td${EM("Valor", b.Tipo_Medida === "Cualitativo" ? { tipo: "select", opciones: ["Cumplido", "Parcial", "Pendiente"], etiqueta: "Resultado" } : { tipo: "number", requerido: true, etiqueta: "Valor medido" })}><b>${fmtValor(b, m.Valor)}</b></td><td${EM("Comentario", { tipo: "textarea", etiqueta: "Comentario" })}>${esc(m.Comentario || "—")}</td><td>${esc(nombreUsuario(m.Registrado_Por))}</td><td>${puede ? `<button class="btn chico" data-bn-med="${esc(m.ID_Medicion)}">Editar</button>` : ""}</td></tr>`; }).join("")}</tbody></table></details>` : ""}
      </div>`;
    return `<div class="card vl-prop"><div class="titulo-fila"><h2>Propuesta de valor para la FSFB</h2>${puede ? `<button class="btn chico" id="b-vl-prop">Editar propuesta</button>` : ""}</div>
        <div class="vl-grid">
          <div><div class="bn-k">Beneficio esperado</div>${p.Beneficio_Esperado ? esc(p.Beneficio_Esperado) : `<span class="sub">Sin registrar. ¿Qué cambia cuando el proyecto entregue?</span>`}</div>
          <div><div class="bn-k">Valor para la FSFB</div>${p.Valor_FSFB ? esc(p.Valor_FSFB) : `<span class="sub">Sin registrar. ¿Por qué le importa a la Fundación?</span>`}</div>
          <div><div class="bn-k">Nivel de valor para la FSFB</div>${chipNivel(V.nivel, !V.manual && !!V.nivel)}${V.manual && V.sugerido && V.sugerido !== V.nivel ? `<div class="sub">La app sugiere: ${esc(V.sugerido)}</div>` : ""}</div>
        </div></div>
      <div class="kpis kpis-3">
        <div class="kpi estatico"><span class="kpi-t">Índice de valor (logro ponderado)</span><span class="kpi-v">${V.indice === null ? "—" : `${V.indice}%`}</span><span class="sub">pesa más lo crítico (4·3·2·1)</span></div>
        <div class="kpi estatico"><span class="kpi-t">Beneficios</span><span class="kpi-v">${V.bs.length}</span><span class="sub">${NIVELES_VALOR.map((n) => (nImp(n) ? `${nImp(n)} ${n.toLowerCase()}` : "")).filter(Boolean).join(" · ") || "ninguno"}</span></div>
        <div class="kpi estatico ${V.riesgo ? "alerta" : ""}"><span class="kpi-t">En riesgo</span><span class="kpi-v">${V.riesgo}</span><span class="sub">${V.pendientes} medición(es) pendiente(s)</span></div>
        ${V.valorFin ? `<div class="kpi estatico"><span class="kpi-t">Valor financiero (meta)</span><span class="kpi-v kpi-chico">${cop(V.valorFin)}</span><span class="sub">${V.roi ? `ROI ${V.roi.toFixed(1).replace(".", ",")}x · payback ${V.payback} meses` : "sin inversión registrada"}</span></div>` : ""}
      </div>
      <div class="card"><div class="titulo-fila"><h2>Impacto por dimensión de valor</h2></div>
        ${filasDim ? `<div class="tabla-scroll"><table class="no-orden vl-dims"><thead><tr><th>Dimensión</th><th>Impacto</th><th>Indicadores y logro</th></tr></thead><tbody>${filasDim}</tbody></table></div>` : vacio("Aún no se ha marcado el impacto por dimensión. Usa «Editar propuesta».")}
        <p class="sub">El valor no es solo financiero: seguridad del paciente, eficiencia operacional, experiencia, cumplimiento y lo estratégico también cuentan.</p></div>
      <div class="card"><div class="titulo-fila"><h2>Beneficios <span class="contador">${V.bs.length}</span></h2>${puede ? `<button class="btn chico primario" id="b-bn">+ Beneficio</button>` : ""}</div>
        ${V.bs.length ? V.bs.map(tarjeta).join("") : vacio("Registra los beneficios que el proyecto debe dejar, con su indicador, línea base y meta.")}
        <p class="sub">Los beneficios se siguen midiendo después del cierre del proyecto.</p></div>`;
  }
  function enlazarValor(el, p) {
    if ($("#b-vl-prop")) $("#b-vl-prop").addEventListener("click", () => formPropuestaValor(p));
    if ($("#b-bn")) $("#b-bn").addEventListener("click", () => formBeneficio(p));
    el.querySelectorAll("[data-bn-edit]").forEach((b) => b.addEventListener("click", () => formBeneficio(p, beneficiosDe(p.ID_Proyecto).find((x) => x.ID_Beneficio === b.dataset.bnEdit))));
    el.querySelectorAll("[data-bn-medir]").forEach((b) => b.addEventListener("click", () => formMedicion(p, beneficiosDe(p.ID_Proyecto).find((x) => x.ID_Beneficio === b.dataset.bnMedir))));
    el.querySelectorAll("[data-bn-med]").forEach((b) => b.addEventListener("click", () => { const m = (S.datos.Mediciones || []).find((x) => x.ID_Medicion === b.dataset.bnMed); if (m) formMedicion(p, beneficiosDe(p.ID_Proyecto).find((x) => x.ID_Beneficio === m.ID_Beneficio), m); }));
  }
  function formPropuestaValor(p) {
    const V = valorProyecto(p), dims = [...new Set([...dimensiones(), ...Object.keys(V.imp)])];
    const campos = [
      { k: "Beneficio_Esperado", label: "Beneficio esperado", tipo: "textarea", placeholder: "Qué cambia cuando el proyecto entregue. Ej.: resultados de patología en menos de 48 h y sin errores de rotulado." },
      { k: "Valor_FSFB", label: "Valor para la FSFB", tipo: "textarea", placeholder: "Por qué le importa a la Fundación: paciente, acreditación, eficiencia, costos, estrategia." },
      { k: "Nivel_Valor", label: "Nivel de valor para la FSFB", tipo: "select", opciones: NIVELES_VALOR, textoVacio: `Automático${V.sugerido ? ` (sugerido: ${V.sugerido})` : ""}`, ayuda: "Si lo dejas en automático, la app lo sugiere según el impacto por dimensión y la importancia de los beneficios." },
      { seccion: "Impacto por dimensión de valor", ayuda: "Alto, medio, bajo o vacío si no aplica" },
      ...dims.map((d, i) => ({ k: `_dim${i}`, label: d, tipo: "select", opciones: IMPACTOS_DIM, textoVacio: "No aplica", def: V.imp[d] || "" })),
    ];
    modal(`Propuesta de valor · ${p.Nombre}`, campos, p, (fd) => {
      const imp = {}; dims.forEach((d, i) => { if (fd[`_dim${i}`]) imp[d] = fd[`_dim${i}`]; });
      guardar(() => S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { Beneficio_Esperado: fd.Beneficio_Esperado, Valor_FSFB: fd.Valor_FSFB, Nivel_Valor: fd.Nivel_Valor, Impacto_Valor: Object.keys(imp).length ? JSON.stringify(imp) : "", ...sello() }), "Propuesta de valor guardada");
    });
  }
  function formBeneficio(p, b) {
    const nuevo = !b;
    const campos = [
      { k: "Importancia", label: "Importancia", tipo: "select", opciones: NIVELES_VALOR.map((n) => [n, `${n} (peso ${PESO_IMP[n]})`]), def: "Alto" },
      { k: "Dimension", label: "Dimensión de valor", tipo: "select", opciones: dimensiones() },
      { k: "Beneficio_Esperado", label: "Beneficio esperado", tipo: "textarea", placeholder: "Qué mejora. Ej.: eliminar muestras mal rotuladas con código de barras en la toma." },
      { k: "Valor_FSFB", label: "Valor para la FSFB", tipo: "textarea", placeholder: "Por qué importa. Ej.: evita repetir biopsias y eventos adversos; requisito de acreditación." },
      { seccion: "Cómo se mide" },
      { k: "Indicador", label: "Indicador (KPI)", ancho: true, placeholder: "Muestras mal rotuladas por mes" },
      { k: "Tipo_Medida", label: "Tipo de medida", tipo: "select", opciones: [["Numérico", "Numérico (línea base → meta)"], ["Cualitativo", "Cualitativo (cumplido / parcial / pendiente)"]], def: "Numérico" },
      { k: "Unidad", label: "Unidad", placeholder: "horas, %, casos/mes, COP", ayuda: "Si es dinero escribe COP: se calcula ROI y payback." },
      { k: "Linea_Base", label: "Línea base (hoy)", tipo: "number", placeholder: "12" },
      { k: "Meta", label: "Meta", tipo: "number", placeholder: "2", ayuda: "Puede ser menor que la línea base (bajar) o mayor (subir)." },
      { k: "Medir_Desde", label: "Medir desde", tipo: "date", def: p.Fecha_Fin_Plan && p.Fecha_Fin_Plan > R.hoyISO() ? p.Fecha_Fin_Plan : R.hoyISO(), ayuda: "Normalmente cuando el proyecto sale a producción." },
      { k: "Fecha_Meta", label: "Fecha meta", tipo: "date" },
      { k: "Frecuencia_Medicion", label: "Medir cada", tipo: "select", opciones: Object.keys(FREC_MED), def: "Mensual" },
      { k: "Dueno", label: "Dueño del beneficio (área o persona)", placeholder: "Patología" },
    ];
    const ms = b ? medicionesDe(b.ID_Beneficio) : [];
    modal(nuevo ? "Nuevo beneficio" : "Editar beneficio", campos, b || {}, (fd) => {
      if (fd.Tipo_Medida === "Cualitativo") { fd.Linea_Base = ""; fd.Meta = ""; }
      if (nuevo) guardar(() => S.api.agregarFila("Beneficios", { ID_Beneficio: R.siguienteIdHijo("BEN", S.datos.Beneficios || [], "ID_Beneficio", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo, Fecha_Registro: R.hoyISO() }), "Beneficio registrado");
      else guardar(() => S.api.actualizarPorId("Beneficios", "ID_Beneficio", b.ID_Beneficio, fd), "Beneficio actualizado");
    }, (fd) => {
      if (!fd.Dimension) return "Elige la dimensión de valor.";
      if (!fd.Beneficio_Esperado && !fd.Indicador) return "Escribe el beneficio esperado o el indicador.";
      if (fd.Tipo_Medida !== "Cualitativo") {
        if (fd.Linea_Base === "" || fd.Meta === "") return "Para una medida numérica escribe la línea base y la meta.";
        if (fd.Linea_Base === fd.Meta) return "La meta debe ser distinta de la línea base.";
      }
      if (fd.Fecha_Meta && fd.Medir_Desde && fd.Fecha_Meta < fd.Medir_Desde) return "La fecha meta no puede ser anterior a «Medir desde».";
      return "";
    }, nuevo ? {} : { eliminar: { texto: "Eliminar beneficio", mensaje: `Se eliminará el beneficio${ms.length ? ` y sus ${ms.length} medición(es)` : ""}.`,
      accion: () => guardar(async () => { if (ms.length) await S.api.eliminarFilas("Mediciones", "ID_Medicion", ms.map((m) => m.ID_Medicion)); await S.api.eliminarFilas("Beneficios", "ID_Beneficio", [b.ID_Beneficio]); }, "Beneficio eliminado") } });
  }
  function formMedicion(p, b, m) {
    if (!b) return;
    const nuevo = !m, cual = b.Tipo_Medida === "Cualitativo";
    const campos = [
      { k: "Fecha", label: "Fecha de la medición", tipo: "date", def: R.hoyISO() },
      cual ? { k: "Valor", label: "Resultado", tipo: "select", opciones: ["Cumplido", "Parcial", "Pendiente"] }
        : { k: "Valor", label: `Valor medido${b.Unidad ? ` (${b.Unidad})` : ""}`, tipo: "number", ayuda: `Línea base ${fmtValor(b, b.Linea_Base)} → meta ${fmtValor(b, b.Meta)}` },
      { k: "Comentario", label: "Comentario / fuente", tipo: "textarea", placeholder: "Fuente del dato o explicación (opcional)" },
    ];
    modal(`${nuevo ? "Registrar medición" : "Editar medición"} · ${b.Indicador || b.Dimension}`, campos, m || {}, (fd) => {
      if (nuevo) guardar(() => S.api.agregarFila("Mediciones", { ID_Medicion: R.siguienteIdHijo("MED", S.datos.Mediciones || [], "ID_Medicion", p.ID_Proyecto), ID_Beneficio: b.ID_Beneficio, ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo }), "Medición registrada");
      else guardar(() => S.api.actualizarPorId("Mediciones", "ID_Medicion", m.ID_Medicion, fd), "Medición actualizada");
    }, (fd) => (!fd.Fecha ? "Escribe la fecha." : fd.Valor === "" ? "Escribe el valor medido." : ""),
    nuevo ? {} : { eliminar: { texto: "Eliminar medición", mensaje: "Se eliminará esta medición.", accion: () => guardar(() => S.api.eliminarFilas("Mediciones", "ID_Medicion", [m.ID_Medicion]), "Medición eliminada") } });
  }

  // Vista de portafolio: valor por proyecto y por dimensión, beneficios en riesgo y mediciones pendientes.
  function vValor(el) {
    const F = S.valF || { nivel: "" };
    const base = filtrados().filter((p) => p.Estado !== "Cancelado").map((p) => ({ p, V: valorProyecto(p) }));
    const lista = base.filter((x) => !F.nivel || (F.nivel === "sin" ? !x.V.nivel : x.V.nivel === F.nivel))
      .sort((a, b) => (PESO_IMP[b.V.nivel] || 0) - (PESO_IMP[a.V.nivel] || 0) || (b.V.indice ?? -1) - (a.V.indice ?? -1) || String(a.p.Nombre).localeCompare(String(b.p.Nombre), "es"));
    const n = (nv) => base.filter((x) => x.V.nivel === nv).length;
    const todosB = lista.flatMap((x) => x.V.bs.map((y) => ({ ...y, p: x.p })));
    const conI = lista.filter((x) => x.V.indice !== null);
    const indiceP = conI.length ? Math.round(conI.reduce((t, x) => t + x.V.indice * (PESO_IMP[x.V.nivel] || 1), 0) / conI.reduce((t, x) => t + (PESO_IMP[x.V.nivel] || 1), 0)) : null;
    const valorFin = lista.reduce((t, x) => t + x.V.valorFin, 0);
    const dims = [...new Set([...dimensiones(), ...todosB.map((x) => x.b.Dimension).filter(Boolean)])];
    const filaDim = (d) => {
      const ps = lista.filter((x) => x.V.imp[d] || x.V.bs.some((y) => y.b.Dimension === d));
      const bs = todosB.filter((x) => x.b.Dimension === d), cl = bs.filter((x) => x.e.logro !== null);
      if (!ps.length) return "";
      return `<tr><td class="nowrap"><span class="dim-ico">${ICO_DIM[d] || "•"}</span> ${esc(d)}</td><td>${ps.length}</td><td>${IMPACTOS_DIM.map((i) => { const k = lista.filter((x) => x.V.imp[d] === i).length; return k ? `${chipImp(i)} ${k}` : ""; }).filter(Boolean).join(" ") || "—"}</td>
        <td>${bs.length}</td><td>${cl.length ? barraLogro(Math.round(cl.reduce((t, x) => t + x.e.logro, 0) / cl.length)) : `<span class="sub">sin medir</span>`}</td><td>${bs.filter((x) => x.e.estado === "En riesgo").length || "—"}</td></tr>`;
    };
    const pend = todosB.filter((x) => x.e.pendiente).sort((a, b) => String(a.e.proxima).localeCompare(String(b.e.proxima)));
    const riesgo = todosB.filter((x) => x.e.estado === "En riesgo").sort((a, b) => (PESO_IMP[b.b.Importancia] || 0) - (PESO_IMP[a.b.Importancia] || 0));
    el.innerHTML = `<h1>Valor del portafolio</h1>
      ${barraFiltros()}
      <div class="filtros"><label class="filtro"><span>Nivel de valor</span><select id="vl-nivel"><option value="">Todos</option>${NIVELES_VALOR.map((v) => `<option ${v === F.nivel ? "selected" : ""}>${v}</option>`).join("")}<option value="sin" ${F.nivel === "sin" ? "selected" : ""}>Sin definir</option></select></label><span class="sub">${lista.length} proyecto(s) · incluye cerrados que siguen midiendo beneficios</span></div>
      <div class="kpis kpis-3">
        <div class="kpi estatico"><span class="kpi-t">Proyectos por nivel de valor</span><span class="kpi-v kpi-chico">${NIVELES_VALOR.map((v) => `<span class="pill nv ${clsNivel(v)}">${v} ${n(v)}</span>`).join(" ")}</span><span class="sub">${base.filter((x) => !x.V.nivel).length} sin definir</span></div>
        <div class="kpi estatico"><span class="kpi-t">Índice de valor del portafolio</span><span class="kpi-v">${indiceP === null ? "—" : `${indiceP}%`}</span><span class="sub">logro de beneficios, ponderado por nivel</span></div>
        <div class="kpi estatico ${riesgo.length ? "alerta" : ""}"><span class="kpi-t">Beneficios en riesgo</span><span class="kpi-v">${riesgo.length}</span><span class="sub">${riesgo.filter((x) => x.b.Importancia === "Crítico").length} crítico(s) · ${todosB.length} beneficio(s) en total</span></div>
        <div class="kpi estatico ${pend.length ? "alerta" : ""}"><span class="kpi-t">Mediciones pendientes</span><span class="kpi-v">${pend.length}</span><span class="sub">${valorFin ? `valor financiero meta ${cop(valorFin)}` : "las registra el PM"}</span></div>
      </div>
      <div class="card"><h2>Aporte por dimensión de valor</h2>
        ${dims.map(filaDim).join("") ? `<div class="tabla-scroll"><table class="no-orden"><thead><tr><th>Dimensión</th><th>Proyectos</th><th>Impacto</th><th>Beneficios</th><th>Logro promedio</th><th>En riesgo</th></tr></thead><tbody>${dims.map(filaDim).join("")}</tbody></table></div>` : vacio("Aún no hay proyectos con propuesta de valor. Se registra en la pestaña «Valor» de cada proyecto.")}</div>
      <div class="card"><h2>Valor por proyecto</h2>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Proyecto</th><th>PM</th><th>Nivel de valor</th><th>Dimensiones con impacto alto</th><th>Índice de valor</th><th>Beneficios</th><th>En riesgo</th><th>Próxima medición</th></tr></thead><tbody>
          ${lista.map(({ p, V }) => { const prox = V.bs.map((x) => x.e.proxima).filter(Boolean).sort()[0]; return `<tr class="clic" data-vl-pid="${esc(p.ID_Proyecto)}"><td><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.ID_Proyecto)} · ${esc(p.Estado)}</div>${p.Valor_FSFB ? `<div class="sub vl-frase">${esc(p.Valor_FSFB)}</div>` : ""}</td><td>${persona(nombreUsuario(p.PM))}</td>
            <td${ed({ puede: R.puedeEditar(S.usuario, p, "editarProyecto"), tabla: "Proyectos", idCol: "ID_Proyecto", id: p.ID_Proyecto, campo: "Nivel_Valor", tipo: "select", opciones: NIVELES_VALOR, vacio: `Automático (${V.sugerido || "sin datos"})`, valor: p.Nivel_Valor || "", extra: () => sello(), etiqueta: "Nivel de valor" })}>${chipNivel(V.nivel, !V.manual && !!V.nivel)}</td><td>${Object.entries(V.imp).filter(([, v]) => v === "Alto").map(([d]) => `<span class="chip-act">${ICO_DIM[d] || "•"} ${esc(d)}</span>`).join(" ") || `<span class="sub">—</span>`}</td>
            <td class="nowrap">${barraLogro(V.indice)}</td><td>${V.bs.length}</td><td>${V.riesgo ? `<b class="baja">${V.riesgo}</b>` : "—"}</td><td class="nowrap">${prox ? `${fecha(prox)}${V.pendientes ? ` <span class="baja">⚠</span>` : ""}` : "—"}</td></tr>`; }).join("")}</tbody></table></div>`
        : vacio("No hay proyectos con estos filtros.")}</div>
      ${pend.length || riesgo.length ? `<div class="card"><h2>Para atender</h2><div class="tabla-scroll"><table class="no-orden"><thead><tr><th>Beneficio</th><th>Proyecto</th><th>PM</th><th>Importancia</th><th>Situación</th></tr></thead><tbody>
        ${[...new Set([...riesgo, ...pend])].slice(0, 30).map((x) => `<tr class="clic" data-vl-pid="${esc(x.p.ID_Proyecto)}"><td>${esc(x.b.Indicador || x.b.Beneficio_Esperado)}<div class="sub">${esc(x.b.Dimension)}</div></td><td>${esc(x.p.Nombre)}</td><td>${esc(nombreUsuario(x.p.PM))}</td><td>${chipNivel(x.b.Importancia)}</td>
          <td>${x.e.estado === "En riesgo" ? chipEstB("En riesgo") : ""} ${x.e.pendiente ? `<span class="pill vencido">Medición pendiente desde ${fecha(x.e.proxima)}</span>` : ""}</td></tr>`).join("")}</tbody></table></div></div>` : ""}`;
    enlazarFiltros();
    $("#vl-nivel").addEventListener("change", (e) => { S.valF = { ...F, nivel: e.target.value }; render(); });
    el.querySelectorAll("[data-vl-pid]").forEach((r) => r.addEventListener("click", () => { S.fichaTab = "valor"; S.fichaPid = r.dataset.vlPid; ir("ficha", r.dataset.vlPid); }));
  }

  // ---------- Informe ejecutivo en PowerPoint (marca FSFB) ----------
  let cargaPptx = null;
  function libPptx() {
    if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
    if (!cargaPptx) cargaPptx = new Promise((ok, falla) => {
      const s = document.createElement("script");
      s.src = new URL("lib/pptxgen.bundle.js", location.href).href;
      s.onload = () => ok(window.PptxGenJS); s.onerror = () => { cargaPptx = null; falla(new Error("No se pudo cargar el generador de PowerPoint.")); };
      document.head.appendChild(s);
    });
    return cargaPptx;
  }
  async function imagenDataURL(ruta) {
    const r = await fetch(new URL(ruta, location.href).href, { cache: "force-cache" });
    if (!r.ok) throw new Error(`No se encontró ${ruta}`);
    const b = await r.blob();
    return new Promise((ok) => { const fr = new FileReader(); fr.onload = () => ok(String(fr.result).replace(/^data:/, "")); fr.readAsDataURL(b); });
  }
  // Colores y fuente de la plantilla oficial FSFB 2026 (azul de títulos 00109E, turquesa de acento 00A3C4, Century Gothic).
  const PPT = { azul: "00109E", claro: "00A3C4", dorado: "00A3C4", suave: "EAF2FB", texto: "333333", sec: "666666", linea: "CCCCCC", blanco: "FFFFFF",
    Verde: "2E8B57", Amarillo: "E0A800", Rojo: "C0392B", gris: "9AA5B1", fuente: "Century Gothic" };
  const MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const fechaLarga = (iso) => { const [y, m, d] = String(iso).split("-").map(Number); return `${d} de ${MESES_LARGOS[m - 1]} de ${y}`; };
  const corta = (s, n) => { s = String(s ?? "").replace(/\s+/g, " ").trim(); return s.length > n ? s.slice(0, n - 1) + "…" : s; };

  function formInformePPT() {
    const vis = visibles().slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
    const pms = [...new Set(vis.map((p) => lc(p.PM)).filter(Boolean))].map((c) => [c, nombreUsuario(c)]).sort((a, b) => String(a[1]).localeCompare(String(b[1]), "es"));
    const yoPM = R.soloPM(S.usuario);
    const pmIni = S.filtros.pm ? lc(S.filtros.pm) : "";
    const ck = (cls, val, txt, sub, on, extra = "") => `<label class="check inf-op" ${extra}><input type="checkbox" class="${cls}" value="${esc(val)}" ${on ? "checked" : ""}> <span>${esc(txt)}${sub ? ` <span class="sub">${esc(sub)}</span>` : ""}</span></label>`;
    const html = `<div class="inf-alc">
        <p class="sub">Elige uno o varios PM y/o uno o varios proyectos. Si no marcas nada, sale todo el portafolio. Si combinas, entran los proyectos de los PM marcados más los proyectos marcados.</p>
        <div class="inf-cols">
          <div class="inf-col"><div class="inf-cab"><b>PM</b> <span class="sub" id="inf-npm"></span> ${pms.length > 1 ? `<button type="button" class="btn enlace" data-inf-todos="inf-pm">Todos</button> <button type="button" class="btn enlace" data-inf-ninguno="inf-pm">Ninguno</button>` : ""}</div>
            <div class="inf-lista">${pms.map(([c, n]) => ck("inf-pm", c, n, `${vis.filter((p) => lc(p.PM) === c && CUENTAN(p)).length} activos`, c === pmIni)).join("")}</div></div>
          <div class="inf-col"><div class="inf-cab"><b>Proyectos</b> <span class="sub" id="inf-npr"></span> <button type="button" class="btn enlace" data-inf-ninguno="inf-pr">Ninguno</button></div>
            <input type="search" id="inf-q" placeholder="Buscar proyecto, ID o PM" aria-label="Buscar proyecto">
            <label class="check inf-op"><input type="checkbox" id="inf-cerrados"> <span class="sub">Mostrar también cerrados y cancelados</span></label>
            <div class="inf-lista inf-lista-pr">${vis.map((p) => ck("inf-pr", p.ID_Proyecto, p.Nombre, `${p.ID_Proyecto} · ${nombreUsuario(p.PM)}${CUENTAN(p) ? "" : ` · ${p.Estado}`}`, false, `data-txt="${esc(normTxt(`${p.Nombre} ${p.ID_Proyecto} ${nombreUsuario(p.PM)}`))}" data-activo="${CUENTAN(p) ? 1 : 0}"`)).join("")}</div></div>
        </div>
        <div class="inf-resumen" id="inf-res"></div></div>`;
    const campos = [
      { html },
      { k: "Corte", label: "Fecha de corte", tipo: "date", def: R.hoyISO(), ayuda: "Seguimientos, compromisos y capacidad se calculan a esta fecha." },
      { k: "Incluir", label: "Láminas a incluir", tipo: "checks", opciones: [["pm", "Resumen y sección por cada PM"], ["proyectos", "Estado de los proyectos"], ["capacidad", "Capacidad del equipo"], ["seguimientos", "Seguimientos"], ["compromisos", "Compromisos y riesgos"], ["gantt", "Cronograma del portafolio (Gantt agrupado por PM)"], ["valor", "Valor para la FSFB (portafolio y, con la lámina por proyecto, el de cada uno)"],
        ["fichas", "Una lámina por proyecto (logros, próximos pasos, hitos)"], ["edt", "EDT y cronograma de cada proyecto (los que tienen cronograma cargado)"]],
        def: "pm, proyectos, gantt, valor, capacidad, seguimientos, compromisos, edt" },
    ];
    const marcados = (m, cls) => [...m.querySelectorAll(`.${cls}:checked`)].map((i) => i.value);
    modal("Informe de estado en PowerPoint", campos, {}, (fd, sel) => { cerrarModal(); generarInformePPT({ pms: new Set(sel.pms), pids: new Set(sel.pids), corte: fd.Corte || R.hoyISO(), incluir: new Set(String(fd.Incluir).split(/,\s*/).filter(Boolean)) }); },
      (fd) => (fd.Corte && !/^\d{4}-\d{2}-\d{2}$/.test(fd.Corte) ? "Fecha de corte no válida." : ""),
      { textoBoton: "Generar PowerPoint", antes: `<p class="sub">Presentación sobre la plantilla FSFB (portada, resumen ejecutivo y las láminas que elijas). Se descarga como .pptx y la puedes editar en PowerPoint.</p>`,
        init: (m) => {
          const q = m.querySelector("#inf-q"), cerr = m.querySelector("#inf-cerrados");
          const filtrar = () => { const t = normTxt(q.value); m.querySelectorAll(".inf-lista-pr .inf-op").forEach((l) => { const i = l.querySelector("input"); l.hidden = !(i.checked || ((!t || l.dataset.txt.includes(t)) && (cerr.checked || l.dataset.activo === "1"))); }); };
          const resumir = () => {
            const a = marcados(m, "inf-pm"), b = marcados(m, "inf-pr");
            m.querySelector("#inf-npm").textContent = a.length ? `(${a.length})` : "";
            m.querySelector("#inf-npr").textContent = b.length ? `(${b.length})` : "";
            const n = vis.filter((p) => (!a.length && !b.length ? CUENTAN(p) : (a.includes(lc(p.PM)) && CUENTAN(p)) || b.includes(p.ID_Proyecto))).length;
            m.querySelector("#inf-res").innerHTML = `<b>Alcance:</b> ${!a.length && !b.length ? "todo el portafolio" : [a.length ? `${a.length === 1 ? "PM" : `${a.length} PM`}: ${a.map(nombreUsuario).join(", ")}` : "", b.length ? `${b.length} proyecto(s) marcado(s)` : ""].filter(Boolean).join(" + ")} · <b>${n}</b> proyecto(s) en el informe`;
          };
          q.addEventListener("input", filtrar); cerr.addEventListener("change", filtrar);
          m.addEventListener("change", (e) => { if (e.target.matches(".inf-pm, .inf-pr")) { resumir(); filtrar(); } });
          m.querySelectorAll("[data-inf-todos]").forEach((b) => b.addEventListener("click", () => { m.querySelectorAll(`.${b.dataset.infTodos}`).forEach((i) => { i.checked = true; }); resumir(); }));
          m.querySelectorAll("[data-inf-ninguno]").forEach((b) => b.addEventListener("click", () => { m.querySelectorAll(`.${b.dataset.infNinguno}`).forEach((i) => { i.checked = false; }); resumir(); filtrar(); }));
          if (yoPM) m.querySelectorAll(".inf-pm").forEach((i) => { i.checked = true; });
          filtrar(); resumir();
        },
        recoger: (m) => ({ datos: { pms: marcados(m, "inf-pm"), pids: marcados(m, "inf-pr") } }) });
  }

  async function generarInformePPT({ pms = new Set(), pids = new Set(), corte, incluir }) {
    cargando(true, "Generando PowerPoint…");
    try {
      // Fondos, detalle de esquina y logo tomados de la plantilla FSFB_Plantilla_2026.pptx.
      const [Pptx, imgPortada, imgOscura, imgCierre, imgEsquina, imgLogo] = await Promise.all([libPptx(), ...["portada.jpg", "oscura.jpg", "cierre.jpg", "esquina.png", "logo_pie.png"].map((f) => imagenDataURL(`assets/ppt/${f}`))]);
      const segs = S.datos.Seguimientos;
      const todos = visibles();
      // Alcance: todo el portafolio, o los proyectos de los PM marcados más los proyectos marcados (estos aunque estén cerrados).
      const libre = !pms.size && !pids.size;
      const alcance = todos.filter((p) => libre || pms.has(lc(p.PM)) || pids.has(p.ID_Proyecto));
      const ps = alcance.filter((p) => CUENTAN(p) || pids.has(p.ID_Proyecto)).sort((a, b) => String(nombreUsuario(a.PM)).localeCompare(String(nombreUsuario(b.PM)), "es") || String(a.Nombre).localeCompare(String(b.Nombre), "es"));
      const pmsEnInforme = [...new Set(ps.map((p) => lc(p.PM)))];
      const pm = !libre && pmsEnInforme.length === 1 ? pmsEnInforme[0] : "";   // un solo PM: informe "del PM" (sin secciones por PM)
      const nomPM = pm ? nombreUsuario(pm) : "";
      const alcanceTxt = libre ? "" : pids.size && !pms.size && ps.length === 1 ? `Proyecto: ${ps[0].Nombre}` : nomPM && !pids.size ? `PM: ${nomPM}`
        : [pms.size ? `PM: ${[...pms].map(nombreUsuario).join(", ")}` : "", pids.size ? `${pids.size} proyecto(s) seleccionado(s)` : ""].filter(Boolean).join(" + ");
      const pptx = new Pptx();
      pptx.layout = "LAYOUT_WIDE";   // 13,33 × 7,5 in
      pptx.author = S.usuario ? S.usuario.Nombre || S.usuario.Correo : "Oficina de Proyectos";
      pptx.company = "Fundación Santa Fe de Bogotá"; pptx.title = "Estado del portafolio de proyectos";
      const F = PPT.fuente, W = 13.333;
      const pie = `Oficina de Proyectos · Fundación Santa Fe de Bogotá · Corte ${fecha(corte)}${alcanceTxt ? ` · ${corta(alcanceTxt, 70)}` : ""}`;
      // Layouts de la plantilla: portada (foto), separador oscuro (curvas), contenido (esquina + logo) y cierre (edificio).
      pptx.defineSlideMaster({ title: "PORTADA", background: { data: imgPortada } });
      pptx.defineSlideMaster({ title: "SECCION", background: { data: imgOscura } });
      pptx.defineSlideMaster({ title: "CIERRE", background: { data: imgCierre } });
      pptx.defineSlideMaster({ title: "FSFB", background: { color: PPT.blanco }, margin: [0.4, 0.6, 0.4, 0.6],
        objects: [
          { image: { x: 11.737, y: 0, w: 1.596, h: 1.94, data: imgEsquina } },
          { image: { x: 11.217, y: 6.745, w: 1.714, h: 0.498, data: imgLogo } },
          { text: { text: pie, options: { x: 0.6, y: 6.95, w: 9.6, h: 0.3, fontFace: F, fontSize: 8, color: PPT.sec, valign: "middle", margin: 0 } } },
        ],
        slideNumber: { x: 10.3, y: 6.95, w: 0.8, h: 0.3, fontFace: F, fontSize: 8, color: PPT.sec, align: "right" } });
      // El contenido se diseñó sobre 0,5–12,83 × 1,5–7,05 in; se lleva a la zona útil de la plantilla (0,6–11,6 × 1,45–6,65 in)
      // y la letra se reduce un poco porque Century Gothic es más ancha.
      const KZ = 0.89, KF = 0.88;
      const fz = (v) => Math.max(7, Math.round(v * KF * 2) / 2);
      const escOpts = (o0) => {
        if (!o0 || typeof o0 !== "object") return o0;
        const o = { ...o0 };
        if (typeof o.x === "number") o.x = 0.6 + (o.x - 0.5) * KZ;
        if (typeof o.y === "number") o.y = 1.45 + (o.y - 1.5) * KZ;
        if (typeof o.w === "number") o.w *= KZ;
        if (typeof o.h === "number") o.h *= KZ;
        if (Array.isArray(o.colW)) o.colW = o.colW.map((c) => c * KZ);
        if (typeof o.rowH === "number") o.rowH *= KZ;
        if (typeof o.rectRadius === "number") o.rectRadius *= KZ;
        Object.keys(o).forEach((k) => { if (/fontsize$/i.test(k) && typeof o[k] === "number") o[k] = fz(o[k]); if (/fontface$/i.test(k)) o[k] = F; });
        return o;
      };
      const escTexto = (t) => (Array.isArray(t) ? t.map((r) => (r && r.options ? { ...r, options: { ...r.options, ...(r.options.fontSize ? { fontSize: fz(r.options.fontSize) } : {}) } } : r)) : t);
      const zona = (sl) => ({
        addText: (t, o) => sl.addText(escTexto(t), escOpts(o)),
        addShape: (tipo, o) => sl.addShape(tipo, escOpts(o)),
        addChart: (tipo, d, o) => sl.addChart(tipo, d, escOpts({ fontFace: F, catAxisLabelFontFace: F, valAxisLabelFontFace: F, legendFontFace: F, titleFontFace: F, dataLabelFontFace: F, ...o })),
        addImage: (o) => sl.addImage(escOpts(o)),
        addTable: (filas, o) => sl.addTable(filas.map((f) => f.map((c) => (c && c.options && c.options.fontSize ? { ...c, options: { ...c.options, fontSize: fz(c.options.fontSize) } } : c))), escOpts(o)),
      });
      const lamina = (titulo, sub) => {
        const s = pptx.addSlide({ masterName: "FSFB" });
        s.addText(titulo, { x: 0.6, y: 0.4, w: 10.9, h: 0.55, fontFace: F, fontSize: 24, bold: true, color: PPT.azul, valign: "middle", margin: 0, fit: "shrink" });
        if (sub) s.addText(sub, { x: 0.6, y: 0.97, w: 10.9, h: 0.38, fontFace: F, fontSize: 11, color: PPT.sec, valign: "top", margin: 0, fit: "shrink" });
        return zona(s);
      };
      const vacioTxt = (s, msg, y = 2.6) => s.addText(msg, { x: 0.5, y, w: W - 1, h: 0.6, fontFace: F, fontSize: 14, color: PPT.sec, italic: true, align: "center" });
      const tile = (s, x, y, w, h, titulo, valor, nota, color = PPT.azul) => {
        s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: PPT.suave }, line: { color: PPT.suave }, rectRadius: 0.08 });
        s.addShape(pptx.ShapeType.rect, { x, y: y + 0.12, w: 0.07, h: h - 0.24, fill: { color }, line: { color } });
        s.addText(titulo, { x: x + 0.2, y: y + 0.08, w: w - 0.3, h: 0.36, fontFace: F, fontSize: 11, color: PPT.sec, margin: 0 });
        s.addText(String(valor), { x: x + 0.2, y: y + 0.4, w: w - 0.3, h: 0.52, fontFace: F, fontSize: 28, bold: true, color, margin: 0, valign: "middle" });
        if (nota) s.addText(nota, { x: x + 0.2, y: y + h - 0.36, w: w - 0.3, h: 0.28, fontFace: F, fontSize: 10, color: PPT.sec, margin: 0 });
      };
      const cab = (t, extra = {}) => ({ text: t, options: { bold: true, color: PPT.blanco, fill: { color: PPT.azul }, fontSize: 10, valign: "middle", ...extra } });
      const celda = (t, i, extra = {}) => ({ text: String(t ?? ""), options: { fill: { color: i % 2 ? PPT.suave : PPT.blanco }, color: PPT.texto, fontSize: 10, valign: "middle", ...extra } });
      const semCelda = (sem, i) => { const c = PPT[sem]; return c ? { text: sem, options: { fill: { color: c }, color: sem === "Amarillo" ? PPT.texto : PPT.blanco, bold: true, fontSize: 10, align: "center", valign: "middle" } } : celda(sem || "—", i, { align: "center" }); };
      const tabla = (s, filas, colW, y = 1.55, extra = {}) => s.addTable(filas, { x: 0.5, y, w: colW.reduce((a, b) => a + b, 0), colW, fontFace: F, border: { type: "solid", pt: 0.5, color: PPT.linea }, rowH: 0.36, margin: [2, 5, 2, 5], ...extra });
      const paginar = (lista, n) => { const out = []; for (let i = 0; i < lista.length; i += n) out.push(lista.slice(i, i + n)); return out.length ? out : [[]]; };

      // Métricas comunes
      const metricas = (lista) => {
        const act = lista.filter((p) => p.Estado === "Activo");
        const prom = (f) => (act.length ? Math.round(act.reduce((t, p) => t + (Number(p[f]) || 0), 0) / act.length) : 0);
        const ids = new Set(lista.map((p) => p.ID_Proyecto));
        const est = act.map((p) => R.estadoSeguimiento(p, segs, corte).estado);
        const comp = S.datos.Compromisos.filter((c) => ids.has(c.ID_Proyecto)).map((c) => ({ c, e: R.estadoCompromiso(c, corte) }));
        const pres = act.reduce((t, p) => t + (Number(p.Presupuesto) || 0), 0), ejec = act.reduce((t, p) => t + (Number(p.Ejecutado) || 0), 0);
        return { act, ids, pausa: lista.filter((p) => p.Estado === "En pausa").length, real: prom("Avance_Real"), plan: prom("Avance_Planeado"),
          sem: ["Verde", "Amarillo", "Rojo"].map((s) => act.filter((p) => p.Semaforo === s).length),
          alDia: act.length ? Math.round((est.filter((e) => e !== "Vencido").length / act.length) * 100) : 100, vencSeg: est.filter((e) => e === "Vencido").length,
          compV: comp.filter((x) => x.e === "Vencido").length, compPV: comp.filter((x) => x.e === "Por vencer").length, compAb: comp.filter((x) => x.e !== "Cerrado").length, pres, ejec };
      };
      const M = metricas(ps);
      const cap = capacidad(corte).filter((x) => x.asign.some((a) => M.ids.has(a.p.ID_Proyecto)) || (pm && lc(x.correo) === pm));
      // ---- Valor para la FSFB ----
      const COL_NIV = { "Crítico": "8E1B3E", Alto: PPT.azul, Medio: PPT.claro, Bajo: "AEB8C6" };
      const nivCelda = (n, i) => (n ? { text: n, options: { fill: { color: COL_NIV[n] || PPT.gris }, color: n === "Medio" || n === "Bajo" ? PPT.texto : PPT.blanco, bold: true, fontSize: 10, align: "center", valign: "middle" } } : celda("—", i, { align: "center" }));
      const estColor = (e) => (e === "Logrado" ? PPT.Verde : e === "En riesgo" ? PPT.Rojo : e === "En camino" ? PPT.azul : PPT.sec);
      const valorPortafolio = () => {
        const base = alcance.filter((p) => p.Estado !== "Cancelado").map((p) => ({ p, V: valorProyecto(p, corte) })).filter((x) => x.V.tieneAlgo);
        const s = lamina("Valor para la FSFB", "Nivel de valor de los proyectos, aporte por dimensión y logro de los beneficios (no solo financiero)");
        if (!base.length) return vacioTxt(s, "Aún no hay proyectos con propuesta de valor. Se registra en la pestaña «Valor» de cada proyecto.");
        NIVELES_VALOR.forEach((n, k) => tile(s, 0.5 + k * 1.55, 1.5, 1.45, 1.25, n, base.filter((x) => x.V.nivel === n).length, "proyecto(s)", COL_NIV[n]));
        const bs = base.flatMap((x) => x.V.bs), conI = base.filter((x) => x.V.indice !== null);
        const idx = conI.length ? Math.round(conI.reduce((t, x) => t + x.V.indice * (PESO_IMP[x.V.nivel] || 1), 0) / conI.reduce((t, x) => t + (PESO_IMP[x.V.nivel] || 1), 0)) : null;
        tile(s, 0.5, 2.9, 3.0, 1.25, "Índice de valor del portafolio", idx === null ? "—" : `${idx}%`, "logro ponderado por nivel", PPT.dorado);
        const rg = bs.filter((x) => x.e.estado === "En riesgo").length;
        tile(s, 3.6, 2.9, 2.6, 1.25, "Beneficios en riesgo", rg, `${bs.filter((x) => x.e.pendiente).length} mediciones pendientes`, rg ? PPT.Rojo : PPT.Verde);
        const dims = dimensiones().filter((d) => base.some((x) => x.V.imp[d]));
        if (dims.length) s.addChart(pptx.ChartType.bar, IMPACTOS_DIM.map((i) => ({ name: `Impacto ${i.toLowerCase()}`, labels: dims, values: dims.map((d) => base.filter((x) => x.V.imp[d] === i).length) })),
          { x: 0.5, y: 4.3, w: 5.8, h: 2.7, barDir: "bar", barGrouping: "stacked", chartColors: [PPT.azul, PPT.claro, "C9D3E0"], catAxisOrientation: "maxMin", catAxisLabelFontSize: 9, valAxisLabelFontSize: 8,
            showLegend: true, legendPos: "t", legendFontSize: 9, showValue: true, dataLabelFontSize: 8, dataLabelFormatCode: "0;;;", dataLabelColor: PPT.blanco, showTitle: true, title: "Proyectos por dimensión de valor", titleFontSize: 11, titleColor: PPT.azul, valGridLine: { color: "E5E5E5", size: 0.5 } });
        const top = base.sort((a, b) => (PESO_IMP[b.V.nivel] || 0) - (PESO_IMP[a.V.nivel] || 0) || (b.V.indice ?? -1) - (a.V.indice ?? -1)).slice(0, 13);
        const filas = [[cab("Proyecto"), cab("Nivel", { align: "center" }), cab("Logro", { align: "center" }), cab("En riesgo", { align: "center" })]];
        top.forEach(({ p, V }, i) => filas.push([celda(corta(p.Nombre, 34), i, { bold: true }), nivCelda(V.nivel, i), celda(V.indice === null ? "—" : `${V.indice}%`, i, { align: "center" }), celda(V.riesgo || "—", i, { align: "center", color: V.riesgo ? PPT.Rojo : PPT.texto, bold: !!V.riesgo })]));
        tabla(s, filas, [3.3, 1.0, 0.8, 0.88], 1.5, { x: 6.85, rowH: 0.36 });
      };
      const valorDeProyecto = (p) => {
        const V = valorProyecto(p, corte);
        if (!V.tieneAlgo) return;
        const s = lamina(`Valor · ${corta(p.Nombre, 52)}`, `Propuesta de valor para la FSFB${V.indice !== null ? ` · índice de valor ${V.indice}%` : ""}${V.valorFin ? ` · valor financiero meta ${cop(V.valorFin)}${V.roi ? ` (ROI ${V.roi.toFixed(1).replace(".", ",")}x)` : ""}` : ""}`);
        const caja = (x, w, tit, txt) => {
          s.addShape(pptx.ShapeType.roundRect, { x, y: 1.5, w, h: 1.55, fill: { color: PPT.suave }, line: { color: PPT.suave }, rectRadius: 0.06 });
          s.addText(tit, { x: x + 0.15, y: 1.55, w: w - 0.3, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: PPT.azul, margin: 0 });
          s.addText(corta(txt || "Sin registrar.", 300), { x: x + 0.15, y: 1.86, w: w - 0.3, h: 1.12, fontFace: F, fontSize: 11, color: txt ? PPT.texto : PPT.sec, italic: !txt, valign: "top", margin: 0, fit: "shrink" });
        };
        caja(0.5, 4.9, "Beneficio esperado", p.Beneficio_Esperado);
        caja(5.55, 4.9, "Valor para la FSFB", p.Valor_FSFB);
        s.addShape(pptx.ShapeType.roundRect, { x: 10.6, y: 1.5, w: 2.23, h: 1.55, fill: { color: COL_NIV[V.nivel] || PPT.gris }, line: { color: COL_NIV[V.nivel] || PPT.gris }, rectRadius: 0.06 });
        s.addText([{ text: "Nivel de valor", options: { fontSize: 10, breakLine: true } }, { text: V.nivel || "Sin definir", options: { fontSize: 24, bold: true, breakLine: true } }, { text: V.nivel && !V.manual ? "sugerido" : "", options: { fontSize: 9 } }],
          { x: 10.6, y: 1.5, w: 2.23, h: 1.55, fontFace: F, color: V.nivel === "Medio" || V.nivel === "Bajo" ? PPT.texto : PPT.blanco, align: "center", valign: "middle" });
        const imps = Object.entries(V.imp);
        let x = 0.5;
        s.addText("Impacto:", { x, y: 3.2, w: 0.8, h: 0.32, fontFace: F, fontSize: 10, bold: true, color: PPT.azul, margin: 0, valign: "middle" }); x += 0.8;
        imps.forEach(([d, i]) => { const w = 0.35 + (d.length + i.length) * 0.075; if (x + w > 12.8) return; s.addShape(pptx.ShapeType.roundRect, { x, y: 3.22, w, h: 0.28, fill: { color: i === "Alto" ? PPT.azul : i === "Medio" ? PPT.claro : "DCE3EC" }, line: { color: PPT.blanco }, rectRadius: 0.1 });
          s.addText(`${d}: ${i}`, { x, y: 3.22, w, h: 0.28, fontFace: F, fontSize: 9, color: i === "Alto" ? PPT.blanco : PPT.texto, align: "center", valign: "middle", margin: 0 }); x += w + 0.1; });
        if (!imps.length) s.addText("Sin marcar", { x, y: 3.2, w: 3, h: 0.32, fontFace: F, fontSize: 10, color: PPT.sec, italic: true, margin: 0, valign: "middle" });
        if (!V.bs.length) return vacioTxt(s, "Sin beneficios medibles registrados.", 4.2);
        const filas = [[cab("Importancia", { align: "center" }), cab("Dimensión"), cab("Beneficio / indicador"), cab("Base → meta", { align: "center" }), cab("Actual", { align: "center" }), cab("Logro", { align: "center" }), cab("Estado", { align: "center" })]];
        V.bs.slice(0, 7).forEach(({ b, e }, i) => filas.push([nivCelda(b.Importancia, i), celda(b.Dimension, i), celda(corta([b.Indicador, b.Beneficio_Esperado].filter(Boolean).join(" — "), 90), i),
          celda(b.Tipo_Medida === "Cualitativo" ? "Cualitativo" : `${fmtValor(b, b.Linea_Base)} → ${fmtValor(b, b.Meta)}`, i, { align: "center" }), celda(e.ult ? fmtValor(b, e.ult.Valor) : "—", i, { align: "center" }),
          celda(e.logro === null ? "—" : `${e.logro}%`, i, { align: "center", bold: true }), celda(e.estado, i, { align: "center", bold: true, color: estColor(e.estado) })]));
        tabla(s, filas, [1.05, 1.9, 4.2, 1.75, 1.3, 0.8, 1.33], 3.7, { rowH: 0.4, fontSize: 9 });
        if (V.bs.length > 7) s.addText(`y ${V.bs.length - 7} beneficio(s) más`, { x: 0.5, y: 6.8, w: 6, h: 0.25, fontFace: F, fontSize: 9, color: PPT.sec });
      };
      // ---- Gantt genérico con figuras (portafolio o proyecto) ----
      const MES3 = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
      const ganttPPT = (s, filas, ini, fin, o = {}) => {
        const X0 = o.x0 || 4.9, X1 = 12.83, Y0 = o.y0 || 1.85, Yf = 6.55, colTxt = X0 - 0.5 - (o.colPct ? 0.75 : 0);
        const total = Math.max(1, R.difDias(fin, ini));
        const xd = (d) => X0 + (Math.min(Math.max(R.difDias(d, ini), 0), total) / total) * (X1 - X0);
        const rh = Math.min(0.36, (Yf - Y0) / Math.max(1, filas.length));
        // Meses (o trimestres si el rango es largo)
        const meses = Math.ceil(total / 30.4), paso = meses > 30 ? 3 : meses > 15 ? 2 : 1;
        let [y, m] = ini.split("-").map(Number); m -= 1;
        for (let k = 0; k < 400; k++) {
          const iso = `${y}-${String(m + 1).padStart(2, "0")}-01`;
          if (iso > fin) break;
          if (iso >= ini && k % paso === 0) {
            const xx = xd(iso);
            s.addShape(pptx.ShapeType.line, { x: xx, y: Y0 - 0.05, w: 0, h: rh * filas.length + 0.05, line: { color: m === 0 ? "B8C4D6" : "E3E8EF", width: m === 0 ? 1 : 0.5 } });
            s.addText(`${MES3[m]}${m === 0 || k === 0 ? ` ${String(y).slice(2)}` : ""}`, { x: xx, y: Y0 - 0.33, w: 0.7, h: 0.25, fontFace: F, fontSize: 8, color: m === 0 ? PPT.azul : PPT.sec, bold: m === 0, margin: 0 });
          }
          m++; if (m > 11) { m = 0; y++; }
        }
        filas.forEach((f, i) => {
          const yy = Y0 + i * rh;
          if (f.grupo) {
            s.addShape(pptx.ShapeType.rect, { x: 0.5, y: yy, w: X1 - 0.5, h: rh, fill: { color: "DCE7F5" }, line: { color: "DCE7F5" } });
            s.addText([{ text: f.txt, options: { bold: true, color: PPT.azul } }, { text: f.sub ? `   ${f.sub}` : "", options: { color: PPT.sec, fontSize: 8 } }], { x: 0.55, y: yy, w: X1 - 0.6, h: rh, fontFace: F, fontSize: 10, valign: "middle", margin: 0 });
            return;
          }
          if (i % 2) s.addShape(pptx.ShapeType.rect, { x: 0.5, y: yy, w: X1 - 0.5, h: rh, fill: { color: "F6F9FD" }, line: { color: "F6F9FD" } });
          s.addText([{ text: f.txt, options: { bold: !!f.negrita, color: f.colorTxt || PPT.texto } }, ...(f.sub ? [{ text: `  ${f.sub}`, options: { color: PPT.sec, fontSize: 7.5 } }] : [])],
            { x: 0.5 + (f.nivel || 0) * 0.18, y: yy, w: colTxt - (f.nivel || 0) * 0.18, h: rh, fontFace: F, fontSize: rh < 0.26 ? 8 : 9, valign: "middle", margin: [0, 2, 0, 2] });
          if (o.colPct && f.pct !== undefined) s.addText(`${f.pct}%`, { x: X0 - 0.8, y: yy, w: 0.7, h: rh, fontFace: F, fontSize: 8.5, align: "right", valign: "middle", color: f.colorPct || PPT.texto, bold: true, margin: 0 });
          if (f.base && f.base[0] && f.base[1] && f.base[1] >= ini && f.base[0] <= fin) {
            const a = xd(f.base[0]), b = xd(f.base[1]);
            s.addShape(pptx.ShapeType.rect, { x: a, y: yy + rh * 0.72, w: Math.max(0.03, b - a), h: rh * 0.14, fill: { color: "B5BFCC" }, line: { color: "B5BFCC" } });
          }
          if (f.a && f.b && f.b >= ini && f.a <= fin && !f.soloHito) {
            const a = xd(f.a), b = xd(f.b), w = Math.max(0.04, b - a), h = f.resumen ? rh * 0.3 : rh * 0.46, ty = yy + (f.base ? rh * 0.14 : (rh - h) / 2);
            s.addShape(pptx.ShapeType.rect, { x: a, y: ty, w, h, fill: { color: f.resumen ? "7C93B5" : PPT.claro }, line: { color: f.resumen ? "7C93B5" : PPT.claro } });
            const pc = Math.min(100, Math.max(0, Number(f.pct) || 0));
            if (pc > 0 && !f.resumen) s.addShape(pptx.ShapeType.rect, { x: a, y: ty, w: Math.max(0.02, (w * pc) / 100), h, fill: { color: f.colorReal || PPT.azul }, line: { color: f.colorReal || PPT.azul } });
            if (f.resumen && pc > 0) s.addShape(pptx.ShapeType.rect, { x: a, y: ty, w: Math.max(0.02, (w * pc) / 100), h, fill: { color: PPT.azul }, line: { color: PPT.azul } });
          }
          (f.hitos || []).filter((h) => h.f >= ini && h.f <= fin).forEach((h) => {
            const d = Math.min(0.2, rh * 0.62);
            s.addShape(pptx.ShapeType.diamond, { x: xd(h.f) - d / 2, y: yy + (rh - d) / 2, w: d, h: d, fill: { color: h.color }, line: { color: PPT.blanco, width: 0.5 } });
          });
        });
        if (corte >= ini && corte <= fin) {
          const xx = xd(corte);
          s.addShape(pptx.ShapeType.line, { x: xx, y: Y0 - 0.08, w: 0, h: rh * filas.length + 0.1, line: { color: PPT.dorado, width: 1.5, dashType: "dash" } });
          s.addText("Hoy", { x: xx - 0.3, y: Y0 + rh * filas.length + 0.02, w: 0.6, h: 0.2, fontFace: F, fontSize: 8, color: PPT.dorado, bold: true, align: "center", margin: 0 });
        }
      };
      const leyenda = (s, items, y = 6.86) => {
        let x = 0.5;
        items.forEach(([c, t, forma]) => {
          s.addShape(forma === "rombo" ? pptx.ShapeType.diamond : pptx.ShapeType.rect, { x, y: y + 0.04, w: 0.16, h: 0.14, fill: { color: c }, line: { color: c } });
          s.addText(t, { x: x + 0.2, y, w: 2.2, h: 0.22, fontFace: F, fontSize: 8, color: PPT.sec, margin: 0 });
          x += 0.35 + t.length * 0.065;
        });
      };
      const colorHitoP = (h) => (h.Estado === "Cumplido" ? PPT.Verde : h.Fecha_Plan < corte ? PPT.Rojo : "8A94A3");

      // ---- Cronograma del portafolio (Gantt agrupado por PM) ----
      const ganttPortafolio = () => {
        const mesDe = (iso, n) => { const [y, m] = iso.split("-").map(Number); const d = new Date(y, m - 1 + n, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`; };
        const ini = mesDe(corte, -3), fin = R.sumarDias(mesDe(corte, 10), -1);
        const conF = ps.filter((p) => p.Fecha_Inicio && p.Fecha_Fin_Plan);
        const filas = [];
        const filaP = (p) => {
          const d = (Number(p.Avance_Real) || 0) - (Number(p.Avance_Planeado) || 0);
          return { txt: corta(p.Nombre, 34), sub: `${Number(p.Avance_Real) || 0}% / ${Number(p.Avance_Planeado) || 0}%`, a: p.Fecha_Inicio, b: p.Fecha_Fin_Plan, pct: Number(p.Avance_Real) || 0,
            colorReal: p.Semaforo === "Rojo" || d < -15 ? PPT.Rojo : p.Semaforo === "Amarillo" ? "C98F00" : PPT.azul,
            hitos: S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto && h.Fecha_Plan).map((h) => ({ f: h.Fecha_Plan, color: colorHitoP(h) })), nivel: pm ? 0 : 1 };
        };
        if (pm) conF.forEach((p) => filas.push(filaP(p)));
        else gruposPM.forEach((g) => { const l = g.lista.filter((p) => p.Fecha_Inicio && p.Fecha_Fin_Plan); if (!l.length) return; filas.push({ grupo: true, txt: g.n, sub: `${l.length} proyecto(s)` }); l.sort((a, b) => String(a.Fecha_Inicio).localeCompare(String(b.Fecha_Inicio))).forEach((p) => filas.push(filaP(p))); });
        const pags = paginar(filas, 16);
        pags.forEach((pag, k) => {
          const s = lamina(pags.length > 1 ? `Cronograma del portafolio (${k + 1}/${pags.length})` : "Cronograma del portafolio", `${pm ? "" : "Agrupado por PM · "}${fecha(ini)} a ${fecha(fin)} · barra = duración planeada; relleno = avance real (color según semáforo)`);
          if (!pag.length) return vacioTxt(s, "No hay proyectos con fechas de inicio y fin.");
          if (k > 0 && !pag[0].grupo && !pm) { const g = [...filas.slice(0, filas.indexOf(pag[0]))].reverse().find((f) => f.grupo); if (g) pag.unshift({ ...g, sub: `${g.sub} (cont.)` }); }
          ganttPPT(s, pag, ini, fin, { x0: 4.6 });
          leyenda(s, [[PPT.claro, "Duración planeada"], [PPT.azul, "Avance real"], [PPT.Rojo, "Avance (rojo / atraso > 15 pp)"], [PPT.Verde, "Hito cumplido", "rombo"], [PPT.Rojo, "Hito atrasado", "rombo"], ["8A94A3", "Hito pendiente", "rombo"], [PPT.dorado, "Hoy"]]);
        });
      };

      // ---- EDT (árbol con avance real vs. plan por fase) y Gantt del cronograma de un proyecto ----
      const COL_EDT = { "ed-comp": ["D6E9D2", "4F8A43"], "ed-ok": ["E6F4E3", "7FB873"], "ed-warn": ["FFF3CC", "D9A400"], "ed-bad": ["FDE3E1", "C62828"], "ed-noini": ["F1F3F5", "B9C0C9"] };
      const laminasCronograma = (p) => {
        const ts = tareasDe(p.ID_Proyecto);
        if (!ts.length) return false;
        const v = versionVigente(p.ID_Proyecto);
        const niv = (i) => Number(ts[i].Nivel) || 1;
        const hijosDe = (i) => { const n = niv(i), out = []; for (let j = i + 1; j < ts.length && niv(j) > n; j++) out.push(j); return out; };
        const avance = (i) => { const desc = hijosDe(i).map((j) => ts[j]); const hojas = desc.length ? desc : [ts[i]]; const c = calcCronograma(hojas.map((t) => ({ ...t, Resumen: desc.length && t.Resumen === "Sí" ? "Sí" : "No" })), corte); return { real: c.real, plan: c.plan }; };
        const minN = Math.min(...ts.map((t, i) => niv(i)));
        const raices = ts.map((t, i) => i).filter((i) => niv(i) === minN);
        let raizTxt, raizAv, fases, edtRaiz = "";
        if (raices.length === 1 && hijosDe(raices[0]).length) { const r = raices[0]; edtRaiz = String(ts[r].EDT || ""); raizTxt = ts[r].Nombre; raizAv = avance(r); fases = hijosDe(r).filter((j) => niv(j) === minN + 1); }
        else { raizTxt = p.Nombre; raizAv = calcCronograma(ts, corte); fases = raices; }
        const nivelFase = fases.length ? niv(fases[0]) : minN;
        const num = (t) => { const e = String(t.EDT || ""); return edtRaiz && e.startsWith(edtRaiz) ? "1" + e.slice(edtRaiz.length) : e; };
        const verTxt = v ? `Cronograma V${v.Numero}${v.Origen ? ` (${v.Origen})` : ""}${v.Fecha_Corte ? ` · corte de la versión ${fecha(v.Fecha_Corte)}` : ""}` : "Cronograma importado";
        const caja = (s, x, y, w, h, t, a, fase) => {
          const [bg, bd] = COL_EDT[claseAvance(a.real, a.plan)];
          const d = a.real - a.plan;
          s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: bg }, line: { color: bd, width: claseAvance(a.real, a.plan) === "ed-bad" ? 1.75 : 1 }, rectRadius: 0.05 });
          s.addText([{ text: `${num(t) ? num(t) + " " : ""}${t.Hito === "Sí" ? "◆ " : ""}${corta(t.Nombre, fase ? 60 : 70)}`, options: { bold: true, color: PPT.texto, fontSize: fase ? 10 : 8.5, breakLine: true } },
            { text: `Real ${a.real}% / ${a.plan}%${d ? `  ${d > 0 ? "+" : ""}${d} pp` : ""}`, options: { color: d < -5 ? "B3261E" : PPT.sec, fontSize: fase ? 9 : 8, bold: d < -15 } }],
            { x: x + 0.04, y, w: w - 0.08, h, fontFace: F, valign: "middle", margin: [1, 3, 1, 3], fit: "shrink" });
        };
        // Lámina EDT (hasta 7 fases por lámina)
        const grupos = paginar(fases, 7);
        grupos.forEach((gf, k) => {
          const s = lamina(`EDT · ${corta(p.Nombre, 50)}${grupos.length > 1 ? ` (${k + 1}/${grupos.length})` : ""}`, `Avance real vs. planeado por fase a ${fecha(corte)} · ${verTxt}`);
          caja(s, 6.67 - 1.75, 1.5, 3.5, 0.62, { EDT: edtRaiz, Nombre: raizTxt }, raizAv, true);
          const n = Math.max(1, gf.length), gap = 0.14, cw = (12.33 - gap * (n - 1)) / n;
          const cx = (i) => 0.5 + i * (cw + gap) + cw / 2;
          s.addShape(pptx.ShapeType.line, { x: 6.67, y: 2.12, w: 0, h: 0.2, line: { color: "9AA8BC", width: 1 } });
          if (n > 1) s.addShape(pptx.ShapeType.line, { x: cx(0), y: 2.32, w: cx(n - 1) - cx(0), h: 0, line: { color: "9AA8BC", width: 1 } });
          gf.forEach((i, c) => {
            const x = 0.5 + c * (cw + gap);
            s.addShape(pptx.ShapeType.line, { x: cx(c), y: 2.32, w: 0, h: 0.18, line: { color: "9AA8BC", width: 1 } });
            caja(s, x, 2.5, cw, 0.78, ts[i], avance(i), true);
            const subs = hijosDe(i).filter((j) => niv(j) === niv(i) + 1);
            const maxS = 8, vis = subs.length > maxS ? subs.slice(0, maxS - 1) : subs;
            const hS = Math.min(0.56, (6.7 - 3.42 - 0.07 * Math.max(0, vis.length - 1)) / Math.max(1, vis.length + (subs.length > maxS ? 1 : 0)));
            vis.forEach((j, r) => caja(s, x + 0.12, 3.42 + r * (hS + 0.07), cw - 0.12, hS, ts[j], avance(j), false));
            if (subs.length > maxS) s.addText(`y ${subs.length - vis.length} más…`, { x: x + 0.12, y: 3.42 + vis.length * (hS + 0.07), w: cw - 0.12, h: 0.25, fontFace: F, fontSize: 8, color: PPT.sec, italic: true, margin: 0 });
            if (vis.length) s.addShape(pptx.ShapeType.line, { x: x + 0.06, y: 3.28, w: 0, h: (vis.length - 0.5) * (hS + 0.07) + 0.14, line: { color: "C3CCD8", width: 0.75 } });
          });
          leyenda(s, [["D6E9D2", "Completado"], ["E6F4E3", "En plazo (≤ 5 pp)"], ["FFF3CC", "Atraso 6–15 pp"], ["FDE3E1", "Atraso ≥ 16 pp"], ["F1F3F5", "No iniciado"]]);
        });
        // Lámina Gantt: fases, subfases e hitos (hasta 2 láminas)
        const cRes = calcCronograma(ts, corte);
        let sel = ts.map((t, i) => i).filter((i) => niv(i) <= nivelFase + 1 || ts[i].Hito === "Sí");
        if (sel.length > 34) sel = ts.map((t, i) => i).filter((i) => niv(i) <= nivelFase || ts[i].Hito === "Sí");
        const filas = sel.slice(0, 34).map((i) => {
          const t = ts[i], a = hijosDe(i).length ? avance(i) : { real: Number(t.Pct_Completado) || 0, plan: 0 };
          const est = estadoTarea(t, corte), hito = t.Hito === "Sí";
          return { txt: `${num(t) ? num(t) + "  " : ""}${corta(t.Nombre, 46 - (niv(i) - minN) * 3)}`, nivel: Math.max(0, niv(i) - minN - (fases === raices ? 0 : 1)), negrita: hijosDe(i).length > 0 && niv(i) <= nivelFase,
            a: t.Inicio, b: t.Fin, pct: a.real, resumen: hijosDe(i).length > 0, soloHito: hito, colorReal: est === "atr" ? PPT.Rojo : PPT.azul, colorPct: est === "atr" ? PPT.Rojo : est === "ok" ? PPT.Verde : PPT.texto,
            base: cRes.conBase && t.Inicio_Base && t.Fin_Base && !hito ? [t.Inicio_Base, t.Fin_Base] : null,
            hitos: hito ? [{ f: t.Fin || t.Inicio, color: (Number(t.Pct_Completado) || 0) >= 100 ? PPT.Verde : (t.Fin || t.Inicio) < corte ? PPT.Rojo : "8A94A3" }] : [] };
        });
        const ini = cRes.inicio, fin = cRes.fin > corte ? cRes.fin : R.sumarDias(corte, 15);
        if (ini && fin) paginar(filas, 17).forEach((pag, k, arr) => {
          const s = lamina(`Cronograma · ${corta(p.Nombre, 46)}${arr.length > 1 ? ` (${k + 1}/${arr.length})` : ""}`, `${verTxt} · avance real ${cRes.real}% vs. ${cRes.plan}% planeado · ${cRes.atrasadas.length} tarea(s) atrasada(s)`);
          s.addText("Fase / tarea", { x: 0.5, y: 1.5, w: 3, h: 0.25, fontFace: F, fontSize: 9, bold: true, color: PPT.azul, margin: 0 });
          s.addText("Real", { x: 4.1, y: 1.5, w: 0.7, h: 0.25, fontFace: F, fontSize: 9, bold: true, color: PPT.azul, align: "right", margin: 0 });
          ganttPPT(s, pag, ini, fin, { x0: 4.95, y0: 1.85, colPct: true });
          leyenda(s, [[PPT.claro, "Planeado"], [PPT.azul, "Avance real"], [PPT.Rojo, "Tarea atrasada"], ...(cRes.conBase ? [["B5BFCC", "Línea base"]] : []), [PPT.Verde, "Hito cumplido", "rombo"], [PPT.Rojo, "Hito atrasado", "rombo"], [PPT.dorado, "Hoy"]]);
        });
        return true;
      };

      // 1. Portada
      const port = pptx.addSlide({ masterName: "PORTADA" });
      port.addText([
        { text: "Estado del portafolio de proyectos", options: { fontSize: 24, bold: true, breakLine: true } },
        { text: alcanceTxt ? corta(alcanceTxt.replace(/^PM: /, nomPM && !pids.size ? "Informe del PM: " : "PM: "), 90) : "Informe general de la Oficina de Proyectos", options: { fontSize: 13, breakLine: true } },
        { text: `Corte: ${fechaLarga(corte)}`, options: { fontSize: 13, breakLine: true } },
        { text: `Proyectos activos y en pausa: ${ps.length}`, options: { fontSize: 13 } },
      ], { x: 0.44, y: 4.74, w: 5.26, h: 1.82, fontFace: F, color: PPT.blanco, valign: "middle", margin: 0.05, fit: "shrink" });

      // 2. Resumen ejecutivo
      const resumen = (lista, m, titulo, sub) => {
        const s = lamina(titulo, sub);
        const desv = m.real - m.plan;
        tile(s, 0.5, 1.55, 2.95, 1.45, "Proyectos activos", m.act.length, m.pausa ? `${m.pausa} en pausa` : "Ninguno en pausa");
        tile(s, 3.6, 1.55, 2.95, 1.45, "Avance real vs. planeado", `${m.real}% / ${m.plan}%`, desv < -5 ? `${-desv} pp por debajo del plan` : desv > 5 ? `${desv} pp por encima del plan` : "En línea con el plan", desv < -15 ? PPT.Rojo : desv < -5 ? PPT.Amarillo : PPT.Verde);
        tile(s, 6.7, 1.55, 2.95, 1.45, "Seguimientos al día", `${m.alDia}%`, m.vencSeg ? `${m.vencSeg} proyecto(s) con seguimiento vencido` : "Todos al día", m.vencSeg ? PPT.Rojo : PPT.Verde);
        tile(s, 9.8, 1.55, 3.03, 1.45, "Compromisos vencidos", m.compV, `${m.compPV} por vencer · ${m.compAb} abiertos`, m.compV ? PPT.Rojo : PPT.Verde);
        if (!m.act.length) return vacioTxt(s, "No hay proyectos activos en este alcance.", 3.8);
        s.addText("Salud de los proyectos activos (semáforo)", { x: 0.5, y: 3.2, w: 4.2, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        s.addChart(pptx.ChartType.doughnut, [{ name: "Semáforo", labels: ["Verde", "Amarillo", "Rojo"], values: m.sem }], { x: 0.5, y: 3.55, w: 4.2, h: 3.4, holeSize: 55,
          chartColors: [PPT.Verde, PPT.Amarillo, PPT.Rojo], showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: F, showPercent: false, showValue: true, dataLabelColor: PPT.blanco, dataLabelFontSize: 12, dataLabelFontBold: true });
        const top = m.act.slice().sort((a, b) => ((Number(b.Avance_Planeado) || 0) - (Number(b.Avance_Real) || 0)) - ((Number(a.Avance_Planeado) || 0) - (Number(a.Avance_Real) || 0))).slice(0, 10);
        s.addText(`Avance real vs. planeado${m.act.length > 10 ? " (10 con mayor desviación)" : ""}`, { x: 5.0, y: 3.2, w: 7.8, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        s.addChart(pptx.ChartType.bar, [{ name: "Planeado", labels: top.map((p) => corta(p.Nombre, 28)), values: top.map((p) => Number(p.Avance_Planeado) || 0) }, { name: "Real", labels: top.map((p) => corta(p.Nombre, 28)), values: top.map((p) => Number(p.Avance_Real) || 0) }],
          { x: 5.0, y: 3.55, w: 7.83, h: 3.45, barDir: "bar", barGrouping: "clustered", chartColors: [PPT.claro, PPT.azul], valAxisMaxVal: 100, valAxisMinVal: 0, valAxisLabelFontSize: 9, catAxisLabelFontSize: 9, catAxisLabelFontFace: F, catAxisOrientation: "maxMin",
            showLegend: true, legendPos: "t", legendFontSize: 10, showValue: true, dataLabelFontSize: 8, dataLabelPosition: "outEnd", valGridLine: { color: "E5E5E5", size: 0.5 } });
        if (m.pres) s.addText(`Presupuesto de proyectos activos: ${cop(m.pres)} · ejecutado ${cop(m.ejec)} (${Math.round((m.ejec / m.pres) * 100)}%)`, { x: 0.5, y: 6.85, w: 4.5, h: 0.25, fontFace: F, fontSize: 8, color: PPT.sec, margin: 0 });
        return s;
      };
      resumen(ps, M, "Resumen ejecutivo", libre ? "Todos los proyectos activos o en pausa del portafolio" : nomPM && !pids.size ? `Proyectos de ${nomPM} activos o en pausa` : `Alcance: ${corta(alcanceTxt, 110)}`);

      // 3. Resumen por PM (general)
      const gruposPM = [...new Set(ps.map((p) => lc(p.PM)))].map((c) => ({ c, n: nombreUsuario(c), lista: ps.filter((p) => lc(p.PM) === c) })).sort((a, b) => String(a.n).localeCompare(String(b.n), "es"));
      if (!pm && incluir.has("pm") && gruposPM.length) {
        const s = lamina("Resumen por PM", "Proyectos, avance, salud, seguimiento y carga de cada PM");
        const filas = [[cab("PM"), cab("Activos", { align: "center" }), cab("En pausa", { align: "center" }), cab("Avance real / plan", { align: "center" }), cab("Amarillo / Rojo", { align: "center" }), cab("Seguimientos al día", { align: "center" }), cab("Compromisos vencidos", { align: "center" }), cab("Carga del PM", { align: "center" })]];
        gruposPM.forEach((g, i) => {
          const m = metricas(g.lista), c = cap.find((x) => lc(x.correo) === g.c);
          filas.push([celda(g.n, i, { bold: true }), celda(m.act.length, i, { align: "center" }), celda(m.pausa, i, { align: "center" }), celda(`${m.real}% / ${m.plan}%`, i, { align: "center" }),
            celda(`${m.sem[1]} / ${m.sem[2]}`, i, { align: "center", color: m.sem[2] ? PPT.Rojo : PPT.texto, bold: !!m.sem[2] }), celda(`${m.alDia}%`, i, { align: "center", color: m.alDia < 80 ? PPT.Rojo : PPT.texto }),
            celda(m.compV, i, { align: "center", color: m.compV ? PPT.Rojo : PPT.texto, bold: !!m.compV }),
            celda(c ? `${c.total}% de ${c.cap}%${c.tOtras ? ` (${c.tOtras}% otras)` : ""}` : "—", i, { align: "center", color: c && c.total > c.cap ? PPT.Rojo : PPT.texto, bold: !!(c && c.total > c.cap) })]);
        });
        tabla(s, filas.slice(0, 13), [2.9, 1.0, 1.0, 1.6, 1.4, 1.5, 1.5, 1.43]);
        const n = Math.min(gruposPM.length, 12), yC = 1.55 + 0.36 * (n + 1) + 0.25;
        if (yC < 5.2) {
          const cs = gruposPM.map((g) => cap.find((x) => lc(x.correo) === g.c) || { tProy: 0, tOtras: 0 });
          s.addChart(pptx.ChartType.bar, [{ name: "Proyectos", labels: gruposPM.map((g) => g.n), values: cs.map((c) => c.tProy) }, { name: "Otras actividades", labels: gruposPM.map((g) => g.n), values: cs.map((c) => c.tOtras) }],
            { x: 0.5, y: yC, w: 12.33, h: 7.0 - yC, barDir: "col", barGrouping: "stacked", chartColors: [PPT.azul, PPT.dorado], showLegend: true, legendPos: "r", legendFontSize: 10, catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
              showValue: true, dataLabelFontSize: 9, dataLabelColor: PPT.blanco, dataLabelFormatCode: "0;;;", valAxisTitle: "% de dedicación", showValAxisTitle: true, valAxisTitleFontSize: 9, valGridLine: { color: "E5E5E5", size: 0.5 } });
        }
      }

      // 4. Estado de los proyectos (tabla paginada)
      const tablaProyectos = (lista, titulo, conPM) => {
        const pags = paginar(lista, 11);
        pags.forEach((pag, k) => {
          const s = lamina(pags.length > 1 ? `${titulo} (${k + 1}/${pags.length})` : titulo, `${lista.length} proyecto(s) activos o en pausa · semáforo y avance del último seguimiento`);
          if (!pag.length) return vacioTxt(s, "No hay proyectos activos o en pausa.");
          const filas = [[cab("Proyecto"), ...(conPM ? [cab("PM")] : []), cab("Fase"), cab("Estado", { align: "center" }), cab("Semáforo", { align: "center" }), cab("Avance real / plan", { align: "center" }), cab("Fin planeado", { align: "center" }), cab("Último seguimiento", { align: "center" }), cab("Seguimiento", { align: "center" })]];
          pag.forEach((p, i) => {
            const e = R.estadoSeguimiento(p, segs, corte), d = (Number(p.Avance_Real) || 0) - (Number(p.Avance_Planeado) || 0);
            filas.push([celda(corta(p.Nombre, 42), i, { bold: true }), ...(conPM ? [celda(corta(nombreUsuario(p.PM), 22), i)] : []), celda(p.Fase || "—", i), celda(p.Estado, i, { align: "center" }), semCelda(p.Semaforo, i),
              celda(`${Number(p.Avance_Real) || 0}% / ${Number(p.Avance_Planeado) || 0}%`, i, { align: "center", color: d < -15 ? PPT.Rojo : PPT.texto, bold: d < -15 }), celda(fecha(p.Fecha_Fin_Plan), i, { align: "center" }),
              celda(e.ultimaFecha ? fecha(e.ultimaFecha) : "Sin registro", i, { align: "center" }),
              celda(e.estado, i, { align: "center", bold: e.estado !== "Al día", color: e.estado === "Vencido" ? PPT.Rojo : e.estado === "Por vencer" ? "9A6B00" : PPT.texto })]);
          });
          tabla(s, filas, conPM ? [2.95, 1.75, 1.2, 0.95, 1.0, 1.25, 1.03, 1.1, 1.1] : [3.95, 1.45, 1.0, 1.05, 1.4, 1.1, 1.18, 1.2]);
        });
      };
      if (incluir.has("proyectos")) tablaProyectos(ps, "Estado de los proyectos", !pm);
      if (incluir.has("gantt")) ganttPortafolio();
      if (incluir.has("valor")) valorPortafolio();

      // 5. Secciones por PM (general)
      if (!pm && incluir.has("pm") && gruposPM.length > 1) gruposPM.forEach((g) => {
        const sec = pptx.addSlide({ masterName: "SECCION" });
        const m = metricas(g.lista), c = cap.find((x) => lc(x.correo) === g.c);
        sec.addText(g.n, { x: 0.44, y: 2.84, w: 7.6, h: 0.9, fontFace: F, fontSize: 32, bold: true, color: PPT.blanco, margin: 0, valign: "bottom", fit: "shrink" });
        sec.addText([`${m.act.length} proyecto(s) activo(s)${m.pausa ? `, ${m.pausa} en pausa` : ""}`, `Avance ${m.real}% frente a ${m.plan}% planeado`, `Seguimientos al día: ${m.alDia}%`, ...(c ? [`Carga del PM: ${c.total}% de ${c.cap}%`] : [])].map((t, i, a) => ({ text: t, options: { breakLine: i < a.length - 1 } })),
          { x: 0.44, y: 3.85, w: 7.6, h: 1.6, fontFace: F, fontSize: 16, color: PPT.blanco, margin: 0, valign: "top" });
        tablaProyectos(g.lista, `Proyectos de ${g.n}`, false);
      });

      // 6. Capacidad
      if (incluir.has("capacidad")) {
        const s = lamina("Capacidad del equipo", `Dedicación a ${fecha(corte)}: proyectos activos o en pausa + otras actividades (operación, comités, demanda…). Las ausencias reducen la capacidad.`);
        const l = cap.slice().sort((a, b) => b.total / Math.max(1, b.cap) - a.total / Math.max(1, a.cap)).slice(0, 14);
        const capT = cap.reduce((t, x) => t + x.capBase, 0), otrasT = cap.reduce((t, x) => t + x.tOtras, 0);
        const sobre = cap.filter((x) => x.total > x.cap), limite = cap.filter((x) => x.total <= x.cap && x.total >= x.cap * 0.85);
        tile(s, 9.3, 1.5, 3.53, 1.3, "Sobreasignados", sobre.length, `${limite.length} al límite (85–100%)`, sobre.length ? PPT.Rojo : PPT.Verde);
        tile(s, 9.3, 2.95, 3.53, 1.3, "Capacidad en otras actividades", `${capT ? Math.round((otrasT / capT) * 100) : 0}%`, `${cap.length} persona(s) en el alcance`, PPT.dorado);
        if (!l.length) vacioTxt(s, "Aún no hay dedicaciones registradas.");
        else {
          s.addChart(pptx.ChartType.bar, [{ name: "Proyectos", labels: l.map((x) => corta(x.nombre, 26)), values: l.map((x) => x.tProy) }, { name: "Otras actividades", labels: l.map((x) => corta(x.nombre, 26)), values: l.map((x) => x.tOtras) }],
            { x: 0.5, y: 1.5, w: 8.6, h: 5.5, barDir: "bar", barGrouping: "stacked", chartColors: [PPT.azul, PPT.dorado], catAxisOrientation: "maxMin", catAxisLabelFontSize: 10, catAxisLabelFontFace: F, valAxisLabelFontSize: 9,
              showLegend: true, legendPos: "t", legendFontSize: 10, showValue: true, dataLabelFontSize: 8, dataLabelColor: PPT.blanco, dataLabelFormatCode: "0;;;", valGridLine: { color: "E5E5E5", size: 0.5 } });
          const filas = [[cab("Sobre su capacidad"), cab("Carga", { align: "center" })]];
          sobre.slice(0, 6).forEach((x, i) => filas.push([celda(corta(x.nombre, 26), i), celda(`${x.total}% de ${x.cap}%`, i, { align: "center", color: PPT.Rojo, bold: true })]));
          if (sobre.length) tabla(s, filas, [1.9, 1.63], 4.45, { x: 9.3, rowH: 0.32 });
          else s.addText("Nadie está sobre su capacidad.", { x: 9.3, y: 4.4, w: 3.5, h: 0.4, fontFace: F, fontSize: 12, color: PPT.Verde, bold: true });
        }
      }

      // 7. Seguimientos
      if (incluir.has("seguimientos")) {
        const s = lamina("Seguimientos", "Seguimientos de avance y sesiones registradas · estado frente a la frecuencia definida de cada proyecto");
        const sg = segs.filter((x) => M.ids.has(x.ID_Proyecto) && x.Fecha_Corte && x.Fecha_Corte <= corte);
        const hace30 = R.sumarDias(corte, -30), ult30 = sg.filter((x) => x.Fecha_Corte > hace30);
        tile(s, 0.5, 1.55, 2.95, 1.3, "Seguimientos (últimos 30 días)", ult30.filter((x) => !R.esSesion(x)).length, "Reportes de avance");
        tile(s, 3.6, 1.55, 2.95, 1.3, "Sesiones (últimos 30 días)", ult30.filter(R.esSesion).length, "Comités, mesas técnicas, reuniones", PPT.dorado);
        tile(s, 6.7, 1.55, 2.95, 1.3, "Proyectos al día", `${M.alDia}%`, `${M.vencSeg} con seguimiento vencido`, M.vencSeg ? PPT.Rojo : PPT.Verde);
        const sems = []; for (let i = 7; i >= 0; i--) sems.push(R.semanaISO(R.sumarDias(corte, -7 * i)));
        s.addText("Registros por semana (últimas 8)", { x: 0.5, y: 3.05, w: 6, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        s.addChart(pptx.ChartType.bar, [{ name: "Seguimientos", labels: sems.map((w) => w.slice(5)), values: sems.map((w) => sg.filter((x) => !R.esSesion(x) && R.semanaISO(x.Fecha_Corte) === w).length) },
          { name: "Sesiones", labels: sems.map((w) => w.slice(5)), values: sems.map((w) => sg.filter((x) => R.esSesion(x) && R.semanaISO(x.Fecha_Corte) === w).length) }],
          { x: 0.5, y: 3.4, w: 6.0, h: 3.6, barDir: "col", barGrouping: "stacked", chartColors: [PPT.azul, PPT.dorado], showLegend: true, legendPos: "t", legendFontSize: 10, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, valGridLine: { color: "E5E5E5", size: 0.5 }, showValue: true, dataLabelFontSize: 8, dataLabelColor: PPT.blanco, dataLabelFormatCode: "0;;;" });
        const pend = M.act.map((p) => ({ p, e: R.estadoSeguimiento(p, segs, corte) })).filter((x) => x.e.estado !== "Al día").sort((a, b) => a.e.faltan - b.e.faltan);
        s.addText("Seguimientos vencidos o por vencer", { x: 6.8, y: 3.05, w: 6, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        if (!pend.length) s.addText("Todos los proyectos activos tienen su seguimiento al día.", { x: 6.8, y: 3.5, w: 6, h: 0.5, fontFace: F, fontSize: 12, color: PPT.Verde, bold: true });
        else {
          const filas = [[cab("Proyecto"), cab("Último", { align: "center" }), cab("Debía", { align: "center" }), cab("Estado", { align: "center" })]];
          pend.slice(0, 8).forEach(({ p, e }, i) => filas.push([celda(corta(p.Nombre, 30), i), celda(e.ultimaFecha ? fecha(e.ultimaFecha) : "—", i, { align: "center" }), celda(fecha(e.proximo), i, { align: "center" }),
            celda(e.estado === "Vencido" ? `Vencido ${-e.faltan} d` : e.estado, i, { align: "center", bold: true, color: e.estado === "Vencido" ? PPT.Rojo : "9A6B00" })]));
          tabla(s, filas, [2.65, 1.1, 1.1, 1.18], 3.45, { x: 6.8, rowH: 0.33 });
          if (pend.length > 8) s.addText(`y ${pend.length - 8} más`, { x: 6.8, y: 6.75, w: 6, h: 0.3, fontFace: F, fontSize: 9, color: PPT.sec });
        }
      }

      // 8. Compromisos y riesgos
      if (incluir.has("compromisos")) {
        const s = lamina("Compromisos y riesgos", "Compromisos vencidos o por vencer y riesgos altos abiertos");
        const comp = S.datos.Compromisos.filter((c) => M.ids.has(c.ID_Proyecto)).map((c) => ({ c, e: R.estadoCompromiso(c, corte) })).filter((x) => x.e === "Vencido" || x.e === "Por vencer")
          .sort((a, b) => (a.e === b.e ? String(a.c.Fecha_Compromiso).localeCompare(String(b.c.Fecha_Compromiso)) : a.e === "Vencido" ? -1 : 1));
        s.addText(`Compromisos (${comp.length})`, { x: 0.5, y: 1.5, w: 6, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        if (!comp.length) s.addText("Sin compromisos vencidos ni por vencer.", { x: 0.5, y: 1.95, w: 6, h: 0.4, fontFace: F, fontSize: 12, color: PPT.Verde, bold: true });
        else {
          const filas = [[cab("Compromiso"), cab("Proyecto"), cab("Responsable"), cab("Fecha", { align: "center" })]];
          comp.slice(0, 10).forEach(({ c, e }, i) => filas.push([celda(corta(c.Compromiso, 48), i), celda(corta((proyecto(c.ID_Proyecto) || {}).Nombre, 20), i), celda(corta(responsablesDe(c, false).map((x) => x.Nombre).join(", ") || "—", 18), i),
            celda(fecha(c.Fecha_Compromiso), i, { align: "center", bold: true, color: e === "Vencido" ? PPT.Rojo : "9A6B00" })]));
          tabla(s, filas, [2.4, 1.5, 1.2, 1.2], 1.9, { rowH: 0.42, fontSize: 9 });
          if (comp.length > 10) s.addText(`y ${comp.length - 10} más`, { x: 0.5, y: 6.75, w: 6, h: 0.3, fontFace: F, fontSize: 9, color: PPT.sec });
        }
        const rg = S.datos.Riesgos.filter((r) => M.ids.has(r.ID_Proyecto) && riesgoActivo(r) && R.nivelRiesgo(califActual(r)) === "Alto").sort((a, b) => califActual(b) - califActual(a));
        s.addText(`Riesgos altos abiertos (${rg.length})`, { x: 6.95, y: 1.5, w: 6, h: 0.35, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        if (!rg.length) s.addText("Sin riesgos altos abiertos.", { x: 6.95, y: 1.95, w: 6, h: 0.4, fontFace: F, fontSize: 12, color: PPT.Verde, bold: true });
        else {
          const filas = [[cab("Riesgo"), cab("Proyecto"), cab("Calif.", { align: "center" }), cab("Mitigación")]];
          rg.slice(0, 8).forEach((r, i) => filas.push([celda(corta(r.Descripcion, 60), i), celda(corta((proyecto(r.ID_Proyecto) || {}).Nombre, 18), i), celda(califActual(r), i, { align: "center", bold: true, color: PPT.Rojo }), celda(corta(r.Plan_Mitigacion || "Sin plan", 50), i, { color: r.Plan_Mitigacion ? PPT.texto : PPT.Rojo })]));
          tabla(s, filas, [2.3, 1.3, 0.6, 1.68], 1.9, { x: 6.95, rowH: 0.5, fontSize: 9 });
          if (rg.length > 8) s.addText(`y ${rg.length - 8} más`, { x: 6.95, y: 6.75, w: 6, h: 0.3, fontFace: F, fontSize: 9, color: PPT.sec });
        }
      }

      // 9. Una lámina por proyecto
      const fichaProyecto = (p) => {
        const s = lamina(corta(p.Nombre, 60), `${p.ID_Proyecto} · PM: ${nombreUsuario(p.PM)} · ${p.Metodologia || ""}${p.Fase ? ` · Fase: ${p.Fase}` : ""}`);
        const ult = R.avancesDe(p.ID_Proyecto, segs).filter((x) => x.Fecha_Corte <= corte).pop();
        const e = R.estadoSeguimiento(p, segs, corte);
        const real = Number(p.Avance_Real) || 0, plan = Number(p.Avance_Planeado) || 0;
        s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 1.55, w: 4.1, h: 5.4, fill: { color: PPT.suave }, line: { color: PPT.suave }, rectRadius: 0.08 });
        const dato = (y, t, v, col) => { s.addText(t, { x: 0.75, y, w: 3.6, h: 0.28, fontFace: F, fontSize: 10, color: PPT.sec, margin: 0 }); s.addText(String(v), { x: 0.75, y: y + 0.27, w: 3.6, h: 0.36, fontFace: F, fontSize: 14, bold: true, color: col || PPT.texto, margin: 0 }); };
        dato(1.7, "Estado · semáforo", `${p.Estado} · ${p.Semaforo || "—"}`, PPT[p.Semaforo] || PPT.texto);
        s.addText("Avance real vs. planeado", { x: 0.75, y: 2.4, w: 3.6, h: 0.28, fontFace: F, fontSize: 10, color: PPT.sec, margin: 0 });
        s.addShape(pptx.ShapeType.rect, { x: 0.75, y: 2.75, w: 3.6, h: 0.22, fill: { color: "D9E2EF" }, line: { color: "D9E2EF" } });
        if (real) s.addShape(pptx.ShapeType.rect, { x: 0.75, y: 2.75, w: 3.6 * Math.min(1, real / 100), h: 0.22, fill: { color: PPT.azul }, line: { color: PPT.azul } });
        s.addShape(pptx.ShapeType.rect, { x: 0.75 + 3.6 * Math.min(1, plan / 100) - 0.015, y: 2.68, w: 0.03, h: 0.36, fill: { color: PPT.dorado }, line: { color: PPT.dorado } });
        s.addText(`${real}% real · ${plan}% plan`, { x: 0.75, y: 3.0, w: 3.6, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: real - plan < -15 ? PPT.Rojo : PPT.texto, margin: 0 });
        dato(3.45, "Fechas", `${fecha(p.Fecha_Inicio)} → ${fecha(p.Fecha_Fin_Plan)}`);
        dato(4.15, "Último seguimiento", e.ultimaFecha ? `${fecha(e.ultimaFecha)} (${e.estado})` : "Sin registro", e.estado === "Vencido" ? PPT.Rojo : PPT.texto);
        const nComp = S.datos.Compromisos.filter((c) => c.ID_Proyecto === p.ID_Proyecto && !R.cerrado(c)).length, nVenc = S.datos.Compromisos.filter((c) => c.ID_Proyecto === p.ID_Proyecto && R.estadoCompromiso(c, corte) === "Vencido").length;
        dato(4.85, "Compromisos abiertos", `${nComp}${nVenc ? ` (${nVenc} vencidos)` : ""}`, nVenc ? PPT.Rojo : PPT.texto);
        const nR = S.datos.Riesgos.filter((r) => r.ID_Proyecto === p.ID_Proyecto && riesgoActivo(r) && R.nivelRiesgo(califActual(r)) === "Alto").length;
        dato(5.55, "Riesgos altos abiertos", nR, nR ? PPT.Rojo : PPT.texto);
        const bloque = (x, y, w, h, titulo, txt, color) => {
          s.addText(titulo, { x, y, w, h: 0.32, fontFace: F, fontSize: 13, bold: true, color, margin: 0 });
          s.addText(corta(txt || "Sin registrar en el último seguimiento.", 520), { x, y: y + 0.34, w, h: h - 0.34, fontFace: F, fontSize: 11, color: txt ? PPT.texto : PPT.sec, italic: !txt, valign: "top", margin: 2, fit: "shrink" });
        };
        bloque(4.9, 1.55, 3.85, 2.55, "Logros", ult && ult.Logros, PPT.Verde);
        bloque(8.95, 1.55, 3.88, 2.55, "Próximos pasos", ult && ult.Proximos_Pasos, PPT.azul);
        bloque(4.9, 4.2, 3.85, 1.1, "Bloqueos", ult && ult.Bloqueos, PPT.Rojo);
        const hs = S.datos.Hitos.filter((h) => h.ID_Proyecto === p.ID_Proyecto && h.Estado !== "Cumplido").sort((a, b) => String(a.Fecha_Plan).localeCompare(String(b.Fecha_Plan))).slice(0, 4);
        s.addText("Próximos hitos", { x: 8.95, y: 4.2, w: 3.88, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: PPT.azul, margin: 0 });
        if (hs.length) tabla(s, [[cab("Hito"), cab("Fecha", { align: "center" })], ...hs.map((h, i) => [celda(corta(h.Hito, 34), i), celda(fecha(h.Fecha_Plan), i, { align: "center", bold: h.Fecha_Plan < corte, color: h.Fecha_Plan < corte ? PPT.Rojo : PPT.texto })])], [2.78, 1.1], 4.58, { x: 8.95, rowH: 0.3, fontSize: 9 });
        else s.addText("Sin hitos pendientes.", { x: 8.95, y: 4.58, w: 3.88, h: 0.3, fontFace: F, fontSize: 11, color: PPT.sec, italic: true });
        if (ult) s.addText(`Fuente: seguimiento del ${fecha(ult.Fecha_Corte)}`, { x: 4.9, y: 6.7, w: 6, h: 0.28, fontFace: F, fontSize: 9, color: PPT.sec, margin: 0 });
      };
      // Por proyecto: ficha, luego su EDT y su cronograma (si tiene uno cargado).
      ps.forEach((p) => { if (incluir.has("fichas")) fichaProyecto(p); if (incluir.has("valor") && incluir.has("fichas")) valorDeProyecto(p); if (incluir.has("edt")) laminasCronograma(p); });

      const fin = pptx.addSlide({ masterName: "CIERRE" });
      fin.addText("Gracias", { x: 7.646, y: 2.838, w: 4.542, h: 1.823, fontFace: F, fontSize: 36, bold: true, color: PPT.blanco, valign: "middle", margin: 0 });
      const bytes = await pptx.write({ outputType: "uint8array" });
      const sufijo = libre ? "" : ps.length === 1 ? ps[0].Nombre : nomPM || (pms.size ? `${pms.size}_PM` : "") + (pids.size ? `${pms.size ? "_" : ""}${pids.size}_proyectos` : "");
      const nombre = `Estado_Portafolio_${sufijo ? normTxt(sufijo).replace(/[^a-z0-9]+/g, "_").slice(0, 40) + "_" : ""}${corte}.pptx`;
      descargar(bytes, nombre, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      toast(`Presentación generada: ${nombre}`);
    } catch (e) { toast("No se pudo generar la presentación: " + e.message + (S.modo === "panel" ? " Si el panel bloquea la descarga, abre la app en ventana grande." : ""), true); }
    finally { cargando(false); }
  }

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
    { t: "Nivel de valor FSFB", v: (p) => valorProyecto(p).nivel, ancho: 14 }, { t: "Índice de valor", v: (p) => valorProyecto(p).indice ?? "", tipo: "pct" },
    { t: "Beneficio esperado", v: (p) => p.Beneficio_Esperado, ancho: 40 }, { t: "Valor para la FSFB", v: (p) => p.Valor_FSFB, ancho: 40 },
    { t: "Comentario de estado", v: (p) => p.Comentario_Estado, ancho: 44 },
  ];
  const colsCompromiso = (conProy) => [
    ...(conProy ? [C.proy] : []), { t: "Compromiso", v: (c) => c.Compromiso, ancho: 46 },
    { t: "Creado", v: (c) => fechaCreacion(c), tipo: "fecha" },
    { t: "Sesión", v: (c) => { const s = sesionDe(c); return s ? `Sesión del ${fecha(s.Fecha_Corte)}` : "Sin sesión"; }, ancho: 18 },
    { t: "Responsable", v: (c) => c.Responsable }, { t: "Correo responsable", v: (c) => correoDe(c), ancho: 26 },
    { t: "Fecha límite", v: (c) => c.Fecha_Compromiso, tipo: "fecha" }, { t: "Estado", v: (c) => R.estadoBase(c), ancho: 12, color: true },
    { t: "Alerta", v: (c) => { const e = R.estadoCompromiso(c); return e === "Vencido" || e === "Por vencer" ? e : ""; }, ancho: 12, color: true },
    { t: "Fecha de cierre", v: (c) => c.Fecha_Cierre, tipo: "fecha" },
    { t: "Último comentario", v: (c) => { const l = comentariosDe(c.ID_Compromiso); const u = l[l.length - 1]; return u ? resumenCom(u) : ""; }, ancho: 46 },
    { t: "Fecha último comentario", v: (c) => { const l = comentariosDe(c.ID_Compromiso); const u = l[l.length - 1]; return u ? u.Fecha_Hora : ""; }, tipo: "fechahora", ancho: 16 },
  ];
  const colsSeguimiento = (conProy) => [
    ...(conProy ? [C.proy] : []), { t: "Fecha de corte", v: (s) => s.Fecha_Corte, tipo: "fecha" }, { t: "Tipo", v: (s) => (R.esSesion(s) ? `Sesión${s.Nombre_Sesion ? `: ${s.Nombre_Sesion}` : ""}` : "Seguimiento"), ancho: 18 }, { t: "Semana", v: (s) => s.Semana, ancho: 10 },
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
    sugerencias: () => exportarExcel("Sugerencias de mejora", [{ nombre: "Sugerencias", columnas: [
      { t: "ID", v: (x) => x.ID_Sugerencia, ancho: 11 }, { t: "Fecha", v: (x) => String(x.Fecha_Hora).slice(0, 10), tipo: "fecha" }, { t: "Sugerida por", v: (x) => nombreUsuario(x.Usuario) },
      { t: "Módulo", v: (x) => x.Modulo }, { t: "Dónde", v: (x) => x.Contexto, ancho: 26 }, { t: "Tipo", v: (x) => x.Tipo, ancho: 11 }, { t: "Prioridad", v: (x) => x.Prioridad, ancho: 10 },
      { t: "Descripción", v: (x) => x.Descripcion, ancho: 60 }, { t: "Adjuntos", v: (x) => adjuntosDe(x).length, tipo: "num", ancho: 9 }, { t: "Enlace", v: (x) => x.URL_Referencia, ancho: 26 },
      { t: "Estado", v: (x) => x.Estado || "Nueva", ancho: 12 }, { t: "Respuesta", v: (x) => x.Respuesta, ancho: 40 }, { t: "Versión implementada", v: (x) => x.Version_Implementada, ancho: 12 }],
      filas: sugerenciasFiltradas() }], [S.sugF && S.sugF.estado ? `Estado: ${S.sugF.estado}` : "", S.sugF && S.sugF.modulo ? `Módulo: ${S.sugF.modulo}` : "", S.sugF && S.sugF.mias ? "Solo las mías" : ""].filter(Boolean)),
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
      const fp = S.recP || "";
    const lista = recursos().filter((r) => (!q || normTxt(`${r.Nombre} ${r.Correo} ${r.Cargo} ${r.Area} ${empresaDe(r)}`).includes(normTxt(q))) && (!fp || proyectosDeRecurso(r).some((x) => x.p.ID_Proyecto === fp)));
      const asig = (r) => cap.find((x) => (r.Correo && lc(x.correo) === lc(r.Correo)) || normTxt(x.nombre) === normTxt(r.Nombre));
      return exportarExcel("Recursos", [{ nombre: "Recursos", columnas: [{ t: "Nombre", v: (r) => r.Nombre, ancho: 30 }, { t: "Cargo", v: (r) => r.Cargo }, { t: "Área", v: (r) => r.Area },
        { t: "Empresa", v: (r) => empresaDe(r) }, { t: "Proyectos", v: (r) => proyectosDeRecurso(r).map(({ p, roles }) => `${p.Nombre} (${roles.join(", ")})`).join("; "), ancho: 40 }, { t: "Correo", v: (r) => r.Correo, ancho: 28 }, { t: "Teléfono", v: (r) => r.Telefono, ancho: 14 },
        { t: "Capacidad", v: (r) => num0(r.Capacidad) ?? 100, tipo: "pct" }, { t: "Asignado", v: (r) => (asig(r) || {}).total || 0, tipo: "pct" }, { t: "Activo", v: (r) => r.Activo || "Sí", ancho: 8 }], filas: lista }],
        q ? [`Búsqueda: «${q}»`] : []);
    },
    valor: () => {
      const ps = filtrados().filter((p) => p.Estado !== "Cancelado").map((p) => ({ p, V: valorProyecto(p) }));
      const bs = ps.flatMap(({ p, V }) => V.bs.map((x) => ({ p, ...x })));
      return exportarExcel("Valor del portafolio", [
        { nombre: "Proyectos", columnas: [{ t: "Proyecto", v: (x) => x.p.Nombre, ancho: 34 }, { t: "PM", v: (x) => nombreUsuario(x.p.PM) }, { t: "Estado", v: (x) => x.p.Estado, ancho: 12 },
          { t: "Nivel de valor", v: (x) => x.V.nivel, ancho: 12 }, { t: "Índice de valor", v: (x) => x.V.indice ?? "", tipo: "pct" }, { t: "Beneficio esperado", v: (x) => x.p.Beneficio_Esperado, ancho: 44 },
          { t: "Valor para la FSFB", v: (x) => x.p.Valor_FSFB, ancho: 44 }, ...dimensiones().map((d) => ({ t: d, v: (x) => x.V.imp[d] || "", ancho: 14 })),
          { t: "Beneficios", v: (x) => x.V.bs.length, tipo: "num" }, { t: "En riesgo", v: (x) => x.V.riesgo, tipo: "num" }, { t: "Mediciones pendientes", v: (x) => x.V.pendientes, tipo: "num" }], filas: ps },
        { nombre: "Beneficios", columnas: [{ t: "Proyecto", v: (x) => x.p.Nombre, ancho: 30 }, { t: "Importancia", v: (x) => x.b.Importancia, ancho: 11 }, { t: "Dimensión", v: (x) => x.b.Dimension, ancho: 22 },
          { t: "Beneficio esperado", v: (x) => x.b.Beneficio_Esperado, ancho: 40 }, { t: "Valor para la FSFB", v: (x) => x.b.Valor_FSFB, ancho: 40 }, { t: "Indicador", v: (x) => x.b.Indicador, ancho: 30 },
          { t: "Unidad", v: (x) => x.b.Unidad, ancho: 10 }, { t: "Línea base", v: (x) => x.b.Linea_Base }, { t: "Meta", v: (x) => x.b.Meta }, { t: "Último valor", v: (x) => (x.e.ult ? x.e.ult.Valor : "") },
          { t: "Fecha última medición", v: (x) => (x.e.ult ? x.e.ult.Fecha : ""), tipo: "fecha" }, { t: "Logro", v: (x) => x.e.logro ?? "", tipo: "pct" }, { t: "Estado", v: (x) => x.e.estado, ancho: 14 },
          { t: "Fecha meta", v: (x) => x.b.Fecha_Meta, tipo: "fecha" }, { t: "Próxima medición", v: (x) => x.e.proxima, tipo: "fecha" }, { t: "Dueño", v: (x) => x.b.Dueno }], filas: bs }], filtrosGenerales());
    },
    capacidad: () => {
      const idsF = new Set(filtrados().map((p) => p.ID_Proyecto));
      let lista = capacidad().filter((x) => x.asign.some((a) => idsF.has(a.p.ID_Proyecto)) || (!x.asign.length && x.otras.length && R.puede(S.usuario, "verTodo")));
      if (S.capVer && S.capVer !== "todos") lista = lista.filter((x) => nivelCap(x.total, x.cap) === S.capVer);
      const det = lista.flatMap((x) => x.asign.map((a) => ({ x, a })));
      return exportarExcel("Capacidad del equipo", [
        { nombre: "Por persona", columnas: [{ t: "Persona", v: (x) => x.nombre, ancho: 30 }, { t: "Correo", v: (x) => x.correo, ancho: 28 }, { t: "Asignado", v: (x) => x.total, tipo: "pct" },
          { t: "Proyectos (%)", v: (x) => x.tProy, tipo: "pct" }, { t: "Otras actividades (%)", v: (x) => x.tOtras, tipo: "pct" },
          { t: "Capacidad", v: (x) => x.cap, tipo: "pct" }, { t: "Estado", v: (x) => nivelCap(x.total, x.cap), ancho: 14, color: true }, { t: "Proyectos", v: (x) => x.asign.length, tipo: "num" }], filas: lista },
        { nombre: "Detalle por proyecto", columnas: [{ t: "Persona", v: (d) => d.x.nombre, ancho: 30 }, { t: "Proyecto", v: (d) => d.a.p.Nombre, ancho: 32 }, { t: "Rol", v: (d) => d.a.rol }, { t: "Dedicación", v: (d) => d.a.ded, tipo: "pct" }], filas: det },
        { nombre: "Otras actividades", columnas: [{ t: "Persona", v: (a) => a.Nombre || a.Correo, ancho: 30 }, { t: "Tipo", v: (a) => a.Tipo, ancho: 24 }, { t: "Descripción", v: (a) => a.Descripcion, ancho: 40 },
          { t: "Dedicación", v: (a) => num0(a.Dedicacion) || 0, tipo: "pct" }, { t: "Desde", v: (a) => a.Fecha_Inicio, tipo: "fecha" }, { t: "Hasta", v: (a) => a.Fecha_Fin || "Permanente", ancho: 12 }, { t: "Estado", v: (a) => estadoAct(a, corteCap()), ancho: 12 }],
          filas: actividades().filter((a) => R.puede(S.usuario, "verTodo") || lista.some((x) => x.k === clavePersona(a.Correo, a.Nombre))) }],
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
        { nombre: "Documentos", columnas: [{ t: "Documento", v: (d) => d.Nombre, ancho: 40 }, { t: "Tipo", v: (d) => d.Tipo, ancho: 18 }, { t: "Descripción", v: (d) => d.Descripcion, ancho: 36 },
          { t: "Enlace", v: (d) => d.URL, ancho: 60 }, { t: "Registrado", v: (d) => d.Fecha_Registro, tipo: "fecha" }], filas: documentosDe(pid) },
        { nombre: "Cambios", columnas: [{ t: "N.º", v: (c) => c.Numero_CC, ancho: 12 }, { t: "Fecha", v: (c) => c.Fecha, tipo: "fecha" }, { t: "Impacta en", v: (c) => impactosDe(c).join(", "), ancho: 20 }, { t: "Cambio", v: (c) => c.Descripcion, ancho: 44 },
          { t: "Justificación", v: (c) => c.Justificacion, ancho: 36 }, { t: "Impacto", v: (c) => c.Impacto, ancho: 36 }, { t: "Nueva fecha fin", v: (c) => c.Nueva_Fecha_Fin, tipo: "fecha" },
          { t: "Nuevo presupuesto", v: (c) => c.Nuevo_Presupuesto, tipo: "cop" }, { t: "Estado", v: (c) => c.Estado, ancho: 12, color: true }, { t: "Decisión", v: (c) => c.Fecha_Decision, tipo: "fecha" }, { t: "Documento", v: (c) => c.URL_Documento, ancho: 50 }], filas: cambiosDe(pid) },
        { nombre: "Cronograma", columnas: [{ t: "EDT", v: (x) => x.EDT, ancho: 10 }, { t: "Nivel", v: (x) => x.Nivel, tipo: "num", ancho: 7 }, { t: "Tarea", v: (x) => `${"   ".repeat(Math.max(0, (Number(x.Nivel) || 1) - 1))}${x.Nombre}`, ancho: 50 },
          { t: "Responsable", v: (x) => x.Responsable, ancho: 24 }, { t: "Inicio", v: (x) => x.Inicio, tipo: "fecha" }, { t: "Fin", v: (x) => x.Fin, tipo: "fecha" }, { t: "% completado", v: (x) => x.Pct_Completado, tipo: "pct" },
          { t: "Hito", v: (x) => x.Hito, ancho: 7 }, { t: "Estado", v: (x) => TXT_EST_TAREA[estadoTarea(x)], ancho: 13 }, { t: "Inicio LB", v: (x) => x.Inicio_Base, tipo: "fecha" }, { t: "Fin LB", v: (x) => x.Fin_Base, tipo: "fecha" }, { t: "Predecesoras", v: (x) => x.Predecesoras, ancho: 14 }], filas: tareasDe(pid) },
        { nombre: "Versiones cronograma", columnas: [{ t: "Versión", v: (v) => "V" + v.Numero, ancho: 9 }, { t: "Corte", v: (v) => v.Fecha_Corte, tipo: "fecha" }, { t: "Sesión", v: (v) => { const sg = v.ID_Seguimiento && S.datos.Seguimientos.find((x) => x.ID_Seguimiento === v.ID_Seguimiento); return sg ? sg.Fecha_Corte : ""; }, tipo: "fecha" },
          { t: "Avance real", v: (v) => v.Avance_Real, tipo: "pct" }, { t: "Avance planeado", v: (v) => v.Avance_Plan, tipo: "pct" }, { t: "Inicio", v: (v) => v.Inicio, tipo: "fecha" }, { t: "Fin", v: (v) => v.Fin, tipo: "fecha" },
          { t: "Atrasadas", v: (v) => v.Atrasadas, tipo: "num", ancho: 10 }, { t: "Origen", v: (v) => v.Origen, ancho: 14 }, { t: "Archivo", v: (v) => v.Archivo, ancho: 30 }, { t: "Nota", v: (v) => v.Nota, ancho: 40 }], filas: versionesDe(pid) },
        { nombre: "Líneas base", columnas: [{ t: "LB", v: (x) => "LB" + x.Numero, ancho: 8 }, { t: "Fijada", v: (x) => x.Fecha, tipo: "fecha" }, { t: "Inicio", v: (x) => x.Fecha_Inicio, tipo: "fecha" }, { t: "Fin", v: (x) => x.Fecha_Fin, tipo: "fecha" },
          { t: "Presupuesto", v: (x) => x.Presupuesto, tipo: "cop" }, { t: "Motivo", v: (x) => x.Motivo, ancho: 44 }, { t: "Control de cambios", v: (x) => { const c = x.ID_Cambio && cambiosDe(pid).find((k) => k.ID_Cambio === x.ID_Cambio); return c ? (c.Numero_CC || c.Descripcion) : ""; }, ancho: 20 }], filas: lineasBaseDe(pid) },
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
  const CRITERIOS_DEM = CRITERIOS;
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
        <td${ed({ puede: R.puede(S.usuario, "verTodo") && d.Estado !== "Convertida", tabla: "Demandas", idCol: "ID_Demanda", id: d.ID_Demanda, campo: "Nombre", tipo: "text", requerido: true, valor: d.Nombre, etiqueta: "Iniciativa" })}><b>${esc(d.Nombre)}</b><div class="sub">${esc(d.Solicitante ? `Solicita: ${d.Solicitante}` : "")}${d.Sponsor ? ` · Sponsor: ${esc(d.Sponsor)}` : ""}</div></td>
        <td>${esc(d.Tipo || "—")}<div class="sub">${esc([d.Categoria, d.Area, d.Tamano ? `Talla ${d.Tamano}` : ""].filter(Boolean).join(" · "))}</div>${obligatoria(d) ? `<span class="pill materializado">Obligatorio</span>` : ""}</td>
        <td>${pill(d.Estado || "Recibida")}${d.ID_Proyecto ? `<div><button class="btn enlace" data-pid="${esc(d.ID_Proyecto)}">Ver proyecto</button></div>` : ""}</td>
        <td class="nowrap">${pts === null ? `<span class="sub">Sin calificar</span>` : `<span class="cap-barra"><span class="cap-lleno" style="width:${pts}%;background:var(--azul)"></span></span> <b>${pts}</b>`}</td>
        <td${ed({ puede: R.puede(S.usuario, "verTodo"), tabla: "Demandas", idCol: "ID_Demanda", id: d.ID_Demanda, campo: "PM_Asignado", tipo: "select", opciones: usuariosConRol("PM").map((x) => [x.Correo, x.Nombre || x.Correo]), vacio: "Sin asignar", valor: d.PM_Asignado, etiqueta: "PM asignado" })}>${d.PM_Asignado ? persona(nombreUsuario(d.PM_Asignado)) : `<span class="sub">—</span>`}</td>
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
      const cap = () => { const c = capacidad().find((x) => lc(x.correo) === lc(pmSel.value)); pmAy.innerHTML = pmSel.value ? `Hoy tiene asignado <b class="${c && c.total > c.cap ? "baja" : ""}">${c ? c.total : 0}%</b> de ${c ? c.cap : 100}% entre proyectos y otras actividades.` : "Quien hace el análisis de la demanda."; };
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
    }, (fd) => {
      if (!String(fd.Decidido_Por || "").trim()) return "Indica quién aprueba o decide en el comité.";
      if (fd.Decision === "Aprobada" && puntajeDemanda(d) === null && !obligatoria(d)) return "Califica la demanda (todos los criterios) antes de aprobarla; solo las obligatorias se aprueban sin puntaje.";
      if (fd.Decision !== "Aprobada" && !String(fd.Comentario_Decision || "").trim()) return `Escribe el motivo: una demanda ${String(fd.Decision || "").toLowerCase()} debe quedar justificada.`;
      return "";
    }, { antes: `<div class="resumen-filtro"><b>${esc(d.Nombre)}</b><br>Puntaje ${puntajeDemanda(d) ?? "sin calificar"}${obligatoria(d) ? " · Obligatoria" : ""} · ${esc(d.Tipo || "")} · ${esc(d.Area || "")}</div>` });
  }
  // Crea el proyecto con los datos de la demanda y deja ambos enlazados.
  function convertirDemanda(d) {
    const pref = { Nombre: d.Nombre, Cliente_Area: d.Area, PM: d.PM_Asignado, Codigo_Almera: d.Codigo_Almera, Estado: "Activo", Fase: "Inicio", Presupuesto: d.Costo_Estimado || 0,
      Valor: d.Valor, Alineacion: d.Alineacion, Impacto_Paciente: d.Impacto_Paciente, Urgencia: d.Urgencia, Riesgo_Prio: d.Riesgo_Prio, Esfuerzo: d.Esfuerzo, Comentario_Estado: `Viene de la demanda ${d.Codigo_Almera}. ${d.Descripcion || ""}`.trim() };
    formProyecto(null, pref, async (idProyecto) => {
      await S.api.actualizarPorId("Demandas", "ID_Demanda", d.ID_Demanda, { Estado: "Convertida", ID_Proyecto: idProyecto, Fecha_Estado: R.hoyISO() });
      if (d.Beneficio) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", idProyecto, { Beneficio_Esperado: d.Beneficio });
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
        <tbody>${lista.map((t) => { const ET = (campo, o) => ed({ puede, tabla: "Tickets", idCol: "ID_Ticket", id: t.ID_Ticket, campo, valor: t[campo] ?? "", ...o }); return `<tr><td${ET("Numero", { tipo: "text", etiqueta: "Número" })}><b>${esc(t.Numero || "—")}</b></td><td${ET("Titulo", { tipo: "text", requerido: true, etiqueta: "Título" })}>${esc(t.Titulo)}</td>
          <td${ET("Estado", { tipo: "select", opciones: S.cat.Estado_Ticket, etiqueta: "Estado del ticket" })}>${pill(t.Estado || "Abierto")}</td><td${ET("Prioridad", { tipo: "select", opciones: S.cat.Prioridad, etiqueta: "Prioridad" })}>${esc(t.Prioridad || "—")}</td>
          <td class="opc sub">${t.Fecha_Registro ? fecha(t.Fecha_Registro) : ""} ${esc(nombreUsuario(t.Registrado_Por))}</td>
          <td class="derecha">${puede ? `<button class="btn chico" data-tk="${esc(t.ID_Ticket)}">Editar</button>` : ""}</td></tr>`; }).join("")}</tbody></table></div>`
        : vacio(todos.length ? "No hay tickets abiertos. Mira «Todos»." : `Aún no hay tickets.${puede ? " Registra el primero con «+ Ticket»." : ""}`)}</div>`;
  }
  // ---------- Documentos del proyecto (varios enlaces) ----------
  const documentosDe = (pid) => (S.datos.Documentos || []).filter((d) => d.ID_Proyecto === pid).sort((a, b) => String(a.Tipo || "~").localeCompare(String(b.Tipo || "~"), "es") || String(a.Nombre).localeCompare(String(b.Nombre), "es"));
  // Nombre legible a partir del enlace (último tramo de la ruta, sin parámetros).
  function nombreDeURL(u) {
    try {
      const url = new URL(u);
      const id = url.searchParams.get("id") || url.searchParams.get("file");
      const seg = (id || url.pathname).split("/").filter(Boolean).pop() || url.hostname;
      return decodeURIComponent(seg).replace(/\+/g, " ").replace(/\.(aspx)$/i, "").trim() || url.hostname;
    } catch (e) { return u; }
  }
  const icoDoc = (d) => {
    const ext = (String(d.Nombre).match(/\.(\w{2,5})$/) || String(d.URL).split("?")[0].match(/\.(\w{2,5})$/) || [])[1] || "";
    const t = /pdf/i.test(ext) ? "PDF" : /docx?/i.test(ext) ? "W" : /xlsx?|csv/i.test(ext) ? "X" : /pptx?/i.test(ext) ? "P" : /carpeta/i.test(d.Tipo || "") ? "📁" : "🔗";
    return `<span class="doc-ico t-${esc(t.replace(/\W/g, "") || "l")}">${t}</span>`;
  };
  function seccionDocumentos(p, puede) {
    const docs = documentosDe(p.ID_Proyecto);
    const tipos = [...new Set(docs.map((d) => d.Tipo || "Otro"))];
    const fijos = [p.URL_Documentos && esURL(p.URL_Documentos) ? `<a class="btn" href="${esc(p.URL_Documentos)}" target="_blank" rel="noopener">📁 Carpeta de documentos ↗</a>` : "",
      p.URL_Repositorio && esURL(p.URL_Repositorio) ? `<a class="btn" href="${esc(p.URL_Repositorio)}" target="_blank" rel="noopener">⌥ Repositorio ↗</a>` : ""].filter(Boolean).join(" ");
    return `<div class="card"><div class="titulo-fila"><h2>Documentos del proyecto</h2>${puede ? `<div class="acciones"><button class="btn" id="b-doc-varios">+ Pegar varios enlaces</button><button class="btn primario" id="b-doc">+ Documento</button></div>` : ""}</div>
      <p class="sub">Enlaces a los documentos del proyecto (SharePoint, OneDrive, Teams, Drive…). También los encuentras desde la búsqueda global <kbd>Ctrl K</kbd>.</p>
      ${fijos ? `<div class="doc-fijos">${fijos}</div>` : ""}
      ${docs.length ? `<div class="filtros"><label class="filtro crece"><span>Buscar documento</span><input id="doc-q" type="search" placeholder="Nombre, tipo o descripción" autocomplete="off"></label></div>
        <div class="seg doc-tipos"><button class="btn chico primario" data-doc-tipo="">Todos (${docs.length})</button>${tipos.map((t) => `<button class="btn chico" data-doc-tipo="${esc(t)}">${esc(t)} (${docs.filter((d) => (d.Tipo || "Otro") === t).length})</button>`).join("")}</div>
        <div class="tabla-scroll"><table class="tabla-docs"><thead><tr><th>Documento</th><th>Tipo</th><th class="opc">Descripción</th><th class="opc">Registrado</th><th></th></tr></thead><tbody>
        ${docs.map((d) => `<tr data-doc="${esc(d.ID_Documento)}" data-buscar="${esc(normTxt(`${d.Nombre} ${d.Tipo} ${d.Descripcion} ${d.URL}`))}" data-tipo="${esc(d.Tipo || "Otro")}">
          <td${ed({ puede, tabla: "Documentos", idCol: "ID_Documento", id: d.ID_Documento, campo: "Nombre", tipo: "text", valor: d.Nombre || "", etiqueta: "Nombre del documento" })}>${icoDoc(d)} <a href="${esc(d.URL)}" target="_blank" rel="noopener" data-doc-ir>${esc(d.Nombre || nombreDeURL(d.URL))}</a></td>
          <td${ed({ puede, tabla: "Documentos", idCol: "ID_Documento", id: d.ID_Documento, campo: "Tipo", tipo: "select", opciones: S.cat.Tipo_Documento, valor: d.Tipo || "Otro", etiqueta: "Tipo" })}>${pill(d.Tipo || "Otro")}</td><td class="opc"${ed({ puede, tabla: "Documentos", idCol: "ID_Documento", id: d.ID_Documento, campo: "Descripcion", tipo: "textarea", valor: d.Descripcion || "", etiqueta: "Descripción" })}>${esc(d.Descripcion || "—")}</td>
          <td class="opc sub nowrap">${d.Fecha_Registro ? fecha(d.Fecha_Registro) : ""}<br>${esc(nombreUsuario(d.Registrado_Por))}</td>
          <td class="derecha nowrap"><button class="btn chico" data-doc-copiar="${esc(d.URL)}" title="Copiar el enlace">Copiar enlace</button>${puede ? ` <button class="btn chico" data-doc-editar="${esc(d.ID_Documento)}">Editar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>
        <p class="sub" id="doc-nada" hidden>Ningún documento coincide.</p>`
        : vacio(`Aún no hay documentos.${puede ? " Agrega el primero con «+ Documento» o pega varios enlaces de una vez." : ""}`)}</div>`;
  }
  function enlazarDocumentos(el, p) {
    if ($("#b-doc")) $("#b-doc").addEventListener("click", () => formDocumento(p));
    if ($("#b-doc-varios")) $("#b-doc-varios").addEventListener("click", () => formDocumentosVarios(p));
    el.querySelectorAll("[data-doc-editar]").forEach((b) => b.addEventListener("click", () => formDocumento(p, (S.datos.Documentos || []).find((d) => d.ID_Documento === b.dataset.docEditar))));
    el.querySelectorAll("[data-doc-copiar]").forEach((b) => b.addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(b.dataset.docCopiar); toast("Enlace copiado"); } catch (e) { const t = document.createElement("textarea"); t.value = b.dataset.docCopiar; document.body.appendChild(t); t.select(); try { document.execCommand("copy"); toast("Enlace copiado"); } catch (e2) { toast("No se pudo copiar", true); } t.remove(); }
    }));
    const q = el.querySelector("#doc-q");
    if (!q) return;
    let tipo = "";
    const filtrar = () => {
      const terms = normTxt(q.value).split(" ").filter(Boolean);
      let n = 0;
      el.querySelectorAll("tr[data-doc]").forEach((tr) => { const ok = (!tipo || tr.dataset.tipo === tipo) && terms.every((t) => tr.dataset.buscar.includes(t)); tr.hidden = !ok; if (ok) n++; });
      el.querySelector("#doc-nada").hidden = n > 0;
    };
    q.addEventListener("input", filtrar);
    el.querySelectorAll("[data-doc-tipo]").forEach((b) => b.addEventListener("click", () => { tipo = b.dataset.docTipo; el.querySelectorAll("[data-doc-tipo]").forEach((x) => x.classList.toggle("primario", x === b)); filtrar(); }));
  }
  const nuevoIdDoc = (pid, i = 0) => R.siguienteIdHijo("DOC", S.datos.Documentos || [], "ID_Documento", pid, i);
  function formDocumento(p, d) {
    const nuevo = !d;
    const campos = [
      { k: "URL", label: "Enlace (URL)", tipo: "url", placeholder: "https://…sharepoint.com/…", ancho: true },
      { k: "Nombre", label: "Nombre del documento", placeholder: "Se llena solo con el nombre del archivo; puedes cambiarlo", ancho: true },
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: S.cat.Tipo_Documento, def: "Otro" },
      { k: "Descripcion", label: "Descripción (opcional)", placeholder: "Para qué sirve o qué versión es" },
    ];
    const init = (m) => {
      const u = m.querySelector("[name=URL]"), n = m.querySelector("[name=Nombre]");
      u.addEventListener("change", () => { if (!n.value.trim() && esURL(u.value.trim())) n.value = nombreDeURL(u.value.trim()); });
      if (nuevo) u.focus();
    };
    const validar = (fd) => (!fd.URL ? "Pega el enlace del documento." : !esURL(fd.URL) ? "El enlace debe empezar por https://" :
      documentosDe(p.ID_Proyecto).some((x) => x.URL === fd.URL && (nuevo || x.ID_Documento !== d.ID_Documento)) ? "Ese enlace ya está registrado en este proyecto." : "");
    modal(nuevo ? `Nuevo documento · ${p.Nombre}` : "Editar documento", campos, nuevo ? {} : d, (fd) => {
      fd.Nombre = fd.Nombre || nombreDeURL(fd.URL);
      if (nuevo) guardar(() => S.api.agregarFila("Documentos", { ID_Documento: nuevoIdDoc(p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd, Registrado_Por: S.usuario.Correo, Fecha_Registro: R.hoyISO() }), "Documento agregado");
      else guardar(() => S.api.actualizarPorId("Documentos", "ID_Documento", d.ID_Documento, fd), "Documento actualizado");
    }, validar, { init, ...(nuevo ? {} : { eliminar: { texto: "Quitar documento", mensaje: `Se quitará «${d.Nombre}» de la lista (el archivo en SharePoint no se borra).`,
      accion: () => guardar(() => S.api.eliminarFilas("Documentos", "ID_Documento", [d.ID_Documento]), "Documento quitado") } }) });
  }
  // Pegar varios enlaces de una vez: uno por línea, opcionalmente «Nombre | enlace».
  function formDocumentosVarios(p) {
    modal(`Pegar varios enlaces · ${p.Nombre}`, [
      { k: "Lista", label: "Enlaces (uno por línea)", tipo: "textarea", placeholder: "https://…/Cronograma.xlsx\nActa de inicio | https://…/Acta inicio.pdf\nhttps://…/Especificación.docx" },
      { k: "Tipo", label: "Tipo para todos", tipo: "select", opciones: S.cat.Tipo_Documento, def: "Otro", ayuda: "Luego puedes cambiar el tipo de cada uno." },
    ], {}, (fd) => {
      const ya = new Set(documentosDe(p.ID_Proyecto).map((x) => x.URL));
      const filas = [];
      String(fd.Lista).split(/\n+/).map((l) => l.trim()).filter(Boolean).forEach((l) => {
        const m = l.match(/^(.*?)\s*[|;\t]\s*(https?:\/\/\S+)\s*$/) || l.match(/^(https?:\/\/\S+)$/);
        if (!m) return;
        const url = (m[2] || m[1]).trim(), nombre = m[2] ? m[1].trim() : nombreDeURL(url);
        if (ya.has(url) || filas.some((f) => f.URL === url)) return;
        filas.push({ ID_Documento: nuevoIdDoc(p.ID_Proyecto, filas.length), ID_Proyecto: p.ID_Proyecto, Nombre: nombre, URL: url, Tipo: fd.Tipo, Descripcion: "", Registrado_Por: S.usuario.Correo, Fecha_Registro: R.hoyISO() });
      });
      if (!filas.length) return toast("No encontré enlaces nuevos (deben empezar por https:// y no estar ya registrados).", true);
      guardar(() => S.api.agregarFilas("Documentos", filas), `${filas.length} documento(s) agregados`);
    }, (fd) => (!String(fd.Lista || "").match(/https?:\/\//) ? "Pega al menos un enlace que empiece por https://" : ""));
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
  // ---------- Otras actividades (fuera de proyectos) ----------
  const actividades = () => S.datos.Actividades || [];
  const esAusencia = (a) => /ausencia|vacaci|incapacid|permiso|licencia/i.test(String(a.Tipo || ""));
  const vigenteAct = (a, corte) => (!a.Fecha_Inicio || String(a.Fecha_Inicio) <= corte) && (!a.Fecha_Fin || String(a.Fecha_Fin) >= corte);
  const estadoAct = (a, corte = R.hoyISO()) => (a.Fecha_Inicio && String(a.Fecha_Inicio) > corte ? "Programada" : a.Fecha_Fin && String(a.Fecha_Fin) < corte ? "Finalizada" : "Vigente");
  const nuevoIdAct = () => `ACT-${String(actividades().reduce((m, a) => Math.max(m, parseInt(String(a.ID_Actividad).slice(4), 10) || 0), 0) + 1).padStart(4, "0")}`;
  const JORNADA_SEMANAL = 42;   // horas/semana para convertir horas → %
  const corteCap = () => S.capCorte || R.hoyISO();

  // Carga de cada persona a la fecha de corte: proyectos activos o en pausa + otras actividades vigentes.
  // Las ausencias no suman carga: restan capacidad disponible en su periodo.
  function capacidad(corte = corteCap()) {
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
    // Tiempo del PM en las sesiones de su agenda (calculado): cuenta como otra actividad.
    const minPM = {};
    S.datos.Proyectos.filter(CUENTAN).forEach((p) => agendaDe(p.ID_Proyecto).filter((a) => agendaActiva(a, corte)).forEach((a) => { if (p.PM) minPM[lc(p.PM)] = (minPM[lc(p.PM)] || 0) + minutosSemana(a); }));
    Object.entries(minPM).forEach(([correo, min]) => {
      const ded = Math.round((min / 60 / JORNADA_SEMANAL) * 100);
      if (!ded) return;
      const k = clavePersona(correo, "");
      const x = (personas[k] = personas[k] || { k, nombre: nombreUsuario(correo), correo, total: 0, asign: [] });
      (x.otras = x.otras || []).push({ ID_Actividad: "", Tipo: "Sesiones de seguimiento (agenda)", Descripcion: `${durTxt(min)} por semana en sesiones programadas`, Dedicacion: ded, Fecha_Inicio: "", Fecha_Fin: "", auto: true });
    });
    actividades().filter((a) => vigenteAct(a, corte) && num0(a.Dedicacion) !== null && (a.Correo || a.Nombre)).forEach((a) => {
      const k = clavePersona(a.Correo, a.Nombre);
      const x = (personas[k] = personas[k] || { k, nombre: a.Nombre || nombreUsuario(a.Correo), correo: a.Correo || "", total: 0, asign: [] });
      (x.otras = x.otras || []).push(a);
    });
    Object.values(personas).forEach((x) => {
      const r = recursoPor({ correo: x.correo, nombre: x.nombre });
      x.otras = x.otras || [];
      x.tProy = x.total;
      x.tOtras = x.otras.filter((a) => !esAusencia(a)).reduce((t, a) => t + num0(a.Dedicacion), 0);
      x.aus = Math.min(100, x.otras.filter(esAusencia).reduce((t, a) => t + num0(a.Dedicacion), 0));
      x.total = x.tProy + x.tOtras;
      x.capBase = (r && num0(r.Capacidad)) ?? 100;
      x.cap = Math.max(0, Math.round(x.capBase * (1 - x.aus / 100)));
      if (r) x.nombre = r.Nombre;
    });
    return Object.values(personas).sort((a, b) => b.total / b.cap - a.total / a.cap || String(a.nombre).localeCompare(String(b.nombre), "es"));
  }
  const nivelCap = (t, cap = 100) => (t > cap ? "Sobreasignado" : t >= cap * 0.85 ? "Al límite" : "Disponible");
  const totalDe = (correo, nombre) => (capacidad().find((x) => x.k === clavePersona(correo, nombre)) || { total: 0, asign: [] });
  const barraCap = (t, cap = 100, otras = 0) => `<span class="cap-barra" title="${otras ? `${t - otras}% proyectos + ${otras}% otras actividades = ${t}%` : `${t}%`} de ${cap}%"><span class="cap-lleno ${t > cap ? "rojo" : t >= cap * 0.85 ? "ambar" : ""}" style="width:${Math.min(100, (t / 150) * 100)}%">${otras && t ? `<span class="cap-otras" style="width:${Math.min(100, (otras / t) * 100)}%"></span>` : ""}</span><span class="cap-100" style="left:${(cap / 150) * 100}%"></span></span> <b>${t}%</b>${cap !== 100 ? ` <span class="sub">de ${cap}%</span>` : ""}`;

  function vCapacidad(el) {
    const u = S.usuario, todo = R.puede(u, "verTodo");
    const ids = new Set(visibles().map((p) => p.ID_Proyecto));
    const idsF = new Set(filtrados().map((p) => p.ID_Proyecto));
    const ver = S.capVer || "todos", corte = corteCap(), hoy = R.hoyISO();
    const filtroProy = S.filtros.pm || S.filtros.proyecto;
    const yo = (x) => u && x.correo && lc(x.correo) === lc(u.Correo);
    // Personas de los proyectos filtrados (su total cuenta todo). Quien solo tiene otras actividades aparece para PMO/Admin, o si es uno mismo.
    const base = capacidad(corte).filter((x) => x.asign.some((a) => idsF.has(a.p.ID_Proyecto)) || (!x.asign.length && x.otras.length && ((todo && !filtroProy) || yo(x))));
    let lista = base;
    const n = (nv) => lista.filter((x) => nivelCap(x.total, x.cap) === nv).length;
    if (ver !== "todos") lista = lista.filter((x) => nivelCap(x.total, x.cap) === ver);
    const capT = base.reduce((t, x) => t + x.capBase, 0), otrasT = base.reduce((t, x) => t + x.tOtras, 0), proyT = base.reduce((t, x) => t + x.tProy, 0);
    const pctOtras = capT ? Math.round((otrasT / capT) * 100) : 0;
    // Otras actividades que puede ver: PMO/Admin todas; el PM las de las personas de sus proyectos y las suyas.
    const llaves = new Set(base.map((x) => x.k));
    const AF = S.actF || { tipo: "", estado: "Vigente" };
    const acts = actividades().filter((a) => todo || llaves.has(clavePersona(a.Correo, a.Nombre)) || (u && lc(a.Correo) === lc(u.Correo)))
      .filter((a) => (!AF.tipo || a.Tipo === AF.tipo) && (!AF.estado || estadoAct(a, corte) === AF.estado))
      .sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es") || String(a.Fecha_Inicio).localeCompare(String(b.Fecha_Inicio)));
    const puedeAct = R.puede(u, "editarProyecto");
    const editable = (a) => todo || (u && (lc(a.Correo) === lc(u.Correo) || lc(a.Registrado_Por) === lc(u.Correo)));
    const tipos = [...new Set([...(S.cat.Tipo_Actividad || []), ...actividades().map((a) => a.Tipo).filter(Boolean)])];
    el.innerHTML = `
      <h1>Capacidad del equipo</h1>
      ${barraFiltros()}
      <div class="filtros"><label class="filtro"><span>Fecha de corte</span><input type="date" id="cap-corte" value="${esc(corte)}"></label>
        ${corte !== hoy ? `<button class="btn chico" id="cap-hoy">Volver a hoy</button>` : ""}<span class="sub cap-corte-ay">Las otras actividades y ausencias cuentan si están vigentes en esta fecha.</span></div>
      <div class="kpis kpis-3">
        <div class="kpi estatico ${n("Sobreasignado") ? "alerta" : ""}"><span class="kpi-t">Sobreasignados (sobre su capacidad)</span><span class="kpi-v">${n("Sobreasignado")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Al límite (85% a 100% de su capacidad)</span><span class="kpi-v">${n("Al límite")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Con capacidad disponible</span><span class="kpi-v">${n("Disponible")}</span></div>
        <div class="kpi estatico"><span class="kpi-t">Capacidad del equipo en otras actividades</span><span class="kpi-v">${pctOtras}%</span><span class="sub">${proyT ? `Proyectos: ${capT ? Math.round((proyT / capT) * 100) : 0}%` : ""}</span></div>
      </div>
      <div class="card"><div class="titulo-fila"><h2>Dedicación por persona</h2>
        <div class="seg">${["todos", "Sobreasignado", "Al límite", "Disponible"].map((v) => `<button class="btn chico ${ver === v ? "primario" : ""}" data-capver="${v}">${v === "todos" ? "Todos" : v}</button>`).join("")}</div></div>
        <p class="sub">Suma el % de dedicación en proyectos activos o en pausa (PM y stakeholders) y en <b>otras actividades</b> vigentes (parte rayada de la barra). Las ausencias reducen la capacidad de la persona mientras duran. La línea marca su capacidad.</p>
        ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Persona</th><th>Total</th><th>Estado</th><th>Proyectos</th><th>Otras actividades</th></tr></thead><tbody>
          ${lista.map((x) => `<tr><td>${persona(x.nombre, x.correo)}</td><td class="nowrap">${barraCap(x.total, x.cap, x.tOtras)}${x.aus ? `<div class="sub">Ausencia ${x.aus}% · capacidad ${x.cap}% de ${x.capBase}%</div>` : ""}</td><td>${pill(nivelCap(x.total, x.cap))}</td>
            <td><div class="cap-asign">${x.asign.sort((a, b) => b.ded - a.ded).map((a) => ids.has(a.p.ID_Proyecto)
              ? `<button class="chip-proy" data-pid="${esc(a.p.ID_Proyecto)}">${esc(a.p.Nombre)} · ${esc(a.rol)} · <b>${a.ded}%</b></button>`
              : `<span class="chip-proy">Otro proyecto · <b>${a.ded}%</b></span>`).join("") || `<span class="sub">—</span>`}</div></td>
            <td><div class="cap-asign">${x.otras.map((a) => `<span class="chip-act ${esAusencia(a) ? "aus" : ""}" title="${esc([a.Descripcion, a.Fecha_Fin ? `hasta ${fecha(a.Fecha_Fin)}` : "permanente"].filter(Boolean).join(" · "))}">${esc(a.Tipo || "Actividad")} · <b>${num0(a.Dedicacion)}%</b></span>`).join("") || `<span class="sub">—</span>`}</div></td></tr>`).join("")}</tbody></table></div>`
        : vacio("Aún no hay dedicaciones registradas. Pon el % de dedicación del PM en «Editar proyecto», el de cada stakeholder en su ficha y las otras actividades abajo.")}
      </div>
      <div class="card" id="otras-actividades"><div class="titulo-fila"><h2>Otras actividades <span class="sub">${acts.length}</span></h2>${puedeAct ? `<button class="btn primario" id="b-act">+ Actividad</button>` : ""}</div>
        <p class="sub">Lo que no es proyecto pero consume tiempo: operación y soporte, comités, gestión de demanda, informes, capacitación y ausencias. Sin fecha de fin = permanente.</p>
        <div class="filtros"><label class="filtro"><span>Tipo</span><select id="act-tipo"><option value="">Todos</option>${tipos.map((t) => `<option ${t === AF.tipo ? "selected" : ""}>${esc(t)}</option>`).join("")}</select></label>
          <label class="filtro"><span>Estado</span><select id="act-estado">${[["Vigente", "Vigentes"], ["Programada", "Programadas"], ["Finalizada", "Finalizadas"], ["", "Todas"]].map(([v, t]) => `<option value="${v}" ${v === AF.estado ? "selected" : ""}>${t}</option>`).join("")}</select></label></div>
        ${acts.length ? `<div class="tabla-scroll"><table><thead><tr><th>Persona</th><th>Tipo</th><th>Descripción</th><th>Dedicación</th><th>Vigencia</th><th>Estado</th><th></th></tr></thead><tbody>
          ${acts.map((a) => `<tr><td>${persona(a.Nombre || nombreUsuario(a.Correo), a.Correo)}</td><td${ed({ puede: editable(a), tabla: "Actividades", idCol: "ID_Actividad", id: a.ID_Actividad, campo: "Tipo", tipo: "select", opciones: S.cat.Tipo_Actividad, valor: a.Tipo, etiqueta: "Tipo" })}><span class="chip-act ${esAusencia(a) ? "aus" : ""}">${esc(a.Tipo || "—")}</span></td>
            <td${ed({ puede: editable(a), tabla: "Actividades", idCol: "ID_Actividad", id: a.ID_Actividad, campo: "Descripcion", tipo: "textarea", valor: a.Descripcion, etiqueta: "Descripción" })}>${esc(a.Descripcion || "—")}</td>
            <td class="nowrap"${ed({ puede: editable(a), tabla: "Actividades", idCol: "ID_Actividad", id: a.ID_Actividad, campo: "Dedicacion", tipo: "number", min: 1, max: 100, requerido: true, valor: num0(a.Dedicacion) ?? "", etiqueta: "Dedicación" })}><b>${num0(a.Dedicacion) ?? "—"}%</b></td><td class="nowrap">${fecha(a.Fecha_Inicio)} → ${a.Fecha_Fin ? fecha(a.Fecha_Fin) : "Permanente"}</td><td>${pill(estadoAct(a, corte))}</td>
            <td class="derecha">${editable(a) ? `<button class="btn chico" data-act="${esc(a.ID_Actividad)}">Editar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`
        : vacio(AF.estado || AF.tipo ? "No hay actividades con ese filtro." : "Aún no hay otras actividades registradas.")}
      </div>`;
    enlazarFiltros();
    el.querySelectorAll("[data-capver]").forEach((b) => b.addEventListener("click", () => { S.capVer = b.dataset.capver; render(); }));
    el.querySelectorAll(".chip-proy[data-pid]").forEach((b) => b.addEventListener("click", (e) => { e.stopPropagation(); ir("ficha", b.dataset.pid); }));
    $("#cap-corte").addEventListener("change", (e) => { S.capCorte = /^\d{4}-\d{2}-\d{2}$/.test(e.target.value) ? e.target.value : ""; render(); });
    if ($("#cap-hoy")) $("#cap-hoy").addEventListener("click", () => { S.capCorte = ""; render(); });
    $("#act-tipo").addEventListener("change", (e) => { S.actF = { ...AF, tipo: e.target.value }; render(); });
    $("#act-estado").addEventListener("change", (e) => { S.actF = { ...AF, estado: e.target.value }; render(); });
    if ($("#b-act")) $("#b-act").addEventListener("click", () => formActividad());
    el.querySelectorAll("[data-act]").forEach((b) => b.addEventListener("click", () => formActividad(actividades().find((a) => a.ID_Actividad === b.dataset.act))));
  }

  // Formulario de otra actividad. PMO/Admin eligen a cualquiera; el PM, a sí mismo o a las personas de sus proyectos.
  function formActividad(a, persona0) {
    const nuevo = !a, u = S.usuario, todo = R.puede(u, "verTodo");
    const ops = new Map();
    const add = (correo, nombre, id) => { if (!(correo || nombre)) return; const k = clavePersona(correo, nombre); if (!ops.has(k)) ops.set(k, { k, correo: correo || "", nombre: nombre || nombreUsuario(correo), id: id || "" }); };
    if (u) { const r = recursoPor({ correo: u.Correo }); add(u.Correo, r ? r.Nombre : u.Nombre, r && r.ID_Recurso); }
    if (todo) {
      recursos().filter((r) => r.Activo !== "No").forEach((r) => add(r.Correo, r.Nombre, r.ID_Recurso));
      S.datos.Usuarios.filter((x) => x.Activo === "Sí" && R.tieneRol(x, "PM")).forEach((x) => add(x.Correo, x.Nombre));
    } else {
      capacidad().filter((x) => x.asign.some((y) => R.esMio(u, y.p))).forEach((x) => { const r = recursoPor({ correo: x.correo, nombre: x.nombre }); add(x.correo, x.nombre, r && r.ID_Recurso); });
    }
    if (a) add(a.Correo, a.Nombre, a.ID_Recurso);
    const lista = [...ops.values()].sort((x, y) => String(x.nombre).localeCompare(String(y.nombre), "es"));
    const actual = a ? clavePersona(a.Correo, a.Nombre) : persona0 ? clavePersona(persona0.Correo, persona0.Nombre) : (u ? clavePersona(u.Correo) : "");
    const campos = [
      { k: "_Persona", label: "Persona", tipo: "select", ancho: true, opciones: lista.map((x) => [x.k, `${x.nombre}${x.correo ? ` · ${x.correo}` : ""}`]), def: actual },
      { k: "Tipo", label: "Tipo de actividad", tipo: "select", opciones: S.cat.Tipo_Actividad || [] },
      { k: "Dedicacion", label: "Dedicación (%)", tipo: "number", min: 1, max: 100, ayuda: "Del tiempo total de la persona. Ausencia: 100 si está fuera todo el periodo." },
      { k: "_Horas", label: "…o en horas por semana", tipo: "number", min: 0, max: JORNADA_SEMANAL, placeholder: `de ${JORNADA_SEMANAL} h`, ayuda: `Opcional: calcula el % con una jornada de ${JORNADA_SEMANAL} h.` },
      { k: "Descripcion", label: "Descripción", tipo: "textarea", placeholder: "Ej.: Soporte de segundo nivel a Kardex; comité de cambios semanal; vacaciones." },
      { k: "Fecha_Inicio", label: "Desde", tipo: "date", def: R.hoyISO() },
      { k: "Fecha_Fin", label: "Hasta (opcional)", tipo: "date", ayuda: "Vacío = permanente." },
    ];
    modal(nuevo ? "Nueva actividad" : "Editar actividad", campos, a || {}, (fd) => {
      const x = ops.get(fd._Persona) || {};
      const fila = { Correo: x.correo || "", Nombre: x.nombre || "", ID_Recurso: x.id || "", Tipo: fd.Tipo, Descripcion: fd.Descripcion, Dedicacion: fd.Dedicacion, Fecha_Inicio: fd.Fecha_Inicio, Fecha_Fin: fd.Fecha_Fin };
      if (nuevo) guardar(() => S.api.agregarFila("Actividades", { ID_Actividad: nuevoIdAct(), ...fila, Registrado_Por: u.Correo, Fecha_Registro: R.hoyISO() }), "Actividad registrada");
      else guardar(() => S.api.actualizarPorId("Actividades", "ID_Actividad", a.ID_Actividad, fila), "Actividad actualizada");
    }, (fd) => {
      if (!fd._Persona) return "Elige la persona.";
      if (!fd.Tipo) return "Elige el tipo de actividad.";
      if (fd.Dedicacion === "") return "Escribe la dedicación en % (o en horas por semana).";
      if (!fd.Fecha_Inicio) return "Escribe desde cuándo.";
      if (fd.Fecha_Fin && fd.Fecha_Fin < fd.Fecha_Inicio) return "La fecha «Hasta» no puede ser anterior a «Desde».";
      return "";
    }, {
      init: (m) => {
        const h = m.querySelector('[name="_Horas"]'), d = m.querySelector('[name="Dedicacion"]');
        h.addEventListener("input", () => { const v = Number(h.value); if (h.value !== "" && v >= 0) d.value = Math.max(1, Math.min(100, Math.round((v / JORNADA_SEMANAL) * 100))); });
      },
      ...(nuevo ? {} : { eliminar: { texto: "Eliminar actividad", mensaje: `Se eliminará «${a.Tipo}${a.Descripcion ? `: ${a.Descripcion}` : ""}» de ${a.Nombre || a.Correo}. Si ya terminó, mejor pon la fecha «Hasta» para conservar el histórico.`,
        accion: () => guardar(() => S.api.eliminarFilas("Actividades", "ID_Actividad", [a.ID_Actividad]), "Actividad eliminada") } }),
    });
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
          <td${ed({ puede, tabla: "Proyectos", idCol: "ID_Proyecto", id: p.ID_Proyecto, campo: "Dedicacion_PM", tipo: "number", min: 0, max: 100, valor: p.Dedicacion_PM ?? "", extra: () => sello(), etiqueta: "Dedicación del PM" })}>${celdaDedicacion(p.PM, nombreUsuario(p.PM), p.Dedicacion_PM)}</td><td class="derecha sub"></td></tr>
        ${lista.map((x) => `<tr><td${ed({ puede, tabla: "Stakeholders", idCol: "ID_Stakeholder", id: x.ID_Stakeholder, campo: "Rol", tipo: "select", opciones: S.cat.Rol_Stakeholder, valor: x.Rol, etiqueta: "Rol" })}><b>${esc(x.Rol)}</b></td><td>${persona(x.Nombre)}</td><td>${esc([x.Cargo, x.Area].filter(Boolean).join(" · ") || "—")}</td>
          <td>${x.ID_Proveedor ? esc((proveedor(x.ID_Proveedor) || {}).Nombre || x.ID_Proveedor) : "Interno"}</td><td>${contacto(x.Correo, x.Telefono)}</td>
          <td${ed({ puede, tabla: "Stakeholders", idCol: "ID_Stakeholder", id: x.ID_Stakeholder, campo: "Dedicacion", tipo: "number", min: 0, max: 100, valor: x.Dedicacion ?? "", etiqueta: "Dedicación" })}>${celdaDedicacion(x.Correo, x.Nombre, x.Dedicacion)}</td>
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
          (S.cat.Rol_Stakeholder = S.cat.Rol_Stakeholder || []).push(rol);
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
        if (seCierra) { S.fichaTab = "lecciones"; if ((S.cat.Fase || []).includes("Cierre")) fd.Fase = "Cierre"; }
        const pend = seCierra ? pendientesCierre(p) : [];
        guardar(async () => { await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...fd, ...sello() }); await guardarEquipo(p.ID_Proyecto); },
          seCierra ? `Proyecto cerrado.${pend.length ? ` Quedaron abiertos: ${pend.join(", ")}.` : ""} Es buen momento para registrar las lecciones aprendidas.` : "Guardado");
      }
    }, (fd) => {
      if (fd.Fecha_Inicio && fd.Fecha_Fin_Plan && fd.Fecha_Fin_Plan < fd.Fecha_Inicio) return "La fecha fin no puede ser anterior a la fecha de inicio.";
      if (!nuevo && fd.Estado === "Cerrado" && p.Estado !== "Cerrado") {
        const sol = cambiosDe(p.ID_Proyecto).filter((c) => (c.Estado || "Solicitado") === "Solicitado").length;
        if (sol) return `No se puede cerrar: hay ${sol} control(es) de cambio por decidir. Apruébalos o recházalos en la pestaña Cambios.`;
      }
      if (fd.Avance_Real !== "" && fd.Avance_Planeado !== "" && (Number(fd.Avance_Real) > 100 || Number(fd.Avance_Planeado) > 100)) return "El avance va de 0 a 100%.";
      return "";
    },
    !nuevo && R.puede(u, "verTodo") ? { init: (m) => { initProv(m); initEquipo(m); }, eliminar: { texto: "Eliminar proyecto", titulo: `Eliminar ${p.ID_Proyecto}`,
      mensaje: `Se borrará «${p.Nombre}» con todos sus seguimientos, compromisos, hitos, riesgos, stakeholders y actas. No se puede deshacer. Si solo terminó, mejor cambia su estado a Cerrado o Cancelado.`,
      accion: () => { S.vista = "proyectos"; guardar(() => eliminarProyecto(p.ID_Proyecto), `Proyecto ${p.ID_Proyecto} eliminado`); } } } : { init: (m) => { initProv(m); initEquipo(m); } });
  }

  // Lo que sigue abierto al cerrar un proyecto (se avisa; no bloquea salvo cambios por decidir).
  function pendientesCierre(p) {
    const pid = p.ID_Proyecto, n = [];
    const comps = S.datos.Compromisos.filter((c) => c.ID_Proyecto === pid && !R.cerrado(c)).length;
    const rsg = (S.datos.Riesgos || []).filter((r) => r.ID_Proyecto === pid && ["Abierto", "Materializado"].includes(r.Estado)).length;
    const hit = (S.datos.Hitos || []).filter((h) => h.ID_Proyecto === pid && h.Estado !== "Cumplido").length;
    const tks = (S.datos.Tickets || []).filter((t) => t.ID_Proyecto === pid && !["Resuelto", "Cerrado"].includes(t.Estado)).length;
    const sinMed = beneficiosDe(pid).filter((b) => !medicionesDe(b.ID_Beneficio).length).length;
    if (sinMed) n.push(`${sinMed} beneficio(s) sin medir (se siguen midiendo después del cierre)`);
    if (comps) n.push(`${comps} compromiso(s)`); if (rsg) n.push(`${rsg} riesgo(s)`); if (hit) n.push(`${hit} hito(s)`); if (tks) n.push(`${tks} ticket(s)`);
    return n;
  }

  // tipoInicial: "Seguimiento" (reporte de avance) o "Sesión" (comité, mesa técnica, reunión: acta y compromisos, sin avance).
  function formSeguimiento(p, tipoInicial = "Seguimiento") {
    S.pidForm = p.ID_Proyecto;
    const e = R.estadoSeguimiento(p, S.datos.Seguimientos);
    const ult = e.ultimo;
    const abiertos = S.datos.Compromisos.filter((c) => c.ID_Proyecto === p.ID_Proyecto && !R.cerrado(c))
      .sort((a, b) => (a.Fecha_Compromiso < b.Fecha_Compromiso ? -1 : 1));
    const campos = [
      { k: "Tipo_Sesion", label: "¿Qué vas a registrar?", tipo: "select", def: tipoInicial, opciones: [["Seguimiento", "Seguimiento de avance (reporta avance, semáforo y presupuesto)"], ["Sesión", "Sesión / reunión (comité, mesa técnica…): acta y compromisos, sin avance"]], ancho: true },
      { k: "Nombre_Sesion", label: "Nombre de la sesión", placeholder: "Ej.: Comité de cambios, Mesa técnica con Sonda, Revisión con el área" },
      { seccion: "1. Avance" },
      { k: "Fecha_Corte", label: "Fecha de corte / de la sesión", tipo: "date", def: R.hoyISO() },
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
      <div><b>¿Tienes el acta?</b><div class="sub">Súbela en PDF o Word (.docx), en cualquier formato: la app lee la fecha y los compromisos (tabla o lista), los llena abajo y guarda el acta con la sesión.</div></div>
      <button type="button" class="btn primario" id="b-acta-pdf">Subir acta (PDF o Word)</button>
      <div id="acta-estado" class="acta-estado" aria-live="polite"></div></div>`;
    const contexto = `${datalistRecursos("dl-recursos")}<div class="contexto">
      <div><span class="sub">Último reporte</span><b>${ult ? `${fecha(ult.Fecha_Corte)} · ${pct(ult.Avance_Real)} · ${chipSemaforo(ult.Semaforo)}` : "Es el primer reporte"}</b></div>
      <div><span class="sub">Plan registrado</span><b>${pct(p.Avance_Planeado)}</b></div>
      <div><span class="sub">Frecuencia</span><b>${esc(p.Frecuencia_Seguimiento)}</b></div></div>
      <div class="crono-seg" id="crono-seg"></div>`;
    const despues = `
      <div class="sugerencia" id="sug-sem" aria-live="polite"></div>
      ${abiertos.length ? `<div class="form-seccion">4. Compromisos anteriores<span class="sub"> · actualiza su estado en esta sesión</span></div>
        <div class="comp-previos">${abiertos.map((c) => `<div class="cp-fila"><span>${esc(c.Compromiso)} <span class="sub">${esc(c.Responsable || "sin responsable")} · ${fecha(c.Fecha_Compromiso)}</span> ${R.estadoCompromiso(c) === "Vencido" ? pill("Vencido") : ""}</span>
          <select class="cp-estado" data-cp="${esc(c.ID_Compromiso)}" aria-label="Estado de ${esc(c.Compromiso)}">${R.ESTADOS_COMPROMISO.map((e) => `<option ${e === R.estadoBase(c) ? "selected" : ""}>${esc(e)}</option>`).join("")}</select></div>`).join("")}</div>` : ""}
      <div class="form-seccion">${abiertos.length ? "5" : "4"}. Compromisos nuevos<span class="sub"> · acuerdos que quedaron en este seguimiento (opcional)</span></div>
      <div class="comp-cab"><span>Compromiso</span><span>Responsable(s)</span><span>Fecha límite</span><span>Estado</span><span></span></div>
      <div id="comp-nuevos">${filaComp()}</div>
      <button type="button" class="btn chico" id="b-add-comp">+ Agregar otro compromiso</button>`;
    const init = (m) => {
      // Seguimiento vs. sesión: en una sesión no se reporta avance, semáforo ni presupuesto.
      const secs = [...m.querySelectorAll(".form-seccion")];
      const lbl = (k) => { const x = m.querySelector(`[name=${k}]`); return x && x.closest("label"); };
      const tipo = () => m.querySelector("[name=Tipo_Sesion]").value;
      const aplicarTipo = () => {
        const ses = tipo() === "Sesión";
        ["Avance_Real", "Avance_Planeado", "Semaforo", "Ejecutado"].forEach((k) => { const l = lbl(k); if (l) l.hidden = ses; });
        const ln = lbl("Nombre_Sesion"); if (ln) ln.hidden = !ses;
        if (secs[0]) secs[0].firstChild.textContent = ses ? "1. Fecha de la sesión" : "1. Avance";
        if (secs[1]) secs[1].firstChild.textContent = ses ? "2. Qué se trató" : "2. Qué pasó en el periodo";
        const lg = lbl("Logros"); if (lg) lg.querySelector(".etq").textContent = ses ? "Temas tratados / conclusiones" : "Logros del periodo";
        ["#sug-sem", "#crono-seg", ".contexto"].forEach((q) => { const x = m.querySelector(q); if (x) x.hidden = ses; });
        m.querySelector(".modal-caja h2").textContent = `${ses ? "Sesión" : "Seguimiento"} · ${p.Nombre}`;
      };
      m.querySelector("[name=Tipo_Sesion]").addEventListener("change", aplicarTipo);
      aplicarTipo();
      const caja = $("#sug-sem");
      let sugerido = "";
      const sug = () => {
        const real = m.querySelector("[name=Avance_Real]").value, plan = m.querySelector("[name=Avance_Planeado]").value;
        let html = "";
        sugerido = "";
        const ejec = m.querySelector("[name=Ejecutado]").value;
        if (real !== "" && plan !== "") {
          const v = R.valorGanado(bacDe(p), real, plan, ejec);
          const s = R.semaforoSugerido(real, plan, v ? v.CPI : null);
          sugerido = s.semaforo;
          const actual = m.querySelector("[name=Semaforo]").value;
          const texto = s.desviacion > 0 ? `${s.desviacion} pts por debajo del plan` : s.desviacion < 0 ? `${-s.desviacion} pts por encima del plan` : "igual al plan";
          const costo = v && v.CPI !== null ? ` · costo: CPI ${v.CPI.toFixed(2).replace(".", ",")}${s.porCosto !== "Verde" ? " (sobrecosto)" : ""}` : "";
          html = `Semáforo sugerido: ${chipSemaforo(s.semaforo)} (cronograma: ${texto}${costo}).${actual !== s.semaforo ? ` <button type="button" class="btn enlace" data-usar-sug>Usar sugerido</button>` : " Coincide con el elegido."}`;
        }
        const avisos = [];
        if (ult && real !== "" && Number(real) < (Number(ult.Avance_Real) || 0)) avisos.push(`El avance real bajó frente al último reporte (${pct(ult.Avance_Real)}).`);
        if (ejec !== "" && ult && ult.Ejecutado !== "" && ult.Ejecutado != null && Number(ejec) < Number(ult.Ejecutado)) avisos.push(`El ejecutado es acumulado y quedó por debajo del último reporte (${cop(ult.Ejecutado)}).`);
        if (ejec !== "" && bacDe(p) && Number(ejec) > bacDe(p)) avisos.push(`El ejecutado supera el presupuesto (${cop(bacDe(p))}): evalúa un control de cambios de costo.`);
        if (real !== "" && Number(real) >= 100 && p.Fase !== "Cierre") avisos.push("Avance del 100%: considera pasar el proyecto a la fase de Cierre.");
        const fc = m.querySelector("[name=Fecha_Corte]").value || R.hoyISO();
        if (p.Fecha_Inicio && p.Fecha_Fin_Plan && p.Fecha_Fin_Plan > p.Fecha_Inicio) {
          const lin = Math.max(0, Math.min(100, Math.round((R.difDias(fc, p.Fecha_Inicio) / R.difDias(p.Fecha_Fin_Plan, p.Fecha_Inicio)) * 100)));
          html += `<div class="sub">Referencia: por fechas (reparto lineal) a ${fecha(fc)} llevaría ~${lin}%. Usa el % de tu cronograma si lo tienes.</div>`;
        }
        if (avisos.length) html += `<div class="aviso-seg">⚠ ${avisos.map(esc).join("<br>⚠ ")}</div>`;
        if (caja.innerHTML !== html) caja.innerHTML = html;   // no redibujar si no cambió (evita perder el clic)
      };
      // Delegado en la caja: el botón puede redibujarse sin perder el evento.
      caja.addEventListener("mousedown", (ev) => { if (ev.target.closest("[data-usar-sug]")) ev.preventDefault(); });
      caja.addEventListener("click", (ev) => {
        if (!ev.target.closest("[data-usar-sug]") || !sugerido) return;
        m.querySelector("[name=Semaforo]").value = sugerido;
        sug();
      });
      ["Avance_Real", "Avance_Planeado", "Semaforo", "Ejecutado", "Fecha_Corte"].forEach((k) => ["input", "change"].forEach((t) => m.querySelector(`[name=${k}]`).addEventListener(t, sug)));
      sug();
      // Cronograma: usar el vigente o cargar una versión nueva que quedará asociada a esta sesión.
      const caja2 = m.querySelector("#crono-seg");
      const fCorte = () => m.querySelector("[name=Fecha_Corte]").value || R.hoyISO();
      const usar = (c) => { m.querySelector("[name=Avance_Real]").value = c.real; m.querySelector("[name=Avance_Planeado]").value = c.plan; sug(); };
      const pintarCrono = () => {
        const vig = versionVigente(p.ID_Proyecto), ts = tareasDe(p.ID_Proyecto);
        if (cronoNuevo) {
          const c = calcCronograma(cronoNuevo.res.tareas, fCorte()), corr = vig && vig.Fin && c.fin ? R.difDias(c.fin, vig.Fin) : null;
          const num = (versionesDe(p.ID_Proyecto).reduce((mx, v) => Math.max(mx, Number(v.Numero) || 0), 0)) + 1;
          caja2.innerHTML = `<span>📅 <b>${esc(cronoNuevo.archivo)}</b> se guardará como <span class="pill vnum">V${num}</span> de esta sesión: real <b>${c.real}%</b> · planeado al corte <b>${c.plan}%</b> · fin ${fecha(c.fin)}${corr ? ` <span class="gt-dif ${corr > 0 ? "mal" : "bien"}">${corr > 0 ? "+" : ""}${corr} d vs V${esc(vig.Numero)}</span>` : ""}${c.atrasadas.length ? ` · <span class="texto-rojo">${c.atrasadas.length} atrasada(s)</span>` : ""}.</span>
            <span class="acciones"><button type="button" class="btn chico" data-cs="usar">Usar estos valores</button><button type="button" class="btn chico" data-cs="quitar">Quitar</button></span>`;
        } else if (ts.length) {
          const c = calcCronograma(ts, fCorte()), dias = R.difDias(R.hoyISO(), (vig && vig.Fecha_Importacion) || p.Cronograma_Fecha || R.hoyISO());
          caja2.innerHTML = `<span>📅 Cronograma vigente ${vig ? `<span class="pill vnum">V${esc(vig.Numero)}</span>` : ""} (cargado el ${fecha((vig && vig.Fecha_Importacion) || p.Cronograma_Fecha)}): real <b>${c.real}%</b> · planeado al corte <b>${c.plan}%</b>${c.atrasadas.length ? ` · <span class="texto-rojo">${c.atrasadas.length} atrasada(s)</span>` : ""}.${dias > 7 ? ` <span class="sub">Tiene ${dias} días: si cambió, carga la versión nueva.</span>` : ""}</span>
            <span class="acciones"><button type="button" class="btn chico" data-cs="usar">Usar estos valores</button><button type="button" class="btn chico primario" data-cs="cargar">⇪ Cargar versión nueva</button></span>`;
        } else {
          caja2.innerHTML = `<span>📅 ¿Tienes el cronograma en Planner (Excel) o MS Project (XML)? Cárgalo y el avance real y planeado se llenan solos; queda como versión asociada a esta sesión.</span>
            <span class="acciones"><button type="button" class="btn chico primario" data-cs="cargar">⇪ Cargar cronograma</button></span>`;
        }
      };
      caja2.addEventListener("click", (ev) => {
        const a = ev.target.closest("[data-cs]"); if (!a) return;
        if (a.dataset.cs === "cargar") elegirCronograma((res0, archivo) => { const res = { ...res0, tareas: conLineaBase(p.ID_Proyecto, res0.tareas, traeBase(res0.tareas) ? "archivo" : "mantener") }; cronoNuevo = { res, archivo }; usar(calcCronograma(res.tareas, fCorte())); pintarCrono(); });
        if (a.dataset.cs === "usar") usar(calcCronograma(cronoNuevo ? cronoNuevo.res.tareas : tareasDe(p.ID_Proyecto), fCorte()));
        if (a.dataset.cs === "quitar") { cronoNuevo = null; pintarCrono(); }
      });
      m.querySelector("[name=Fecha_Corte]").addEventListener("change", pintarCrono);
      pintarCrono();
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
            (res.compromisos.length ? `${msgActa(res)} y los cargué abajo${sinFecha ? ` (${sinFecha} sin fecha en el acta)` : ""}; revísalos antes de guardar.` : `${msgActa(res)}; escríbelos abajo si hay.`);
        } catch (e) {
          acta = null;
          est.className = "acta-estado error"; est.textContent = "No se pudo leer el acta: " + e.message;
        }
      }));
    };
    let acta = null, cronoNuevo = null;
    const recoger = (m, fd) => {
      const nuevos = [];
      for (const f of m.querySelectorAll(".comp-fila")) {
        const t = f.querySelector(".c-texto").value.trim(), r = f.querySelector(".c-resp").value.trim(), d = f.querySelector(".c-fecha").value;
        if (!t && !r && !d) continue;
        if (d && fd.Fecha_Corte && d < fd.Fecha_Corte) return { error: `La fecha del compromiso «${t || "sin descripción"}» no puede ser anterior a la fecha de corte.` };
        const inp = f.querySelector(".c-resp");
        nuevos.push({ Compromiso: t || "(sin descripción)", Responsable: r, Fecha_Compromiso: d, Estado: f.querySelector(".c-estado").value, _nuevos: inp._nuevos || new Set() });
      }
      const cambiosEstado = [...m.querySelectorAll(".cp-estado")].map((sel) => ({ c: abiertos.find((x) => x.ID_Compromiso === sel.dataset.cp), estado: sel.value })).filter((x) => x.c && x.estado !== R.estadoBase(x.c));
      return { datos: { nuevos, cambiosEstado } };
    };
    modal(`Seguimiento o sesión · ${p.Nombre}`, campos, {}, (fd, extra) => {
      if (!fd.Fecha_Corte) fd.Fecha_Corte = R.hoyISO();
      const esSes = fd.Tipo_Sesion === "Sesión";
      if (esSes) { fd.Avance_Real = ""; fd.Avance_Planeado = ""; fd.Semaforo = ""; fd.Ejecutado = ""; }
      const idSeg = R.siguienteIdHijo("SEG", S.datos.Seguimientos, "ID_Seguimiento", p.ID_Proyecto);
      const fila = {
        ID_Seguimiento: idSeg, ID_Proyecto: p.ID_Proyecto, Fecha_Corte: fd.Fecha_Corte, Semana: R.semanaISO(fd.Fecha_Corte),
        Avance_Real: fd.Avance_Real, Avance_Planeado: fd.Avance_Planeado, Semaforo: fd.Semaforo, Ejecutado: fd.Ejecutado, Logros: fd.Logros, Proximos_Pasos: fd.Proximos_Pasos,
        Bloqueos: fd.Bloqueos, Reportado_Por: S.usuario.Correo, Fecha_Acta: fd.Fecha_Acta, URL_Acta: fd.URL_Acta,
        Acta_Archivo: "",   // se marca solo cuando el PDF quedó completo en Excel
        Tipo_Sesion: esSes ? "Sesión" : "Seguimiento", Nombre_Sesion: fd.Nombre_Sesion || "",
      };
      const actaElegida = acta, cronoElegido = esSes ? null : cronoNuevo;
      const comps = extra.nuevos.map((c, i) => ({
        ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", p.ID_Proyecto, i),
        ID_Proyecto: p.ID_Proyecto, ID_Seguimiento: idSeg, ...c, Correo_Responsable: "", ID_Recurso: "", Estado: c.Estado || "Pendiente", Fecha_Cierre: c.Estado === "Cerrado" ? fd.Fecha_Corte : "", Registrado_Por: S.usuario.Correo,
      }));
      guardar(async () => {
        await S.api.agregarFila("Seguimientos", fila);
        if (!esSes) await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}),
          ...(fd.Avance_Planeado !== "" ? { Avance_Planeado: fd.Avance_Planeado } : {}),
          ...(fd.Ejecutado !== "" ? { Ejecutado: fd.Ejecutado } : {}),
          ...(fd.Semaforo ? { Semaforo: fd.Semaforo } : {}),
          ...(fd.Bloqueos || fd.Logros ? { Comentario_Estado: fd.Bloqueos ? `Bloqueo: ${fd.Bloqueos}` : fd.Logros } : {}), ...sello(),
        });
        for (const c of comps) { const nv = c._nuevos; delete c._nuevos; Object.assign(c, await resolverResponsables(separarNombres(c.Responsable).map((n) => ({ Nombre: n })), nv)); }
        if (comps.length) await S.api.agregarFilas("Compromisos", comps);
        if (extra.cambiosEstado.length) {
          await S.api.actualizarVarios("Compromisos", "ID_Compromiso", extra.cambiosEstado.map(({ c, estado }) => ({ id: c.ID_Compromiso, cambios: { Estado: estado, Fecha_Cierre: estado === "Cerrado" ? fd.Fecha_Corte : "" } })));
          const txt = { Cerrado: "Cerrado", "En curso": "Pasó a En curso", Pendiente: "Quedó Pendiente" };
          await S.api.agregarFilas("Comentarios", extra.cambiosEstado.map(({ c, estado }) => nuevoComentario(c, `${txt[estado] || estado} en la sesión de seguimiento del ${fecha(fd.Fecha_Corte)}.`)));
        }
        if (cronoElegido) {
          try { await guardarVersionCronograma(p, cronoElegido.res, cronoElegido.archivo, { corte: fd.Fecha_Corte, idSeg, avance: false, hitos: true }); }
          catch (e) { toast("El seguimiento quedó registrado, pero el cronograma no se pudo guardar: " + e.message + " Cárgalo en la pestaña Cronograma.", true); }
        }
        if (actaElegida) {
          try { await subirActa(idSeg, actaElegida.bytes, actaElegida.nombre); await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", idSeg, { Acta_Archivo: actaElegida.nombre }); }
          catch (e) { toast("La sesión quedó registrada, pero el acta no se pudo guardar: " + e.message + " Usa «Adjuntar acta» en la sesión.", true); }
        }
      }, `${esSes ? "Sesión registrada" : "Seguimiento registrado"}${comps.length ? ` con ${comps.length} compromiso(s)` : ""}${actaElegida ? " y acta guardada" : ""}${cronoElegido ? " · cronograma versionado" : ""}`);
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
      else if (fd.Estado === "Cumplido") fd.Fecha_Real = R.hoyISO();
      if (h) guardar(() => S.api.actualizarPorId("Hitos", "ID_Hito", h.ID_Hito, fd));
      else guardar(() => S.api.agregarFila("Hitos", { ID_Hito: R.siguienteIdHijo("HIT", S.datos.Hitos, "ID_Hito", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }), "Hito agregado");
    }, (fd) => (!String(fd.Hito || "").trim() ? "Escribe el nombre del hito." : !fd.Fecha_Plan ? "La fecha planeada es obligatoria." : fd.Fecha_Real && fd.Fecha_Real > R.hoyISO() ? "La fecha real no puede ser futura." : ""),
    h ? { eliminar: { texto: "Eliminar hito", mensaje: `Se borrará el hito «${h.Hito}».`, accion: () => guardar(() => S.api.eliminarFilas("Hitos", "ID_Hito", [h.ID_Hito]), "Hito eliminado") } } : {});
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
    }, (fd) => {
      if (!String(fd.Descripcion || "").trim()) return "Describe el riesgo.";
      const ci = R.calificacion(fd.Probabilidad_Inherente, fd.Impacto_Inherente), cr = R.calificacion(fd.Probabilidad_Residual, fd.Impacto_Residual);
      if (fd.Tipo !== "Oportunidad" && ci && cr && cr > ci) return `La calificación residual (${cr}) no puede ser mayor que la inherente (${ci}): mitigar debe reducir el riesgo.`;
      return "";
    }, r ? { init, eliminar: { texto: "Eliminar riesgo", mensaje: `Se borrará el riesgo «${r.Descripcion}».`, accion: () => guardar(() => S.api.eliminarFilas("Riesgos", "ID_Riesgo", [r.ID_Riesgo]), "Riesgo eliminado") } } : { init });
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
    const actas = segs.filter((x) => x.Acta_Archivo).map((x) => x.ID_Seguimiento);
    if (actas.length) await S.api.borrarArchivos(actas);
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
    await borrar("Documentos", "ID_Documento", (S.datos.Documentos || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("LineasBase", "ID_LB", (S.datos.LineasBase || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Tareas", "ID_Tarea", (S.datos.Tareas || []).filter((x) => x.ID_Proyecto === pid));
    const archCrono = (S.datos.Cronogramas || []).filter((x) => x.ID_Proyecto === pid && x.Archivo_Guardado === "Sí").map((x) => x.ID_Version);
    if (archCrono.length) await S.api.borrarArchivos(archCrono);
    await borrar("Cronogramas", "ID_Version", (S.datos.Cronogramas || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Agenda", "ID_Sesion", (S.datos.Agenda || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Mediciones", "ID_Medicion", (S.datos.Mediciones || []).filter((x) => x.ID_Proyecto === pid));
    await borrar("Beneficios", "ID_Beneficio", (S.datos.Beneficios || []).filter((x) => x.ID_Proyecto === pid));
    await S.api.eliminarFilas("Proyectos", "ID_Proyecto", [pid]);
  }

  // Responsables del compromiso: uno o varios, del directorio de Recursos; desde aquí también se crean y editan los stakeholders del proyecto.
  function selectorResponsables(m, pid, inicial) {
    const cont = m.querySelector("#fc-resp");
    const elegidos = inicial.map((x) => ({ Nombre: x.Nombre, Correo: x.Correo || "", ID_Recurso: x.ID_Recurso || "" }));
    const yaEsta = (x) => elegidos.some((e) => (x.ID_Recurso && e.ID_Recurso === x.ID_Recurso) || normTxt(e.Nombre) === normTxt(x.Nombre));
    const agregar = (x) => { if (!String(x.Nombre || "").trim() || yaEsta(x)) return; elegidos.push(x); pintar(); };
    const deStk = (st) => { const r = recursoPor({ id: st.ID_Recurso, correo: st.Correo, nombre: st.Nombre }); return { Nombre: r ? r.Nombre : st.Nombre, Correo: st.Correo || (r && r.Correo) || "", ID_Recurso: r ? r.ID_Recurso : "" }; };
    const opcionesRol = (sel) => (S.cat.Rol_Stakeholder || []).map((r) => `<option ${r === sel ? "selected" : ""}>${esc(r)}</option>`).join("");
    cont.innerHTML = `<span class="etq">Responsables <span class="sub">(uno o varios, del directorio de Recursos)</span></span>
      <div class="rs-chips" id="rs-chips"></div>
      <input id="rs-buscar" data-recurso autocomplete="off" placeholder="+ Agregar responsable: busca en Recursos o escribe un nombre nuevo" aria-label="Agregar responsable">
      <div class="rs-sug" id="rs-sug"></div>
      <details class="rs-stk" id="rs-stk"><summary>Stakeholders del proyecto <span class="contador" id="rs-n"></span> · crear, cambiar rol o asignar</summary>
        <ul class="eq-lista" id="rs-lista"></ul>
        <div class="acciones"><button type="button" class="btn chico" id="rs-nuevo">+ Nuevo stakeholder</button><button type="button" class="btn chico" id="rs-rol-abrir">+ Crear rol nuevo</button></div>
        <div id="rs-rol-form" class="eq-form" hidden><input id="rs-rol-nombre" placeholder="Nombre del rol (ej.: Líder técnico, QA)" aria-label="Nuevo rol">
          <button type="button" class="btn primario chico" id="rs-rol-crear">Crear rol</button><button type="button" class="btn chico" id="rs-rol-cancelar">Cancelar</button></div>
        <div id="rs-form" class="eq-form" hidden><div class="pn-grid">
          <select id="rs-f-rol" aria-label="Rol">${opcionesRol()}</select>
          <input id="rs-f-nombre" placeholder="Persona * (elige del directorio)" aria-label="Persona" autocomplete="off">
          <input id="rs-f-correo" type="email" placeholder="Correo" aria-label="Correo"></div>
          <div class="sub" id="rs-f-rec"></div>
          <label class="check"><input type="checkbox" id="rs-f-resp" checked> Dejarlo también como responsable de este compromiso</label>
          <div class="acciones"><button type="button" class="btn primario chico" id="rs-f-guardar">Agregar al proyecto</button><button type="button" class="btn chico" id="rs-f-cancelar">Cancelar</button></div></div>
        <div class="error-campo" id="rs-error"></div>
      </details>`;
    const $c = (id) => cont.querySelector(id);
    const err = $c("#rs-error");
    function pintar() {
      $c("#rs-chips").innerHTML = elegidos.map((x, i) => { if (esGrupo(x.Nombre)) { const g = integrantesGrupo(pid);
          return `<span class="rs-chip grupo" title="${esc(g.map((y) => `${y.Nombre}${y.Rol ? ` (${y.Rol})` : ""}`).join(", "))}"><span class="avatar">👥</span><span><b>${GRUPO_PROY}</b><br><span class="sub">${g.length} stakeholder(s) del proyecto</span></span><button type="button" class="btn enlace" data-rs-quitar="${i}" aria-label="Quitar el grupo">✕</button></span>`; }
        const rol = rolEnProyecto(pid, x);
        return `<span class="rs-chip">${avatar(x.Nombre)}<span><b>${esc(x.Nombre)}</b>${rol ? ` <span class="sub">· ${esc(rol)}</span>` : ""}<br><span class="sub ${x.Correo ? "" : "baja"}">${esc(x.Correo || "sin correo")}</span></span><button type="button" class="btn enlace" data-rs-quitar="${i}" aria-label="Quitar a ${esc(x.Nombre)}">✕</button></span>`; }).join("")
        || `<span class="sub">Sin responsable asignado.</span>`;
      cont.querySelectorAll("[data-rs-quitar]").forEach((b) => b.addEventListener("click", () => { elegidos.splice(Number(b.dataset.rsQuitar), 1); pintar(); }));
      const stks = stakeholdersDe(pid);
      const libres = stks.filter((st) => !yaEsta(deStk(st)));
      const grupoBtn = stks.length && !elegidos.some((x) => esGrupo(x.Nombre)) ? `<button type="button" class="btn chico" id="rs-grupo">+ 👥 ${GRUPO_PROY}</button> ` : "";
      $c("#rs-sug").innerHTML = libres.length || grupoBtn ? `<span class="sub">Del proyecto:</span> ${grupoBtn}${libres.map((st) => `<button type="button" class="btn chico" data-rs-stk="${esc(st.ID_Stakeholder)}">+ ${esc(st.Nombre)}${st.Rol ? ` · ${esc(st.Rol)}` : ""}</button>`).join(" ")}` : "";
      $c("#rs-n").textContent = stks.length;
      $c("#rs-lista").innerHTML = stks.map((st) => `<li><b>${esc(st.Nombre)}</b> <span class="sub">${esc(st.Correo || "")}</span>
          <select data-rs-rol="${esc(st.ID_Stakeholder)}" aria-label="Rol de ${esc(st.Nombre)}">${(S.cat.Rol_Stakeholder || []).includes(st.Rol) ? "" : `<option selected>${esc(st.Rol || "")}</option>`}${opcionesRol(st.Rol)}</select>
          ${yaEsta(deStk(st)) ? `<span class="pill encurso">responsable</span>` : `<button type="button" class="btn chico" data-rs-stk="${esc(st.ID_Stakeholder)}">+ Responsable</button>`}</li>`).join("")
        || `<li class="sub">Este proyecto aún no tiene stakeholders. Crea el primero con «+ Nuevo stakeholder».</li>`;
      const gb = $c("#rs-grupo"); if (gb) gb.addEventListener("click", () => agregar({ Nombre: GRUPO_PROY, Correo: "", ID_Recurso: "" }));
      cont.querySelectorAll("[data-rs-stk]").forEach((b) => b.addEventListener("click", () => { const st = stks.find((x) => x.ID_Stakeholder === b.dataset.rsStk); if (st) agregar(deStk(st)); }));
      cont.querySelectorAll("[data-rs-rol]").forEach((sel) => sel.addEventListener("change", async () => {
        const st = stks.find((x) => x.ID_Stakeholder === sel.dataset.rsRol);
        if (!st || sel.value === st.Rol) return;
        cargando(true, "Guardando rol…");
        try { await S.api.actualizarPorId("Stakeholders", "ID_Stakeholder", st.ID_Stakeholder, { Rol: sel.value }); toast(`${st.Nombre}: rol «${sel.value}»`); pintar(); }
        catch (e) { err.textContent = "No se pudo cambiar el rol: " + e.message; } finally { cargando(false); }
      }));
    }
    pintar();
    cont.addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target.tagName === "INPUT") e.preventDefault(); });
    // Buscar en el directorio
    const buscar = $c("#rs-buscar");
    buscar._pid = pid;
    buscar._excluir = () => new Set(elegidos.map((x) => normTxt(x.Nombre)));
    buscar.addEventListener("elegido", (e) => {
      const { r, nuevo, grupo } = e.detail;
      if (grupo) agregar({ Nombre: GRUPO_PROY, Correo: "", ID_Recurso: "" });
      else if (r) agregar({ Nombre: r.Nombre, Correo: r.Correo || "", ID_Recurso: r.ID_Recurso });
      else if (nuevo) agregar({ Nombre: nuevo, Correo: correoSuelto(nuevo), ID_Recurso: "" });
      buscar.value = ""; buscar.focus();
    });
    buscar.addEventListener("keydown", (e) => { if (e.key === "Enter") e.preventDefault(); });
    // Crear rol
    $c("#rs-rol-abrir").addEventListener("click", () => { $c("#rs-rol-form").hidden = false; $c("#rs-rol-nombre").focus(); });
    $c("#rs-rol-cancelar").addEventListener("click", () => { $c("#rs-rol-form").hidden = true; });
    $c("#rs-rol-crear").addEventListener("click", async () => {
      const rol = $c("#rs-rol-nombre").value.trim();
      if (!rol) { err.textContent = "Escribe el nombre del rol."; return; }
      if ((S.cat.Rol_Stakeholder || []).some((r) => lc(r) === lc(rol))) { err.textContent = "Ese rol ya existe."; return; }
      cargando(true, "Creando rol…");
      try {
        await S.api.agregarFila("Catalogos", { Lista: "Rol_Stakeholder", Valor: rol }); armarCatalogos();
        $c("#rs-f-rol").innerHTML = opcionesRol(rol); $c("#rs-rol-nombre").value = ""; $c("#rs-rol-form").hidden = true; err.textContent = "";
        pintar(); toast(`Rol «${rol}» creado`);
      } catch (e) { err.textContent = "No se pudo crear el rol: " + e.message; } finally { cargando(false); }
    });
    // Nuevo stakeholder del proyecto
    selectorPersona(m, "#rs-f-nombre", { Correo: "#rs-f-correo" }, "#rs-f-rec");
    $c("#rs-nuevo").addEventListener("click", () => { $c("#rs-form").hidden = false; $c("#rs-f-nombre").focus(); });
    $c("#rs-f-cancelar").addEventListener("click", () => { $c("#rs-form").hidden = true; });
    $c("#rs-f-guardar").addEventListener("click", async () => {
      const nombre = $c("#rs-f-nombre").value.trim(), correo = $c("#rs-f-correo").value.trim(), rol = $c("#rs-f-rol").value;
      if (!nombre) { err.textContent = "Elige o escribe la persona."; return; }
      if (!rol) { err.textContent = "Elige el rol (o crea uno nuevo)."; return; }
      if (correo && !validoCorreo(correo)) { err.textContent = "Correo no válido."; return; }
      const r0 = recursoPor({ correo, nombre });
      if (stakeholdersDe(pid).some((st) => (r0 && st.ID_Recurso === r0.ID_Recurso) || normTxt(st.Nombre) === normTxt(nombre))) { err.textContent = "Esa persona ya es stakeholder del proyecto: cámbiale el rol en la lista."; return; }
      cargando(true, "Agregando stakeholder…");
      try {
        const r = await asegurarRecurso({ Nombre: nombre, Correo: correo });
        await S.api.agregarFila("Stakeholders", { ID_Stakeholder: R.siguienteIdHijo("STK", S.datos.Stakeholders || [], "ID_Stakeholder", pid), ID_Proyecto: pid, Rol: rol,
          Nombre: r.Nombre, Cargo: r.Cargo || "", Area: r.Area || "", Correo: correo || r.Correo || "", Telefono: r.Telefono || "", ID_Proveedor: r.ID_Proveedor || "", Dedicacion: "", ID_Recurso: r.ID_Recurso });
        if ($c("#rs-f-resp").checked) agregar({ Nombre: r.Nombre, Correo: correo || r.Correo || "", ID_Recurso: r.ID_Recurso });
        ["#rs-f-nombre", "#rs-f-correo"].forEach((id) => { $c(id).value = ""; }); $c("#rs-f-rec").textContent = ""; $c("#rs-form").hidden = true; err.textContent = "";
        pintar(); toast(`${r.Nombre} agregado como ${rol}`);
      } catch (e) { err.textContent = "No se pudo agregar: " + e.message; } finally { cargando(false); }
    });
    return () => elegidos.slice();
  }

  // Cambiar solo el/los responsable(s) de un compromiso, eligiendo del directorio de Recursos.
  const celdaResp = (c, puede) => `${esc(c.Responsable || "—")}${puede ? ` <button class="btn enlace cambiar-resp" data-cambiar-resp="${esc(c.ID_Compromiso)}" title="Cambiar responsable" aria-label="Cambiar responsable de ${esc(c.Compromiso)}">✎</button>` : ""}`;
  function formResponsables(c, alTerminar) {
    let leerResp = () => [];
    const antes = responsablesDe(c, false).map((x) => x.Nombre).join(", ") || "sin responsable";
    modal(`Responsable · ${c.Compromiso}`, [{ html: `<div id="fc-resp" class="fc-resp"></div>` }], {}, (fd, resp) => {
      guardar(async () => {
        const cambios = await resolverResponsables(resp, true);
        await S.api.actualizarPorId("Compromisos", "ID_Compromiso", c.ID_Compromiso, cambios);
        const ahora = partesLista(cambios.Responsable).filter(Boolean).join(", ") || "sin responsable";
        if (ahora !== antes) await S.api.agregarFila("Comentarios", nuevoComentario(c, `Responsable cambiado: ${antes} → ${ahora}.`));
      }, "Responsable actualizado").then(() => { if (alTerminar) alTerminar(); });
    }, null, {
      init: (m) => { leerResp = selectorResponsables(m, c.ID_Proyecto, responsablesDe(c, false)); setTimeout(() => m.querySelector("#rs-buscar").focus(), 0); },
      recoger: (m) => {
        const pend = m.querySelector("#rs-buscar").value.trim();
        if (pend) return { error: `Elige «${pend}» de la lista (o bórralo) antes de guardar.` };
        return { datos: leerResp() };
      },
    });
  }
  document.addEventListener("click", (e) => {
    const b = e.target.closest && e.target.closest("[data-cambiar-resp]");
    if (!b) return;
    e.preventDefault(); e.stopPropagation();
    const c = S.datos.Compromisos.find((x) => x.ID_Compromiso === b.dataset.cambiarResp);
    if (!c) return;
    const desdeDetalle = !!b.closest("#modal");
    formResponsables(c, desdeDetalle ? () => detalleCompromiso(c.ID_Compromiso) : null);
  }, true);

  function formCompromiso(c, pNuevo, segSugerida) {
    const nuevo = !c;
    const pid = nuevo ? pNuevo.ID_Proyecto : c.ID_Proyecto;
    const sesiones = R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse()
      .map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)}${s.Fecha_Acta && s.Fecha_Acta !== s.Fecha_Corte ? ` (acta ${fecha(s.Fecha_Acta)})` : ""}`]);
    let leerResp = () => [];
    const extra = {
      init: (m) => { leerResp = selectorResponsables(m, pid, nuevo ? [] : responsablesDe(c, false)); },
      recoger: (m) => {
        const pend = m.querySelector("#rs-buscar").value.trim();
        if (pend) return { error: `Elige «${pend}» de la lista de responsables (o bórralo) antes de guardar.` };
        return { datos: leerResp() };
      },
    };
    const campos = [
      { k: "Compromiso", label: "Compromiso", tipo: "textarea", placeholder: "Qué se hará" },
      { k: "ID_Seguimiento", label: "¿De qué sesión salió?", tipo: "select", opciones: sesiones, ancho: true, textoVacio: "Sin sesión (solo del proyecto)",
        ayuda: sesiones.length ? "Elige la sesión de seguimiento donde se acordó, o déjalo sin sesión." : "Este proyecto aún no tiene sesiones; quedará atado solo al proyecto." },
      { html: `<div id="fc-resp" class="fc-resp"></div>` },
      { k: "Fecha_Compromiso", label: "Fecha límite", tipo: "date" },
      { k: "Dias_Alerta", label: "Avisar con (días de anticipación)", tipo: "number", min: 0, max: 60, def: R.DIAS_COMPROMISO_POR_VENCER, ayuda: "Desde ese día aparece en «Por vencer» y en la campana 🔔." },
      ...(nuevo ? [{ k: "Comentario", label: "Comentario inicial (opcional)", tipo: "textarea", placeholder: "Contexto o primer avance" }] : [
        { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Compromiso, def: "Pendiente" },
        { k: "Fecha_Creacion", label: "Fecha de creación", tipo: "date", ayuda: "Se pone sola al crearlo; corrígela si lo registraste después." },
        { k: "Fecha_Cierre", label: "Fecha de cierre", tipo: "date", ayuda: "Se llena sola al marcarlo cumplido; se borra si lo reabres." }]),
    ];
    if (nuevo) {
      modal(`Nuevo compromiso · ${pNuevo.Nombre}`, campos, { ID_Seguimiento: segSugerida || "" }, (fd, resp) => {
        const fila = { ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", pNuevo.ID_Proyecto), ID_Proyecto: pNuevo.ID_Proyecto, ID_Seguimiento: fd.ID_Seguimiento || "",
          Compromiso: fd.Compromiso || "(sin descripción)", Dias_Alerta: fd.Dias_Alerta === "" ? "" : fd.Dias_Alerta, Fecha_Compromiso: fd.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo };
        guardar(async () => {
          Object.assign(fila, await resolverResponsables(resp, true));
          await S.api.agregarFila("Compromisos", fila);
          if (fd.Comentario) await S.api.agregarFila("Comentarios", nuevoComentario(fila, fd.Comentario));
        }, "Compromiso agregado");
      }, null, extra);
      return;
    }
    const coms = comentariosDe(c.ID_Compromiso);
    modal("Editar compromiso", campos, { ...c, Estado: R.estadoBase(c), Fecha_Creacion: fechaCreacion(c) }, (fd, resp) => {
      if (fd.Estado === "Cerrado" && !fd.Fecha_Cierre) fd.Fecha_Cierre = R.hoyISO();
      if (fd.Estado !== "Cerrado") fd.Fecha_Cierre = "";
      const cambioEstado = fd.Estado !== R.estadoBase(c);
      guardar(async () => {
        Object.assign(fd, await resolverResponsables(resp, true));
        await S.api.actualizarPorId("Compromisos", "ID_Compromiso", c.ID_Compromiso, fd);
        if (cambioEstado) await S.api.agregarFila("Comentarios", nuevoComentario(c, TEXTO_ESTADO[fd.Estado] || `Estado: ${fd.Estado}.`));
      }, "Compromiso actualizado");
    }, null, { ...extra, eliminar: { texto: "Eliminar compromiso", mensaje: `Se borrará «${c.Compromiso}»${coms.length ? ` con sus ${coms.length} comentario(s)` : ""}.`,
      accion: () => guardar(async () => {
        await borrarAdjuntos(coms);
        if (coms.length) await S.api.eliminarFilas("Comentarios", "ID_Comentario", coms.map((x) => x.ID_Comentario));
        await S.api.eliminarFilas("Compromisos", "ID_Compromiso", [c.ID_Compromiso]);
      }, "Compromiso eliminado") } });
  }

  // Fila rápida de compromiso (registrar o editar una sesión).
  const filaComp = (c = {}) => `<div class="comp-fila">
      <textarea class="c-texto" rows="1" placeholder="Compromiso (qué se hará)" aria-label="Compromiso">${esc(String(c.Compromiso || "").replace(/\s*\n+\s*/g, " "))}</textarea>
      <input class="c-resp" placeholder="Responsable(s) de Recursos" title="Puedes elegir varias personas: se separan con «;»" aria-label="Responsables" data-recurso data-multi autocomplete="off" value="${esc(separarNombres(c.Responsable).join("; "))}">
      <input class="c-fecha" type="date" aria-label="Fecha límite" value="${esc(c.Fecha_Compromiso || "")}" title="${esc(c.Fecha_Texto && !c.Fecha_Compromiso ? `En el acta: ${c.Fecha_Texto}` : "")}">
      <select class="c-estado" aria-label="Estado">${R.ESTADOS_COMPROMISO.map((e) => `<option ${e === (c.Estado || "Pendiente") ? "selected" : ""}>${esc(e)}</option>`).join("")}</select>
      <button type="button" class="btn chico c-quitar" aria-label="Quitar compromiso">✕</button></div>`;
  function formEditarSeguimiento(s) {
    S.pidForm = s.ID_Proyecto;
    const p = proyecto(s.ID_Proyecto);
    const segs = R.avancesDe(s.ID_Proyecto, S.datos.Seguimientos);
    const esUltimo = segs.length && segs[segs.length - 1].ID_Seguimiento === s.ID_Seguimiento;
    const comps = S.datos.Compromisos.filter((c) => c.ID_Seguimiento === s.ID_Seguimiento).sort(porCreacion);
    const campos = [
      { k: "Tipo_Sesion", label: "Tipo", tipo: "select", opciones: [["Seguimiento", "Seguimiento de avance"], ["Sesión", "Sesión / reunión (sin avance)"]], def: "Seguimiento" },
      { k: "Nombre_Sesion", label: "Nombre de la sesión", placeholder: "Ej.: Comité de cambios" },
      { seccion: "1. Avance" },
      { k: "Fecha_Corte", label: "Fecha de corte / de la sesión", tipo: "date" },
      { k: "Avance_Real", label: "Avance real (%)", tipo: "number", min: 0, max: 100 },
      { k: "Avance_Planeado", label: "Avance planeado a la fecha (%)", tipo: "number", min: 0, max: 100 },
      { k: "Semaforo", label: "Semáforo", tipo: "select", opciones: S.cat.Semaforo },
      { k: "Ejecutado", label: "Presupuesto ejecutado a la fecha (COP)", tipo: "number", min: 0 },
      { seccion: "2. Qué pasó en el periodo" },
      { k: "Logros", label: "Logros del periodo", tipo: "textarea" },
      { k: "Proximos_Pasos", label: "Próximos pasos", tipo: "textarea" },
      { k: "Bloqueos", label: "Bloqueos", tipo: "textarea" },
      { seccion: "3. Acta de la sesión" },
      { k: "Fecha_Acta", label: "Fecha del acta", tipo: "date" },
      { k: "URL_Acta", label: "Enlace al acta (opcional)", tipo: "url", placeholder: "https://…" },
      { seccion: `4. Compromisos de esta sesión (${comps.length})`, ayuda: "edita aquí el texto, responsables, fecha y estado; ábrelos para comentar" },
      { html: `${comps.length ? `<div><table class="no-orden ses-comps"><thead><tr><th>Compromiso</th><th>Responsable(s)</th><th>Fecha límite</th><th>Estado</th><th></th></tr></thead><tbody>
        ${comps.map((c) => `<tr data-sc="${esc(c.ID_Compromiso)}"><td><textarea class="c-texto sc-texto" rows="1" aria-label="Compromiso">${esc(c.Compromiso || "")}</textarea>${comentariosDe(c.ID_Compromiso).length ? `<div class="sub">💬 ${comentariosDe(c.ID_Compromiso).length} comentario(s)</div>` : ""}</td>
          <td><input class="sc-resp" data-recurso data-multi autocomplete="off" aria-label="Responsables" placeholder="Responsable(s) de Recursos" value="${esc(responsablesDe(c, false).map((x) => x.Nombre).join("; "))}"></td>
          <td><input type="date" class="sc-fecha" value="${esc(c.Fecha_Compromiso || "")}" aria-label="Fecha límite"></td>
          <td><select class="sc-estado" aria-label="Estado">${R.ESTADOS_COMPROMISO.map((e) => `<option ${e === R.estadoBase(c) ? "selected" : ""}>${esc(e)}</option>`).join("")}</select>${R.estadoCompromiso(c) === "Vencido" ? ` ${pill("Vencido")}` : ""}</td>
          <td class="derecha"><button type="button" class="btn chico" data-sc-abrir="${esc(c.ID_Compromiso)}">Abrir</button></td></tr>`).join("")}</tbody></table></div>` : `<p class="sub">Esta sesión aún no tiene compromisos.</p>`}
        <div id="ses-nuevos"></div><button type="button" class="btn chico" id="b-ses-comp">+ Agregar compromiso a esta sesión</button>` },
    ];
    const bloqueActa = `<div class="carga-acta">
      <div><b>Archivo del acta</b><div class="sub" id="acta-actual">${s.Acta_Archivo ? `Guardada: ${esc(s.Acta_Archivo)}` : "Esta sesión no tiene acta guardada."}</div></div>
      <div class="acciones">${s.Acta_Archivo ? `<button type="button" class="btn" id="b-ver-acta">Ver</button><button type="button" class="btn" id="b-quitar-acta">Quitar</button>` : ""}
        <button type="button" class="btn primario" id="b-reemplazar-acta">${s.Acta_Archivo ? "Reemplazar PDF" : "Subir PDF"}</button></div>
      <div id="acta-estado" class="acta-estado" aria-live="polite"></div></div>`;
    const despues = esUltimo ? `<p class="sub">Es la última sesión: al guardar también se actualiza el avance y el semáforo del proyecto.</p>` : "";
    let accionActa = null;   // null = sin cambio · { tipo: "subir", bytes, nombre } · { tipo: "quitar" }
    const init = (m) => {
      const lblE = (k) => { const x = m.querySelector(`[name=${k}]`); return x && x.closest("label"); };
      const aplicar = () => { const ses = m.querySelector("[name=Tipo_Sesion]").value === "Sesión"; ["Avance_Real", "Avance_Planeado", "Semaforo", "Ejecutado"].forEach((k) => { const l = lblE(k); if (l) l.hidden = ses; }); const n = lblE("Nombre_Sesion"); if (n) n.hidden = !ses; };
      m.querySelector("[name=Tipo_Sesion]").addEventListener("change", aplicar); aplicar();
      const est = $("#acta-estado");
      // Compromisos de la sesión: abrir uno y agregar nuevos.
      m.querySelectorAll("[data-sc-abrir]").forEach((b) => b.addEventListener("click", () => detalleCompromiso(b.dataset.scAbrir)));
      const contN = m.querySelector("#ses-nuevos");
      const enlazarN = () => contN.querySelectorAll(".c-quitar").forEach((b) => { b.onclick = () => b.parentElement.remove(); });
      m.querySelector("#b-ses-comp").addEventListener("click", () => { contN.insertAdjacentHTML("beforeend", filaComp()); enlazarN(); contN.lastElementChild.querySelector("input").focus(); });
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
          const nuevosActa = res.compromisos.filter((c) => !yaExisteComp(s.ID_Proyecto, c));
          nuevosActa.forEach((c) => contN.insertAdjacentHTML("beforeend", filaComp(c))); enlazarN();
          if (typeof enlazarRecursos === "function") enlazarRecursos(contN);
          est.textContent = `${archivo.name} se guardará al dar Guardar. ${msgActa(res)}${res.compromisos.length ? `: ${nuevosActa.length} nuevo(s) cargados abajo en «Compromisos nuevos» para revisar${res.compromisos.length - nuevosActa.length ? ` (${res.compromisos.length - nuevosActa.length} ya existían)` : ""}` : ""}.`;
          accionActa.compromisos = [];
        } catch (e) { est.className = "acta-estado error"; est.textContent = "No se pudo leer el acta: " + e.message; }
      }));
    };
    const recoger = (m) => {
      const cambiosC = [...m.querySelectorAll("tr[data-sc]")].map((tr) => {
        const c = comps.find((x) => x.ID_Compromiso === tr.dataset.sc);
        const estado = tr.querySelector(".sc-estado").value, f = tr.querySelector(".sc-fecha").value;
        const texto = tr.querySelector(".sc-texto").value.replace(/\s+/g, " ").trim(), inpR = tr.querySelector(".sc-resp");
        const resp = separarNombres(inpR.value);
        const ch = {};
        if (texto && texto !== String(c.Compromiso || "").trim()) ch.Compromiso = texto;
        if (resp.join("|") !== responsablesDe(c, false).map((x) => x.Nombre).join("|")) ch._resp = { nombres: resp, nuevos: inpR._nuevos || new Set() };
        if (estado !== R.estadoBase(c)) { ch.Estado = estado; ch.Fecha_Cierre = estado === "Cerrado" ? R.hoyISO() : ""; }
        if (f !== String(c.Fecha_Compromiso || "")) ch.Fecha_Compromiso = f;
        return Object.keys(ch).length ? { c, cambios: ch } : null;
      }).filter(Boolean);
      const nuevosC = [];
      for (const f of m.querySelectorAll("#ses-nuevos .comp-fila")) {
        const t = f.querySelector(".c-texto").value.trim(), inp = f.querySelector(".c-resp"), r = inp.value.trim(), d = f.querySelector(".c-fecha").value;
        if (!t && !r && !d) continue;
        if (!t) return { error: "Escribe la descripción del compromiso nuevo (o quita la fila)." };
        nuevosC.push({ Compromiso: t, Responsable: r, Fecha_Compromiso: d, Estado: f.querySelector(".c-estado").value, _nuevos: inp._nuevos || new Set() });
      }
      return { datos: { cambiosC, nuevosC } };
    };
    modal(`Editar sesión del ${fecha(s.Fecha_Corte)} · ${p ? p.Nombre : ""}`, campos, { ...s, Tipo_Sesion: s.Tipo_Sesion || "Seguimiento" }, (fd, extra) => {
      if (!fd.Fecha_Corte) fd.Fecha_Corte = s.Fecha_Corte || R.hoyISO();
      if (fd.Tipo_Sesion === "Sesión") { fd.Avance_Real = ""; fd.Avance_Planeado = ""; fd.Semaforo = ""; fd.Ejecutado = ""; }
      fd.Semana = R.semanaISO(fd.Fecha_Corte);
      const accion = accionActa;
      guardar(async () => {
        if (extra.cambiosC.length) {
          for (const x of extra.cambiosC) {
            if (!x.cambios._resp) continue;
            const { nombres, nuevos } = x.cambios._resp; delete x.cambios._resp;
            Object.assign(x.cambios, await resolverResponsables(nombres.map((n) => ({ Nombre: n })), nuevos));
          }
          await S.api.actualizarVarios("Compromisos", "ID_Compromiso", extra.cambiosC.map(({ c, cambios }) => ({ id: c.ID_Compromiso, cambios })));
          const coms = extra.cambiosC.filter((x) => x.cambios.Estado).map(({ c, cambios }) => nuevoComentario(c, `${TEXTO_ESTADO[cambios.Estado] || `Estado: ${cambios.Estado}.`} (desde la edición de la sesión del ${fecha(fd.Fecha_Corte)})`));
          if (coms.length) await S.api.agregarFilas("Comentarios", coms);
        }
        if (extra.nuevosC.length) {
          const filas = [];
          for (const [i, c] of extra.nuevosC.entries()) {
            const resp = await resolverResponsables(separarNombres(c.Responsable).map((n) => ({ Nombre: n })), c._nuevos);
            filas.push({ ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", s.ID_Proyecto, i), ID_Proyecto: s.ID_Proyecto, ID_Seguimiento: s.ID_Seguimiento,
              Compromiso: c.Compromiso, ...resp, Fecha_Compromiso: c.Fecha_Compromiso, Estado: c.Estado || "Pendiente", Fecha_Cierre: c.Estado === "Cerrado" ? R.hoyISO() : "", Registrado_Por: S.usuario.Correo });
          }
          await S.api.agregarFilas("Compromisos", filas);
        }
        const cambios = { ...fd };
        if (accion && accion.tipo === "quitar") { await S.api.borrarArchivo(s.ID_Seguimiento); cambios.Acta_Archivo = ""; }
        if (accion && accion.tipo === "subir") { await subirActa(s.ID_Seguimiento, accion.bytes, accion.nombre); cambios.Acta_Archivo = accion.nombre; }
        await S.api.actualizarPorId("Seguimientos", "ID_Seguimiento", s.ID_Seguimiento, cambios);
        if (accion && accion.compromisos && accion.compromisos.length) await S.api.agregarFilas("Compromisos", accion.compromisos.map((c, i) => ({
          ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", s.ID_Proyecto, i), ID_Proyecto: s.ID_Proyecto, ID_Seguimiento: s.ID_Seguimiento,
          Compromiso: c.Compromiso, Responsable: c.Responsable, Fecha_Compromiso: c.Fecha_Compromiso, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo })));
        if (esUltimo && p && fd.Tipo_Sesion !== "Sesión") await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}), ...(fd.Avance_Planeado !== "" && fd.Avance_Planeado !== undefined ? { Avance_Planeado: fd.Avance_Planeado } : {}), ...(fd.Semaforo ? { Semaforo: fd.Semaforo } : {}), ...(fd.Ejecutado !== "" ? { Ejecutado: fd.Ejecutado } : {}), ...sello() });
      }, "Sesión actualizada");
    }, (fd) => (fd.Fecha_Corte && fd.Fecha_Corte > R.hoyISO() ? "La fecha de corte no puede ser futura." : ""), {
      antes: bloqueActa, despues, init, recoger,
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

  // ---------- Sugerencias de mejora de la app ----------
  const TIPOS_SUG = ["Mejora", "Error", "Idea nueva"];
  const PRIOS_SUG = ["Alta", "Media", "Baja"];
  const ESTADOS_SUG = ["Nueva", "En análisis", "Planeada", "Hecha", "Descartada"];
  const gestionaSug = () => R.puede(S.usuario, "verTodo");
  const sugerencias = () => (S.datos.Sugerencias || []).slice().sort((a, b) => String(b.Fecha_Hora).localeCompare(String(a.Fecha_Hora)));
  // Dónde está el usuario cuando sugiere (módulo, proyecto y pestaña).
  function contextoActual() {
    const v = VISTAS.find((x) => x.id === S.vista);
    if (S.vista === "ficha") { const p = proyecto(S.pid); return { Modulo: "Ficha del proyecto", Contexto: `${p ? p.Nombre : ""} › ${NOMBRE_TAB[S.fichaTab] || "Resumen"}`, ID_Proyecto: p ? p.ID_Proyecto : "" }; }
    return { Modulo: v ? v.t : "General", Contexto: "", ID_Proyecto: "" };
  }
  const MODULOS_SUG = () => ["General", "Ficha del proyecto", ...VISTAS.map((v) => v.t)];
  function formSugerencia(sg) {
    const nuevo = !sg;
    const ctx = nuevo ? contextoActual() : sg;
    const yaAdj = nuevo ? [] : adjuntosDe(sg);
    const pend = [];
    const campos = [
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: TIPOS_SUG, def: "Mejora" },
      { k: "Prioridad", label: "Prioridad", tipo: "select", opciones: PRIOS_SUG, def: "Media" },
      { k: "Modulo", label: "Módulo", tipo: "select", opciones: MODULOS_SUG(), def: ctx.Modulo },
      { k: "Contexto", label: "Dónde exactamente (opcional)", placeholder: "Pestaña, botón o campo", def: ctx.Contexto },
      { k: "Descripcion", label: "¿Qué cambio necesitas?", tipo: "textarea", placeholder: "Describe la mejora o el error: qué pasa hoy y qué esperas. Puedes pegar aquí una captura con Ctrl+V." },
      { html: `<div class="adj-barra"><button type="button" class="btn chico" id="sg-adjuntar">📎 Adjuntar captura o archivo</button><span class="sub">o pega una imagen en la descripción con Ctrl+V · máx. 5 MB</span></div><div id="sg-pend" class="adj-pend"></div>${yaAdj.length ? `<div class="sub">Ya adjuntos: ${yaAdj.map((a) => esc(a.n)).join(", ")}</div>` : ""}` },
      { k: "URL_Referencia", label: "Enlace de referencia (opcional)", tipo: "url", placeholder: "https://…sharepoint.com/… (captura, documento o ejemplo)" },
    ];
    const init = (m) => {
      const err = m.querySelector("#m-error");
      const pintar = () => {
        m.querySelector("#sg-pend").innerHTML = pend.map((a, i) => `<span class="adj-chip">${esImagen(a.tipo) ? `<img src="${a.url}" alt="">` : "📄"} ${esc(a.nombre)} <span class="sub">${tamano(a.tam)}</span><button type="button" class="btn enlace" data-sg-quitar="${i}" aria-label="Quitar">✕</button></span>`).join("");
        m.querySelectorAll("[data-sg-quitar]").forEach((b) => b.addEventListener("click", () => { pend.splice(Number(b.dataset.sgQuitar), 1); pintar(); }));
      };
      const agregar = async (files) => {
        for (const f of files) {
          const a = await prepararArchivo(f);
          if (a.tam > MAX_ADJ) { err.textContent = `«${a.nombre}» pesa más de 5 MB.`; continue; }
          a.url = esImagen(a.tipo) ? URL.createObjectURL(new Blob([a.bytes], { type: a.tipo })) : "";
          pend.push(a);
        }
        pintar();
      };
      m.querySelector("[name=Descripcion]").addEventListener("paste", (ev) => {
        const files = [...(ev.clipboardData ? ev.clipboardData.items : [])].filter((i) => i.kind === "file").map((i) => i.getAsFile()).filter(Boolean)
          .map((f, k) => (f.name && f.name !== "image.png" ? f : new File([f], `captura-${Date.now()}-${k + 1}.${(f.type.split("/")[1] || "png")}`, { type: f.type })));
        if (files.length) { ev.preventDefault(); agregar(files); }
      });
      m.querySelector("#sg-adjuntar").addEventListener("click", () => { const i = document.createElement("input"); i.type = "file"; i.multiple = true; i.addEventListener("change", () => agregar([...i.files])); i.click(); });
      m.querySelector("[name=Descripcion]").focus();
    };
    const subirAdj = async (idSug, base) => {
      const adj = base.slice();
      for (let k = 0; k < pend.length; k++) {
        const a = pend[k], id = `ADJ-${idSug}-${Date.now().toString(36)}-${k + 1}`;
        await guardarArchivo(id, a.bytes, a.nombre, a.tipo, `el adjunto ${k + 1} de ${pend.length}`);
        adj.push({ id, n: a.nombre, t: a.tipo, s: a.tam });
      }
      return adj.length ? JSON.stringify(adj) : "";
    };
    const validar = (fd) => (!fd.Descripcion ? "Describe la mejora o el error." : "");
    if (nuevo) {
      modal("💡 Sugerir una mejora", campos, {}, (fd) => {
        const n = (S.datos.Sugerencias || []).reduce((mx, x) => Math.max(mx, parseInt(String(x.ID_Sugerencia).replace(/\D/g, ""), 10) || 0), 0) + 1;
        const fila = { ID_Sugerencia: `SUG-${String(n).padStart(4, "0")}`, Fecha_Hora: ahora(), Usuario: S.usuario.Correo, ...fd, ID_Proyecto: fd.Modulo === ctx.Modulo ? ctx.ID_Proyecto : "",
          Adjuntos: "", Estado: "Nueva", Respuesta: "", Version_App: window.APP_VERSION || "", Version_Implementada: "", Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() };
        guardar(async () => { fila.Adjuntos = await subirAdj(fila.ID_Sugerencia, []); await S.api.agregarFila("Sugerencias", fila); }, `¡Gracias! Sugerencia ${fila.ID_Sugerencia} registrada`);
      }, validar, { init });
      return;
    }
    const puedeBorrar = gestionaSug() || (lc(sg.Usuario) === lc(S.usuario.Correo) && (sg.Estado || "Nueva") === "Nueva");
    modal(`Editar ${sg.ID_Sugerencia}`, campos, sg, (fd) => {
      guardar(async () => { const Adjuntos = await subirAdj(sg.ID_Sugerencia, yaAdj); await S.api.actualizarPorId("Sugerencias", "ID_Sugerencia", sg.ID_Sugerencia, { ...fd, Adjuntos, Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() }); }, "Sugerencia actualizada");
    }, validar, { init, ...(puedeBorrar ? { eliminar: { texto: "Eliminar sugerencia", mensaje: `Se borrará ${sg.ID_Sugerencia}${yaAdj.length ? " con sus adjuntos" : ""}.`,
      accion: () => guardar(async () => { await borrarAdjuntos([sg]); await S.api.eliminarFilas("Sugerencias", "ID_Sugerencia", [sg.ID_Sugerencia]); }, "Sugerencia eliminada") } } : {}) });
  }
  // Admin/PMO: estado, respuesta y versión en que quedó.
  function gestionarSugerencia(sg) {
    modal(`Gestionar ${sg.ID_Sugerencia}`, [
      { html: `<div class="sg-resumen">${pill(sg.Tipo || "Mejora")} ${pill(sg.Prioridad || "Media")} <b>${esc(sg.Modulo)}</b>${sg.Contexto ? ` · ${esc(sg.Contexto)}` : ""}<p>${esc(sg.Descripcion)}</p><div class="sub">${esc(nombreUsuario(sg.Usuario))} · ${fechaHora(sg.Fecha_Hora)}</div></div>` },
      { k: "Estado", label: "Estado", tipo: "select", opciones: ESTADOS_SUG, def: "Nueva" },
      { k: "Version_Implementada", label: "Versión en que quedó", placeholder: "Ej.: 34", def: "" },
      { k: "Respuesta", label: "Respuesta para quien la sugirió", tipo: "textarea", placeholder: "Qué se hizo, por qué se descartó o cuándo se hará" },
    ], { ...sg, Estado: sg.Estado || "Nueva" }, (fd) => {
      if (fd.Estado === "Hecha" && !fd.Version_Implementada) fd.Version_Implementada = window.APP_VERSION || "";
      guardar(() => S.api.actualizarPorId("Sugerencias", "ID_Sugerencia", sg.ID_Sugerencia, { ...fd, Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() }), "Sugerencia actualizada");
    });
  }
  function detalleSugerencia(sg) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    const mia = lc(sg.Usuario) === lc(S.usuario.Correo);
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="Sugerencia">
      <div class="titulo-fila"><div><h2>${esc(sg.ID_Sugerencia)} · ${esc(sg.Tipo || "Mejora")}</h2><div class="sub">${esc(sg.Modulo)}${sg.Contexto ? ` · ${esc(sg.Contexto)}` : ""}</div></div><button class="btn enlace" id="sg-cerrar" aria-label="Cerrar">✕</button></div>
      <div class="contexto">
        <div><span class="sub">Estado</span><b>${pill(sg.Estado || "Nueva")}</b></div>
        <div><span class="sub">Prioridad</span><b>${pill(sg.Prioridad || "Media")}</b></div>
        <div><span class="sub">Sugerida por</span><b>${esc(nombreUsuario(sg.Usuario))}</b><div class="sub">${fechaHora(sg.Fecha_Hora)} · versión ${esc(sg.Version_App || "—")}</div></div>
        ${sg.Version_Implementada ? `<div><span class="sub">Implementada en</span><b>versión ${esc(sg.Version_Implementada)}</b></div>` : ""}
      </div>
      <div class="form-seccion">Descripción</div><p class="sg-desc">${esc(sg.Descripcion)}</p>
      ${sg.URL_Referencia ? `<p><a href="${esc(sg.URL_Referencia)}" target="_blank" rel="noopener">🔗 Enlace de referencia</a></p>` : ""}
      ${htmlAdjuntos(sg)}
      ${sg.Respuesta ? `<div class="form-seccion">Respuesta</div><p class="sg-desc">${esc(sg.Respuesta)}</p>` : ""}
      <div class="acciones derecha">${mia || gestionaSug() ? `<button class="btn" id="sg-editar">Editar</button>` : ""}${gestionaSug() ? `<button class="btn primario" id="sg-gestionar">Cambiar estado / responder</button>` : ""}</div></div>`;
    document.body.appendChild(m);
    $("#sg-cerrar").addEventListener("click", cerrarModal);
    m.addEventListener("keydown", (e) => { if (e.key === "Escape") cerrarModal(); });
    if ($("#sg-editar")) $("#sg-editar").addEventListener("click", () => formSugerencia(sg));
    if ($("#sg-gestionar")) $("#sg-gestionar").addEventListener("click", () => gestionarSugerencia(sg));
    enlazarAdjuntos(m);
  }
  function sugerenciasFiltradas() {
    const F = S.sugF || (S.sugF = { estado: "abiertas", modulo: "", tipo: "", prio: "", mias: false, q: "" });
    return sugerencias().filter((x) => {
      const e = x.Estado || "Nueva";
      return (F.estado === "abiertas" ? !["Hecha", "Descartada"].includes(e) : !F.estado || e === F.estado) && (!F.modulo || x.Modulo === F.modulo) && (!F.tipo || x.Tipo === F.tipo) &&
        (!F.prio || x.Prioridad === F.prio) && (!F.mias || lc(x.Usuario) === lc(S.usuario.Correo)) && (!F.q || normTxt(`${x.ID_Sugerencia} ${x.Descripcion} ${x.Contexto} ${x.Respuesta}`).includes(normTxt(F.q)));
    });
  }
  // Texto ordenado por módulo para pedir los cambios a Claude.
  function textoParaClaude(lista) {
    const porMod = {};
    lista.forEach((x) => (porMod[x.Modulo || "General"] = porMod[x.Modulo || "General"] || []).push(x));
    const ordenP = { Alta: 0, Media: 1, Baja: 2 };
    return [`Mejoras pedidas para la app Portafolio PMO (versión actual ${window.APP_VERSION || ""}) — ${lista.length} sugerencia(s):`, "",
      ...Object.keys(porMod).sort().flatMap((mod) => [`## ${mod}`, ...porMod[mod].sort((a, b) => (ordenP[a.Prioridad] ?? 1) - (ordenP[b.Prioridad] ?? 1)).map((x) =>
        `- [${x.ID_Sugerencia}] (${x.Tipo || "Mejora"} · prioridad ${x.Prioridad || "Media"})${x.Contexto ? ` En «${x.Contexto}»:` : ""} ${String(x.Descripcion || "").replace(/\s+/g, " ").trim()}` +
        `${adjuntosDe(x).length ? ` [${adjuntosDe(x).length} captura(s) en la app]` : ""}${x.URL_Referencia ? ` Referencia: ${x.URL_Referencia}` : ""} — ${nombreUsuario(x.Usuario)}, ${fecha(String(x.Fecha_Hora).slice(0, 10))}`), ""]),
      "Al terminar, indícame la versión para marcarlas como «Hecha»."].join("\n");
  }
  function mostrarTexto(titulo, texto) {
    cerrarModal();
    const m = document.createElement("div");
    m.id = "modal";
    m.innerHTML = `<div class="modal-caja" role="dialog" aria-modal="true" aria-label="${esc(titulo)}"><div class="titulo-fila"><h2>${esc(titulo)}</h2><button class="btn enlace" id="tx-cerrar">✕</button></div>
      <p class="sub">Cópialo y pégalo en el chat con Claude. Si tienen capturas, adjúntalas también.</p><textarea id="tx-texto" rows="16" readonly>${esc(texto)}</textarea>
      <div class="acciones derecha"><button class="btn primario" id="tx-copiar">Copiar</button></div></div>`;
    document.body.appendChild(m);
    $("#tx-cerrar").addEventListener("click", cerrarModal);
    const ta = $("#tx-texto"); ta.focus(); ta.select();
    $("#tx-copiar").addEventListener("click", async () => {
      ta.select();
      try { await navigator.clipboard.writeText(texto); toast("Texto copiado"); } catch (e) { try { document.execCommand("copy"); toast("Texto copiado"); } catch (e2) { toast("Selecciona el texto y cópialo con Ctrl+C", true); } }
    });
  }
  function vSugerencias(el) {
    const F = S.sugF || (S.sugF = { estado: "abiertas", modulo: "", tipo: "", prio: "", mias: false, q: "" });
    const todas = sugerencias(), lista = sugerenciasFiltradas(), gest = gestionaSug();
    const cuenta = (e) => todas.filter((x) => (x.Estado || "Nueva") === e).length;
    const modulos = [...new Set(todas.map((x) => x.Modulo).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
    const sel = (id, val, ops, vacioTxt) => `<select id="${id}"><option value="">${vacioTxt}</option>${ops.map((o) => { const [v, t] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(v)}" ${v === val ? "selected" : ""}>${esc(t)}</option>`; }).join("")}</select>`;
    el.innerHTML = `
      <div class="titulo-fila"><h1>Sugerencias de mejora</h1><button class="btn primario" id="sg-nueva">💡 Nueva sugerencia</button></div>
      <div class="kpis sg-kpis">${ESTADOS_SUG.map((e) => `<button class="kpi sg-kpi ${F.estado === e ? "activo" : ""}" data-sg-estado="${esc(e)}"><span class="kpi-v">${cuenta(e)}</span><span class="kpi-t">${esc(e)}</span></button>`).join("")}</div>
      <div class="filtros">
        <label class="filtro crece"><span>Buscar</span><input id="sg-q" type="search" placeholder="ID, descripción o respuesta" value="${esc(F.q)}"></label>
        <label class="filtro"><span>Estado</span>${sel("sg-f-estado", F.estado, [["abiertas", "Abiertas (sin cerrar)"], ...ESTADOS_SUG], "Todas")}</label>
        <label class="filtro"><span>Módulo</span>${sel("sg-f-mod", F.modulo, modulos, "Todos")}</label>
        <label class="filtro"><span>Tipo</span>${sel("sg-f-tipo", F.tipo, TIPOS_SUG, "Todos")}</label>
        <label class="filtro"><span>Prioridad</span>${sel("sg-f-prio", F.prio, PRIOS_SUG, "Todas")}</label>
        <label class="check"><input type="checkbox" id="sg-mias" ${F.mias ? "checked" : ""}> Solo las mías</label>
      </div>
      ${gest && lista.length ? `<div class="sg-barra"><label class="check"><input type="checkbox" id="sg-todas"> Seleccionar todas</label><span class="sub" id="sg-nsel">0 seleccionadas</span>
        <button class="btn chico" id="sg-claude" disabled>📋 Copiar para Claude</button>
        <select id="sg-bulk-estado" aria-label="Nuevo estado">${ESTADOS_SUG.map((e) => `<option>${esc(e)}</option>`).join("")}</select><input id="sg-bulk-ver" placeholder="Versión" aria-label="Versión" style="width:90px">
        <button class="btn chico" id="sg-bulk" disabled>Cambiar estado</button></div>` : ""}
      <div class="card">${lista.length ? `<div class="tabla-scroll"><table class="tabla-sug"><thead><tr>${gest ? "<th></th>" : ""}<th>ID</th><th>Módulo</th><th>Sugerencia</th><th>Tipo</th><th>Prioridad</th><th>Estado</th><th class="opc">Por</th><th></th></tr></thead>
        <tbody>${lista.map((x) => `<tr>${gest ? `<td><input type="checkbox" class="sg-sel" value="${esc(x.ID_Sugerencia)}" aria-label="Seleccionar ${esc(x.ID_Sugerencia)}"></td>` : ""}
          <td class="nowrap"><b>${esc(x.ID_Sugerencia)}</b><div class="sub">${fecha(String(x.Fecha_Hora).slice(0, 10))}</div></td>
          <td>${esc(x.Modulo)}${x.Contexto ? `<div class="sub">${esc(x.Contexto)}</div>` : ""}</td>
          <td class="sg-txt">${esc(String(x.Descripcion || "").slice(0, 220))}${String(x.Descripcion || "").length > 220 ? "…" : ""}${adjuntosDe(x).length ? ` <span class="sub">📎 ${adjuntosDe(x).length}</span>` : ""}${x.Respuesta ? `<div class="sub">↳ ${esc(String(x.Respuesta).slice(0, 140))}</div>` : ""}</td>
          <td${ed({ puede: gest, tabla: "Sugerencias", idCol: "ID_Sugerencia", id: x.ID_Sugerencia, campo: "Tipo", tipo: "select", opciones: TIPOS_SUG, valor: x.Tipo || "Mejora", extra: () => ({ Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() }), etiqueta: "Tipo" })}>${pill(x.Tipo || "Mejora")}</td>
          <td${ed({ puede: gest, tabla: "Sugerencias", idCol: "ID_Sugerencia", id: x.ID_Sugerencia, campo: "Prioridad", tipo: "select", opciones: PRIOS_SUG, valor: x.Prioridad || "Media", extra: () => ({ Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() }), etiqueta: "Prioridad" })}>${pill(x.Prioridad || "Media")}</td>
          <td${ed({ puede: gest, tabla: "Sugerencias", idCol: "ID_Sugerencia", id: x.ID_Sugerencia, campo: "Estado", tipo: "select", opciones: ESTADOS_SUG, valor: x.Estado || "Nueva", extra: (v) => ({ Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO(), ...(v === "Hecha" && !x.Version_Implementada ? { Version_Implementada: window.APP_VERSION || "" } : {}) }), etiqueta: "Estado" })}>${pill(x.Estado || "Nueva")}${x.Version_Implementada ? `<div class="sub">v${esc(x.Version_Implementada)}</div>` : ""}</td>
          <td class="opc">${esc(nombreUsuario(x.Usuario))}</td>
          <td class="derecha nowrap"><button class="btn chico" data-sg-ver="${esc(x.ID_Sugerencia)}">Ver</button>${gest ? ` <button class="btn chico" data-sg-gest="${esc(x.ID_Sugerencia)}">Gestionar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`
        : vacio(todas.length ? "No hay sugerencias con estos filtros." : "Aún no hay sugerencias. Usa «💡 Sugerir mejora» en cualquier pantalla.")}</div>`;
    const buscar = (id) => (S.datos.Sugerencias || []).find((x) => x.ID_Sugerencia === id);
    const set = (k, v) => { F[k] = v; render(); };
    $("#sg-nueva").addEventListener("click", () => formSugerencia());
    el.querySelectorAll("[data-sg-estado]").forEach((b) => b.addEventListener("click", () => set("estado", F.estado === b.dataset.sgEstado ? "abiertas" : b.dataset.sgEstado)));
    $("#sg-q").addEventListener("change", (e) => set("q", e.target.value.trim()));
    $("#sg-f-estado").addEventListener("change", (e) => set("estado", e.target.value));
    $("#sg-f-mod").addEventListener("change", (e) => set("modulo", e.target.value));
    $("#sg-f-tipo").addEventListener("change", (e) => set("tipo", e.target.value));
    $("#sg-f-prio").addEventListener("change", (e) => set("prio", e.target.value));
    $("#sg-mias").addEventListener("change", (e) => set("mias", e.target.checked));
    el.querySelectorAll("[data-sg-ver]").forEach((b) => b.addEventListener("click", () => detalleSugerencia(buscar(b.dataset.sgVer))));
    el.querySelectorAll("[data-sg-gest]").forEach((b) => b.addEventListener("click", () => gestionarSugerencia(buscar(b.dataset.sgGest))));
    if (!gest || !lista.length) return;
    const elegidas = () => [...el.querySelectorAll(".sg-sel:checked")].map((i) => buscar(i.value)).filter(Boolean);
    const act = () => { const n = elegidas().length; $("#sg-nsel").textContent = `${n} seleccionada${n === 1 ? "" : "s"}`; $("#sg-claude").disabled = !n; $("#sg-bulk").disabled = !n; };
    el.querySelectorAll(".sg-sel").forEach((c) => c.addEventListener("change", act));
    $("#sg-todas").addEventListener("change", (e) => { el.querySelectorAll(".sg-sel").forEach((c) => { c.checked = e.target.checked; }); act(); });
    $("#sg-claude").addEventListener("click", () => mostrarTexto("Copiar para Claude", textoParaClaude(elegidas())));
    $("#sg-bulk").addEventListener("click", () => {
      const l = elegidas(), estado = $("#sg-bulk-estado").value, ver = $("#sg-bulk-ver").value.trim() || (estado === "Hecha" ? window.APP_VERSION || "" : "");
      guardar(() => S.api.actualizarVarios("Sugerencias", "ID_Sugerencia", l.map((x) => ({ id: x.ID_Sugerencia, cambios: { Estado: estado, ...(ver ? { Version_Implementada: ver } : {}), Actualizado_Por: S.usuario.Correo, Actualizado_El: R.hoyISO() } }))),
        `${l.length} sugerencia(s) en «${estado}»`);
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
      S.api = conAuditoria(conMedicion(await DATOS.crear(S.modo)));
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
