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
    { id: "proyectos", t: "Proyectos" },
    { id: "cronograma", t: "Cronograma" },
    { id: "riesgos", t: "Riesgos" },
    { id: "catalogos", t: "Catálogos", permiso: "catalogos" },
    { id: "usuarios", t: "Usuarios", permiso: "usuarios" },
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
    riesgos: ["Mapa de calor de riesgos abiertos por probabilidad e impacto.",
      "Cambia entre inherente (antes de mitigar) y residual (después de mitigar)."],
    catalogos: ["Valores de los desplegables de los formularios.",
      "Agrega o quita valores; los proyectos que ya usan un valor lo conservan.",
      "Aquí también configuras los roles de stakeholder (Sponsor, Líder funcional, Product Owner…) y el maestro de proveedores."],
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
  }
  // Listas nuevas que se crean solas en la tabla Catalogos la primera vez (luego se editan desde Catálogos).
  const CATALOGOS_BASE = { Rol_Stakeholder: ["Sponsor", "Líder funcional", "Product Owner"], Estado_Ticket: ["Abierto", "En curso", "Resuelto", "Cerrado"] };
  let sembrado = false;
  // La primera vez, deja en el Excel los roles de stakeholder base para poder editarlos desde Catálogos.
  async function sembrarCatalogos() {
    // Compromisos de versiones anteriores: «Cumplido» pasa a «Cerrado».
    const viejos = S.datos.Compromisos.filter((c) => c.Estado === "Cumplido");
    if (viejos.length) {
      try { await S.api.actualizarVarios("Compromisos", "ID_Compromiso", viejos.map((c) => ({ id: c.ID_Compromiso, cambios: { Estado: "Cerrado" } }))); viejos.forEach((c) => { c.Estado = "Cerrado"; }); }
      catch (e) { /* sin permiso de edición: la app igual los trata como cerrados */ }
    }
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
          <div class="marca"><span class="logo">P</span> Portafolio de Proyectos</div>
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
    app().innerHTML = `
      <div class="layout">
        <aside class="lateral">
          <div class="marca"><span class="logo">P</span> Portafolio PMO</div>
          <nav aria-label="Módulos">${menu.map((v) => `<button class="nav ${v.id === activa ? "activa" : ""}" data-vista="${v.id}">${esc(v.t)}</button>`).join("")}</nav>
          <select id="nav-movil" aria-label="Módulo">${menu.map((v) => `<option value="${v.id}" ${v.id === activa ? "selected" : ""}>${esc(v.t)}</option>`).join("")}</select>
        </aside>
        <main>
          <header class="barra">
            <span class="modo ${S.modo}">${esc(etiquetaModo)}</span><span class="version" title="Versión de la app">v${esc(window.APP_VERSION || "")}</span>
            <span class="quien">${esc(S.usuario.Nombre || S.usuario.Correo)} · <b>${esc(R.roles(S.usuario).join(", "))}</b></span>
            ${(() => { const a = alertas(); const n = a.vencidos.length + a.pronto.length; return `<button class="btn campana ${a.vencidos.length ? "roja" : n ? "ambar" : ""}" id="b-alertas" title="Compromisos vencidos o por vencer">🔔 <span>${n}</span></button>`; })()}
            <button class="btn" id="b-ayuda" title="Mostrar la ayuda de esta pantalla">? Ayuda</button>
            <button class="btn" id="b-recargar" title="Volver a leer el Excel">Actualizar</button>
            <button class="btn" id="b-salir">Salir</button>
          </header>
          <section id="vista"></section>
        </main>
      </div>`;
    document.querySelectorAll(".nav").forEach((b) => b.addEventListener("click", () => ir(b.dataset.vista)));
    $("#nav-movil").addEventListener("change", (e) => ir(e.target.value));
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
    const vistas = { dashboard: vDashboard, avances: vAvances, seguimiento: vSeguimiento, proyectos: vProyectos, ficha: vFicha, cronograma: vCronograma, riesgos: vRiesgos, catalogos: vCatalogos, usuarios: vUsuarios };
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
    el.querySelectorAll("[data-pid]").forEach((r) => r.addEventListener("click", (ev) => {
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
  const vacio = (msg) => `<div class="vacio">${esc(msg)}</div>`;

  function grafico(canvas, config) {
    if (!canvas) return;
    if (typeof Chart === "undefined") { canvas.replaceWith(Object.assign(document.createElement("div"), { className: "vacio", textContent: "No se pudo cargar la librería de gráficos." })); return; }
    Chart.defaults.font.family = "Calibri, Arial, sans-serif";
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
      <div class="card"><h2>Portafolio por estado</h2>${barrasH(porEstado, { color: COLORES.azulClaro })}</div>`;

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
      <h1>Avances de la semana</h1>
      ${barraFiltros()}
      <div class="filtros"><label class="filtro"><span>Semana</span><select id="sel-semana">${semanas.map((w) => `<option ${w === S.semana ? "selected" : ""}>${w}</option>`).join("")}</select></label>
        <span class="sub">${deSemana.length} seguimiento(s) reportado(s) · ${sinReporte.length} proyecto(s) semanal(es) sin reporte</span></div>
      <div class="card"><div class="tabla-scroll"><table>
        <thead><tr><th>Proyecto</th><th>Avance</th><th>Semáforo</th><th>Logros</th><th class="opc">Próximos pasos</th><th>Bloqueos</th><th class="opc">Compromisos</th></tr></thead>
        <tbody>${filas || `<tr><td colspan="7">${vacio("No hay seguimientos reportados en esta semana con los filtros actuales.")}</td></tr>`}</tbody></table></div></div>
      <div class="card"><h2>Proyectos semanales sin reporte en ${esc(S.semana)}</h2>
        ${sinReporte.length ? `<ul class="lista">${sinReporte.map((p) => `<li class="clic" data-pid="${esc(p.ID_Proyecto)}"><b>${esc(p.Nombre)}</b> · ${esc(nombreUsuario(p.PM))}</li>`).join("")}</ul>` : vacio("Todos los proyectos semanales reportaron.")}</div>`;
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
  function correoRecordatorio(items) {
    const para = [...new Set(items.map((x) => correoDe(x.c)).filter(Boolean))];
    const uno = items.length === 1 ? items[0].c : null;
    const asunto = uno ? `Recordatorio: «${uno.Compromiso}» ${R.estadoCompromiso(uno) === "Vencido" ? "está vencido" : `vence el ${fecha(uno.Fecha_Compromiso)}`}` : `Recordatorio: ${items.length} compromisos pendientes`;
    const cuerpo = `Hola,\n\nTe recuerdo ${items.length === 1 ? "este compromiso" : "estos compromisos"} del portafolio de proyectos:\n\n` +
      items.map((x) => { const p = proyecto(x.c.ID_Proyecto) || {}; return `• ${x.c.Compromiso}\n  Proyecto: ${p.Nombre || x.c.ID_Proyecto}\n  Fecha límite: ${x.c.Fecha_Compromiso ? fecha(x.c.Fecha_Compromiso) : "sin fecha"} (${R.estadoCompromiso(x.c)})`; }).join("\n\n") +
      `\n\nPor favor cuéntame cómo va o si necesitas apoyo.\n\nGracias,\n${S.usuario.Nombre || S.usuario.Correo}`;
    return { para, asunto, cuerpo, href: `mailto:${para.join(";")}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}` };
  }
  // Abre el correo en Outlook y deja constancia en el hilo de cada compromiso.
  async function enviarRecordatorio(items) {
    const m = correoRecordatorio(items);
    const a = document.createElement("a"); a.href = m.href; a.target = "_blank"; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove();
    try {
      await S.api.agregarFilas("Comentarios", items.map((x) => nuevoComentario(x.c, `Recordatorio enviado por correo${m.para.length ? ` a ${m.para.join(", ")}` : ""}.`)));
      await recargar();
    } catch (e) { /* el correo ya se abrió; la constancia es opcional */ }
    return m;
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
      b.textContent = "✓ Correo abierto"; b.disabled = true;
      toast(r.para.length ? `Correo listo para ${r.para.join(", ")}` : "Correo abierto: escribe el destinatario (el compromiso no tiene correo del responsable).");
    }));
    m.querySelectorAll("[data-al-grupo]").forEach((b) => b.addEventListener("click", async () => {
      const r = await enviarRecordatorio(grupos[Number(b.dataset.alGrupo)][1]);
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
      <div class="hilo">${coms.map((x) => `<div class="com"><div class="com-cab"><b>${esc(nombreUsuario(x.Autor))}</b><span class="sub">${fechaHora(x.Fecha_Hora)}</span></div>${x.Texto ? `<div class="com-texto">${esc(x.Texto)}</div>` : ""}${htmlAdjuntos(x)}</div>`).join("") || `<div class="vacio">Aún no hay comentarios. Escribe el primero: avances, dudas o bloqueos de este compromiso.</div>`}</div>
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
    if ($("#d-recordar")) $("#d-recordar").addEventListener("click", async () => { await enviarRecordatorio([{ c }]); render(); detalleCompromiso(id); toast("Correo de recordatorio abierto en tu Outlook"); });
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
          <td>${esc(nombreUsuario(p.PM))}</td><td class="opc">${esc(p.Metodologia)}</td><td class="opc">${esc(p.Fase)}</td><td>${esc(p.Estado)}</td><td>${chipSemaforo(p.Semaforo)}</td>
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
      ["seguimientos", `Seguimientos <span class="contador">${segs.length}</span>`], ["hitos", `Hitos <span class="contador">${hitos.length}</span>`], ["riesgos", `Riesgos <span class="contador">${riesgos.length}</span>`],
      ["tickets", `Tickets <span class="contador">${ticketsDe(p.ID_Proyecto).filter(ticketAbierto).length}</span>`]];
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
      <div class="card"><h2>Curva de avance</h2>${segs.length ? `<div class="grafico-alto"><canvas id="g-curva"></canvas></div>` : vacio("Aún no hay seguimientos. La curva aparece con el primer reporte.")}</div>
`,
      compromisos: () => seccionCompromisos(p, comps, segs, puedeP("compromisos")),
      tickets: () => seccionTickets(p, puedeP("compromisos")),
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
      riesgos: () => `      <div class="card"><div class="titulo-fila"><h2>Riesgos</h2>${puedeP("riesgos") ? `<button class="btn" id="b-riesgo">+ Riesgo</button>` : ""}</div><div class="tabla-scroll"><table>
        <thead><tr><th>Riesgo</th><th class="opc">Tipo</th><th>Estado</th><th>Inherente</th><th>Residual</th><th class="opc">Mitigación</th><th></th></tr></thead>
        <tbody>${riesgos.map((r) => `<tr><td>${esc(r.Descripcion)}</td><td class="opc">${esc(r.Tipo)}</td><td>${esc(r.Estado)}</td>
          <td>${nivel(r.Calificacion_Inherente)}</td><td>${nivel(r.Calificacion_Residual)}</td><td class="opc">${esc(r.Plan_Mitigacion)}</td>
          <td class="derecha">${puedeP("riesgos") ? `<button class="btn chico" data-riesgo="${esc(r.ID_Riesgo)}">Editar</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="7">${vacio("Sin riesgos registrados.")}</td></tr>`}</tbody></table></div></div>`,
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
          return `<div class="g-fila clic" data-pid="${esc(p.ID_Proyecto)}"><div class="g-nombre"><b>${esc(p.Nombre)}</b><div class="sub">${esc(p.Estado)} · ${pct(p.Avance_Real)} · ${fecha(p.Fecha_Inicio)} → ${fecha(p.Fecha_Fin_Plan)}</div></div>
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
    const abiertos = S.datos.Riesgos.filter((r) => ids.has(r.ID_Proyecto) && r.Estado === "Abierto");
    const res = S.riesgoVista === "residual";
    const P = res ? "Probabilidad_Residual" : "Probabilidad_Inherente";
    const I = res ? "Impacto_Residual" : "Impacto_Inherente";
    const C = res ? "Calificacion_Residual" : "Calificacion_Inherente";
    let celdas = "";
    for (let pr = 5; pr >= 1; pr--) {
      celdas += `<div class="hm-eje">${pr}</div>`;
      for (let im = 1; im <= 5; im++) {
        const n = abiertos.filter((r) => Number(r[P]) === pr && Number(r[I]) === im).length;
        celdas += `<div class="hm-celda riesgo-${R.nivelRiesgo(pr * im).toLowerCase()}" title="Probabilidad ${pr} × Impacto ${im}: ${n} riesgo(s)">${n || ""}</div>`;
      }
    }
    celdas += `<div></div>${[1, 2, 3, 4, 5].map((i) => `<div class="hm-eje">${i}</div>`).join("")}`;
    const orden = abiertos.slice().sort((a, b) => (Number(b[C]) || 0) - (Number(a[C]) || 0));
    el.innerHTML = `
      <h1>Riesgos</h1>
      ${barraFiltros()}
      <div class="filtros"><div class="seg" role="group" aria-label="Tipo de calificación">
        <button class="btn ${!res ? "primario" : ""}" data-rv="inherente">Inherente</button>
        <button class="btn ${res ? "primario" : ""}" data-rv="residual">Residual</button></div>
        <span class="sub">${abiertos.length} riesgo(s) abierto(s) · Bajo ≤ 6 · Medio 7–12 · Alto > 12</span></div>
      <div class="grid2">
        <div class="card"><h2>Mapa de calor (${res ? "residual" : "inherente"})</h2>
          <div class="heatmap">${celdas}</div><div class="sub centro">Horizontal: impacto (1 a 5) · Vertical: probabilidad (1 a 5)</div></div>
        <div class="card"><h2>Riesgos abiertos más críticos</h2><div class="tabla-scroll"><table>
          <thead><tr><th>Proyecto</th><th>Riesgo</th><th>Calificación</th></tr></thead>
          <tbody>${orden.slice(0, 10).map((r) => `<tr class="clic" data-pid="${esc(r.ID_Proyecto)}"><td>${esc((proyecto(r.ID_Proyecto) || {}).Nombre)}</td><td>${esc(r.Descripcion)}</td><td>${nivel(r[C])}</td></tr>`).join("") || `<tr><td colspan="3">${vacio("Sin riesgos abiertos en la selección.")}</td></tr>`}</tbody>
        </table></div></div>
      </div>`;
    enlazarFiltros();
    el.querySelectorAll("[data-rv]").forEach((b) => b.addEventListener("click", () => { S.riesgoVista = b.dataset.rv; render(); }));
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
            <p class="sub">${c.tabla ? "Entre paréntesis: registros que usan cada valor." : "Entre paréntesis: proyectos que usan cada valor."}</p>
            <ul class="cat-lista">${valores.map((v) => `<li><span class="cat-valor">${esc(v)} <span class="contador" title="En uso">${enUso(c, v)}</span></span>
              <span class="cat-acc"><button class="btn chico" data-renombrar="${esc(c.lista)}" data-valor="${esc(v)}" title="Renombrar">✎ Renombrar</button><button class="btn chico" data-quitar="${esc(c.lista)}" data-valor="${esc(v)}" title="Quitar">✕</button></span></li>`).join("") || `<li>${vacio("Sin valores.")}</li>`}</ul>
            <form class="cat-agregar" data-lista="${esc(c.lista)}" novalidate>
              <input name="valor" placeholder="Nuevo valor" aria-label="Nuevo valor para ${esc(c.t)}">
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
        if (valor.length > 60) { err.textContent = "Máximo 60 caracteres."; return; }
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
      if ((S.cat[lista] || []).length <= 1) { err.textContent = "La lista debe tener al menos un valor."; return; }
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
      return dl + `<input name="${c.k}" ${c.sugerencias ? `list="dl-${c.k}" autocomplete="off"` : ""} type="${c.tipo || "text"}" value="${esc(v)}" placeholder="${esc(c.placeholder || "")}" ${c.min !== undefined ? `min="${c.min}"` : ""} ${c.max !== undefined ? `max="${c.max}"` : ""} ${dis}>`;
    };
    const cuerpo = campos.map((c) => (c.seccion
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
      campos.filter((c) => !c.seccion).forEach((c) => {
        if (c.tipo === "checks") { fd[c.k] = [...m.querySelectorAll(`[data-checks="${c.k}"] input:checked`)].map((i) => i.value).join(", "); return; }
        let v = m.querySelector(`[name="${c.k}"]`).value.trim();
        if (c.tipo === "number" && v !== "") v = Number(v);
        fd[c.k] = v;
      });
      let ok = true;
      const err = (k, msg) => { ok = false; const e = m.querySelector(`[data-e="${k}"]`); if (e) e.textContent = msg; };
      campos.filter((c) => !c.seccion).forEach((c) => {
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
      <div class="hv-marca" aria-hidden="true"><span>FSFB</span></div>
      <header class="hv-cab">
        <div class="hv-logo">Fundación<br>Santa Fe de Bogotá</div>
        <div class="hv-cab-t"><div class="hv-tipo">Hoja de vida del proyecto</div><h1>${esc(p.Nombre)}</h1>
          <div class="hv-sub">${esc(pid)}${p.Codigo_Almera ? ` · Almera ${esc(p.Codigo_Almera)}` : ""} · ${esc(p.Cliente_Area || "")}</div></div>
        <div class="hv-gen">Generado el ${fecha(R.hoyISO())}<br>por ${esc(S.usuario.Nombre || S.usuario.Correo)}</div>
      </header>
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
      ${tabla(["Riesgo", "Tipo", "Estado", "Inherente", "Residual", "Mitigación"], riesgos.map((r) => `<tr><td>${t(r.Descripcion)}</td><td>${t(r.Tipo)}</td><td>${t(r.Estado)}</td><td>${R.nivelRiesgo(Number(r.Calificacion_Inherente) || 0)} (${Number(r.Calificacion_Inherente) || 0})</td><td>${R.nivelRiesgo(Number(r.Calificacion_Residual) || 0)} (${Number(r.Calificacion_Residual) || 0})</td><td>${t(r.Plan_Mitigacion)}</td></tr>`), "Sin riesgos.")}
      <h2>8. Tickets del helpdesk</h2>
      ${tabla(["N.º", "Título", "Estado", "Prioridad", "Registrado"], tks.map((k) => `<tr><td>${t(k.Numero)}</td><td>${t(k.Titulo)}</td><td>${t(k.Estado)}</td><td>${t(k.Prioridad)}</td><td>${fecha(k.Fecha_Registro)}</td></tr>`), "Sin tickets.")}
      <footer class="hv-pie">Fundación Santa Fe de Bogotá · Oficina de Proyectos · Hoja de vida ${esc(pid)} · Documento generado desde el Portafolio PMO</footer>`;
    cerrarModal();
    const o = document.createElement("div");
    o.className = "hv-overlay";
    o.innerHTML = `<div class="hv-barra no-print"><b>Hoja de vida · ${esc(p.Nombre)}</b><div class="acciones">
      <button class="btn primario" id="hv-imprimir">🖨 Imprimir / Guardar como PDF</button><button class="btn" id="hv-cerrar">Cerrar</button></div>
      <div class="sub">En la ventana de impresión elige «Guardar como PDF» como destino.</div></div>
      <div class="hv-hoja"><article class="hv">${doc}</article></div>`;
    document.body.appendChild(o);
    document.body.classList.add("con-hv");
    const cerrar = () => { o.remove(); document.body.classList.remove("con-hv"); };
    o.querySelector("#hv-cerrar").addEventListener("click", cerrar);
    o.querySelector("#hv-imprimir").addEventListener("click", () => {
      const tituloAnt = document.title;
      document.title = `Hoja de vida ${pid} - ${p.Nombre}`;
      window.print();
      setTimeout(() => { document.title = tituloAnt; }, 1000);
    });
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

  // ---------- Stakeholders y proveedores ----------
  const proveedores = () => (S.datos.Proveedores || []).slice().sort((a, b) => String(a.Nombre).localeCompare(String(b.Nombre), "es"));
  const proveedor = (id) => (S.datos.Proveedores || []).find((x) => x.ID_Proveedor === id);
  const idsLista = (v) => String(v || "").split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  // Activos, más los inactivos que el proyecto ya tenía (para no perderlos al editar).
  const proveedoresActivos = (actuales) => proveedores().filter((x) => x.Activo !== "No" || idsLista(actuales).includes(x.ID_Proveedor));
  const stakeholdersDe = (pid) => (S.datos.Stakeholders || []).filter((x) => x.ID_Proyecto === pid);
  const contacto = (correo, tel) => [correo ? `<a href="mailto:${esc(correo)}">${esc(correo)}</a>` : "", tel ? esc(tel) : ""].filter(Boolean).join("<br>") || "—";

  function cardStakeholders(p, puede) {
    const roles = S.cat.Rol_Stakeholder || [];
    const orden = (r) => { const i = roles.indexOf(r); return i < 0 ? 99 : i; };
    const lista = stakeholdersDe(p.ID_Proyecto).sort((a, b) => orden(a.Rol) - orden(b.Rol) || String(a.Nombre).localeCompare(String(b.Nombre), "es"));
    const faltan = roles.filter((r) => !lista.some((x) => x.Rol === r));
    const provs = idsLista(p.Proveedores).map(proveedor).filter(Boolean);
    return `<div class="card"><div class="titulo-fila"><h2>Stakeholders y proveedores</h2>${puede ? `<button class="btn" id="b-stk">+ Stakeholder</button>` : ""}</div>
      ${faltan.length ? `<div class="faltan-stk"><span class="sub">Sin asignar:</span> ${faltan.map((r) => puede ? `<button class="btn chico" data-rol-stk="${esc(r)}">+ ${esc(r)}</button>` : `<span class="pill noaplica">${esc(r)}</span>`).join(" ")}</div>` : ""}
      ${lista.length ? `<div class="tabla-scroll"><table><thead><tr><th>Rol</th><th>Nombre</th><th>Cargo / área</th><th>Empresa</th><th>Contacto</th><th></th></tr></thead>
        <tbody>${lista.map((x) => `<tr><td><b>${esc(x.Rol)}</b></td><td>${esc(x.Nombre)}</td><td>${esc([x.Cargo, x.Area].filter(Boolean).join(" · ") || "—")}</td>
          <td>${x.ID_Proveedor ? esc((proveedor(x.ID_Proveedor) || {}).Nombre || x.ID_Proveedor) : "Interno"}</td><td>${contacto(x.Correo, x.Telefono)}</td>
          <td class="derecha">${puede ? `<button class="btn chico" data-stk="${esc(x.ID_Stakeholder)}">Editar</button>` : ""}</td></tr>`).join("")}</tbody></table></div>` : ""}
      <div class="provs"><b>Proveedor(es):</b> ${provs.length ? provs.map((x) => `<span class="prov">${esc(x.Nombre)}${x.Servicio ? ` <span class="sub">· ${esc(x.Servicio)}</span>` : ""}${x.Contacto || x.Correo ? ` <span class="sub">· ${esc(x.Contacto || "")} ${x.Correo ? `<a href="mailto:${esc(x.Correo)}">${esc(x.Correo)}</a>` : ""}</span>` : ""}</span>`).join("")
        : `<span class="sub">Ninguno${puede ? " — asígnalos con «Editar proyecto»." : "."}</span>`}</div></div>`;
  }

  function formStakeholder(p, x, rolSugerido) {
    const nuevo = !x;
    const empresas = [["", "Interno (FSFB)"], ...proveedoresActivos(x && x.ID_Proveedor).map((v) => [v.ID_Proveedor, v.Nombre])];
    const campos = [
      { k: "Rol", label: "Rol en el proyecto", tipo: "select", opciones: S.cat.Rol_Stakeholder, ayuda: "Los roles se configuran en Catálogos." },
      { k: "Nombre", label: "Nombre", placeholder: "Nombre y apellido" },
      { k: "Cargo", label: "Cargo", placeholder: "Ej.: Director médico" },
      { k: "Area", label: "Área", tipo: "select", opciones: S.cat.Cliente_Area },
      { k: "ID_Proveedor", label: "Empresa", tipo: "select", opciones: empresas, ayuda: "Si la persona es del proveedor, elígelo aquí." },
      { k: "Correo", label: "Correo", tipo: "email", placeholder: "nombre@empresa.com" },
      { k: "Telefono", label: "Teléfono", placeholder: "Opcional" },
    ];
    modal(nuevo ? `Nuevo stakeholder · ${p.Nombre}` : `Editar stakeholder · ${x.Nombre || x.Rol}`, campos, nuevo ? { Rol: rolSugerido || "" } : x, (fd) => {
      if (nuevo) {
        const id = R.siguienteIdHijo("STK", S.datos.Stakeholders || [], "ID_Stakeholder", p.ID_Proyecto);
        guardar(() => S.api.agregarFila("Stakeholders", { ID_Stakeholder: id, ID_Proyecto: p.ID_Proyecto, ...fd }), "Stakeholder agregado");
      } else guardar(() => S.api.actualizarPorId("Stakeholders", "ID_Stakeholder", x.ID_Stakeholder, fd));
    }, null, nuevo ? {} : { eliminar: { texto: "Quitar del proyecto", mensaje: `${x.Nombre || "Esta persona"} dejará de aparecer como ${x.Rol || "stakeholder"} de este proyecto.`,
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

  function formProyecto(p) {
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
      { seccion: "5. Proveedores", ayuda: "márcalos o crea uno nuevo aquí mismo" },
      { k: "Proveedores", label: "Proveedor(es) del proyecto", tipo: "checks", vacio: "Aún no hay proveedores. Créalos en Catálogos → Proveedores.",
        opciones: proveedoresActivos(p ? p.Proveedores : "").map((x) => [x.ID_Proveedor, x.Nombre]) },
      { seccion: "6. Enlaces", ayuda: "opcionales; pega la dirección completa" },
      { k: "URL_Repositorio", label: "URL del repositorio", tipo: "url", placeholder: "https://github.com/… o https://dev.azure.com/…" },
      { k: "URL_Documentos", label: "URL de documentos", tipo: "url", placeholder: "https://…sharepoint.com/…" },
      { seccion: "7. Comentario" },
      { k: "Comentario_Estado", label: "Comentario de estado", tipo: "textarea", placeholder: "Novedad principal del proyecto" },
    ];
    const valores = p ? { ...p } : { PM: pmFijo ? u.Correo : "" };
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
        guardar(() => S.api.agregarFila("Proyectos", fila), `Iniciativa ${fila.ID_Proyecto} registrada`);
      } else {
        guardar(() => S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, { ...fd, ...sello() }));
      }
    }, (fd) => (fd.Fecha_Inicio && fd.Fecha_Fin_Plan && fd.Fecha_Fin_Plan < fd.Fecha_Inicio ? "La fecha fin no puede ser anterior a la fecha de inicio." : ""),
    !nuevo && R.puede(u, "verTodo") ? { init: initProv, eliminar: { texto: "Eliminar proyecto", titulo: `Eliminar ${p.ID_Proyecto}`,
      mensaje: `Se borrará «${p.Nombre}» con todos sus seguimientos, compromisos, hitos, riesgos, stakeholders y actas. No se puede deshacer. Si solo terminó, mejor cambia su estado a Cerrado o Cancelado.`,
      accion: () => { S.vista = "proyectos"; guardar(() => eliminarProyecto(p.ID_Proyecto), `Proyecto ${p.ID_Proyecto} eliminado`); } } } : { init: initProv });
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
    const contexto = `<div class="contexto">
      <div><span class="sub">Último reporte</span><b>${ult ? `${fecha(ult.Fecha_Corte)} · ${pct(ult.Avance_Real)} · ${esc(ult.Semaforo)}` : "Es el primer reporte"}</b></div>
      <div><span class="sub">Plan registrado</span><b>${pct(p.Avance_Planeado)}</b></div>
      <div><span class="sub">Frecuencia</span><b>${esc(p.Frecuencia_Seguimiento)}</b></div></div>`;
    const filaComp = (c = {}) => `<div class="comp-fila">
      <input class="c-texto" placeholder="Compromiso (qué se hará)" aria-label="Compromiso" value="${esc(c.Compromiso || "")}">
      <input class="c-resp" placeholder="Responsable" aria-label="Responsable" value="${esc(c.Responsable || "")}">
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
        Avance_Real: fd.Avance_Real, Semaforo: fd.Semaforo, Logros: fd.Logros, Proximos_Pasos: fd.Proximos_Pasos,
        Bloqueos: fd.Bloqueos, Reportado_Por: S.usuario.Correo, Fecha_Acta: fd.Fecha_Acta, URL_Acta: fd.URL_Acta,
        Acta_Archivo: acta ? acta.nombre : "",
      };
      const actaElegida = acta;
      const comps = extra.nuevos.map((c, i) => ({
        ID_Compromiso: R.siguienteIdHijo("CMP", S.datos.Compromisos, "ID_Compromiso", p.ID_Proyecto, i),
        ID_Proyecto: p.ID_Proyecto, ID_Seguimiento: idSeg, ...c, Estado: "Pendiente", Fecha_Cierre: "", Registrado_Por: S.usuario.Correo,
      }));
      guardar(async () => {
        await S.api.agregarFila("Seguimientos", fila);
        await S.api.actualizarPorId("Proyectos", "ID_Proyecto", p.ID_Proyecto, {
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}),
          ...(fd.Avance_Planeado !== "" ? { Avance_Planeado: fd.Avance_Planeado } : {}),
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
    const esc5 = [[1, "1 · Muy baja"], [2, "2 · Baja"], [3, "3 · Media"], [4, "4 · Alta"], [5, "5 · Muy alta"]];
    const campos = [
      { seccion: "1. Riesgo" },
      { k: "Tipo", label: "Tipo", tipo: "select", opciones: S.cat.Tipo_Riesgo, def: "Amenaza" },
      { k: "Estado", label: "Estado", tipo: "select", opciones: S.cat.Estado_Riesgo, def: "Abierto" },
      { k: "Responsable", label: "Responsable", def: S.usuario.Nombre || S.usuario.Correo },
      { k: "Descripcion", label: "Descripción del riesgo", tipo: "textarea", placeholder: "Si ocurre X, entonces Y" },
      { seccion: "2. Antes de mitigar (inherente)" },
      { k: "Probabilidad_Inherente", label: "Probabilidad", tipo: "select", opciones: esc5 },
      { k: "Impacto_Inherente", label: "Impacto", tipo: "select", opciones: esc5 },
      { seccion: "3. Planes" },
      { k: "Plan_Mitigacion", label: "Plan de mitigación", tipo: "textarea", placeholder: "Qué hacemos para que no ocurra o afecte menos" },
      { k: "Plan_Contingencia", label: "Plan de contingencia", tipo: "textarea", placeholder: "Qué hacemos si ocurre" },
      { seccion: "4. Después de mitigar (residual)" },
      { k: "Probabilidad_Residual", label: "Probabilidad", tipo: "select", opciones: esc5 },
      { k: "Impacto_Residual", label: "Impacto", tipo: "select", opciones: esc5 },
    ];
    modal(r ? "Editar riesgo" : `Nuevo riesgo · ${p.Nombre}`, campos, r || {}, (fd) => {
      ["Probabilidad_Inherente", "Impacto_Inherente", "Probabilidad_Residual", "Impacto_Residual"].forEach((k) => { fd[k] = Number(fd[k]); });
      fd.Calificacion_Inherente = R.calificacion(fd.Probabilidad_Inherente, fd.Impacto_Inherente);
      fd.Calificacion_Residual = R.calificacion(fd.Probabilidad_Residual, fd.Impacto_Residual);
      if (r) guardar(() => S.api.actualizarPorId("Riesgos", "ID_Riesgo", r.ID_Riesgo, fd));
      else guardar(() => S.api.agregarFila("Riesgos", { ID_Riesgo: R.siguienteIdHijo("RSG", S.datos.Riesgos, "ID_Riesgo", p.ID_Proyecto), ID_Proyecto: p.ID_Proyecto, ...fd }), "Riesgo agregado");
    }, null, r ? { eliminar: { texto: "Eliminar riesgo", mensaje: `Se borrará el riesgo «${r.Descripcion}».`, accion: () => guardar(() => S.api.eliminarFilas("Riesgos", "ID_Riesgo", [r.ID_Riesgo]), "Riesgo eliminado") } } : {});
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
    await S.api.eliminarFilas("Proyectos", "ID_Proyecto", [pid]);
  }

  function formCompromiso(c, pNuevo, segSugerida) {
    const nuevo = !c;
    const pid = nuevo ? pNuevo.ID_Proyecto : c.ID_Proyecto;
    const sesiones = R.seguimientosDe(pid, S.datos.Seguimientos).slice().reverse()
      .map((s) => [s.ID_Seguimiento, `Sesión del ${fecha(s.Fecha_Corte)}${s.Fecha_Acta && s.Fecha_Acta !== s.Fecha_Corte ? ` (acta ${fecha(s.Fecha_Acta)})` : ""}`]);
    const gente = [...S.datos.Usuarios.filter((u) => u.Activo === "Sí").map((u) => [u.Nombre || u.Correo, u.Correo]),
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
      { k: "Responsable", label: "Responsable", sugerencias: personas, placeholder: "Escribe o elige" },
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
          ...(fd.Avance_Real !== "" ? { Avance_Real: fd.Avance_Real } : {}), ...(fd.Semaforo ? { Semaforo: fd.Semaforo } : {}), ...sello() });
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
      S.api = await DATOS.crear(S.modo);
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
