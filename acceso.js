/*
 * Control de acceso del sitio Viajes Horizonte.
 *
 * Los precios y los documentos/recursos solo se muestran a quien ya llenó el
 * cuestionario (nombre, correo y por qué quiere viajar). El registro se guarda
 * en el navegador de la persona.
 *
 * IMPORTANTE: el sitio es estático (GitHub Pages), así que esto es un filtro
 * "suave": alguien con conocimientos técnicos podría saltárselo.
 *
 * Para RECIBIR las respuestas, pon aquí la dirección de un servicio de
 * formularios, por ejemplo:
 *   Formspree:  'https://formspree.io/f/TU_ID'
 *   FormSubmit: 'https://formsubmit.co/ajax/tu-correo@ejemplo.com'
 * Si se deja vacío, las respuestas solo se guardan en el navegador del visitante.
 */
var ENDPOINT_FORMULARIO = '';

var Acceso = (function () {
  var CLAVE = 'viajes-horizonte-registro';

  function leer() {
    try {
      var r = JSON.parse(localStorage.getItem(CLAVE) || 'null');
      if (r) return r;
    } catch (e) {}
    try {
      return JSON.parse(sessionStorage.getItem(CLAVE) || 'null');
    } catch (e) { return null; }
  }

  function guardar(datos) {
    var texto = JSON.stringify(datos);
    try { localStorage.setItem(CLAVE, texto); return; } catch (e) {}
    try { sessionStorage.setItem(CLAVE, texto); } catch (e) {}
  }

  function almacenamientoDisponible() {
    try { localStorage.setItem('_p', '1'); localStorage.removeItem('_p'); return true; } catch (e) {}
    try { sessionStorage.setItem('_p', '1'); sessionStorage.removeItem('_p'); return true; } catch (e) {}
    return false;
  }

  // Ruta al cuestionario desde la página actual (raíz o subcarpeta).
  function rutaCuestionario() {
    var s = document.querySelector('script[src$="acceso.js"]');
    var base = s ? s.getAttribute('src').replace(/acceso\.js$/, '') : '';
    return base + 'cuestionario.html';
  }

  function automatizado() {
    try { return navigator.webdriver === true; } catch (e) { return false; }
  }

  function urlCuestionario() {
    var volver = location.pathname + location.search + location.hash;
    return rutaCuestionario() + '?volver=' + encodeURIComponent(volver);
  }

  return {
    registro: leer,
    // Navegador controlado por un programa (Playwright, Selenium, agentes de IA…).
    automatizado: automatizado,
    // Si el navegador no permite guardar nada, no bloqueamos a nadie.
    // La respuesta a la pregunta para IA es obligatoria para desbloquear el contenido.
    registrado: function () {
      var r = leer();
      if (r) return !!r.asistente_ia;
      return !almacenamientoDisponible();
    },
    guardar: guardar,
    urlCuestionario: urlCuestionario,
    enviar: function (datos) {
      if (!ENDPOINT_FORMULARIO) return Promise.resolve(false);
      var cuerpo = new FormData();
      Object.keys(datos).forEach(function (k) { cuerpo.append(k, datos[k]); });
      cuerpo.append('_subject', 'Nuevo registro – Viajes Horizonte');
      return fetch(ENDPOINT_FORMULARIO, {
        method: 'POST', body: cuerpo, headers: { Accept: 'application/json' }
      }).then(function (r) { return r.ok; }).catch(function () { return false; });
    }
  };
})();
