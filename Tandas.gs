/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Tandas.gs
 *  Qué reporte descargar hoy, con qué rango, y qué avisos salen cada día.
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  Regla: "5 días" y "1 día" son DÍAS HÁBILES contados contra la fecha de pago
 *  efectiva (la fecha nominal pasada al siguiente día hábil). Un vencimiento en
 *  domingo se paga el lunes y se avisa como lunes.
 *
 *  Si hoy (día D0) se descarga el Rep1 y la tanda dura L días hábiles, el envío
 *  del día k (k = 0..L-1) cubre:
 *      aviso 1: pagos que caen en D0 + k + 5 días hábiles
 *      aviso 2: pagos que caen en D0 + k + 1 día hábil
 *  Por eso el Rep1 debe traer los pagos efectivos de D0+1 hábil a D0+(L+4) hábiles.
 *  Con L = 1 son los siguientes 5 días hábiles; el rango en fechas nominales
 *  incluye sábado y domingo cuando corresponde (por eso a veces son 7 días).
 */

function diasAviso_() {
  return {
    a1: Math.max(1, Math.round(cfgNum_('DIAS_TIPO_T_MENOS_5', 5))),
    a2: Math.max(1, Math.round(cfgNum_('DIAS_TIPO_T_MENOS_1', 1))),
    separacion: Math.max(0, Math.round(cfgNum_('SEPARACION_MINIMA_DIAS', 2)))
  };
}

function longitudTanda_() {
  return Math.max(1, Math.round(cfgNum_('TANDA_DIAS_HABILES', 1)));
}

/** Día hábil que se toma como D0: hoy si es hábil, si no el siguiente. */
function diaBaseTanda_(hoy) { return esHabil_(hoy) ? hoy : siguienteHabil_(hoy); }

/** ¿D0 es día de inicio de tanda (toca descargar Rep1)? */
function esInicioTanda_(d0, L, anchor) {
  if (L <= 1) return true;
  if (!anchor || !esHabil_(anchor)) return true;       // sin ancla válida: se avisa todos los días
  const n = habilesEntre_(anchor, d0);
  return n >= 0 && n % L === 0;
}

/** Ventana de descarga para una tanda que inicia en d0 y dura L días hábiles. */
function ventanaTanda_(d0, L, a1, a2) {
  const aviso1 = a1 || 5, aviso2 = a2 || 1;
  const pagoDesde = sumarHabiles_(d0, Math.min(aviso1, aviso2));
  const pagoHasta = sumarHabiles_(d0, (L - 1) + Math.max(aviso1, aviso2));
  const nominalDesde = addDays_(anteriorHabil_(addDays_(pagoDesde, -1)), 1);
  const nominalHasta = pagoHasta;
  const dias = [];
  let envio = d0;
  for (let k = 0; k < L; k++) {
    if (k > 0) envio = sumarHabiles_(envio, 1);
    dias.push({ envio: envio, aviso1Pago: sumarHabiles_(envio, aviso1), aviso2Pago: sumarHabiles_(envio, aviso2) });
  }
  return { d0: d0, longitud: L, pagoDesde: pagoDesde, pagoHasta: pagoHasta,
           nominalDesde: nominalDesde, nominalHasta: nominalHasta, dias: dias };
}

/** Línea de tiempo para la "regla" del front: desde hoy hasta pagoHasta + 2 días naturales. */
function lineaDeTiempo_(hoy, d0, vent, a1, a2) {
  const fin = addDays_(vent.pagoHasta, 2);
  const out = [];
  for (let k = hoy; k <= fin; k = addDays_(k, 1)) {
    const hab = esHabil_(k);
    const item = {
      fecha: k, dow: dow_(k), habil: hab, motivo: hab ? '' : motivoInhabil_(k),
      idx: (hab && k >= d0) ? habilesEntre_(d0, k) : null,
      esHoy: k === hoy, enRep1: k >= vent.nominalDesde && k <= vent.nominalHasta,
      aviso1: false, aviso2: false
    };
    vent.dias.forEach(dd => {
      if (dd.aviso1Pago === k) item.aviso1 = dd.envio === d0;
      if (dd.aviso2Pago === k) item.aviso2 = dd.envio === d0;
    });
    out.push(item);
  }
  return out;
}

/** Estado completo de la tanda de hoy (lo consume el front). */
function estadoTanda_() {
  const hoy = hoyKey_();
  const L = longitudTanda_();
  const dias = diasAviso_();
  const d0 = diaBaseTanda_(hoy);
  const anchor = cfgStr_('TANDA_ANCHOR_FECHA', '');
  const vent = ventanaTanda_(d0, L, dias.a1, dias.a2);
  return {
    hoy: hoy,
    esHabil: esHabil_(hoy),
    motivoInhabil: motivoInhabil_(hoy),
    d0: d0,
    longitud: L,
    inicioTanda: esHabil_(hoy) && esInicioTanda_(d0, L, anchor),
    sinAncla: L > 1 && (!anchor || !esHabil_(anchor)),
    ventana: vent,
    dias: dias,
    linea: lineaDeTiempo_(hoy, d0, vent, dias.a1, dias.a2)
  };
}

/** aaaa-mm-dd → dd/mm/aaaa (para copiar al filtro del reporte). */
function fechaDMA_(k) { return k.slice(8, 10) + '/' + k.slice(5, 7) + '/' + k.slice(0, 4); }
