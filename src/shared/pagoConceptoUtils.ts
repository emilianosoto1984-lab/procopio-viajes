/**
 * Formato unificado de Concepto para pagos asociados a un viaje
 * (Registro de Pagos y Registro de Viajes).
 *
 * "{Servicio|Liquidación} - {Nombre del viaje}"
 */

export function buildConceptoPagoConViaje(nombreBase: string, nombreViaje: string): string {
  const base = (nombreBase || '').trim();
  const viaje = (nombreViaje || '').trim();
  let conceptoGenerado = '';
  if (!base) {
    conceptoGenerado = '';
  } else if (!viaje) {
    conceptoGenerado = base;
  } else {
    conceptoGenerado = base + ' - ' + viaje;
  }

  // TEMP diagnóstico Concepto — quitar cuando se confirme la causa
  console.log('[Concepto pago] servicio:', base);
  console.log('[Concepto pago] viaje:', viaje);
  console.log('[Concepto pago] concepto generado:', conceptoGenerado);
  if (base && !viaje) {
    console.warn(
      '[Concepto pago] nombreViaje vacío: el concepto quedará solo con el servicio'
    );
  }

  return conceptoGenerado;
}

/**
 * True si el concepto guardado corresponde al servicio
 * (exacto, con sufijo de viaje, o legacy solo nombre de servicio).
 */
export function servicioCoincideConConceptoGuardado(
  conceptoServicio: string,
  conceptoGuardado: string,
  nombreViaje?: string
): boolean {
  const servicio = (conceptoServicio || '').trim();
  const guardado = (conceptoGuardado || '').trim();
  if (!servicio || !guardado) {
    return false;
  }
  if (servicio === guardado) {
    return true;
  }
  const conViaje = buildConceptoPagoConViaje(servicio, nombreViaje || '');
  if (conViaje && conViaje === guardado) {
    return true;
  }
  return guardado.indexOf(servicio + ' - ') === 0;
}
