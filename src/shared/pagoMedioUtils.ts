/**
 * Constantes y reglas de medio de pago / estado / cuenta bancaria / motivo
 * compartidas entre Registro de Pagos y Registro de Viajes.
 */

/** Choice canónico de Motivo cuando el pago está asociado a un viaje. */
export const MOTIVO_VIAJE = 'Viaje';

/**
 * Resuelve Motivo = "Viaje" solo si existe en las opciones Choice de SharePoint.
 * No inventa otro valor si la opción no está configurada.
 */
export function resolverMotivoViaje(opcionesMotivo: string[]): string {
  if ((opcionesMotivo || []).indexOf(MOTIVO_VIAJE) >= 0) {
    return MOTIVO_VIAJE;
  }
  return '';
}

export const MEDIOS_PAGO = ['Efectivo', 'Transferencia', 'Tarjeta de Credito'] as const;
export type MedioPago = (typeof MEDIOS_PAGO)[number];

export const MEDIO_PAGO_TRANSFERENCIA: MedioPago = 'Transferencia';
export const MEDIO_PAGO_EFECTIVO: MedioPago = 'Efectivo';
export const MEDIO_PAGO_TARJETA: MedioPago = 'Tarjeta de Credito';

/** True cuando el medio exige selector de Cuenta Bancaria. */
export function requiereCuentaBancaria(medioPago: string): boolean {
  return (medioPago || '').trim() === MEDIO_PAGO_TRANSFERENCIA;
}

/**
 * Estado inicial del pago según medio (misma regla que Registro de Pagos).
 * Acepta la variante con tilde por compatibilidad con datos legacy de Viajes.
 */
export function getEstadoByMedioPago(medioPago: string): string {
  switch ((medioPago || '').trim()) {
    case MEDIO_PAGO_EFECTIVO:
      return 'Aprobado';
    case MEDIO_PAGO_TRANSFERENCIA:
    case MEDIO_PAGO_TARJETA:
    case 'Tarjeta de Crédito':
      return 'Pendiente';
    default:
      return 'Pendiente';
  }
}

/**
 * Normaliza el valor de medio de pago al Choice canónico de SharePoint / Pagos.
 * Unifica la tilde de Tarjeta de Crédito → Tarjeta de Credito.
 */
export function normalizarMedioPago(medioPago: string): string {
  const trimmed = (medioPago || '').trim();
  if (trimmed === 'Tarjeta de Crédito' || trimmed === MEDIO_PAGO_TARJETA) {
    return MEDIO_PAGO_TARJETA;
  }
  if (trimmed === MEDIO_PAGO_EFECTIVO || trimmed === MEDIO_PAGO_TRANSFERENCIA) {
    return trimmed;
  }
  return trimmed;
}

/**
 * En edición: conserva Estado si el medio no cambió y ya había uno;
 * si el medio cambió o el Estado legacy está vacío, recalcula.
 */
export function resolverEstadoPagoAlGuardar(
  medioPagoActual: string,
  medioPagoOriginal: string | undefined,
  estadoOriginal: string | undefined,
  esEdicion: boolean
): string {
  const medioActual = normalizarMedioPago(medioPagoActual);
  const medioOriginal = normalizarMedioPago(medioPagoOriginal || '');
  const estadoPrev = (estadoOriginal || '').trim();

  if (esEdicion && medioActual === medioOriginal && estadoPrev) {
    return estadoPrev;
  }
  return getEstadoByMedioPago(medioActual);
}
