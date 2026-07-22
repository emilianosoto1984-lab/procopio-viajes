import { getDateOnlyFromSharePoint } from './sharePointDateUtils';
import { normalizarMonedaPago } from './pagoMonedaUtils';

export type MotivoGeneracionRecibo = 'creacion' | 'aprobacion' | 'edicion';

/**
 * Reglas de cuándo generar o regenerar el recibo PDF (Registro de Pagos / Viajes).
 * - Egreso: nunca
 * - Creación: solo Ingreso con Estado Aprobado
 * - Aprobación: solo Ingreso
 * - Edición: solo Ingreso con Estado final Aprobado (regenera PDF existente)
 */
export function debeGenerarReciboPago(params: {
  tipoPago: string;
  estado: string;
  motivo: MotivoGeneracionRecibo;
}): boolean {
  const tipo = (params.tipoPago || '').trim();
  if (tipo !== 'Ingreso') {
    return false;
  }
  if (params.motivo === 'aprobacion') {
    return true;
  }
  // creacion y edicion: solo si el estado final es Aprobado
  return (params.estado || '').trim().toLowerCase() === 'aprobado';
}

export function buildNumeroRecibo(itemId: number, fechaPago: string): string {
  const year = _extractYearFromFechaPago(fechaPago);
  return 'RP-' + year + '-' + itemId;
}

export function buildReciboFileName(numeroRecibo: string): string {
  return 'Recibo-' + numeroRecibo + '.pdf';
}

export function formatMontoRecibo(monto: number, moneda: string): string {
  const valor = isNaN(monto) ? 0 : monto;
  const formatted = valor.toLocaleString('es-AR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return normalizarMonedaPago(moneda) === 'Dólares' ? 'USD ' + formatted : '$ ' + formatted;
}

function _extractYearFromFechaPago(fechaPago: string): number {
  const datePart = getDateOnlyFromSharePoint(fechaPago);
  if (datePart.length >= 4) {
    const year = parseInt(datePart.substring(0, 4), 10);
    if (!isNaN(year)) {
      return year;
    }
  }
  return new Date().getFullYear();
}
