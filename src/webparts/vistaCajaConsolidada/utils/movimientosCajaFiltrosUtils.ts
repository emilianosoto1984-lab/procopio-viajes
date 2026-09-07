import { normalizarMonedaPago } from '../../../shared/pagoMonedaUtils';
import { toDateInput } from '../../../shared/sharePointDateUtils';
import { IFiltrosMovimientosCaja } from '../models/IFiltrosMovimientosCaja';
import { IMovimientoCajaLinea } from './cajaMovimientosUtils';

function pad2(value: number): string {
  return value < 10 ? '0' + value : String(value);
}

export function toLocalDateInput(date: Date): string {
  return date.getFullYear() + '-' + pad2(date.getMonth() + 1) + '-' + pad2(date.getDate());
}

/**
 * Período inicial: hace 3 meses hasta hoy (fecha local).
 */
export function createDefaultFiltrosMovimientosCaja(): IFiltrosMovimientosCaja {
  const now = new Date();
  return {
    fechaDesde: toLocalDateInput(
      new Date(now.getFullYear(), now.getMonth() - 3, now.getDate())
    ),
    fechaHasta: toLocalDateInput(now),
    moneda: ''
  };
}

export const MENSAJE_PERIODO_MOVIMIENTOS_INVALIDO =
  'La fecha hasta no puede ser menor que la fecha desde.';

export function esPeriodoMovimientosInvalido(filtros: IFiltrosMovimientosCaja): boolean {
  const desde = (filtros && filtros.fechaDesde ? filtros.fechaDesde : '').trim();
  const hasta = (filtros && filtros.fechaHasta ? filtros.fechaHasta : '').trim();
  return !!(desde && hasta && hasta < desde);
}

/**
 * Filtro inclusivo por FechaPago y moneda.
 * FechaDesde <= FechaPago <= FechaHasta. moneda vacío = todas.
 */
export function filtrarMovimientosCaja(
  movimientos: IMovimientoCajaLinea[],
  filtros: IFiltrosMovimientosCaja
): IMovimientoCajaLinea[] {
  const desde = (filtros && filtros.fechaDesde ? filtros.fechaDesde : '').trim();
  const hasta = (filtros && filtros.fechaHasta ? filtros.fechaHasta : '').trim();
  const moneda = (filtros && filtros.moneda ? filtros.moneda : '').trim();
  return (movimientos || []).filter((item: IMovimientoCajaLinea) => {
    const fecha = toDateInput(item.fechaPago || '');
    if (!fecha) {
      return false;
    }
    if (desde && fecha < desde) {
      return false;
    }
    if (hasta && fecha > hasta) {
      return false;
    }
    if (moneda && normalizarMonedaPago(item.moneda || '') !== moneda) {
      return false;
    }
    return true;
  });
}
