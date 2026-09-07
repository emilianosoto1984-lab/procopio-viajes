import {
  esPagoAprobado,
  esPagoPendiente,
  esPagoSinEstado
} from '../../../shared/pagoSaldoUtils';
import { IMovimientoCajaLinea } from './cajaMovimientosUtils';

export type VistaEstadoMovimientosCaja = 'aprobados' | 'pendientes';

export function contarMovimientosCajaAprobados(items: IMovimientoCajaLinea[]): number {
  return (items || []).filter(esPagoAprobado).length;
}

export function contarMovimientosCajaPendientes(items: IMovimientoCajaLinea[]): number {
  return (items || []).filter(esPagoPendiente).length;
}

export function contarMovimientosCajaSinEstado(items: IMovimientoCajaLinea[]): number {
  return (items || []).filter(esPagoSinEstado).length;
}

/**
 * Aprobados: Estado Aprobado + históricos sin estado.
 * Pendientes: solo Estado Pendiente.
 */
export function getMovimientosCajaDeVista(
  items: IMovimientoCajaLinea[],
  vista: VistaEstadoMovimientosCaja
): IMovimientoCajaLinea[] {
  if (vista === 'aprobados') {
    return (items || []).filter(
      (item: IMovimientoCajaLinea) => esPagoAprobado(item) || esPagoSinEstado(item)
    );
  }
  return (items || []).filter(esPagoPendiente);
}

export function resolverVistaMovimientosCajaTrasCambio(
  vistaActual: VistaEstadoMovimientosCaja,
  items: IMovimientoCajaLinea[]
): VistaEstadoMovimientosCaja {
  if (vistaActual === 'pendientes' && contarMovimientosCajaPendientes(items) === 0) {
    return 'aprobados';
  }
  return vistaActual;
}
