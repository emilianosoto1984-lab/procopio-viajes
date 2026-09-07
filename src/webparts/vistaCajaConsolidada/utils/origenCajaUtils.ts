import { getCuentaBancariaLabel, ICuentaBancaria } from '../../../shared/cuentasBancariasUtils';
import { MEDIO_PAGO_EFECTIVO, normalizarMedioPago } from '../../../shared/pagoMedioUtils';
import { OrigenCaja } from '../models/OrigenCaja';
import { IPagoParaCaja } from './cajaMovimientosUtils';

export const ORIGEN_EFECTIVO: OrigenCaja = {
  tipo: 'efectivo',
  id: 'efectivo',
  nombre: 'Efectivo'
};

export function getOrigenCajaKey(origen: OrigenCaja): string {
  if (origen.tipo === 'efectivo') {
    return 'efectivo';
  }
  return 'banco:' + String(origen.id);
}

/**
 * Efectivo siempre primero; las cuentas usan el label de la lista, sin nombres fijos.
 */
export function construirOrigenesCaja(cuentas: ICuentaBancaria[]): OrigenCaja[] {
  const origenes: OrigenCaja[] = [ORIGEN_EFECTIVO];
  (cuentas || []).forEach((cuenta: ICuentaBancaria) => {
    if (!cuenta || !(cuenta.id > 0)) {
      return;
    }
    origenes.push({
      tipo: 'banco',
      id: cuenta.id,
      nombre: getCuentaBancariaLabel(cuenta)
    });
  });
  return origenes;
}

export function findOrigenCajaByKey(
  origenes: OrigenCaja[],
  key: string
): OrigenCaja | undefined {
  const valor = (key || '').trim();
  if (!valor) {
    return undefined;
  }
  return (origenes || []).filter((origen: OrigenCaja) => getOrigenCajaKey(origen) === valor)[0];
}

export function esPagoOrigenEfectivo(pago: IPagoParaCaja): boolean {
  return normalizarMedioPago(pago.medioPago || '') === MEDIO_PAGO_EFECTIVO;
}

export function esPagoOrigenBanco(pago: IPagoParaCaja, cuentaId: number): boolean {
  if (!(cuentaId > 0) || esPagoOrigenEfectivo(pago)) {
    return false;
  }
  return (pago.cuentaBancariaId || 0) === cuentaId;
}

/**
 * Efectivo: MedioPago = Efectivo.
 * Banco: lookup CuentaBancaria.Id = cuenta seleccionada (excluye Efectivo).
 */
export function filtrarPagosPorOrigenCaja(
  pagos: IPagoParaCaja[],
  origen: OrigenCaja
): IPagoParaCaja[] {
  if (!origen) {
    return [];
  }
  if (origen.tipo === 'efectivo') {
    return (pagos || []).filter((pago: IPagoParaCaja) => esPagoOrigenEfectivo(pago));
  }
  return (pagos || []).filter((pago: IPagoParaCaja) => esPagoOrigenBanco(pago, origen.id));
}
