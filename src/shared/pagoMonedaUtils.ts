export type MonedaPago = 'Dólares' | 'Pesos';
export type MonedaServicio = 'Dólares' | 'Pesos';

const MONEDAS_PAGO: MonedaPago[] = ['Dólares', 'Pesos'];
const MONEDAS_SERVICIO: MonedaServicio[] = ['Dólares', 'Pesos'];

export function normalizarMonedaPago(moneda: string): MonedaPago {
  return MONEDAS_PAGO.indexOf(moneda as MonedaPago) >= 0 ? (moneda as MonedaPago) : 'Pesos';
}

export function normalizarMonedaServicio(moneda: string): MonedaServicio {
  return MONEDAS_SERVICIO.indexOf(moneda as MonedaServicio) >= 0 ? (moneda as MonedaServicio) : 'Pesos';
}

/**
 * Misma regla que _requiereCotizacionMovimiento en Registro de Viajes:
 * cotización obligatoria cuando la moneda del servicio difiere de la moneda del pago.
 */
export function requiereCotizacionPago(monedaServicio: string, monedaPago: string): boolean {
  const monedaRelacionada = (monedaServicio || '').trim();
  if (!monedaRelacionada) {
    return false;
  }
  return normalizarMonedaServicio(monedaRelacionada) !== normalizarMonedaPago(monedaPago);
}
