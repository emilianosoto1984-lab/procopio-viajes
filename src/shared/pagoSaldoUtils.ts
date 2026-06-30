import {
  normalizarMonedaPago,
  normalizarMonedaServicio
} from './pagoMonedaUtils';

export interface IPagoSaldoItem {
  id: number;
  tipoPago: string;
  monto: number;
  moneda: string;
  cotizacion?: number;
  servicioAsociadoId?: number;
  concepto?: string;
}

export interface IServicioSaldoItem {
  id: number;
  concepto: string;
  precioCliente: number;
  moneda: string;
}

/**
 * Convierte un ingreso a la moneda del servicio asociado.
 * Cotización se interpreta como ARS por 1 USD (misma regla que Registro de Viajes).
 */
export function convertirIngresoAMonedaServicio(
  monto: number,
  monedaPago: string,
  cotizacion: number | undefined,
  monedaServicio: string
): number {
  const montoNumerico = Number(monto) || 0;
  if (montoNumerico <= 0) {
    return 0;
  }

  const monedaPagoNormalizada = normalizarMonedaPago(monedaPago);
  const monedaServicioNormalizada = normalizarMonedaServicio(monedaServicio);
  if (monedaPagoNormalizada === monedaServicioNormalizada) {
    return montoNumerico;
  }

  const cotizacionNumerica = Number(cotizacion);
  if (!isFinite(cotizacionNumerica) || cotizacionNumerica <= 0) {
    return 0;
  }

  if (monedaPagoNormalizada === 'Pesos' && monedaServicioNormalizada === 'Dólares') {
    return montoNumerico / cotizacionNumerica;
  }
  if (monedaPagoNormalizada === 'Dólares' && monedaServicioNormalizada === 'Pesos') {
    return montoNumerico * cotizacionNumerica;
  }

  return montoNumerico;
}

export function pagoMatchesServicio(pago: IPagoSaldoItem, servicio: IServicioSaldoItem): boolean {
  if (pago.servicioAsociadoId && pago.servicioAsociadoId > 0) {
    return pago.servicioAsociadoId === servicio.id;
  }
  return (pago.concepto || '').trim() === (servicio.concepto || '').trim();
}

export function getTotalIngresosPorServicioEnMonedaServicio(
  servicio: IServicioSaldoItem,
  pagos: IPagoSaldoItem[],
  pagoIdExcluir?: number
): number {
  return pagos
    .filter(
      (pago: IPagoSaldoItem) =>
        (pago.tipoPago || '').trim() === 'Ingreso' &&
        pago.monto > 0 &&
        pagoMatchesServicio(pago, servicio) &&
        (pagoIdExcluir === undefined || pago.id !== pagoIdExcluir)
    )
    .reduce(
      (acc: number, pago: IPagoSaldoItem) =>
        acc +
        convertirIngresoAMonedaServicio(
          pago.monto,
          pago.moneda,
          pago.cotizacion,
          servicio.moneda
        ),
      0
    );
}

export function getSaldoPendienteServicio(
  servicio: IServicioSaldoItem,
  pagos: IPagoSaldoItem[],
  pagoIdExcluir?: number
): number {
  const presupuestoServicio = Number(servicio.precioCliente) || 0;
  const totalIngresado = getTotalIngresosPorServicioEnMonedaServicio(servicio, pagos, pagoIdExcluir);
  return presupuestoServicio - totalIngresado;
}

export function montoExcedeSaldoPendiente(
  servicio: IServicioSaldoItem,
  pagos: IPagoSaldoItem[],
  montoPago: number,
  monedaPago: string,
  cotizacion: number | undefined,
  pagoIdExcluir?: number
): boolean {
  const totalYaIngresado = getTotalIngresosPorServicioEnMonedaServicio(servicio, pagos, pagoIdExcluir);
  const ingresoConvertido = convertirIngresoAMonedaServicio(
    montoPago,
    monedaPago,
    cotizacion,
    servicio.moneda
  );
  const presupuestoServicio = Number(servicio.precioCliente) || 0;
  return totalYaIngresado + ingresoConvertido > presupuestoServicio;
}

export function formatMontoMoneda(monto: number, moneda: string): string {
  if (normalizarMonedaServicio(moneda) === 'Dólares') {
    return monto.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  }
  return monto.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
}

export function formatSaldoPendienteDisplay(
  servicio: IServicioSaldoItem,
  saldoPendiente: number,
  monedaPago?: string,
  cotizacion?: number
): string {
  const monedaServicio = normalizarMonedaServicio(servicio.moneda);
  const saldoTexto = formatMontoMoneda(saldoPendiente, servicio.moneda);
  let display = 'Saldo pendiente: ' + saldoTexto;

  if (
    monedaPago &&
    monedaServicio === 'Dólares' &&
    normalizarMonedaPago(monedaPago) === 'Pesos'
  ) {
    const cotizacionNumerica = Number(cotizacion);
    if (isFinite(cotizacionNumerica) && cotizacionNumerica > 0) {
      const equivalentePesos = saldoPendiente * cotizacionNumerica;
      display +=
        ' (equiv. ' +
        equivalentePesos.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) +
        ' en Pesos)';
    }
  }

  return display;
}
