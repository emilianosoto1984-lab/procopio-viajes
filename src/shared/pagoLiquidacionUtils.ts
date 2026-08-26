import {
  convertirIngresoAMonedaServicio,
  formatMontoMoneda,
  IPagoSaldoItem,
  isPagoConsideradoEnTotales
} from './pagoSaldoUtils';
import { normalizarMonedaServicio } from './pagoMonedaUtils';

/** Liquidación de operador tal como se usa en formularios de pago. */
export interface ILiquidacionOperador {
  id: number;
  /** Código de referencia (columna Title en SharePoint). */
  title: string;
  viajeId?: number;
  operadorId?: number;
  operadorNombre?: string;
  moneda: string;
  importeTotal: number;
  importePagado?: number;
  saldoPendiente?: number;
}

export function pagoMatchesLiquidacion(
  pago: IPagoSaldoItem,
  liquidacionId: number
): boolean {
  if (!liquidacionId || liquidacionId <= 0) {
    return false;
  }
  if (pago.liquidacionOperadorId && pago.liquidacionOperadorId > 0) {
    return pago.liquidacionOperadorId === liquidacionId;
  }
  return false;
}

/**
 * Suma egresos aprobados (y históricos sin estado) asociados a una liquidación,
 * convertidos a la moneda de la liquidación.
 */
export function getTotalEgresosPorLiquidacionEnMonedaLiquidacion(
  liquidacion: Pick<ILiquidacionOperador, 'id' | 'moneda'>,
  pagos: IPagoSaldoItem[],
  pagoIdExcluir?: number
): number {
  return pagos
    .filter(
      (pago: IPagoSaldoItem) =>
        isPagoConsideradoEnTotales(pago) &&
        (pago.tipoPago || '').trim() === 'Egreso' &&
        (Number(pago.monto) || 0) > 0 &&
        pagoMatchesLiquidacion(pago, liquidacion.id) &&
        (pagoIdExcluir === undefined || pago.id !== pagoIdExcluir)
    )
    .reduce(
      (acc: number, pago: IPagoSaldoItem) =>
        acc +
        convertirIngresoAMonedaServicio(
          Number(pago.monto) || 0,
          pago.moneda,
          pago.cotizacion,
          liquidacion.moneda
        ),
      0
    );
}

export function getSaldoPendienteLiquidacion(
  liquidacion: Pick<ILiquidacionOperador, 'id' | 'moneda' | 'importeTotal'>,
  pagos: IPagoSaldoItem[],
  pagoIdExcluir?: number
): number {
  const totalLiquidacion = Number(liquidacion.importeTotal) || 0;
  const totalPagado = getTotalEgresosPorLiquidacionEnMonedaLiquidacion(
    liquidacion,
    pagos,
    pagoIdExcluir
  );
  return totalLiquidacion - totalPagado;
}

export function enriquecerLiquidacionConSaldo(
  liquidacion: ILiquidacionOperador,
  pagos: IPagoSaldoItem[],
  pagoIdExcluir?: number
): ILiquidacionOperador {
  const importePagado = getTotalEgresosPorLiquidacionEnMonedaLiquidacion(
    liquidacion,
    pagos,
    pagoIdExcluir
  );
  const saldoPendiente = getSaldoPendienteLiquidacion(liquidacion, pagos, pagoIdExcluir);
  return {
    ...liquidacion,
    importePagado,
    saldoPendiente
  };
}

export function filtrarLiquidacionesDisponibles(
  liquidaciones: ILiquidacionOperador[],
  pagos: IPagoSaldoItem[],
  liquidacionIncluidaId?: number,
  pagoIdExcluir?: number
): ILiquidacionOperador[] {
  const incluida = liquidacionIncluidaId && liquidacionIncluidaId > 0 ? liquidacionIncluidaId : 0;
  return liquidaciones
    .map((liquidacion: ILiquidacionOperador) =>
      enriquecerLiquidacionConSaldo(liquidacion, pagos, pagoIdExcluir)
    )
    .filter((liquidacion: ILiquidacionOperador) => {
      if (incluida > 0 && liquidacion.id === incluida) {
        return true;
      }
      return (liquidacion.saldoPendiente ?? 0) > 0.0001;
    });
}

export function montoExcedeSaldoPendienteLiquidacion(
  liquidacion: Pick<ILiquidacionOperador, 'id' | 'moneda' | 'importeTotal'>,
  pagos: IPagoSaldoItem[],
  montoEgreso: number,
  monedaPago: string,
  cotizacion: number | undefined,
  pagoIdExcluir?: number
): boolean {
  const saldoPendiente = getSaldoPendienteLiquidacion(liquidacion, pagos, pagoIdExcluir);
  const egresoConvertido = convertirIngresoAMonedaServicio(
    montoEgreso,
    monedaPago,
    cotizacion,
    liquidacion.moneda
  );
  return egresoConvertido > saldoPendiente + 0.0001;
}

export function formatLiquidacionOperadorLabel(liquidacion: ILiquidacionOperador): string {
  const operador = (liquidacion.operadorNombre || '').trim();
  const referencia = (liquidacion.title || '').trim();
  const moneda = normalizarMonedaServicio(liquidacion.moneda);
  const saldo =
    liquidacion.saldoPendiente !== undefined
      ? liquidacion.saldoPendiente
      : liquidacion.importeTotal;
  const saldoTexto = formatMontoMoneda(saldo, moneda);

  if (operador && referencia) {
    return 'Operador: ' + operador + ' — ' + referencia + ' — Saldo ' + saldoTexto;
  }
  if (referencia) {
    return referencia + ' — Saldo ' + saldoTexto;
  }
  return 'Liquidación #' + liquidacion.id + ' — Saldo ' + saldoTexto;
}

export function formatSaldoLiquidacionDisplay(
  liquidacion: ILiquidacionOperador,
  monedaPago?: string,
  cotizacion?: number
): string {
  const monedaLiquidacion = normalizarMonedaServicio(liquidacion.moneda);
  const saldo = liquidacion.saldoPendiente ?? 0;
  const total = liquidacion.importeTotal ?? 0;
  const pagado = liquidacion.importePagado ?? total - saldo;
  const operador = (liquidacion.operadorNombre || '').trim();
  const lineas: string[] = [];

  if (operador) {
    lineas.push('Operador: ' + operador);
  }
  lineas.push('Total liquidación: ' + formatMontoMoneda(total, monedaLiquidacion));
  lineas.push('Pagado: ' + formatMontoMoneda(pagado, monedaLiquidacion));
  lineas.push('Saldo pendiente: ' + formatMontoMoneda(saldo, monedaLiquidacion));

  if (
    monedaPago &&
    monedaLiquidacion === 'Dólares' &&
    monedaPago.trim().toLowerCase().indexOf('pes') >= 0
  ) {
    const cotizacionNumerica = Number(cotizacion);
    if (isFinite(cotizacionNumerica) && cotizacionNumerica > 0) {
      const equivalentePesos = saldo * cotizacionNumerica;
      lineas.push(
        '(equiv. ' +
          equivalentePesos.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }) +
          ' en Pesos)'
      );
    }
  }

  return lineas.join('\n');
}
