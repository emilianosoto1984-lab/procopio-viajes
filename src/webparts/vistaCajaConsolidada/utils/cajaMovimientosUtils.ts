import { isPagoConsideradoEnTotales } from '../../../shared/pagoTotalesUtils';
import { getMontoQueAplicaAlViaje } from '../../../shared/pagoRecuperoUtils';
import { normalizarMonedaPago } from '../../../shared/pagoMonedaUtils';
import { IImportesMoneda } from '../models/IImportesMoneda';
import { ITotalesGeneralesCaja } from '../models/ITotalesGeneralesCaja';
import { createEmptyImportesMoneda } from './monedaUtils';

/** Concepto estable para agrupaciones de recupero. */
export const CONCEPTO_RECUPERO_GASTOS_BANCARIOS = 'Recupero de gastos bancarios';

export type TipoLineaCaja = 'principal' | 'recupero';

/**
 * Pago mínimo para armar líneas de caja y sumar recibidos/recuperos/egresos.
 * Los totales de abonado al viaje siguen usando getTotalesIngresosPorMoneda.
 */
export interface IPagoParaCaja {
  id: number;
  /** 0 = sin viaje asociado. */
  viajeId?: number;
  tipoPago?: string;
  monto: number;
  montoAplicadoViaje?: number | null;
  montoGastosBancarios?: number | null;
  porcentajeRecupero?: number | null;
  moneda: string;
  cotizacion?: number;
  estado?: string;
  servicioAsociadoId?: number;
  concepto?: string;
  motivo?: string;
  cuentaBancariaId?: number | null;
  cuentaBancariaTitulo?: string;
  banco?: string;
  fechaPago?: string;
  medioPago?: string;
  observaciones?: string;
  /** Título del lookup ViajeAsociado (solo presentación). */
  viajeTitulo?: string;
}

export interface IMovimientoCajaLinea {
  key: string;
  tipoLinea: TipoLineaCaja;
  pagoId: number;
  viajeId: number;
  servicioAsociadoId?: number;
  fechaPago: string;
  tipoPago: string;
  concepto: string;
  monto: number;
  moneda: string;
  cotizacion?: number;
  estado?: string;
  cuentaBancariaId?: number | null;
  cuentaBancariaLabel: string;
  medioPago?: string;
  viajeTitulo?: string;
}

export interface IEgresoSinViajeCaja {
  id: number;
  fechaPago: string;
  motivo: string;
  medioPago: string;
  cuentaBancariaLabel: string;
  moneda: string;
  monto: number;
  estado: string;
  observaciones: string;
}

export function esIngreso(pago: IPagoParaCaja): boolean {
  return (pago.tipoPago || '').trim() === 'Ingreso';
}

export function esEgreso(pago: IPagoParaCaja): boolean {
  return (pago.tipoPago || '').trim() === 'Egreso';
}

export function tieneViajeAsociado(pago: IPagoParaCaja): boolean {
  return !!(pago.viajeId && pago.viajeId > 0);
}

function _esConsiderado(pago: IPagoParaCaja): boolean {
  return isPagoConsideradoEnTotales(pago);
}

function _esIngresoConsiderado(pago: IPagoParaCaja): boolean {
  return _esConsiderado(pago) && esIngreso(pago);
}

function _esEgresoConsiderado(pago: IPagoParaCaja): boolean {
  return _esConsiderado(pago) && esEgreso(pago);
}

export function getMontoAplicadoViajeCaja(pago: IPagoParaCaja): number {
  return getMontoQueAplicaAlViaje(pago);
}

/**
 * Recupero bancario. Preferir MontoGastosBancarios; fallback Monto - MontoAplicadoViaje.
 * Informativo: ya está incluido dentro de Monto (no sumar encima del recibido).
 */
export function getMontoRecuperoCaja(pago: IPagoParaCaja): number {
  const raw = pago.montoGastosBancarios;
  if (raw !== null && raw !== undefined) {
    const n = Number(raw);
    if (isFinite(n) && n > 0) {
      return n;
    }
  }

  const aplicadoRaw = pago.montoAplicadoViaje;
  if (aplicadoRaw === null || aplicadoRaw === undefined) {
    return 0;
  }
  const aplicado = Number(aplicadoRaw);
  const bruto = Number(pago.monto) || 0;
  if (!isFinite(aplicado) || !isFinite(bruto)) {
    return 0;
  }
  const diff = Math.round((bruto - aplicado) * 100) / 100;
  return diff > 0.009 ? diff : 0;
}

export function getMontoEgresoCaja(pago: IPagoParaCaja): number {
  return Number(pago.monto) || 0;
}

export function resolverCuentaBancariaLabelCaja(pago: IPagoParaCaja): string {
  const titulo = (pago.cuentaBancariaTitulo || '').trim();
  if (titulo) {
    return titulo;
  }
  return (pago.banco || '').trim();
}

function _conceptoPrincipal(pago: IPagoParaCaja): string {
  const concepto = (pago.concepto || '').trim();
  return concepto || 'Pago';
}

function _motivoEgresoSinViaje(pago: IPagoParaCaja): string {
  const motivo = (pago.motivo || '').trim();
  return motivo || 'Sin motivo';
}

/**
 * Transforma un Registro de Pago en 1 o 2 líneas de caja (solo presentación).
 *
 * Anti doble-suma: totales de "recibido" = sum(Monto), NUNCA Monto + MontoGastosBancarios.
 */
export function crearMovimientosCajaDesdePago(pago: IPagoParaCaja): IMovimientoCajaLinea[] {
  if (!_esIngresoConsiderado(pago) || !(pago.id > 0)) {
    return [];
  }

  const montoRecibido = Number(pago.monto) || 0;
  if (montoRecibido <= 0 && getMontoAplicadoViajeCaja(pago) <= 0) {
    return [];
  }

  const recupero = getMontoRecuperoCaja(pago);
  const aplicado = getMontoAplicadoViajeCaja(pago);
  const cuentaBancariaLabel = resolverCuentaBancariaLabelCaja(pago);
  const base = {
    pagoId: pago.id,
    viajeId: pago.viajeId && pago.viajeId > 0 ? pago.viajeId : 0,
    servicioAsociadoId: pago.servicioAsociadoId,
    fechaPago: pago.fechaPago || '',
    tipoPago: 'Ingreso',
    moneda: pago.moneda,
    cotizacion: pago.cotizacion,
    estado: pago.estado,
    cuentaBancariaId: pago.cuentaBancariaId,
    cuentaBancariaLabel,
    medioPago: pago.medioPago
  };

  const lineas: IMovimientoCajaLinea[] = [
    {
      ...base,
      key: pago.id + '-principal',
      tipoLinea: 'principal',
      concepto: _conceptoPrincipal(pago),
      monto: aplicado
    }
  ];

  if (recupero > 0) {
    lineas.push({
      ...base,
      key: pago.id + '-recupero',
      tipoLinea: 'recupero',
      concepto: CONCEPTO_RECUPERO_GASTOS_BANCARIOS,
      monto: recupero
    });
  }

  return lineas;
}

export function crearMovimientosCajaDesdePagos(
  pagos: IPagoParaCaja[]
): IMovimientoCajaLinea[] {
  const result: IMovimientoCajaLinea[] = [];
  (pagos || []).forEach((pago: IPagoParaCaja) => {
    const lineas = crearMovimientosCajaDesdePago(pago);
    for (let i = 0; i < lineas.length; i++) {
      result.push(lineas[i]);
    }
  });
  return result;
}

export function listarEgresosSinViaje(pagos: IPagoParaCaja[]): IEgresoSinViajeCaja[] {
  return (pagos || [])
    .filter(
      (pago: IPagoParaCaja) =>
        pago.id > 0 && _esEgresoConsiderado(pago) && !tieneViajeAsociado(pago)
    )
    .map((pago: IPagoParaCaja) => ({
      id: pago.id,
      fechaPago: pago.fechaPago || '',
      motivo: _motivoEgresoSinViaje(pago),
      medioPago: (pago.medioPago || '').trim(),
      cuentaBancariaLabel: resolverCuentaBancariaLabelCaja(pago) || '—',
      moneda: pago.moneda || '',
      monto: getMontoEgresoCaja(pago),
      estado: (pago.estado || '').trim() || '—',
      observaciones: (pago.observaciones || '').trim()
    }))
    .sort((a: IEgresoSinViajeCaja, b: IEgresoSinViajeCaja) => {
      const fa = a.fechaPago || '';
      const fb = b.fechaPago || '';
      if (fa === fb) {
        return b.id - a.id;
      }
      return fa < fb ? 1 : -1;
    });
}

function _acumularPorMonedaPago(
  acc: IImportesMoneda,
  moneda: string,
  monto: number
): void {
  const valor = Number(monto) || 0;
  if (valor === 0) {
    return;
  }
  if (normalizarMonedaPago(moneda) === 'Dólares') {
    acc.usd += valor;
  } else {
    acc.ars += valor;
  }
}

/**
 * Dinero realmente recibido: sum(Monto) de ingresos considerados.
 * El recupero YA está dentro de Monto — no sumar MontoGastosBancarios encima.
 */
export function sumTotalRecibidoPorMonedaPago(pagos: IPagoParaCaja[]): IImportesMoneda {
  return (pagos || [])
    .filter((pago: IPagoParaCaja) => _esIngresoConsiderado(pago))
    .reduce(
      (acc: IImportesMoneda, pago: IPagoParaCaja) => {
        _acumularPorMonedaPago(acc, pago.moneda, Number(pago.monto) || 0);
        return acc;
      },
      createEmptyImportesMoneda()
    );
}

/** Recupero informativo: sum(MontoGastosBancarios). No sumar al recibido. */
export function sumTotalRecuperoPorMonedaPago(pagos: IPagoParaCaja[]): IImportesMoneda {
  return (pagos || [])
    .filter((pago: IPagoParaCaja) => _esIngresoConsiderado(pago))
    .reduce(
      (acc: IImportesMoneda, pago: IPagoParaCaja) => {
        _acumularPorMonedaPago(acc, pago.moneda, getMontoRecuperoCaja(pago));
        return acc;
      },
      createEmptyImportesMoneda()
    );
}

export function sumEgresosAsociadosPorMonedaPago(pagos: IPagoParaCaja[]): IImportesMoneda {
  return (pagos || [])
    .filter(
      (pago: IPagoParaCaja) =>
        _esEgresoConsiderado(pago) && tieneViajeAsociado(pago)
    )
    .reduce(
      (acc: IImportesMoneda, pago: IPagoParaCaja) => {
        _acumularPorMonedaPago(acc, pago.moneda, getMontoEgresoCaja(pago));
        return acc;
      },
      createEmptyImportesMoneda()
    );
}

export function sumEgresosSinViajePorMonedaPago(pagos: IPagoParaCaja[]): IImportesMoneda {
  return (pagos || [])
    .filter(
      (pago: IPagoParaCaja) =>
        _esEgresoConsiderado(pago) && !tieneViajeAsociado(pago)
    )
    .reduce(
      (acc: IImportesMoneda, pago: IPagoParaCaja) => {
        _acumularPorMonedaPago(acc, pago.moneda, getMontoEgresoCaja(pago));
        return acc;
      },
      createEmptyImportesMoneda()
    );
}

function _sumImportes(a: IImportesMoneda, b: IImportesMoneda): IImportesMoneda {
  return {
    usd: (a.usd || 0) + (b.usd || 0),
    ars: (a.ars || 0) + (b.ars || 0)
  };
}

function _restarImportes(a: IImportesMoneda, b: IImportesMoneda): IImportesMoneda {
  return {
    usd: (a.usd || 0) - (b.usd || 0),
    ars: (a.ars || 0) - (b.ars || 0)
  };
}

/**
 * Totales generales de caja desde la colección ORIGINAL de pagos.
 * Estrategia anti doble-suma del recupero: totalRecibido = sum(Monto).
 */
export function calcularTotalesGeneralesCaja(pagos: IPagoParaCaja[]): ITotalesGeneralesCaja {
  const totalRecibido = sumTotalRecibidoPorMonedaPago(pagos);
  const totalRecupero = sumTotalRecuperoPorMonedaPago(pagos);
  const egresosAsociados = sumEgresosAsociadosPorMonedaPago(pagos);
  const egresosSinViaje = sumEgresosSinViajePorMonedaPago(pagos);
  const totalEgresos = _sumImportes(egresosAsociados, egresosSinViaje);
  return {
    totalRecibido,
    totalRecupero,
    egresosAsociados,
    egresosSinViaje,
    totalEgresos,
    resultadoGeneral: _restarImportes(totalRecibido, totalEgresos)
  };
}

export function createEmptyTotalesGeneralesCaja(): ITotalesGeneralesCaja {
  return {
    totalRecibido: createEmptyImportesMoneda(),
    totalRecupero: createEmptyImportesMoneda(),
    egresosAsociados: createEmptyImportesMoneda(),
    egresosSinViaje: createEmptyImportesMoneda(),
    totalEgresos: createEmptyImportesMoneda(),
    resultadoGeneral: createEmptyImportesMoneda()
  };
}

export interface ITotalesMovimientosCaja {
  ingresos: IImportesMoneda;
  egresos: IImportesMoneda;
  saldo: IImportesMoneda;
  cantidad: number;
}

/**
 * Totales de la vista Movimientos a partir de las líneas visibles.
 * Sobre una colección vacía devuelve ceros.
 */
export function calcularTotalesMovimientosCaja(
  lineas: IMovimientoCajaLinea[]
): ITotalesMovimientosCaja {
  const ingresos = createEmptyImportesMoneda();
  const egresos = createEmptyImportesMoneda();
  (lineas || []).forEach((linea: IMovimientoCajaLinea) => {
    if ((linea.tipoPago || '').trim() === 'Egreso') {
      _acumularPorMonedaPago(egresos, linea.moneda, linea.monto);
    } else {
      _acumularPorMonedaPago(ingresos, linea.moneda, linea.monto);
    }
  });
  return {
    ingresos,
    egresos,
    saldo: _restarImportes(ingresos, egresos),
    cantidad: (lineas || []).length
  };
}

/**
 * Una fila de la vista Movimientos por cada Registro de Pago (ingreso o egreso).
 * No usa el split de recupero: el importe es el Monto del ítem.
 */
export function mapearPagosALineasMovimientosCaja(
  pagos: IPagoParaCaja[]
): IMovimientoCajaLinea[] {
  return (pagos || [])
    .filter((pago: IPagoParaCaja) => pago.id > 0)
    .map((pago: IPagoParaCaja) => {
      const tipoPago = (pago.tipoPago || '').trim() === 'Egreso' ? 'Egreso' : 'Ingreso';
      const concepto = (pago.concepto || '').trim() || 'Pago';
      return {
        key: String(pago.id),
        tipoLinea: 'principal' as TipoLineaCaja,
        pagoId: pago.id,
        viajeId: pago.viajeId && pago.viajeId > 0 ? pago.viajeId : 0,
        servicioAsociadoId: pago.servicioAsociadoId,
        fechaPago: pago.fechaPago || '',
        tipoPago,
        concepto,
        monto: Number(pago.monto) || 0,
        moneda: pago.moneda,
        cotizacion: pago.cotizacion,
        estado: pago.estado,
        cuentaBancariaId: pago.cuentaBancariaId,
        cuentaBancariaLabel: resolverCuentaBancariaLabelCaja(pago),
        medioPago: pago.medioPago,
        viajeTitulo: (pago.viajeTitulo || '').trim()
      };
    })
    .sort((a: IMovimientoCajaLinea, b: IMovimientoCajaLinea) => {
      const fa = a.fechaPago || '';
      const fb = b.fechaPago || '';
      if (fa === fb) {
        return b.pagoId - a.pagoId;
      }
      return fa < fb ? 1 : -1;
    });
}
