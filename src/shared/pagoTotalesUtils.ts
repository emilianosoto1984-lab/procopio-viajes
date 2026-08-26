import {
  normalizarMonedaServicio
} from './pagoMonedaUtils';
import {
  convertirIngresoAMonedaServicio,
  isPagoConsideradoEnTotales
} from './pagoSaldoUtils';
import { getMontoQueAplicaAlViaje } from './pagoRecuperoUtils';

export { isPagoConsideradoEnTotales };

export interface IImportesPorMoneda {
  usd: number;
  ars: number;
}

export interface IServicioTotalItem {
  id: number;
  precioCliente: number;
  moneda: string;
}

export interface IPagoTotalItem {
  id?: number;
  tipoPago?: string;
  monto: number;
  montoAplicadoViaje?: number | null;
  moneda: string;
  cotizacion?: number;
  estado?: string;
  servicioAsociadoId?: number;
}

/**
 * Solo pagos Aprobados (y históricos sin Estado) impactan saldos/totales.
 * Misma regla que Registro de Viajes (ProcopioForms).
 */
// isPagoConsideradoEnTotales se reexporta desde pagoSaldoUtils.

export function createEmptyImportesPorMoneda(): IImportesPorMoneda {
  return { usd: 0, ars: 0 };
}

/**
 * Total del viaje = suma de PrecioCliente de ServiciosViaje, separado por moneda.
 */
export function getTotalesServiciosPorMoneda(
  servicios: IServicioTotalItem[]
): IImportesPorMoneda {
  return (servicios || []).reduce(
    (acc: IImportesPorMoneda, servicio: IServicioTotalItem) => {
      const monto = Number(servicio.precioCliente) || 0;
      if (normalizarMonedaServicio(servicio.moneda) === 'Dólares') {
        acc.usd += monto;
      } else {
        acc.ars += monto;
      }
      return acc;
    },
    createEmptyImportesPorMoneda()
  );
}

/**
 * Total abonado = ingresos considerados en totales, convertidos a la moneda del servicio asociado.
 * Excluye egresos y pendientes. Requiere servicioAsociadoId (misma regla que Viajes).
 */
export function getTotalesIngresosPorMoneda(
  pagos: IPagoTotalItem[],
  servicios: IServicioTotalItem[]
): IImportesPorMoneda {
  const serviciosById: { [id: number]: IServicioTotalItem } = {};
  (servicios || []).forEach((servicio: IServicioTotalItem) => {
    if (servicio.id > 0) {
      serviciosById[servicio.id] = servicio;
    }
  });

  return (pagos || [])
    .filter(
      (pago: IPagoTotalItem) =>
        isPagoConsideradoEnTotales(pago) &&
        (pago.tipoPago || '').trim() === 'Ingreso' &&
        getMontoQueAplicaAlViaje(pago) > 0 &&
        !!(pago.servicioAsociadoId && pago.servicioAsociadoId > 0)
    )
    .reduce(
      (acc: IImportesPorMoneda, pago: IPagoTotalItem) => {
        const servicio = serviciosById[pago.servicioAsociadoId as number];
        if (!servicio) {
          return acc;
        }
        const convertido = convertirIngresoAMonedaServicio(
          getMontoQueAplicaAlViaje(pago),
          pago.moneda,
          pago.cotizacion,
          servicio.moneda
        );
        if (normalizarMonedaServicio(servicio.moneda) === 'Dólares') {
          acc.usd += convertido;
        } else {
          acc.ars += convertido;
        }
        return acc;
      },
      createEmptyImportesPorMoneda()
    );
}

export function getSaldosPendientesPorMoneda(
  totalViaje: IImportesPorMoneda,
  totalAbonado: IImportesPorMoneda
): IImportesPorMoneda {
  return {
    usd: (totalViaje.usd || 0) - (totalAbonado.usd || 0),
    ars: (totalViaje.ars || 0) - (totalAbonado.ars || 0)
  };
}
