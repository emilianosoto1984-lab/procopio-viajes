import { IImportesMoneda } from './IImportesMoneda';

/**
 * Estado financiero agregado de un viaje en la vista de caja.
 */
export type EstadoCajaConsolidada =
  | 'AlDia'
  | 'Parcial'
  | 'Pendiente'
  | 'SinMovimientos'
  | 'Desconocido';

/**
 * Fila consolidada por viaje.
 *
 * Distinción de importes:
 * - totalViaje: precio de servicios (moneda del servicio).
 * - totalAbonado / ingresosAplicados: MontoAplicadoViaje ?? Monto (moneda del servicio).
 * - egresosAsociados: Monto de egresos del viaje (moneda del pago).
 * - totalRecibido: sum(Monto) ingresos (incluye recupero; moneda del pago).
 * - totalRecupero: informativo; ya incluido en totalRecibido.
 * - saldoPendiente: cobro pendiente = totalViaje - ingresosAplicados (NO resta egresos ni recupero).
 * - resultadoViaje: ingresosAplicados - egresosAsociados (resultado económico del viaje).
 */
export interface ICajaConsolidadaItem {
  viajeId: number;
  nombreViaje: string;
  destino: string;
  fechaSalida: string;
  /** Reservado: aún no hay columna Vendedor en Registro de Viajes. */
  vendedor: string;
  totalViaje: IImportesMoneda;
  /** Alias de negocio: ingresos aplicados al viaje. */
  totalAbonado: IImportesMoneda;
  egresosAsociados: IImportesMoneda;
  totalRecibido: IImportesMoneda;
  totalRecupero: IImportesMoneda;
  /** Cobro pendiente del cliente (precio - ingresos aplicados). */
  saldoPendiente: IImportesMoneda;
  /** Resultado económico del viaje (ingresos aplicados - egresos asociados). */
  resultadoViaje: IImportesMoneda;
  estadoFinanciero: EstadoCajaConsolidada;
}
