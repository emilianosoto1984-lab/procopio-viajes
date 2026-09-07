import { ICuentaBancaria } from '../../../shared/cuentasBancariasUtils';
import { ICajaConsolidadaVistaData } from '../models/ICajaConsolidadaVistaData';
import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';
import { OrigenCaja } from '../models/OrigenCaja';
import { IMovimientoCajaLinea } from '../utils/cajaMovimientosUtils';

/**
 * Contrato del servicio de datos de la vista consolidada.
 */
export interface ICajaConsolidadaService {
  /**
   * Obtiene viajes consolidados, egresos sin viaje y totales generales de caja.
   */
  getVistaData(filtros?: IFiltrosCajaConsolidada): Promise<ICajaConsolidadaVistaData>;

  /**
   * Cuentas activas de la lista CuentasBancarias.
   */
  getCuentasBancarias(): Promise<ICuentaBancaria[]>;

  /**
   * Movimientos de Registro de Pagos filtrados por origen de fondos.
   */
  getMovimientos(origen: OrigenCaja): Promise<IMovimientoCajaLinea[]>;

  /**
   * URL del formulario Display de un ítem de Registro de Pagos (PageType=4).
   */
  getPagoDisplayFormUrl(pagoId: number): Promise<string>;

  /**
   * Actualiza únicamente Estado = Aprobado en Registro de Pagos.
   */
  aprobarPago(pagoId: number): Promise<void>;
}
