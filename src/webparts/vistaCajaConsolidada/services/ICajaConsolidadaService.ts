import { ICajaConsolidadaVistaData } from '../models/ICajaConsolidadaVistaData';
import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';

/**
 * Contrato del servicio de datos de la vista consolidada.
 */
export interface ICajaConsolidadaService {
  /**
   * Obtiene viajes consolidados, egresos sin viaje y totales generales de caja.
   */
  getVistaData(filtros?: IFiltrosCajaConsolidada): Promise<ICajaConsolidadaVistaData>;
}
