import { IImportesMoneda } from './IImportesMoneda';
import { ICajaConsolidadaItem } from './ICajaConsolidadaItem';
import { ITotalesGeneralesCaja } from './ITotalesGeneralesCaja';
import { IEgresoSinViajeCaja } from '../utils/cajaMovimientosUtils';

/**
 * Payload completo de la vista: viajes + egresos generales + totales de caja.
 */
export interface ICajaConsolidadaVistaData {
  items: ICajaConsolidadaItem[];
  egresosSinViaje: IEgresoSinViajeCaja[];
  totalesGenerales: ITotalesGeneralesCaja;
  /** Totales económicos agregados de la grilla de viajes (ingresos aplicados). */
  totalIngresosAplicados: IImportesMoneda;
  totalEgresosAsociadosViajes: IImportesMoneda;
  totalSaldoPendienteCobro: IImportesMoneda;
}
