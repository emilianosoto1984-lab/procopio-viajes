import { IImportesMoneda } from '../models/IImportesMoneda';

export interface ICajaConsolidadaKpisProps {
  cantidadViajes: number;
  totalRecibido: IImportesMoneda;
  /** Total de egresos (asociados + sin viaje). */
  totalEgresos: IImportesMoneda;
  totalRecupero: IImportesMoneda;
}
