import { IImportesMoneda } from '../models/IImportesMoneda';

export interface IImporteMonedaProps {
  importes: IImportesMoneda;
  /** Si true, muestra ARS y USD en líneas separadas. */
  stacked?: boolean;
}
