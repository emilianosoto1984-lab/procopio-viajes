import { IImportesMoneda } from '../models/IImportesMoneda';

export interface IMovimientosCajaKpisProps {
  ingresos: IImportesMoneda;
  egresos: IImportesMoneda;
  saldo: IImportesMoneda;
}
