import { IImportesMoneda } from './IImportesMoneda';

/**
 * Totales generales de caja (dinero real), independientes del resultado por viaje.
 *
 * Anti doble-suma: totalRecibido = sum(Monto). totalRecupero es informativo
 * (ya incluido dentro de Monto).
 */
export interface ITotalesGeneralesCaja {
  totalRecibido: IImportesMoneda;
  totalRecupero: IImportesMoneda;
  egresosAsociados: IImportesMoneda;
  egresosSinViaje: IImportesMoneda;
  totalEgresos: IImportesMoneda;
  resultadoGeneral: IImportesMoneda;
}
