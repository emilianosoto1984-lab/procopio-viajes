import { EstadoCajaConsolidada } from './ICajaConsolidadaItem';

/**
 * Filtros de la vista consolidada.
 */
export interface IFiltrosCajaConsolidada {
  textoBusqueda: string;
  fechaSalidaDesde: string;
  fechaSalidaHasta: string;
  vendedor: string;
  estadoFinanciero: EstadoCajaConsolidada | '';
  /** Si es false (default), oculta viajes con saldo pendiente en 0. */
  mostrarSaldosEnCero: boolean;
}
