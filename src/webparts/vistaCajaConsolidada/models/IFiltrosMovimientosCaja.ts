/**
 * Filtros de la vista Movimientos.
 * moneda vacío = todas; si tiene valor: Pesos o Dólares.
 */
export interface IFiltrosMovimientosCaja {
  fechaDesde: string;
  fechaHasta: string;
  moneda: string;
}
