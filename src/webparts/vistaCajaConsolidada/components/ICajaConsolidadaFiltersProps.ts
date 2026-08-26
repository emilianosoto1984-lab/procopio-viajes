import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';

export interface ICajaConsolidadaFiltersProps {
  filtros: IFiltrosCajaConsolidada;
  disabled?: boolean;
  onChange: (filtros: IFiltrosCajaConsolidada) => void;
  onApply: () => void;
  onClear: () => void;
}
