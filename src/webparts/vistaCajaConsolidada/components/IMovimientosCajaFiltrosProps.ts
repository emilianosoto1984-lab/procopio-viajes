import { IFiltrosMovimientosCaja } from '../models/IFiltrosMovimientosCaja';

export interface IMovimientosCajaFiltrosProps {
  filtros: IFiltrosMovimientosCaja;
  disabled?: boolean;
  errorPeriodo?: string;
  onChange: (filtros: IFiltrosMovimientosCaja) => void;
  onClear: () => void;
}
