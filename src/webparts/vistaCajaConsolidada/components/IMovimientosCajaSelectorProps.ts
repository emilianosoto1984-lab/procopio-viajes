import { OrigenCaja } from '../models/OrigenCaja';

export interface IMovimientosCajaSelectorProps {
  origenes: OrigenCaja[];
  seleccionado: OrigenCaja;
  disabled?: boolean;
  onChange: (origen: OrigenCaja) => void;
}
