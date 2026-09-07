import { IMovimientoCajaLinea } from '../utils/cajaMovimientosUtils';
import { VistaEstadoMovimientosCaja } from '../utils/movimientosEstadoUtils';

export interface IMovimientosCajaGridProps {
  items: IMovimientoCajaLinea[];
  vista: VistaEstadoMovimientosCaja;
  aprobandoPagoId: number;
  onVerDetalle: (pagoId: number) => void;
  onAprobar: (pagoId: number) => void;
}
