import * as React from 'react';
import { IEstadoCajaBadgeProps } from './IEstadoCajaBadgeProps';
import { getEstadoCajaLabel } from '../utils/cajaConsolidadaUtils';
import styles from './VistaCajaConsolidada.module.scss';

/**
 * Badge visual del estado financiero del viaje.
 */
export default class EstadoCajaBadge extends React.Component<IEstadoCajaBadgeProps> {
  public render(): React.ReactElement<IEstadoCajaBadgeProps> {
    const { estado } = this.props;
    const label = getEstadoCajaLabel(estado);
    const toneClass = this._getToneClass(estado);

    return <span className={`${styles.estadoBadge} ${toneClass}`}>{label}</span>;
  }

  private _getToneClass(estado: IEstadoCajaBadgeProps['estado']): string {
    switch (estado) {
      case 'AlDia':
        return styles.estadoAlDia;
      case 'Parcial':
        return styles.estadoParcial;
      case 'Pendiente':
        return styles.estadoPendiente;
      case 'SinMovimientos':
        return styles.estadoSinMovimientos;
      default:
        return styles.estadoDesconocido;
    }
  }
}
