import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react';
import styles from './VistaCajaConsolidada.module.scss';
import type { IVistaCajaConsolidadaProps } from './IVistaCajaConsolidadaProps';
import ResumenPorViaje from './ResumenPorViaje';
import MovimientosCaja from './MovimientosCaja';

type VistaCajaConsolidadaTab = 'resumen' | 'movimientos';

export interface IVistaCajaConsolidadaState {
  vistaSeleccionada: VistaCajaConsolidadaTab;
}

/**
 * Contenedor principal del Web Part Vista Caja Consolidada.
 * Alterna entre Resumen por viaje y Movimientos.
 */
export default class VistaCajaConsolidada extends React.Component<
  IVistaCajaConsolidadaProps,
  IVistaCajaConsolidadaState
> {
  public constructor(props: IVistaCajaConsolidadaProps) {
    super(props);
    this.state = {
      vistaSeleccionada: 'resumen'
    };
  }

  public render(): React.ReactElement<IVistaCajaConsolidadaProps> {
    const { vistaSeleccionada } = this.state;

    return (
      <section className={styles.vistaCajaConsolidada}>
        <div className={styles.header}>
          <h2 className={styles.title}>Vista caja consolidada</h2>
        </div>

        <div className={styles.viewSelector}>
          <Pivot
            selectedKey={vistaSeleccionada}
            onLinkClick={this._onCambiarVista}
            headersOnly={true}
            aria-label="Vistas de caja consolidada"
          >
            <PivotItem headerText="Resumen por viaje" itemKey="resumen" />
            <PivotItem headerText="Movimientos" itemKey="movimientos" />
          </Pivot>
        </div>

        <div hidden={vistaSeleccionada !== 'resumen'}>
          <ResumenPorViaje context={this.props.context} />
        </div>

        <div hidden={vistaSeleccionada !== 'movimientos'}>
          <MovimientosCaja
            context={this.props.context}
            activa={vistaSeleccionada === 'movimientos'}
          />
        </div>
      </section>
    );
  }

  private _onCambiarVista = (item?: PivotItem): void => {
    if (!item || !item.props.itemKey) {
      return;
    }
    const key = item.props.itemKey;
    if (key === 'resumen' || key === 'movimientos') {
      this.setState({ vistaSeleccionada: key });
    }
  };
}
