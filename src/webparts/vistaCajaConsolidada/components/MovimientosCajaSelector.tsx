import * as React from 'react';
import { IMovimientosCajaSelectorProps } from './IMovimientosCajaSelectorProps';
import { OrigenCaja } from '../models/OrigenCaja';
import { findOrigenCajaByKey, getOrigenCajaKey } from '../utils/origenCajaUtils';
import styles from './VistaCajaConsolidada.module.scss';

/**
 * Selector de origen de fondos: Efectivo + cuentas de CuentasBancarias.
 */
export default class MovimientosCajaSelector extends React.Component<IMovimientosCajaSelectorProps> {
  public render(): React.ReactElement<IMovimientosCajaSelectorProps> {
    const { origenes, seleccionado, disabled } = this.props;

    return (
      <div className={styles.filtersRoot}>
        <div className={styles.filtersRow}>
          <div className={`${styles.fieldGroup} ${styles.originField}`}>
            <label className={styles.label} htmlFor="origen-fondos-caja">
              Origen de fondos
            </label>
            <select
              id="origen-fondos-caja"
              className={styles.select}
              value={getOrigenCajaKey(seleccionado)}
              disabled={!!disabled}
              onChange={this._onChange}
            >
              {(origenes || []).map((origen: OrigenCaja) => (
                <option key={getOrigenCajaKey(origen)} value={getOrigenCajaKey(origen)}>
                  {origen.nombre}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>
    );
  }

  private _onChange = (ev: React.ChangeEvent<HTMLSelectElement>): void => {
    const origen = findOrigenCajaByKey(this.props.origenes, ev.target.value);
    if (origen) {
      this.props.onChange(origen);
    }
  };
}
