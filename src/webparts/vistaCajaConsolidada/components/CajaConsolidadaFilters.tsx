import * as React from 'react';
import { Toggle } from '@fluentui/react';
import { ICajaConsolidadaFiltersProps } from './ICajaConsolidadaFiltersProps';
import styles from './VistaCajaConsolidada.module.scss';

/**
 * Barra de filtros locales sobre la grilla consolidada.
 */
export default class CajaConsolidadaFilters extends React.Component<ICajaConsolidadaFiltersProps> {
  public render(): React.ReactElement<ICajaConsolidadaFiltersProps> {
    const { filtros, disabled } = this.props;

    return (
      <div className={styles.filtersRoot}>
        <div className={styles.filtersRow}>
          <div className={styles.fieldGroup}>
            <label className={styles.label}>Buscar viaje</label>
            <input
              type="text"
              className={styles.input}
              value={filtros.textoBusqueda}
              disabled={disabled}
              placeholder="Nombre del viaje..."
              onChange={this._onTextoBusqueda}
            />
          </div>
          <div className={styles.fieldGroup}>
            <label className={styles.label}>Salida desde</label>
            <input
              type="date"
              className={styles.input}
              value={filtros.fechaSalidaDesde}
              disabled={disabled}
              onChange={this._onFechaDesde}
            />
          </div>
          <div className={styles.fieldGroup}>
            <label className={styles.label}>Salida hasta</label>
            <input
              type="date"
              className={styles.input}
              value={filtros.fechaSalidaHasta}
              disabled={disabled}
              onChange={this._onFechaHasta}
            />
          </div>
        </div>
        <div className={styles.filtersActions}>
          <div className={styles.switchGroup}>
            <Toggle
              label="Mostrar saldos en 0"
              checked={!!filtros.mostrarSaldosEnCero}
              disabled={disabled}
              onText="On"
              offText="Off"
              inlineLabel={true}
              onChange={this._onMostrarSaldosEnCero}
            />
          </div>
          <button
            type="button"
            className={styles.btnPrimary}
            disabled={disabled}
            onClick={this.props.onApply}
          >
            Aplicar
          </button>
          <button
            type="button"
            className={styles.btnDefault}
            disabled={disabled}
            onClick={this.props.onClear}
          >
            Limpiar
          </button>
        </div>
      </div>
    );
  }

  private _patch = (partial: Partial<ICajaConsolidadaFiltersProps['filtros']>): void => {
    this.props.onChange({ ...this.props.filtros, ...partial });
  };

  private _onTextoBusqueda = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    this._patch({ textoBusqueda: ev.target.value });
  };

  private _onFechaDesde = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    this._patch({ fechaSalidaDesde: ev.target.value });
  };

  private _onFechaHasta = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    this._patch({ fechaSalidaHasta: ev.target.value });
  };

  private _onMostrarSaldosEnCero = (
    _ev: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    this._patch({ mostrarSaldosEnCero: !!checked });
  };
}
