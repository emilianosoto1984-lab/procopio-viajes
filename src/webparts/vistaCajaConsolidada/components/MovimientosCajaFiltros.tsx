import * as React from 'react';
import { DatePicker, DayOfWeek } from '@fluentui/react';
import {
  dateToInputString,
  formatDateForPicker,
  FORM_DATE_PICKER_STRINGS,
  parseDateFromPickerString,
  parseDateStringToDate
} from '../../../shared/formDatePickerUtils';
import { IMovimientosCajaFiltrosProps } from './IMovimientosCajaFiltrosProps';
import styles from './VistaCajaConsolidada.module.scss';

const DATE_PICKER_STYLES = {
  root: { width: '100%' },
  textField: {
    fieldGroup: {
      height: 32,
      borderRadius: 4,
      border: '1px solid #c8c6c4'
    }
  }
};

/**
 * Filtro de período sobre FechaPago. Las fechas se muestran en dd/mm/aaaa.
 */
export default class MovimientosCajaFiltros extends React.Component<IMovimientosCajaFiltrosProps> {
  public render(): React.ReactElement<IMovimientosCajaFiltrosProps> {
    const { filtros, disabled, errorPeriodo } = this.props;
    const fechaDesde = parseDateStringToDate(filtros.fechaDesde);
    const fechaHasta = parseDateStringToDate(filtros.fechaHasta);
    const tieneError = !!(errorPeriodo && errorPeriodo.trim());

    return (
      <div className={styles.filtersRoot}>
        <div className={styles.filtersRow}>
          <div className={styles.fieldGroup}>
            <label className={styles.label} htmlFor="movimientos-fecha-desde">
              Fecha desde
            </label>
            <DatePicker
              id="movimientos-fecha-desde"
              value={fechaDesde}
              onSelectDate={this._onFechaDesde}
              formatDate={formatDateForPicker}
              parseDateFromString={parseDateFromPickerString}
              placeholder="dd/mm/aaaa"
              allowTextInput={true}
              strings={FORM_DATE_PICKER_STRINGS}
              firstDayOfWeek={DayOfWeek.Monday}
              disabled={!!disabled}
              maxDate={fechaHasta}
              styles={DATE_PICKER_STYLES}
            />
          </div>
          <div className={styles.fieldGroup}>
            <label className={styles.label} htmlFor="movimientos-fecha-hasta">
              Fecha hasta
            </label>
            <DatePicker
              id="movimientos-fecha-hasta"
              value={fechaHasta}
              onSelectDate={this._onFechaHasta}
              formatDate={formatDateForPicker}
              parseDateFromString={parseDateFromPickerString}
              placeholder="dd/mm/aaaa"
              allowTextInput={true}
              strings={FORM_DATE_PICKER_STRINGS}
              firstDayOfWeek={DayOfWeek.Monday}
              disabled={!!disabled}
              minDate={fechaDesde}
              textField={{
                errorMessage: tieneError ? errorPeriodo : undefined
              }}
              styles={DATE_PICKER_STYLES}
            />
          </div>
          <div className={styles.fieldGroup}>
            <label className={styles.label} htmlFor="movimientos-moneda">
              Moneda
            </label>
            <select
              id="movimientos-moneda"
              className={styles.select}
              value={filtros.moneda || ''}
              disabled={!!disabled}
              onChange={this._onMoneda}
            >
              <option value="">Todas</option>
              <option value="Pesos">Pesos</option>
              <option value="Dólares">Dólares</option>
            </select>
          </div>
        </div>
        {tieneError && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{errorPeriodo}</div>
        )}
        <div className={styles.filtersActions}>
          <button
            type="button"
            className={styles.btnDefault}
            disabled={!!disabled}
            onClick={this.props.onClear}
          >
            Limpiar filtros
          </button>
        </div>
      </div>
    );
  }

  private _onFechaDesde = (date: Date | null | undefined): void => {
    this.props.onChange({
      ...this.props.filtros,
      fechaDesde: date ? dateToInputString(date) : ''
    });
  };

  private _onFechaHasta = (date: Date | null | undefined): void => {
    this.props.onChange({
      ...this.props.filtros,
      fechaHasta: date ? dateToInputString(date) : ''
    });
  };

  private _onMoneda = (ev: React.ChangeEvent<HTMLSelectElement>): void => {
    this.props.onChange({
      ...this.props.filtros,
      moneda: ev.target.value
    });
  };
}
