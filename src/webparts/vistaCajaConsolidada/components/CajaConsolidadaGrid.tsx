import * as React from 'react';
import { formatDateDisplay } from '../../../shared/sharePointDateUtils';
import { ICajaConsolidadaGridProps } from './ICajaConsolidadaGridProps';
import { ICajaConsolidadaItem } from '../models/ICajaConsolidadaItem';
import ImporteMoneda from './ImporteMoneda';
import styles from './VistaCajaConsolidada.module.scss';

/**
 * Grilla por viaje: ingresos aplicados, egresos, recupero, recibido y saldos.
 */
export default class CajaConsolidadaGrid extends React.Component<ICajaConsolidadaGridProps> {
  public render(): React.ReactElement<ICajaConsolidadaGridProps> {
    const { items } = this.props;

    return (
      <div className={styles.tableOuter}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Viaje</th>
              <th className={styles.th}>Fecha de salida</th>
              <th className={styles.th}>Total del viaje</th>
              <th className={styles.th}>Ingresos aplicados</th>
              <th className={styles.th}>Egresos asociados</th>
              <th className={styles.th}>Recuperos bancarios</th>
              <th className={styles.th}>Total recibido</th>
              <th className={styles.th}>Saldo pendiente</th>
              <th className={styles.th}>Resultado viaje</th>
            </tr>
          </thead>
          <tbody>
            {items.map((item: ICajaConsolidadaItem) => (
              <tr key={item.viajeId}>
                <td className={styles.td}>{item.nombreViaje || '—'}</td>
                <td className={styles.td}>
                  {item.fechaSalida ? formatDateDisplay(item.fechaSalida) : '—'}
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.totalViaje} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.totalAbonado} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.egresosAsociados} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.totalRecupero} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.totalRecibido} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.saldoPendiente} stacked={true} />
                </td>
                <td className={styles.td}>
                  <ImporteMoneda importes={item.resultadoViaje} stacked={true} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}
