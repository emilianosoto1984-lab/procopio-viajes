import * as React from 'react';
import { IImporteMonedaProps } from './IImporteMonedaProps';
import { formatImporteArs, formatImporteUsd } from '../utils/monedaUtils';
import styles from './VistaCajaConsolidada.module.scss';

/**
 * Presentación de importes ARS / USD.
 */
export default class ImporteMoneda extends React.Component<IImporteMonedaProps> {
  public render(): React.ReactElement<IImporteMonedaProps> {
    const { importes, stacked } = this.props;
    const className = stacked ? styles.importeStacked : styles.importeInline;

    return (
      <div className={className}>
        <div className={styles.importeLine}>{formatImporteUsd(importes.usd)}</div>
        <div className={styles.importeLine}>{formatImporteArs(importes.ars)}</div>
      </div>
    );
  }
}
