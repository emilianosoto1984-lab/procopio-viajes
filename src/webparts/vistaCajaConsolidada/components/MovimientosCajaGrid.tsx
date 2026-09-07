import * as React from 'react';
import { formatDateDisplay } from '../../../shared/sharePointDateUtils';
import { normalizarMonedaPago } from '../../../shared/pagoMonedaUtils';
import { formatImporteArs, formatImporteUsd } from '../utils/monedaUtils';
import {
  GridIconActionButton,
  GridIconCheck,
  GridIconExternalLink
} from '../../../shared/gridIconActions';
import { IMovimientoCajaLinea } from '../utils/cajaMovimientosUtils';
import { IMovimientosCajaGridProps } from './IMovimientosCajaGridProps';
import styles from './VistaCajaConsolidada.module.scss';

function formatImporteMovimiento(monto: number, moneda: string): string {
  return normalizarMonedaPago(moneda) === 'Dólares'
    ? formatImporteUsd(monto)
    : formatImporteArs(monto);
}

/**
 * Estructura de la grilla de movimientos de caja (ingresos y egresos).
 */
export default class MovimientosCajaGrid extends React.Component<IMovimientosCajaGridProps> {
  public render(): React.ReactElement<IMovimientosCajaGridProps> {
    const { items, vista, aprobandoPagoId } = this.props;
    const mostrarAprobar = vista === 'pendientes';

    return (
      <div className={styles.tableOuter}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th className={styles.th}>Fecha</th>
              <th className={styles.th}>Tipo</th>
              <th className={styles.th}>Concepto</th>
              <th className={styles.th}>Viaje</th>
              <th className={styles.th}>Moneda</th>
              <th className={styles.th}>Importe</th>
              <th className={styles.th}>Acciones</th>
            </tr>
          </thead>
          <tbody>
            {(items || []).map((item: IMovimientoCajaLinea) => (
              <tr
                key={item.key}
                className={mostrarAprobar ? styles.rowPendiente : undefined}
              >
                <td className={styles.td}>
                  {item.fechaPago ? formatDateDisplay(item.fechaPago) : '—'}
                </td>
                <td className={styles.td}>{item.tipoPago || '—'}</td>
                <td className={styles.td}>{item.concepto || '—'}</td>
                <td className={styles.td}>{item.viajeTitulo || '—'}</td>
                <td className={styles.td}>{item.moneda || '—'}</td>
                <td className={styles.td}>{formatImporteMovimiento(item.monto, item.moneda)}</td>
                <td className={styles.td}>
                  <div className={styles.actionsCell}>
                    {item.pagoId > 0 ? (
                      <GridIconActionButton
                        title="Ver detalle"
                        disabled={aprobandoPagoId > 0}
                        onClick={() => this.props.onVerDetalle(item.pagoId)}
                      >
                        <GridIconExternalLink />
                      </GridIconActionButton>
                    ) : (
                      '—'
                    )}
                    {mostrarAprobar && item.pagoId > 0 && (
                      <GridIconActionButton
                        title={aprobandoPagoId === item.pagoId ? 'Aprobando...' : 'Aprobar'}
                        disabled={aprobandoPagoId > 0}
                        onClick={() => this.props.onAprobar(item.pagoId)}
                      >
                        <GridIconCheck />
                      </GridIconActionButton>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
}
