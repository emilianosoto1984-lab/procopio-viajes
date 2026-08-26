import * as React from 'react';
import { formatDateDisplay } from '../../../shared/sharePointDateUtils';
import { formatImporteArs, formatImporteUsd } from '../utils/monedaUtils';
import { IEgresoSinViajeCaja } from '../utils/cajaMovimientosUtils';
import { normalizarMonedaPago } from '../../../shared/pagoMonedaUtils';
import styles from './VistaCajaConsolidada.module.scss';

export interface ICajaEgresosSinViajeProps {
  items: IEgresoSinViajeCaja[];
}

function formatImporteEgreso(monto: number, moneda: string): string {
  return normalizarMonedaPago(moneda) === 'Dólares'
    ? formatImporteUsd(monto)
    : formatImporteArs(monto);
}

/**
 * Sección independiente: egresos sin ViajeAsociado.
 */
export default class CajaEgresosSinViaje extends React.Component<ICajaEgresosSinViajeProps> {
  public render(): React.ReactElement<ICajaEgresosSinViajeProps> {
    const { items } = this.props;

    return (
      <div style={{ marginTop: 24 }}>
        <h3 className={styles.title} style={{ fontSize: 16, marginBottom: 8 }}>
          Egresos generales sin viaje asociado
        </h3>
        {(!items || items.length === 0) && (
          <div className={`${styles.statusMessage} ${styles.statusEmpty}`}>
            No hay egresos generales sin viaje asociado.
          </div>
        )}
        {!!items && items.length > 0 && (
          <div className={styles.tableOuter}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th className={styles.th}>Fecha</th>
                  <th className={styles.th}>Concepto</th>
                  <th className={styles.th}>Medio de pago</th>
                  <th className={styles.th}>Cuenta bancaria</th>
                  <th className={styles.th}>Moneda</th>
                  <th className={styles.th}>Importe</th>
                  <th className={styles.th}>Estado</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item: IEgresoSinViajeCaja) => (
                  <tr key={item.id}>
                    <td className={styles.td}>
                      {item.fechaPago ? formatDateDisplay(item.fechaPago) : '—'}
                    </td>
                    <td className={styles.td}>{item.concepto}</td>
                    <td className={styles.td}>{item.medioPago || '—'}</td>
                    <td className={styles.td}>{item.cuentaBancariaLabel || '—'}</td>
                    <td className={styles.td}>{item.moneda || '—'}</td>
                    <td className={styles.td}>
                      {formatImporteEgreso(item.monto, item.moneda)}
                    </td>
                    <td className={styles.td}>{item.estado || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  }
}
