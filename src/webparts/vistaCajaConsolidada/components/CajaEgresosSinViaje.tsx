import * as React from 'react';
import { formatImporteArs, formatImporteUsd } from '../utils/monedaUtils';
import { IEgresoSinViajeCaja } from '../utils/cajaMovimientosUtils';
import { normalizarMonedaPago } from '../../../shared/pagoMonedaUtils';
import styles from './VistaCajaConsolidada.module.scss';

export interface ICajaEgresosSinViajeProps {
  items: IEgresoSinViajeCaja[];
}

interface IEgresoAgrupado {
  motivo: string;
  totalArs: number;
  totalUsd: number;
}

function agruparPorMotivo(items: IEgresoSinViajeCaja[]): IEgresoAgrupado[] {
  const mapa: { [motivo: string]: IEgresoAgrupado } = {};

  (items || []).forEach((item: IEgresoSinViajeCaja) => {
    const key = item.motivo || 'Sin motivo';
    if (!mapa[key]) {
      mapa[key] = { motivo: key, totalArs: 0, totalUsd: 0 };
    }
    if (normalizarMonedaPago(item.moneda) === 'Dólares') {
      mapa[key].totalUsd += item.monto;
    } else {
      mapa[key].totalArs += item.monto;
    }
  });

  return Object.keys(mapa)
    .sort()
    .map((k: string) => mapa[k]);
}

export default class CajaEgresosSinViaje extends React.Component<ICajaEgresosSinViajeProps> {
  public render(): React.ReactElement<ICajaEgresosSinViajeProps> {
    const { items } = this.props;
    const agrupados = agruparPorMotivo(items);

    return (
      <div style={{ marginTop: 24 }}>
        <h3 className={styles.title} style={{ fontSize: 16, marginBottom: 8 }}>
          Egresos generales sin viaje asociado
        </h3>
        {agrupados.length === 0 && (
          <div className={`${styles.statusMessage} ${styles.statusEmpty}`}>
            No hay egresos generales sin viaje asociado.
          </div>
        )}
        {agrupados.length > 0 && (
          <div className={styles.tableOuter}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th className={styles.th}>Motivo</th>
                  <th className={styles.th}>Total pesos</th>
                  <th className={styles.th}>Total dólares</th>
                </tr>
              </thead>
              <tbody>
                {agrupados.map((g: IEgresoAgrupado) => (
                  <tr key={g.motivo}>
                    <td className={styles.td}>{g.motivo}</td>
                    <td className={styles.td}>{g.totalArs ? formatImporteArs(g.totalArs) : '—'}</td>
                    <td className={styles.td}>{g.totalUsd ? formatImporteUsd(g.totalUsd) : '—'}</td>
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
