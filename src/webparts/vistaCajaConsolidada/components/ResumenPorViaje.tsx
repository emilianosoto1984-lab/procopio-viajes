import * as React from 'react';
import styles from './VistaCajaConsolidada.module.scss';
import type { IResumenPorViajeProps } from './IResumenPorViajeProps';
import { ICajaConsolidadaItem } from '../models/ICajaConsolidadaItem';
import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';
import { ITotalesGeneralesCaja } from '../models/ITotalesGeneralesCaja';
import CajaConsolidadaService from '../services/CajaConsolidadaService';
import { ICajaConsolidadaService } from '../services/ICajaConsolidadaService';
import { createEmptyFiltrosCajaConsolidada } from '../utils/cajaConsolidadaUtils';
import {
  createEmptyTotalesGeneralesCaja,
  IEgresoSinViajeCaja
} from '../utils/cajaMovimientosUtils';
import CajaConsolidadaFilters from './CajaConsolidadaFilters';
import CajaConsolidadaGrid from './CajaConsolidadaGrid';
import CajaConsolidadaKpis from './CajaConsolidadaKpis';
import CajaEgresosSinViaje from './CajaEgresosSinViaje';

export interface IResumenPorViajeState {
  cargando: boolean;
  error: string;
  items: ICajaConsolidadaItem[];
  egresosSinViaje: IEgresoSinViajeCaja[];
  totalesGenerales: ITotalesGeneralesCaja;
  filtros: IFiltrosCajaConsolidada;
}

/**
 * Vista Resumen por viaje: indicadores, filtros y listado consolidado (solo lectura).
 */
export default class ResumenPorViaje extends React.Component<
  IResumenPorViajeProps,
  IResumenPorViajeState
> {
  private readonly _service: ICajaConsolidadaService;

  public constructor(props: IResumenPorViajeProps) {
    super(props);
    this._service = new CajaConsolidadaService(props.context);
    this.state = {
      cargando: true,
      error: '',
      items: [],
      egresosSinViaje: [],
      totalesGenerales: createEmptyTotalesGeneralesCaja(),
      filtros: createEmptyFiltrosCajaConsolidada()
    };
  }

  public componentDidMount(): void {
    void this._cargarItems();
  }

  public render(): React.ReactElement<IResumenPorViajeProps> {
    const {
      cargando,
      error,
      items,
      egresosSinViaje,
      totalesGenerales,
      filtros
    } = this.state;
    const mostrarVacio = !cargando && !error && items.length === 0;
    const mostrarGrilla = !cargando && !error && items.length > 0;
    const mostrarKpis = !cargando && !error;
    const mostrarEgresos = !cargando && !error;

    return (
      <div>
        <div className={styles.header}>
          <p className={styles.subtitle}>
            Resultado de caja (recibido − egresos) separado del cobro pendiente por viaje.
            El recupero bancario ya está incluido en el total recibido.
          </p>
        </div>

        {mostrarKpis && (
          <CajaConsolidadaKpis
            cantidadViajes={items.length}
            totalRecibido={totalesGenerales.totalRecibido}
            totalEgresos={totalesGenerales.totalEgresos}
            totalRecupero={totalesGenerales.totalRecupero}
          />
        )}

        <CajaConsolidadaFilters
          filtros={filtros}
          disabled={cargando}
          onChange={this._onCambiarFiltros}
          onApply={this._onAplicarFiltros}
          onClear={this._onLimpiarFiltros}
        />

        {cargando && (
          <div className={`${styles.statusMessage} ${styles.statusLoading}`}>
            Cargando viajes...
          </div>
        )}

        {!!error && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{error}</div>
        )}

        {mostrarVacio && (
          <div className={`${styles.statusMessage} ${styles.statusEmpty}`}>
            No hay viajes para mostrar.
          </div>
        )}

        {mostrarGrilla && <CajaConsolidadaGrid items={items} />}

        {mostrarEgresos && <CajaEgresosSinViaje items={egresosSinViaje} />}
      </div>
    );
  }

  private _onCambiarFiltros = (filtros: IFiltrosCajaConsolidada): void => {
    const cambioMostrarCero =
      filtros.mostrarSaldosEnCero !== this.state.filtros.mostrarSaldosEnCero;
    this.setState({ filtros }, () => {
      if (cambioMostrarCero) {
        void this._cargarItems();
      }
    });
  };

  private _onAplicarFiltros = (): void => {
    void this._cargarItems();
  };

  private _onLimpiarFiltros = (): void => {
    this.setState({ filtros: createEmptyFiltrosCajaConsolidada() }, () => {
      void this._cargarItems();
    });
  };

  private async _cargarItems(): Promise<void> {
    this.setState({ cargando: true, error: '' });
    try {
      const data = await this._service.getVistaData(this.state.filtros);
      this.setState({
        items: data.items || [],
        egresosSinViaje: data.egresosSinViaje || [],
        totalesGenerales: data.totalesGenerales || createEmptyTotalesGeneralesCaja(),
        cargando: false
      });
    } catch (error) {
      console.error('[ResumenPorViaje] Error al cargar items:', error);
      this.setState({
        cargando: false,
        items: [],
        egresosSinViaje: [],
        totalesGenerales: createEmptyTotalesGeneralesCaja(),
        error: 'No se pudieron cargar los datos de la vista consolidada.'
      });
    }
  }
}
