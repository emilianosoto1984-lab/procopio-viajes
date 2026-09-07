import * as React from 'react';
import { MessageBar, MessageBarType, Pivot, PivotItem } from '@fluentui/react';
import styles from './VistaCajaConsolidada.module.scss';
import type { IMovimientosCajaProps } from './IMovimientosCajaProps';
import { ICuentaBancaria } from '../../../shared/cuentasBancariasUtils';
import { IFiltrosMovimientosCaja } from '../models/IFiltrosMovimientosCaja';
import { OrigenCaja } from '../models/OrigenCaja';
import CajaConsolidadaService from '../services/CajaConsolidadaService';
import { ICajaConsolidadaService } from '../services/ICajaConsolidadaService';
import {
  calcularTotalesMovimientosCaja,
  IMovimientoCajaLinea
} from '../utils/cajaMovimientosUtils';
import {
  createDefaultFiltrosMovimientosCaja,
  esPeriodoMovimientosInvalido,
  filtrarMovimientosCaja,
  MENSAJE_PERIODO_MOVIMIENTOS_INVALIDO
} from '../utils/movimientosCajaFiltrosUtils';
import {
  contarMovimientosCajaAprobados,
  contarMovimientosCajaPendientes,
  contarMovimientosCajaSinEstado,
  getMovimientosCajaDeVista,
  resolverVistaMovimientosCajaTrasCambio,
  VistaEstadoMovimientosCaja
} from '../utils/movimientosEstadoUtils';
import { construirOrigenesCaja, ORIGEN_EFECTIVO } from '../utils/origenCajaUtils';
import MovimientosCajaFiltros from './MovimientosCajaFiltros';
import MovimientosCajaGrid from './MovimientosCajaGrid';
import MovimientosCajaKpis from './MovimientosCajaKpis';
import MovimientosCajaSelector from './MovimientosCajaSelector';

export interface IMovimientosCajaState {
  cargandoCuentas: boolean;
  errorCuentas: string;
  cuentasBancarias: ICuentaBancaria[];
  origenSeleccionado: OrigenCaja;
  filtros: IFiltrosMovimientosCaja;
  filtrosAplicados: IFiltrosMovimientosCaja;
  cargandoMovimientos: boolean;
  errorMovimientos: string;
  errorDetalle: string;
  errorAprobacion: string;
  vistaEstado: VistaEstadoMovimientosCaja;
  aprobandoPagoId: number;
  movimientos: IMovimientoCajaLinea[];
}

/**
 * Vista Movimientos: selector de origen, indicadores y grilla de Registro de Pagos.
 */
export default class MovimientosCaja extends React.Component<
  IMovimientosCajaProps,
  IMovimientosCajaState
> {
  private readonly _service: ICajaConsolidadaService;
  private _cuentasSolicitadas: boolean = false;
  private _movimientosRequestId: number = 0;

  public constructor(props: IMovimientosCajaProps) {
    super(props);
    this._service = new CajaConsolidadaService(props.context);
    const filtrosIniciales = createDefaultFiltrosMovimientosCaja();
    this.state = {
      cargandoCuentas: false,
      errorCuentas: '',
      cuentasBancarias: [],
      origenSeleccionado: ORIGEN_EFECTIVO,
      filtros: filtrosIniciales,
      filtrosAplicados: filtrosIniciales,
      cargandoMovimientos: false,
      errorMovimientos: '',
      errorDetalle: '',
      errorAprobacion: '',
      vistaEstado: 'aprobados',
      aprobandoPagoId: 0,
      movimientos: []
    };
  }

  public componentDidMount(): void {
    this._ensureCuentasBancarias();
    if (this.props.activa) {
      void this._cargarMovimientos();
    }
  }

  public componentDidUpdate(prevProps: IMovimientosCajaProps): void {
    if (this.props.activa && !prevProps.activa) {
      this._ensureCuentasBancarias();
      void this._cargarMovimientos();
    }
  }

  public render(): React.ReactElement<IMovimientosCajaProps> {
    const {
      cargandoCuentas,
      errorCuentas,
      cuentasBancarias,
      origenSeleccionado,
      filtros,
      filtrosAplicados,
      cargandoMovimientos,
      errorMovimientos,
      errorDetalle,
      errorAprobacion,
      vistaEstado,
      aprobandoPagoId,
      movimientos
    } = this.state;
    const origenes = construirOrigenesCaja(cuentasBancarias);
    const periodoInvalido = esPeriodoMovimientosInvalido(filtros);
    const filtrosActivos = periodoInvalido
      ? { ...filtrosAplicados, moneda: filtros.moneda }
      : filtros;
    const movimientosFiltrados = filtrarMovimientosCaja(movimientos, filtrosActivos);
    const cantidadAprobados = contarMovimientosCajaAprobados(movimientosFiltrados);
    const cantidadPendientes = contarMovimientosCajaPendientes(movimientosFiltrados);
    const cantidadSinEstado = contarMovimientosCajaSinEstado(movimientosFiltrados);
    const movimientosVista = getMovimientosCajaDeVista(movimientosFiltrados, vistaEstado);
    const totales = calcularTotalesMovimientosCaja(movimientosVista);
    const mostrarAvisoPendientes = vistaEstado === 'aprobados' && cantidadPendientes > 0;
    const mostrarVacio =
      !cargandoCuentas &&
      !cargandoMovimientos &&
      !errorCuentas &&
      !errorMovimientos &&
      !periodoInvalido &&
      movimientosVista.length === 0;
    const mensajeVacio =
      vistaEstado === 'pendientes'
        ? 'No hay pagos pendientes de aprobación.'
        : 'No hay movimientos aprobados para el origen y el período seleccionados.';

    return (
      <div>
        <div className={styles.header}>
          <p className={styles.subtitle}>
            Ingresos y egresos de caja según el origen de fondos seleccionado.
          </p>
        </div>

        <MovimientosCajaSelector
          origenes={origenes}
          seleccionado={origenSeleccionado}
          disabled={cargandoCuentas}
          onChange={this._onCambiarOrigen}
        />

        <MovimientosCajaFiltros
          filtros={filtros}
          errorPeriodo={periodoInvalido ? MENSAJE_PERIODO_MOVIMIENTOS_INVALIDO : ''}
          onChange={this._onCambiarFiltros}
          onClear={this._onLimpiarFiltros}
        />

        <MovimientosCajaKpis
          ingresos={totales.ingresos}
          egresos={totales.egresos}
          saldo={totales.saldo}
        />

        {cargandoCuentas && (
          <div className={`${styles.statusMessage} ${styles.statusLoading}`}>
            Cargando cuentas bancarias...
          </div>
        )}

        {!!errorCuentas && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{errorCuentas}</div>
        )}

        {!cargandoCuentas && cargandoMovimientos && (
          <div className={`${styles.statusMessage} ${styles.statusLoading}`}>
            Cargando movimientos...
          </div>
        )}

        {!!errorMovimientos && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{errorMovimientos}</div>
        )}

        {!!errorDetalle && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{errorDetalle}</div>
        )}

        {!!errorAprobacion && (
          <div className={`${styles.statusMessage} ${styles.statusError}`}>{errorAprobacion}</div>
        )}

        <div className={styles.estadoTabs}>
          <Pivot
            selectedKey={vistaEstado}
            onLinkClick={this._onCambiarVistaEstado}
            headersOnly={true}
            getTabId={(itemKey: string) => 'movimientos-caja-vista-' + itemKey}
            aria-label="Vistas de movimientos por estado"
          >
            <PivotItem
              headerText={'Aprobados (' + cantidadAprobados + ')'}
              itemKey="aprobados"
            />
            <PivotItem
              headerText={'Pendientes (' + cantidadPendientes + ')'}
              itemKey="pendientes"
            />
          </Pivot>
        </div>

        {mostrarAvisoPendientes && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline={false}
            styles={{ root: { marginBottom: 10 } }}
          >
            {cantidadPendientes === 1
              ? 'Hay 1 pago pendiente de aprobación.'
              : 'Hay ' + cantidadPendientes + ' pagos pendientes de aprobación.'}
          </MessageBar>
        )}

        {vistaEstado === 'aprobados' && cantidadSinEstado > 0 && (
          <div className={styles.legacyHint}>
            Además, hay {cantidadSinEstado} movimiento
            {cantidadSinEstado === 1 ? '' : 's'} histórico
            {cantidadSinEstado === 1 ? '' : 's'} sin estado.
          </div>
        )}

        {mostrarVacio && (
          <div className={`${styles.statusMessage} ${styles.statusEmpty}`}>
            {mensajeVacio}
          </div>
        )}

        <MovimientosCajaGrid
          items={movimientosVista}
          vista={vistaEstado}
          aprobandoPagoId={aprobandoPagoId}
          onVerDetalle={this._onVerDetalle}
          onAprobar={this._onAprobar}
        />
      </div>
    );
  }

  private _ensureCuentasBancarias = (): void => {
    if (!this.props.activa || this._cuentasSolicitadas) {
      return;
    }
    this._cuentasSolicitadas = true;
    void this._cargarCuentasBancarias();
  };

  private _onCambiarOrigen = (origen: OrigenCaja): void => {
    this.setState({ origenSeleccionado: origen, movimientos: [] }, () => {
      void this._cargarMovimientos();
    });
  };

  private _onCambiarFiltros = (filtros: IFiltrosMovimientosCaja): void => {
    if (esPeriodoMovimientosInvalido(filtros)) {
      this.setState({ filtros });
      return;
    }
    this.setState({ filtros, filtrosAplicados: filtros });
  };

  private _onCambiarVistaEstado = (item?: PivotItem): void => {
    if (!item || !item.props.itemKey) {
      return;
    }
    const key = item.props.itemKey;
    if (key === 'aprobados' || key === 'pendientes') {
      this.setState({ vistaEstado: key, errorAprobacion: '' });
    }
  };

  private _onVerDetalle = (pagoId: number): void => {
    void this._abrirDetallePago(pagoId);
  };

  private _onAprobar = (pagoId: number): void => {
    void this._aprobarMovimiento(pagoId);
  };

  private async _aprobarMovimiento(pagoId: number): Promise<void> {
    if (!pagoId || pagoId <= 0 || this.state.aprobandoPagoId > 0) {
      return;
    }
    const movimiento = this.state.movimientos.filter(
      (item: IMovimientoCajaLinea) => item.pagoId === pagoId
    )[0];
    if (!movimiento) {
      return;
    }
    const confirmar = window.confirm('¿Confirma la aprobación de este movimiento?');
    if (!confirmar) {
      return;
    }
    this.setState({ aprobandoPagoId: pagoId, errorAprobacion: '' });
    try {
      await this._service.aprobarPago(pagoId);
      this.setState((prev) => {
        const movimientos = prev.movimientos.map((item: IMovimientoCajaLinea) => {
          if (item.pagoId !== pagoId) {
            return item;
          }
          return { ...item, estado: 'Aprobado' };
        });
        const filtrados = filtrarMovimientosCaja(movimientos, {
          ...prev.filtrosAplicados,
          moneda: prev.filtros.moneda
        });
        return {
          movimientos,
          vistaEstado: resolverVistaMovimientosCajaTrasCambio(prev.vistaEstado, filtrados),
          aprobandoPagoId: 0,
          errorAprobacion: ''
        };
      });
    } catch (error) {
      console.error('[MovimientosCaja] No se pudo aprobar el movimiento:', error);
      this.setState({
        aprobandoPagoId: 0,
        errorAprobacion: 'No se pudo aprobar el movimiento.'
      });
    }
  };

  private async _abrirDetallePago(pagoId: number): Promise<void> {
    if (!pagoId || pagoId <= 0) {
      return;
    }
    this.setState({ errorDetalle: '' });
    try {
      const urlDetallePago = await this._service.getPagoDisplayFormUrl(pagoId);
      window.open(urlDetallePago, '_blank', 'noopener,noreferrer');
    } catch (error) {
      console.error('[MovimientosCaja] No se pudo abrir el detalle del pago:', error);
      this.setState({
        errorDetalle: 'No se pudo abrir el detalle del pago en Registro de Pagos.'
      });
    }
  }

  private _onLimpiarFiltros = (): void => {
    const filtros = createDefaultFiltrosMovimientosCaja();
    this.setState({ filtros, filtrosAplicados: filtros });
  };

  private async _cargarCuentasBancarias(): Promise<void> {
    this.setState({ cargandoCuentas: true, errorCuentas: '' });
    try {
      const cuentas = await this._service.getCuentasBancarias();
      this.setState({
        cuentasBancarias: cuentas || [],
        cargandoCuentas: false,
        errorCuentas: ''
      });
    } catch (error) {
      console.error('[MovimientosCaja] Error al cargar cuentas bancarias:', error);
      this.setState({
        cuentasBancarias: [],
        cargandoCuentas: false,
        errorCuentas: 'No se pudieron cargar las cuentas bancarias.'
      });
    }
  }

  private async _cargarMovimientos(): Promise<void> {
    const requestId = ++this._movimientosRequestId;
    const origen = this.state.origenSeleccionado;
    this.setState({ cargandoMovimientos: true, errorMovimientos: '' });
    try {
      const movimientos = await this._service.getMovimientos(origen);
      if (requestId !== this._movimientosRequestId) {
        return;
      }
      this.setState({
        movimientos: movimientos || [],
        cargandoMovimientos: false,
        errorMovimientos: ''
      });
    } catch (error) {
      console.error('[MovimientosCaja] Error al cargar movimientos:', error);
      if (requestId !== this._movimientosRequestId) {
        return;
      }
      this.setState({
        movimientos: [],
        cargandoMovimientos: false,
        errorMovimientos: 'No se pudieron cargar los movimientos.'
      });
    }
  }
}
