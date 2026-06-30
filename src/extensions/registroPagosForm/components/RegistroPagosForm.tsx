import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import SharePointPagosService, {
  getEstadoByMedioPago,
  IComprobanteItem,
  IPagoSaldoItem,
  IPasajeroLookupItem,
  IServicioViajeItem,
  IViajeLookupItem
} from '../services/SharePointPagosService';
import {
  formatDateDisplay,
  getDateOnlyFromSharePoint
} from '../../../shared/sharePointDateUtils';
import {
  normalizarMonedaServicio,
  requiereCotizacionPago
} from '../../../shared/pagoMonedaUtils';
import {
  formatSaldoPendienteDisplay,
  getSaldoPendienteServicio,
  montoExcedeSaldoPendiente
} from '../../../shared/pagoSaldoUtils';
import styles from './RegistroPagosForm.module.scss';

export interface IRegistroPagosFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'RegistroPagosForm';
const MOTIVO_VIAJE = 'Viaje';

const MEDIOS_PAGO = ['Efectivo', 'Transferencia', 'Tarjeta de Credito'] as const;
const MONEDAS = ['Dólares', 'Pesos'] as const;
const TIPOS_PAGO = ['Ingreso', 'Egreso'] as const;
const TIPOS_INGRESO = [
  { key: 'asociado_viaje', label: 'Asociado a viaje' },
  { key: 'sin_viaje', label: 'Sin viaje asociado' }
] as const;

type MedioPago = typeof MEDIOS_PAGO[number];
type Moneda = typeof MONEDAS[number];
type TipoPago = typeof TIPOS_PAGO[number];
type TipoIngresoPago = typeof TIPOS_INGRESO[number]['key'];

interface IFieldErrors {
  tipoPago: string;
  viaje: string;
  pasajero: string;
  pasajeroNombre: string;
  pasajeroDni: string;
  concepto: string;
  medioPago: string;
  banco: string;
  fechaPago: string;
  monto: string;
  moneda: string;
  cotizacion: string;
}

interface IRegistroPagosFormState {
  pagoId: number | null;
  tipoPago: TipoPago | '';
  tipoIngreso: TipoIngresoPago;
  viajeId: number | null;
  viajeTitulo: string;
  viajeBusquedaTexto: string;
  viajesResultados: IViajeLookupItem[];
  viajeBuscando: boolean;
  pasajeroId: number | null;
  pasajeroTitulo: string;
  pasajerosViaje: IPasajeroLookupItem[];
  pasajerosViajeCargando: boolean;
  pasajeroNombre: string;
  pasajeroDni: string;
  servicioAsociadoId: number | null;
  concepto: string;
  conceptoMoneda: string;
  serviciosViaje: IServicioViajeItem[];
  serviciosViajeCargando: boolean;
  pagosViaje: IPagoSaldoItem[];
  medioPago: string;
  banco: string;
  opcionesBanco: string[];
  motivo: string;
  opcionesMotivo: string[];
  fechaPago: string;
  monto: string;
  moneda: string;
  cotizacion: string;
  estado: string;
  cargando: boolean;
  guardando: boolean;
  aprobando: boolean;
  mostrarDialogoAprobacion: boolean;
  mensajeExito: string;
  error: string;
  fieldErrors: IFieldErrors;
  comprobantes: IComprobanteItem[];
  comprobantesPendientes: File[];
  comprobantesSubiendo: boolean;
  comprobantesError: string;
  mostrarEditorComprobante: boolean;
}

const layoutStyles: { [key: string]: React.CSSProperties } = {
  page: {
    fontFamily: 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif',
    fontSize: 14,
    backgroundColor: '#f3f2f1',
    padding: 16,
    boxSizing: 'border-box'
  },
  container: { maxWidth: 1100, margin: '0 auto' },
  header: { marginBottom: 16 },
  title: { fontSize: 24, fontWeight: 600, marginBottom: 4 },
  subtitle: { fontSize: 14, color: '#605e5c' },
  section: {
    backgroundColor: '#ffffff',
    borderRadius: 8,
    border: '1px solid #e1dfdd',
    padding: 16,
    marginBottom: 16,
    boxShadow: '0 1px 2px rgba(0,0,0,0.04)'
  },
  sectionHeader: { fontSize: 16, fontWeight: 600, marginBottom: 12 },
  fieldGroup: { display: 'flex', flexDirection: 'column', marginBottom: 12 },
  label: { fontWeight: 600, marginBottom: 4 },
  input: {
    padding: '6px 8px',
    borderRadius: 4,
    border: '1px solid #c8c6c4',
    fontSize: 14
  },
  fieldInput: {
    width: '100%',
    height: 38,
    padding: '0 14px',
    borderRadius: 4,
    border: '1.4px solid #CDD0D7',
    background: '#FAFAFC',
    fontSize: 15,
    color: '#232529',
    outline: 'none',
    boxSizing: 'border-box',
    lineHeight: 1.2
  },
  fieldInputError: {
    borderColor: '#a4262c'
  },
  select: {
    width: '100%',
    height: 38,
    padding: '0 14px',
    borderRadius: 4,
    border: '1.4px solid #CDD0D7',
    background: '#FAFAFC',
    fontSize: 15,
    color: '#232529',
    outline: 'none',
    boxSizing: 'border-box'
  },
  primaryButton: {
    padding: '6px 14px',
    borderRadius: 4,
    border: '1px solid #0078d4',
    backgroundColor: '#0078d4',
    color: '#ffffff',
    cursor: 'pointer',
    fontSize: 14,
    fontWeight: 600,
    marginRight: 8
  },
  defaultButton: {
    padding: '6px 14px',
    borderRadius: 4,
    border: '1px solid #c8c6c4',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    fontSize: 14,
    fontWeight: 400,
    marginRight: 8
  },
  buttonDisabled: { opacity: 0.5, cursor: 'default' },
  bottomActions: { display: 'flex', justifyContent: 'flex-end', marginTop: 16 },
  error: { color: '#a4262c', marginBottom: 10, fontSize: 13 },
  fieldError: { color: '#a4262c', fontSize: 12, marginTop: 4 },
  autocompleteList: {
    border: '1px solid #c8c6c4',
    borderRadius: 6,
    backgroundColor: '#ffffff',
    maxHeight: 140,
    overflowY: 'auto',
    marginTop: 4,
    position: 'relative',
    zIndex: 2
  },
  autocompleteItem: {
    padding: '8px 12px',
    fontSize: 14,
    cursor: 'pointer',
    borderBottom: '1px solid #f3f2f1'
  },
  formGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: 16
  },
  fieldFullWidth: {
    gridColumn: '1 / -1'
  },
  choiceGroup: {
    display: 'flex',
    flexWrap: 'wrap' as const,
    gap: 16
  },
  choiceOption: {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    fontSize: 14,
    color: '#232529',
    cursor: 'pointer'
  },
  choiceOptionDisabled: {
    opacity: 0.6,
    cursor: 'default'
  },
  tipoPagoGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
    gap: 12
  },
  tipoPagoCard: {
    padding: '22px 18px',
    borderRadius: 8,
    border: '2px solid #e1dfdd',
    background: '#faf9f8',
    cursor: 'pointer',
    textAlign: 'center' as const,
    transition: 'border-color 0.18s, background-color 0.18s, box-shadow 0.18s',
    boxSizing: 'border-box'
  },
  tipoPagoCardDisabled: {
    opacity: 0.6,
    cursor: 'default'
  },
  tipoPagoCardTitle: {
    fontSize: 18,
    fontWeight: 600,
    color: '#323130',
    marginBottom: 4
  },
  tipoPagoCardHint: {
    fontSize: 13,
    color: '#605e5c',
    lineHeight: 1.35
  },
  tipoPagoCardIngresoActive: {
    borderColor: '#107c10',
    background: '#f3faf3',
    boxShadow: '0 0 0 1px #107c10'
  },
  tipoPagoCardEgresoActive: {
    borderColor: '#c50f1f',
    background: '#fdf3f4',
    boxShadow: '0 0 0 1px #c50f1f'
  },
  tipoPagoDisplayBadge: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    minWidth: 160,
    padding: '12px 24px',
    borderRadius: 8,
    fontSize: 18,
    fontWeight: 600,
    letterSpacing: '-0.2px'
  },
  tipoPagoDisplayIngreso: {
    color: '#0b6a0b',
    background: '#f3faf3',
    border: '2px solid #9fd89f'
  },
  tipoPagoDisplayEgreso: {
    color: '#a4262c',
    background: '#fdf3f4',
    border: '2px solid #f1bbbc'
  },
  saldoPendienteInfo: {
    fontSize: 13,
    color: '#323130',
    marginTop: 6,
    lineHeight: 1.4
  },
  saldoPendienteWarning: {
    fontSize: 13,
    color: '#8a6d00',
    marginTop: 6,
    lineHeight: 1.4
  },
  readOnlyValue: {
    minHeight: 38,
    display: 'flex',
    alignItems: 'center',
    padding: '0 14px',
    borderRadius: 4,
    border: '1.4px solid #e1dfdd',
    background: '#f3f2f1',
    fontSize: 15,
    color: '#323130'
  },
  readOnlyDateValue: {
    width: '100%',
    height: 38,
    padding: '0 14px',
    borderRadius: 8,
    border: '1.4px solid #CDD0D7',
    background: '#F3F2F1',
    fontSize: 15,
    color: '#605E5C',
    boxSizing: 'border-box',
    display: 'flex',
    alignItems: 'center'
  },
  dateInputWrap: {
    position: 'relative' as const
  },
  dateInput: {
    width: '100%',
    height: 38,
    padding: '0 14px',
    borderRadius: 8,
    border: '1.4px solid #CDD0D7',
    background: '#FAFAFC',
    fontSize: 15,
    color: 'transparent',
    caretColor: 'transparent',
    outline: 'none',
    transition: 'border-color 0.18s',
    boxSizing: 'border-box'
  },
  dateInputError: {
    borderColor: '#a4262c'
  },
  dateInputOverlay: {
    position: 'absolute' as const,
    left: 14,
    right: 42,
    top: 0,
    height: 38,
    display: 'flex',
    alignItems: 'center',
    fontSize: 15,
    color: '#232529',
    pointerEvents: 'none' as const,
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  approvalCardPending: {
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'space-between',
    gap: 16,
    background: '#FFFBF0',
    border: '1px solid #F2D189',
    borderRadius: 8,
    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
    padding: '16px 18px',
    marginBottom: 16
  },
  approvalCardApproved: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    background: '#F3FAF3',
    border: '1px solid #9FD89F',
    borderRadius: 8,
    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
    padding: '14px 18px',
    marginBottom: 16
  },
  approvalCardTitle: { fontSize: 16, fontWeight: 600, marginBottom: 4, color: '#323130' },
  approvalCardText: { fontSize: 14, color: '#605e5c', lineHeight: 1.4 },
  approvalCardActions: { flexShrink: 0, alignSelf: 'center' },
  success: { color: '#107c10', marginBottom: 10, fontSize: 13 },
  dialogOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 0, 0, 0.4)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 1000
  },
  dialogBox: {
    backgroundColor: '#ffffff',
    borderRadius: 8,
    border: '1px solid #e1dfdd',
    boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
    padding: 24,
    maxWidth: 440,
    width: 'calc(100% - 32px)'
  },
  dialogTitle: { fontSize: 18, fontWeight: 600, marginBottom: 12, color: '#323130' },
  dialogText: { fontSize: 14, color: '#605e5c', lineHeight: 1.5, marginBottom: 20 },
  dialogActions: { display: 'flex', justifyContent: 'flex-end', gap: 8 },
  sectionToolbar: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 },
  toolbarTitle: { fontSize: 16, fontWeight: 600 },
  info: { marginBottom: 10, color: '#605e5c', fontSize: 13 },
  fileInputWrap: { display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 },
  fileInputHidden: {
    position: 'absolute' as const,
    width: 1,
    height: 1,
    padding: 0,
    margin: -1,
    overflow: 'hidden',
    clip: 'rect(0, 0, 0, 0)',
    whiteSpace: 'nowrap',
    border: 0
  },
  fileInputButton: {
    display: 'inline-block',
    padding: '6px 14px',
    borderRadius: 4,
    border: '1px solid #c8c6c4',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    fontSize: 14,
    fontWeight: 600,
    color: '#323130',
    lineHeight: 1.35
  },
  fileInputHint: { fontSize: 12, color: '#605e5c' },
  attachmentZone: {
    border: '1px solid #e1dfdd',
    borderRadius: 6,
    backgroundColor: '#faf9f8',
    overflow: 'hidden'
  },
  attachmentEmpty: { padding: '22px 16px', color: '#7b8190', fontSize: 13.5, textAlign: 'center' as const },
  attachmentRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    padding: '12px 16px',
    borderBottom: '1px solid #eceef3',
    backgroundColor: '#ffffff'
  },
  attachmentRowLast: { borderBottom: 'none' },
  attachmentMeta: { flex: 1, minWidth: 0 },
  attachmentFileName: { fontSize: 14, fontWeight: 600, color: '#2b2f38', wordBreak: 'break-word' as const },
  attachmentFileSize: { fontSize: 12, color: '#605e5c', marginTop: 2 },
  attachmentActions: { display: 'flex', gap: 4, flexShrink: 0 },
  attachmentActionButton: {
    padding: '4px 10px',
    borderRadius: 4,
    border: '1px solid #c8c6c4',
    backgroundColor: '#ffffff',
    cursor: 'pointer',
    fontSize: 12,
    fontWeight: 600,
    color: '#323130'
  },
  uploadingRow: { display: 'flex', alignItems: 'center', gap: 8, padding: '12px 16px', color: '#605e5c', fontSize: 13 }
};

interface IFileInputEspanolProps {
  inputId: string;
  multiple?: boolean;
  disabled?: boolean;
  buttonLabel?: string;
  onChange: (ev: React.ChangeEvent<HTMLInputElement>) => void;
}

const FileInputEspanol: React.FC<IFileInputEspanolProps> = (props: IFileInputEspanolProps) => {
  const { inputId, multiple, disabled, buttonLabel, onChange } = props;
  const etiqueta = buttonLabel || (multiple ? 'Seleccionar archivos' : 'Seleccionar archivo');
  return (
    <div style={layoutStyles.fileInputWrap}>
      <input
        id={inputId}
        type="file"
        multiple={multiple}
        disabled={disabled}
        onChange={onChange}
        style={layoutStyles.fileInputHidden}
        aria-label={etiqueta}
      />
      <label
        htmlFor={inputId}
        style={
          disabled
            ? { ...layoutStyles.fileInputButton, ...layoutStyles.buttonDisabled, pointerEvents: 'none' as const }
            : layoutStyles.fileInputButton
        }
      >
        {etiqueta}
      </label>
    </div>
  );
};

const emptyFieldErrors = (): IFieldErrors => ({
  tipoPago: '',
  viaje: '',
  pasajero: '',
  pasajeroNombre: '',
  pasajeroDni: '',
  concepto: '',
  medioPago: '',
  banco: '',
  fechaPago: '',
  monto: '',
  moneda: '',
  cotizacion: ''
});

export default class RegistroPagosForm extends React.Component<IRegistroPagosFormProps, IRegistroPagosFormState> {
  private readonly _service: SharePointPagosService;
  private _viajeSearchTimer: number | undefined;

  public constructor(props: IRegistroPagosFormProps) {
    super(props);
    this._service = new SharePointPagosService(props.context);
    this.state = {
      pagoId: null,
      tipoPago: 'Ingreso',
      tipoIngreso: 'asociado_viaje',
      viajeId: null,
      viajeTitulo: '',
      viajeBusquedaTexto: '',
      viajesResultados: [],
      viajeBuscando: false,
      pasajeroId: null,
      pasajeroTitulo: '',
      pasajerosViaje: [],
      pasajerosViajeCargando: false,
      pasajeroNombre: '',
      pasajeroDni: '',
      servicioAsociadoId: null,
      concepto: '',
      conceptoMoneda: '',
      serviciosViaje: [],
      serviciosViajeCargando: false,
      pagosViaje: [],
      medioPago: '',
      banco: '',
      opcionesBanco: [],
      motivo: '',
      opcionesMotivo: [],
      fechaPago: '',
      monto: '',
      moneda: '',
      cotizacion: '',
      estado: '',
      cargando: true,
      guardando: false,
      aprobando: false,
      mostrarDialogoAprobacion: false,
      mensajeExito: '',
      error: '',
      fieldErrors: emptyFieldErrors(),
      comprobantes: [],
      comprobantesPendientes: [],
      comprobantesSubiendo: false,
      comprobantesError: '',
      mostrarEditorComprobante: false
    };
  }

  public componentDidMount(): void {
    console.log('=== RegistroPagosForm inicializado ===');
    Log.info(LOG_SOURCE, 'RegistroPagosForm mounted');
    void this._inicializarFormulario();
  }

  public componentWillUnmount(): void {
    if (this._viajeSearchTimer) {
      window.clearTimeout(this._viajeSearchTimer);
    }
    Log.info(LOG_SOURCE, 'RegistroPagosForm unmounted');
  }

  private _normalizarTipoPago(value: string): TipoPago | '' {
    const trimmed = (value || '').trim();
    if (trimmed === 'Ingreso' || trimmed === 'Egreso') {
      return trimmed;
    }
    return '';
  }

  private _resolverMotivoViaje(opcionesMotivo: string[]): string {
    if (opcionesMotivo.indexOf(MOTIVO_VIAJE) >= 0) {
      console.log('Tipo asociado a viaje: Motivo seteado en Viaje');
      return MOTIVO_VIAJE;
    }
    console.warn('Motivo Viaje no encontrado en opciones de SharePoint');
    return '';
  }

  private _esAsociadoAViaje(): boolean {
    return this.state.tipoIngreso === 'asociado_viaje';
  }

  private _getEtiquetaTipoIngreso(tipo: TipoIngresoPago): string {
    const opcion = TIPOS_INGRESO.filter((item) => item.key === tipo)[0];
    return opcion ? opcion.label : tipo;
  }

  private _getPasajeroSeleccionadoTitulo(): string {
    if (this.state.pasajeroTitulo) {
      return this.state.pasajeroTitulo;
    }
    if (this.state.pasajeroId) {
      const pasajero = this.state.pasajerosViaje.filter(
        (item: IPasajeroLookupItem) => item.id === this.state.pasajeroId
      )[0];
      if (pasajero) {
        return pasajero.nombreApellido;
      }
    }
    if (this.state.pasajeroNombre) {
      return this.state.pasajeroNombre;
    }
    return '';
  }

  private async _cargarPasajerosViaje(viajeId: number, pasajeroIdPrecargado?: number | null): Promise<void> {
    if (!viajeId || viajeId <= 0) {
      this.setState({ pasajerosViaje: [], pasajerosViajeCargando: false });
      return;
    }

    this.setState({ pasajerosViajeCargando: true });
    try {
      const pasajerosViaje = await this._service.getPasajerosByViaje(viajeId);
      const pasajeroId = pasajeroIdPrecargado && pasajeroIdPrecargado > 0 ? pasajeroIdPrecargado : null;
      let pasajeroTitulo = '';
      if (pasajeroId) {
        const pasajero = pasajerosViaje.filter((item: IPasajeroLookupItem) => item.id === pasajeroId)[0];
        pasajeroTitulo = pasajero ? pasajero.nombreApellido : '';
      }
      this.setState({
        pasajerosViaje,
        pasajerosViajeCargando: false,
        pasajeroId,
        pasajeroTitulo
      });
    } catch (error) {
      this.setState({
        pasajerosViaje: [],
        pasajerosViajeCargando: false,
        error: 'No se pudieron cargar los pasajeros del viaje seleccionado.'
      });
    }
  }

  private _requiereCotizacion(): boolean {
    if (!this._esAsociadoAViaje()) {
      return false;
    }
    return requiereCotizacionPago(this.state.conceptoMoneda, this.state.moneda);
  }

  private _getServicioSeleccionado(): IServicioViajeItem | undefined {
    if (!this.state.servicioAsociadoId) {
      return undefined;
    }
    return this.state.serviciosViaje.filter(
      (item: IServicioViajeItem) => item.id === this.state.servicioAsociadoId
    )[0];
  }

  private _getSaldoPendienteServicio(): number | null {
    const servicio = this._getServicioSeleccionado();
    if (!servicio || !this._esAsociadoAViaje()) {
      return null;
    }
    const pagoIdExcluir = this.state.pagoId && this.state.pagoId > 0 ? this.state.pagoId : undefined;
    return getSaldoPendienteServicio(servicio, this.state.pagosViaje, pagoIdExcluir);
  }

  private async _cargarPagosViaje(viajeId: number): Promise<void> {
    if (!viajeId || viajeId <= 0) {
      this.setState({ pagosViaje: [] });
      return;
    }

    try {
      const pagosViaje = await this._service.getPagosSaldoByViaje(viajeId);
      this.setState({ pagosViaje });
    } catch (error) {
      this.setState({ pagosViaje: [] });
    }
  }

  private _renderSaldoPendienteServicio(): React.ReactNode {
    if (!this._esAsociadoAViaje()) {
      return null;
    }

    const servicio = this._getServicioSeleccionado();
    if (!servicio) {
      return null;
    }

    const saldoPendiente = this._getSaldoPendienteServicio();
    if (saldoPendiente === null) {
      return null;
    }

    if (saldoPendiente <= 0) {
      return (
        <div style={layoutStyles.saldoPendienteWarning}>Este servicio no tiene saldo pendiente.</div>
      );
    }

    const cotizacionNumerica = Number(this.state.cotizacion);
    const cotizacion =
      isFinite(cotizacionNumerica) && cotizacionNumerica > 0 ? cotizacionNumerica : undefined;

    return (
      <div style={layoutStyles.saldoPendienteInfo}>
        {formatSaldoPendienteDisplay(servicio, saldoPendiente, this.state.moneda, cotizacion)}
      </div>
    );
  }

  private _limpiarConceptoSeleccionado(): {
    servicioAsociadoId: null;
    concepto: string;
    conceptoMoneda: string;
    cotizacion: string;
    serviciosViaje: IServicioViajeItem[];
    serviciosViajeCargando: boolean;
  } {
    return {
      servicioAsociadoId: null,
      concepto: '',
      conceptoMoneda: '',
      cotizacion: '',
      serviciosViaje: [],
      serviciosViajeCargando: false
    };
  }

  private async _cargarServiciosViaje(
    viajeId: number,
    conceptoPrecargado?: string,
    servicioIdPrecargado?: number | null
  ): Promise<void> {
    if (!viajeId || viajeId <= 0) {
      this.setState(this._limpiarConceptoSeleccionado());
      return;
    }

    this.setState({ serviciosViajeCargando: true });
    try {
      const serviciosViaje = await this._service.getServiciosViajeByViaje(viajeId);
      let servicioAsociadoId: number | null = null;
      let concepto = '';
      let conceptoMoneda = '';

      if (servicioIdPrecargado && servicioIdPrecargado > 0) {
        const servicio = serviciosViaje.filter((item: IServicioViajeItem) => item.id === servicioIdPrecargado)[0];
        if (servicio) {
          servicioAsociadoId = servicio.id;
          concepto = servicio.concepto;
          conceptoMoneda = normalizarMonedaServicio(servicio.moneda);
        }
      } else if (conceptoPrecargado) {
        const servicio = serviciosViaje.filter(
          (item: IServicioViajeItem) => item.concepto === conceptoPrecargado
        )[0];
        if (servicio) {
          servicioAsociadoId = servicio.id;
          concepto = servicio.concepto;
          conceptoMoneda = normalizarMonedaServicio(servicio.moneda);
        } else {
          concepto = conceptoPrecargado;
        }
      }

      this.setState({
        serviciosViaje,
        serviciosViajeCargando: false,
        servicioAsociadoId,
        concepto,
        conceptoMoneda
      });
    } catch (error) {
      this.setState({
        ...this._limpiarConceptoSeleccionado(),
        error: 'No se pudieron cargar los servicios del viaje seleccionado.'
      });
    }
  }

  private _requiereCuentaBancaria(): boolean {
    return this.state.medioPago === 'Transferencia';
  }

  private _logVisibilidadCuentaBancaria(medioPago: string): void {
    if (medioPago === 'Transferencia') {
      console.log('MedioPago Transferencia: mostrando Cuenta Bancaria');
      return;
    }
    if (medioPago) {
      console.log('MedioPago no transferencia: ocultando Cuenta Bancaria');
    }
  }

  private _logEstadoCotizacion(moneda: string): void {
    if (this._requiereCotizacion()) {
      console.log('Monedas distintas entre servicio y pago: mostrando Cotizacion');
      return;
    }
    if (moneda) {
      console.log('Cotizacion no requerida para la combinacion actual de monedas');
    }
  }

  private _logEstadoPorMedioPago(medioPago: string): void {
    switch ((medioPago || '').trim()) {
      case 'Transferencia':
        console.log('MedioPago Transferencia: Estado Pendiente');
        break;
      case 'Efectivo':
        console.log('MedioPago Efectivo: Estado Aprobado');
        break;
      case 'Tarjeta de Credito':
        console.log('MedioPago Tarjeta de Credito: Estado Pendiente');
        break;
      default:
        break;
    }
  }

  private _esModoExistente(): boolean {
    return (
      this.props.displayMode === FormDisplayMode.Edit ||
      this.props.displayMode === FormDisplayMode.Display
    );
  }

  private _puedeAprobarTransferencia(): boolean {
    return (
      this._esModoExistente() &&
      !!this.state.pagoId &&
      this.state.medioPago === 'Transferencia' &&
      this.state.estado === 'Pendiente'
    );
  }

  private _mostrarTransferenciaAprobada(): boolean {
    return this._esModoExistente() && this.state.medioPago === 'Transferencia' && this.state.estado === 'Aprobado';
  }

  private _mostrarEstadoAprobadoGeneral(): boolean {
    return (
      this._esModoExistente() &&
      this.state.estado === 'Aprobado' &&
      this.state.medioPago !== 'Transferencia'
    );
  }

  private _onSolicitarAprobacion = (): void => {
    this.setState({ mostrarDialogoAprobacion: true, mensajeExito: '', error: '' });
  };

  private _onCancelarDialogoAprobacion = (): void => {
    if (this.state.aprobando) {
      return;
    }
    this.setState({ mostrarDialogoAprobacion: false });
  };

  private _approveTransfer = async (): Promise<void> => {
    const pagoId = this.state.pagoId;
    if (!pagoId) {
      return;
    }

    try {
      console.log('Aprobando transferencia...');
      this.setState({ aprobando: true, error: '', mensajeExito: '' });
      await this._service.approveTransfer(pagoId);
      console.log('Estado actualizado a Aprobado');
      console.log('Transferencia aprobada correctamente');
      this.setState({
        estado: 'Aprobado',
        aprobando: false,
        mostrarDialogoAprobacion: false,
        mensajeExito: 'Transferencia aprobada correctamente.'
      });
    } catch (error) {
      this.setState({
        aprobando: false,
        mostrarDialogoAprobacion: false,
        error: 'No se pudo aprobar la transferencia. Intenta nuevamente.'
      });
    }
  };

  private _getItemId(): number | null {
    const fromContext = this.props.context.itemId;
    if (fromContext !== undefined && fromContext > 0) {
      return fromContext;
    }
    const possibleId = (this.props.context as { itemId?: number }).itemId;
    const parsed = Number(possibleId);
    return parsed > 0 ? parsed : null;
  }

  private _esSoloLectura(): boolean {
    return this.props.displayMode === FormDisplayMode.Display;
  }

  private _getSubtitle(): string {
    switch (this.props.displayMode) {
      case FormDisplayMode.New:
        return 'Nuevo pago';
      case FormDisplayMode.Edit:
        return 'Editar pago';
      default:
        return 'Ver pago';
    }
  }

  private async _inicializarFormulario(): Promise<void> {
    try {
      console.log('Aplicando lógica de Fecha de Partida a FechaPago');
      console.log('Cargando opciones del campo Banco...');
      console.log('Cargando opciones del campo Motivo...');
      const [opcionesBanco, opcionesMotivo] = await Promise.all([
        this._service.getBancoChoices(),
        this._service.getMotivoChoices()
      ]);
      console.log('Opciones Banco cargadas correctamente');
      console.log('Opciones Motivo cargadas correctamente');

      const itemId = this._getItemId();
      const esEdicion =
        (this.props.displayMode === FormDisplayMode.Edit || this.props.displayMode === FormDisplayMode.Display) &&
        !!itemId;

      if (!esEdicion) {
        const motivo =
          this.state.tipoIngreso === 'asociado_viaje'
            ? this._resolverMotivoViaje(opcionesMotivo)
            : this.state.motivo;
        this.setState({
          opcionesBanco,
          opcionesMotivo,
          motivo,
          cargando: false
        });
        return;
      }

      console.log('Cargando pago existente...');
      const pago = await this._service.getPagoById(itemId as number);
      let viajeTitulo = pago.viajeTitulo || '';
      const viajeId = pago.viajeId && pago.viajeId > 0 ? pago.viajeId : null;
      const tipoIngreso: TipoIngresoPago = viajeId ? 'asociado_viaje' : 'sin_viaje';

      if (viajeId && !viajeTitulo) {
        const viaje = await this._service.getViajeById(viajeId);
        viajeTitulo = viaje.titulo;
      }

      let pasajeroId: number | null = pago.pasajeroId > 0 ? pago.pasajeroId : null;
      let pasajeroTitulo = '';
      let pasajeroNombre = '';
      let pasajeroDni = '';
      let pasajerosViaje: IPasajeroLookupItem[] = [];
      let serviciosViaje: IServicioViajeItem[] = [];
      let servicioAsociadoId: number | null = null;
      let concepto = '';
      let conceptoMoneda = '';
      let pagosViaje: IPagoSaldoItem[] = [];

      if (pasajeroId) {
        const pasajero = await this._service.getPasajeroById(pasajeroId);
        pasajeroTitulo = pasajero.nombreApellido;
        if (tipoIngreso === 'sin_viaje') {
          pasajeroNombre = pasajero.nombreApellido;
          pasajeroDni = pasajero.dni;
        }
      }

      if (tipoIngreso === 'asociado_viaje' && viajeId) {
        [pasajerosViaje, serviciosViaje, pagosViaje] = await Promise.all([
          this._service.getPasajerosByViaje(viajeId),
          this._service.getServiciosViajeByViaje(viajeId),
          this._service.getPagosSaldoByViaje(viajeId)
        ]);
        if (pasajeroId) {
          const pasajeroViaje = pasajerosViaje.filter((item: IPasajeroLookupItem) => item.id === pasajeroId)[0];
          if (pasajeroViaje) {
            pasajeroTitulo = pasajeroViaje.nombreApellido;
          }
        }
        const conceptoGuardado = (pago.concepto || '').trim();
        const servicioIdGuardado =
          pago.servicioAsociadoId && pago.servicioAsociadoId > 0 ? pago.servicioAsociadoId : null;
        if (servicioIdGuardado) {
          const servicio = serviciosViaje.filter(
            (item: IServicioViajeItem) => item.id === servicioIdGuardado
          )[0];
          if (servicio) {
            servicioAsociadoId = servicio.id;
            concepto = servicio.concepto;
            conceptoMoneda = normalizarMonedaServicio(servicio.moneda);
          } else if (conceptoGuardado) {
            concepto = conceptoGuardado;
          }
        } else if (conceptoGuardado) {
          const servicio = serviciosViaje.filter(
            (item: IServicioViajeItem) => item.concepto === conceptoGuardado
          )[0];
          if (servicio) {
            servicioAsociadoId = servicio.id;
            concepto = servicio.concepto;
            conceptoMoneda = normalizarMonedaServicio(servicio.moneda);
          } else {
            concepto = conceptoGuardado;
          }
        }
      }

      const fechaPago = getDateOnlyFromSharePoint(pago.fechaPago);
      if (fechaPago) {
        console.log('FechaPago parseada correctamente');
      }

      const moneda = pago.moneda;
      this._logEstadoCotizacion(moneda);
      this._logVisibilidadCuentaBancaria(pago.medioPago);

      const comprobantes = await this._service.getAttachments(pago.id);
      const motivo =
        tipoIngreso === 'asociado_viaje'
          ? this._resolverMotivoViaje(opcionesMotivo) || pago.motivo || ''
          : pago.motivo || '';

      this.setState({
        pagoId: pago.id,
        tipoPago: this._normalizarTipoPago(pago.tipoPago) || 'Ingreso',
        tipoIngreso,
        viajeId,
        viajeTitulo,
        viajeBusquedaTexto: viajeTitulo,
        pasajeroId,
        pasajeroTitulo,
        pasajerosViaje,
        pasajerosViajeCargando: false,
        pasajeroNombre,
        pasajeroDni,
        servicioAsociadoId,
        concepto,
        conceptoMoneda,
        serviciosViaje,
        serviciosViajeCargando: false,
        pagosViaje,
        medioPago: pago.medioPago,
        banco: pago.medioPago === 'Transferencia' ? (pago.banco || '') : '',
        motivo: motivo,
        opcionesBanco,
        opcionesMotivo,
        fechaPago,
        monto: pago.monto > 0 ? String(pago.monto) : '',
        moneda,
        cotizacion: pago.cotizacion !== undefined && pago.cotizacion > 0 ? String(pago.cotizacion) : '',
        estado: pago.estado,
        comprobantes,
        cargando: false,
        error: '',
        mensajeExito: '',
        comprobantesError: ''
      });
    } catch (error) {
      this.setState({
        cargando: false,
        error: 'No se pudieron cargar los datos del formulario desde SharePoint.'
      });
    }
  }

  private _onCambiarTipoPago = (tipoPago: TipoPago): void => {
    if (this._esSoloLectura() || this.state.tipoPago === tipoPago) {
      return;
    }
    this.setState({
      tipoPago,
      fieldErrors: { ...this.state.fieldErrors, tipoPago: '' }
    });
  };

  private _onCambiarTipoIngreso = (tipoIngreso: TipoIngresoPago): void => {
    if (this._esSoloLectura() || this.state.tipoIngreso === tipoIngreso) {
      return;
    }

    if (tipoIngreso === 'sin_viaje') {
      this.setState({
        tipoIngreso,
        viajeId: null,
        viajeTitulo: '',
        viajeBusquedaTexto: '',
        viajesResultados: [],
        pasajeroId: null,
        pasajeroTitulo: '',
        pasajerosViaje: [],
        pasajerosViajeCargando: false,
        pasajeroNombre: '',
        pasajeroDni: '',
        motivo: '',
        pagosViaje: [],
        ...this._limpiarConceptoSeleccionado(),
        fieldErrors: {
          ...this.state.fieldErrors,
          viaje: '',
          pasajero: '',
          concepto: ''
        }
      });
      return;
    }

    this.setState({
      tipoIngreso,
      motivo: this._resolverMotivoViaje(this.state.opcionesMotivo),
      pasajeroNombre: '',
      pasajeroDni: '',
      pasajeroId: null,
      pasajeroTitulo: '',
      pasajerosViaje: [],
      ...this._limpiarConceptoSeleccionado(),
      fieldErrors: {
        ...this.state.fieldErrors,
        pasajeroNombre: '',
        pasajeroDni: '',
        concepto: ''
      }
    });
  };

  private _onCambiarBusquedaViaje = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura()) {
      return;
    }

    const texto = event.target.value;
    this.setState({
      viajeBusquedaTexto: texto,
      viajeId: null,
      viajeTitulo: '',
      pasajeroId: null,
      pasajeroTitulo: '',
      pasajerosViaje: [],
      pagosViaje: [],
      ...this._limpiarConceptoSeleccionado(),
      fieldErrors: { ...this.state.fieldErrors, viaje: '', pasajero: '', concepto: '' }
    });

    if (this._viajeSearchTimer) {
      window.clearTimeout(this._viajeSearchTimer);
    }

    if (!texto.trim()) {
      this.setState({ viajesResultados: [], viajeBuscando: false });
      return;
    }

    this.setState({ viajeBuscando: true });
    this._viajeSearchTimer = window.setTimeout(() => {
      void this._buscarViajes(texto);
    }, 300);
  };

  private async _buscarViajes(texto: string): Promise<void> {
    try {
      const resultados = await this._service.searchViajesByTitle(texto);
      this.setState({ viajesResultados: resultados, viajeBuscando: false });
    } catch (error) {
      this.setState({ viajesResultados: [], viajeBuscando: false });
    }
  }

  private _onSeleccionarViaje = (viaje: IViajeLookupItem): void => {
    this.setState({
      viajeId: viaje.id,
      viajeTitulo: viaje.titulo,
      viajeBusquedaTexto: viaje.titulo,
      viajesResultados: [],
      pasajeroId: null,
      pasajeroTitulo: '',
      ...this._limpiarConceptoSeleccionado(),
      fieldErrors: { ...this.state.fieldErrors, viaje: '', pasajero: '', concepto: '' }
    });
    void this._cargarPasajerosViaje(viaje.id);
    void this._cargarServiciosViaje(viaje.id);
    void this._cargarPagosViaje(viaje.id);
  };

  private _onCambiarServicioConcepto = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const servicioId = Number(event.target.value) || null;
    const servicio = this.state.serviciosViaje.filter((item: IServicioViajeItem) => item.id === servicioId)[0];
    const conceptoMoneda = servicio ? normalizarMonedaServicio(servicio.moneda) : '';
    const requiereCotizacion = servicio
      ? requiereCotizacionPago(conceptoMoneda, this.state.moneda)
      : false;
    this.setState({
      servicioAsociadoId: servicioId,
      concepto: servicio ? servicio.concepto : '',
      conceptoMoneda,
      cotizacion: requiereCotizacion ? this.state.cotizacion : '',
      fieldErrors: {
        ...this.state.fieldErrors,
        concepto: '',
        cotizacion: requiereCotizacion ? this.state.fieldErrors.cotizacion : '',
        monto: ''
      }
    });
  };

  private _onCambiarPasajeroViaje = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const pasajeroId = Number(event.target.value) || null;
    const pasajero = this.state.pasajerosViaje.filter((item: IPasajeroLookupItem) => item.id === pasajeroId)[0];
    this.setState({
      pasajeroId,
      pasajeroTitulo: pasajero ? pasajero.nombreApellido : '',
      fieldErrors: { ...this.state.fieldErrors, pasajero: '' }
    });
  };

  private _onCambiarPasajeroNombre = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({
      pasajeroNombre: event.target.value,
      fieldErrors: { ...this.state.fieldErrors, pasajeroNombre: '' }
    });
  };

  private _onCambiarPasajeroDni = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({
      pasajeroDni: event.target.value,
      fieldErrors: { ...this.state.fieldErrors, pasajeroDni: '' }
    });
  };

  private _onCambiarMedioPago = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const medioPago = event.target.value;
    this._logVisibilidadCuentaBancaria(medioPago);
    this.setState({
      medioPago,
      banco: medioPago === 'Transferencia' ? this.state.banco : '',
      fieldErrors: {
        ...this.state.fieldErrors,
        medioPago: '',
        banco: medioPago === 'Transferencia' ? this.state.fieldErrors.banco : ''
      }
    });
  };

  private _onCambiarBanco = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({
      banco: event.target.value,
      fieldErrors: { ...this.state.fieldErrors, banco: '' }
    });
  };

  private _onCambiarMotivo = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ motivo: event.target.value });
  };

  private _onCambiarFechaPago = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({
      fechaPago: getDateOnlyFromSharePoint(event.target.value),
      fieldErrors: { ...this.state.fieldErrors, fechaPago: '' }
    });
  };

  private _onCambiarMonto = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({
      monto: event.target.value,
      fieldErrors: { ...this.state.fieldErrors, monto: '' }
    });
  };

  private _onCambiarMoneda = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const moneda = event.target.value;
    const requiereCotizacion = requiereCotizacionPago(this.state.conceptoMoneda, moneda);
    this._logEstadoCotizacion(moneda);
    this.setState({
      moneda,
      cotizacion: requiereCotizacion ? this.state.cotizacion : '',
      fieldErrors: {
        ...this.state.fieldErrors,
        moneda: '',
        cotizacion: requiereCotizacion ? this.state.fieldErrors.cotizacion : '',
        monto: ''
      }
    });
  };

  private _onCambiarCotizacion = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({
      cotizacion: event.target.value,
      fieldErrors: { ...this.state.fieldErrors, cotizacion: '', monto: '' }
    });
  };

  private _formatFileSize(bytes: number | undefined): string {
    if (!bytes || bytes <= 0) {
      return '';
    }
    if (bytes < 1024) {
      return bytes + ' B';
    }
    if (bytes < 1024 * 1024) {
      return (bytes / 1024).toFixed(1) + ' KB';
    }
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  private _sanitizeNombreArchivo(fileName: string): string {
    const trimmed = (fileName || '').trim();
    const dot = trimmed.lastIndexOf('.');
    let base = dot >= 0 ? trimmed.substring(0, dot) : trimmed;
    const ext = dot >= 0 ? trimmed.substring(dot) : '';
    base = base.replace(/[#%?&'"<>\\\/:*|]/g, '-');
    base = base.replace(/-+/g, '-').replace(/^[\s-]+|[\s-]+$/g, '');
    if (!base) {
      base = 'documento';
    }
    return base + ext;
  }

  private _abrirComprobante = (serverRelativeUrl: string): void => {
    if (!serverRelativeUrl) {
      return;
    }
    if (serverRelativeUrl.indexOf('http') === 0) {
      window.open(serverRelativeUrl, '_blank');
      return;
    }
    const origin = new URL(this.props.context.pageContext.web.absoluteUrl).origin;
    window.open(origin + serverRelativeUrl, '_blank');
  };

  private async _refrescarComprobantes(itemId: number): Promise<void> {
    const comprobantes = await this._service.getAttachments(itemId);
    this.setState({ comprobantes });
  }

  private _abrirEditorComprobante = (): void => {
    if (this._esSoloLectura()) {
      return;
    }
    this.setState({ mostrarEditorComprobante: true, comprobantesError: '' });
  };

  private _cerrarEditorComprobante = (): void => {
    this.setState({ mostrarEditorComprobante: false });
  };

  private _onSeleccionarComprobantes = (event: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura()) {
      return;
    }

    const archivos =
      event.target.files && event.target.files.length > 0
        ? (Array.prototype.slice.call(event.target.files) as File[])
        : [];
    event.target.value = '';

    if (archivos.length === 0) {
      return;
    }

    const pagoId = this.state.pagoId;
    if (pagoId) {
      void this._subirComprobantesExistentes(pagoId, archivos);
      return;
    }

    this._agregarComprobantesPendientes(archivos);
  };

  private _agregarComprobantesPendientes(archivos: File[]): void {
    const nombresActuales = this.state.comprobantesPendientes.map((file: File) =>
      this._sanitizeNombreArchivo(file.name).toLowerCase()
    );
    const nuevos: File[] = [];
    const errores: string[] = [];

    archivos.forEach((archivo: File) => {
      if (!(archivo.name || '').trim()) {
        errores.push('Uno de los archivos no tiene nombre válido.');
        return;
      }
      const nombreSanitizado = this._sanitizeNombreArchivo(archivo.name);
      if (nombresActuales.indexOf(nombreSanitizado.toLowerCase()) >= 0) {
        errores.push("Ya existe un archivo con el nombre '" + nombreSanitizado + "'.");
        return;
      }
      nombresActuales.push(nombreSanitizado.toLowerCase());
      nuevos.push(archivo);
    });

    if (nuevos.length > 0) {
      this.setState((prev) => ({
        comprobantesPendientes: prev.comprobantesPendientes.concat(nuevos),
        mostrarEditorComprobante: false
      }));
    }

    if (errores.length > 0) {
      this.setState({ comprobantesError: errores.join(' ') });
    } else {
      this.setState({ comprobantesError: '' });
    }
  };

  private _eliminarComprobantePendiente = (index: number): void => {
    this.setState((prev) => ({
      comprobantesPendientes: prev.comprobantesPendientes.filter((_: File, i: number) => i !== index),
      comprobantesError: ''
    }));
  };

  private _subirComprobantesExistentes = async (itemId: number, archivos: File[]): Promise<void> => {
    this.setState({ comprobantesSubiendo: true, guardando: true, comprobantesError: '' });
    const fallidos: string[] = [];
    let subidos = 0;

    try {
      for (const archivo of archivos) {
        try {
          await this._service.uploadAttachment(itemId, archivo);
          subidos++;
        } catch (error) {
          const detalle = error instanceof Error && error.message ? error.message : '';
          fallidos.push(detalle || archivo.name);
        }
      }

      await this._refrescarComprobantes(itemId);

      if (fallidos.length === 0) {
        this.setState({ mostrarEditorComprobante: false, comprobantesError: '' });
        return;
      }

      const mensaje =
        subidos > 0
          ? 'Se adjuntaron ' + subidos + ' de ' + archivos.length + ' comprobante(s). Errores: ' + fallidos.join(' | ')
          : 'No se pudieron adjuntar los comprobantes: ' + fallidos.join(' | ');
      this.setState({ comprobantesError: mensaje });
    } finally {
      this.setState({ comprobantesSubiendo: false, guardando: false });
    }
  };

  private _subirComprobantesPendientes = async (itemId: number): Promise<boolean> => {
    const pendientes = this.state.comprobantesPendientes;
    if (pendientes.length === 0) {
      return true;
    }

    const fallidos: string[] = [];
    const pendientesRestantes: File[] = [];

    for (const archivo of pendientes) {
      try {
        await this._service.uploadAttachment(itemId, archivo);
      } catch (error) {
        const detalle = error instanceof Error && error.message ? error.message : '';
        fallidos.push(detalle || archivo.name);
        pendientesRestantes.push(archivo);
      }
    }

    await this._refrescarComprobantes(itemId);

    if (fallidos.length > 0) {
      this.setState({
        comprobantesPendientes: pendientesRestantes,
        comprobantesError:
          'El pago se guardó, pero no se pudieron subir todos los comprobantes: ' + fallidos.join(' | ')
      });
      return false;
    }

    this.setState({ comprobantesPendientes: [], comprobantesError: '' });
    return true;
  };

  private _eliminarComprobante = async (fileName: string): Promise<void> => {
    if (this._esSoloLectura() || !this.state.pagoId) {
      return;
    }

    const confirmar = window.confirm('¿Desea eliminar el comprobante "' + fileName + '"?');
    if (!confirmar) {
      return;
    }

    const itemId = this.state.pagoId;
    try {
      this.setState({ guardando: true, comprobantesError: '' });
      await this._service.deleteAttachment(itemId, fileName);
      await this._refrescarComprobantes(itemId);
    } catch (error) {
      this.setState({ comprobantesError: 'No se pudo eliminar el comprobante.' });
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _validarFormulario(): boolean {
    const fieldErrors = emptyFieldErrors();
    let valido = true;

    if (!this._normalizarTipoPago(this.state.tipoPago)) {
      fieldErrors.tipoPago = 'El tipo de movimiento es obligatorio.';
      valido = false;
    }

    if (this._esAsociadoAViaje()) {
      if (!this.state.viajeId || this.state.viajeId <= 0) {
        fieldErrors.viaje = 'El viaje es obligatorio.';
        valido = false;
      }
      if (!this.state.pasajeroId || this.state.pasajeroId <= 0) {
        fieldErrors.pasajero = 'El pasajero es obligatorio.';
        valido = false;
      }
      if (!this.state.servicioAsociadoId || !(this.state.concepto || '').trim()) {
        fieldErrors.concepto = 'El concepto es obligatorio.';
        valido = false;
      }
    } else {
      if (!(this.state.pasajeroNombre || '').trim()) {
        fieldErrors.pasajeroNombre = 'El nombre es obligatorio.';
        valido = false;
      }
      if (!(this.state.pasajeroDni || '').trim()) {
        fieldErrors.pasajeroDni = 'El DNI es obligatorio.';
        valido = false;
      }
    }

    if (!(this.state.medioPago || '').trim()) {
      fieldErrors.medioPago = 'El medio de pago es obligatorio.';
      valido = false;
    }

    if (this._requiereCuentaBancaria()) {
      if (!(this.state.banco || '').trim()) {
        fieldErrors.banco = 'La cuenta bancaria es obligatoria cuando el medio de pago es Transferencia.';
        valido = false;
      }
    }

    if (!(this.state.fechaPago || '').trim()) {
      fieldErrors.fechaPago = 'La fecha es obligatoria.';
      valido = false;
    }

    const montoNumerico = Number(this.state.monto);
    if (!this.state.monto.trim()) {
      fieldErrors.monto = 'El monto es obligatorio.';
      valido = false;
    } else if (isNaN(montoNumerico) || montoNumerico <= 0) {
      fieldErrors.monto = 'El monto debe ser mayor a 0.';
      valido = false;
    } else if (
      this.state.tipoPago === 'Ingreso' &&
      this._esAsociadoAViaje() &&
      this._getServicioSeleccionado()
    ) {
      const servicio = this._getServicioSeleccionado() as IServicioViajeItem;
      const pagoIdExcluir =
        this.state.pagoId && this.state.pagoId > 0 ? this.state.pagoId : undefined;
      const cotizacionValidacion = this._requiereCotizacion()
        ? Number(this.state.cotizacion)
        : undefined;
      if (
        montoExcedeSaldoPendiente(
          servicio,
          this.state.pagosViaje,
          montoNumerico,
          this.state.moneda,
          cotizacionValidacion,
          pagoIdExcluir
        )
      ) {
        fieldErrors.monto = 'El monto ingresado supera el saldo pendiente del servicio.';
        valido = false;
      }
    }

    if (!(this.state.moneda || '').trim()) {
      fieldErrors.moneda = 'La moneda es obligatoria.';
      valido = false;
    }

    if (this._requiereCotizacion()) {
      const cotizacionNumerica = Number(this.state.cotizacion);
      if (!this.state.cotizacion.trim()) {
        fieldErrors.cotizacion =
          'La cotización es obligatoria cuando la moneda del servicio difiere de la moneda del pago.';
        valido = false;
      } else if (isNaN(cotizacionNumerica) || cotizacionNumerica <= 0) {
        fieldErrors.cotizacion = 'La cotización debe ser mayor a 0.';
        valido = false;
      }
    }

    this.setState({ fieldErrors });
    return valido;
  }

  private _resolverServicioAsociadoIdParaGuardar(): number | undefined {
    const esEdicion = !!(this.state.pagoId && this.state.pagoId > 0);
    const esIngreso = this._normalizarTipoPago(this.state.tipoPago) === 'Ingreso';

    if (esIngreso && this._esAsociadoAViaje()) {
      const servicioId = this.state.servicioAsociadoId;
      return servicioId && servicioId > 0 ? servicioId : esEdicion ? 0 : undefined;
    }

    return esEdicion ? 0 : undefined;
  }

  private _onGuardar = async (): Promise<void> => {
    if (this._esSoloLectura()) {
      return;
    }

    if (!this._validarFormulario()) {
      return;
    }

    try {
      console.log('Guardando Registro de Pago...');
      this.setState({ guardando: true, error: '' });

      const estado = getEstadoByMedioPago(this.state.medioPago);
      this._logEstadoPorMedioPago(this.state.medioPago);

      const banco = this._requiereCuentaBancaria() ? this.state.banco : '';
      console.log('Guardando Banco: ' + (banco || '(vacío)'));

      let pasajeroId = 0;
      let pasajeroNombre = '';
      if (this._esAsociadoAViaje()) {
        pasajeroId = this.state.pasajeroId as number;
        const pasajero = this.state.pasajerosViaje.filter(
          (item: IPasajeroLookupItem) => item.id === pasajeroId
        )[0];
        pasajeroNombre = pasajero ? pasajero.nombreApellido : this.state.pasajeroTitulo;
      } else {
        const pasajero = await this._service.findOrCreatePasajero(
          this.state.pasajeroNombre,
          this.state.pasajeroDni
        );
        pasajeroId = pasajero.id;
        pasajeroNombre = pasajero.nombreApellido;
        console.log('Pasajero resuelto para pago sin viaje: ' + pasajeroId);
      }

      const data = {
        tipoPago: this._normalizarTipoPago(this.state.tipoPago) || 'Ingreso',
        viajeId: this._esAsociadoAViaje() ? (this.state.viajeId as number) : null,
        viajeTitulo: this._esAsociadoAViaje() ? this.state.viajeTitulo : '',
        pasajeroId,
        pasajeroNombre,
        concepto: this._esAsociadoAViaje() ? this.state.concepto : undefined,
        medioPago: this.state.medioPago,
        fechaPago: this.state.fechaPago,
        monto: Number(this.state.monto),
        moneda: this.state.moneda,
        estado,
        banco,
        motivo: this.state.motivo,
        cotizacion: this._requiereCotizacion() ? Number(this.state.cotizacion) || 0 : undefined,
        servicioAsociadoId: this._resolverServicioAsociadoIdParaGuardar()
      };

      let itemId = this.state.pagoId;
      if (itemId) {
        await this._service.updatePago(itemId, data);
      } else {
        const creado = await this._service.createPago(data);
        itemId = creado.id;
        this.setState({ pagoId: creado.id });
      }

      const comprobantesOk = await this._subirComprobantesPendientes(itemId);
      if (!comprobantesOk) {
        return;
      }

      console.log('Registro de Pago guardado correctamente');
      this.props.onSave();
    } catch (error) {
      this.setState({
        error: 'No se pudo guardar el registro de pago en SharePoint. Verifica los datos ingresados.'
      });
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _onCancelar = (): void => {
    this.props.onClose();
  };

  private _navegarAFormularioEdicion = (): void => {
    const itemId = this._getItemId();
    const listGuid = this.props.context.list?.guid?.toString();
    if (!itemId || !listGuid) {
      return;
    }
    const webUrl = this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const url =
      webUrl +
      '/_layouts/15/listform.aspx?PageType=6&ListId=' +
      encodeURIComponent(listGuid) +
      '&ID=' +
      itemId;
    window.location.href = url;
  };

  private _renderFieldError(message: string): React.ReactNode {
    if (!message) {
      return null;
    }
    return <div style={layoutStyles.fieldError}>{message}</div>;
  }

  private _renderSeccionAprobacion(): React.ReactNode {
    if (this.props.displayMode === FormDisplayMode.New) {
      return null;
    }

    if (this._puedeAprobarTransferencia()) {
      return (
        <div style={layoutStyles.approvalCardPending}>
          <div>
            <div style={layoutStyles.approvalCardTitle}>🟡 Transferencia pendiente de aprobación</div>
            <div style={layoutStyles.approvalCardText}>Esta transferencia todavía no fue confirmada.</div>
          </div>
          <div style={layoutStyles.approvalCardActions}>
            <button
              type="button"
              style={
                this.state.aprobando || this.state.guardando
                  ? { ...layoutStyles.primaryButton, ...layoutStyles.buttonDisabled }
                  : layoutStyles.primaryButton
              }
              onClick={this._onSolicitarAprobacion}
              disabled={this.state.aprobando || this.state.guardando}
            >
              Aprobar
            </button>
          </div>
        </div>
      );
    }

    if (this._mostrarTransferenciaAprobada()) {
      return (
        <div style={layoutStyles.approvalCardApproved}>
          <div style={{ ...layoutStyles.approvalCardTitle, marginBottom: 0 }}>🟢 Transferencia aprobada</div>
        </div>
      );
    }

    if (this._mostrarEstadoAprobadoGeneral()) {
      return (
        <div style={layoutStyles.approvalCardApproved}>
          <div style={{ ...layoutStyles.approvalCardTitle, marginBottom: 0 }}>🟢 Estado: Aprobado</div>
        </div>
      );
    }

    return null;
  }

  private _renderSeccionComprobantes(soloLectura: boolean): React.ReactNode {
    const listaVacia =
      this.state.comprobantes.length === 0 && this.state.comprobantesPendientes.length === 0;

    return (
      <div style={layoutStyles.section}>
        <div style={layoutStyles.sectionToolbar}>
          <div style={layoutStyles.toolbarTitle}>Comprobantes</div>
          {!soloLectura && (
            <button
              type="button"
              style={
                this.state.guardando || this.state.comprobantesSubiendo
                  ? { ...layoutStyles.defaultButton, ...layoutStyles.buttonDisabled }
                  : layoutStyles.defaultButton
              }
              onClick={this._abrirEditorComprobante}
              disabled={this.state.guardando || this.state.comprobantesSubiendo}
            >
              Subir comprobantes
            </button>
          )}
        </div>

        {this.state.comprobantesError && <div style={layoutStyles.error}>{this.state.comprobantesError}</div>}

        {!soloLectura && !this.state.pagoId && this.state.comprobantesPendientes.length > 0 && (
          <div style={layoutStyles.info}>
            Los comprobantes seleccionados se adjuntarán al guardar el registro de pago.
          </div>
        )}

        <div style={layoutStyles.attachmentZone}>
          {this.state.comprobantesSubiendo && (
            <div style={layoutStyles.uploadingRow}>Adjuntando comprobantes...</div>
          )}

          {!this.state.comprobantesSubiendo && listaVacia && (
            <div style={layoutStyles.attachmentEmpty}>No hay comprobantes adjuntos.</div>
          )}

          {this.state.comprobantes.map((comprobante: IComprobanteItem, index: number) => {
            const esUltimo =
              index === this.state.comprobantes.length - 1 && this.state.comprobantesPendientes.length === 0;
            return (
              <div
                key={comprobante.fileName}
                style={
                  esUltimo
                    ? { ...layoutStyles.attachmentRow, ...layoutStyles.attachmentRowLast }
                    : layoutStyles.attachmentRow
                }
              >
                <div style={layoutStyles.attachmentMeta}>
                  <div style={layoutStyles.attachmentFileName}>{comprobante.fileName}</div>
                  {comprobante.fileSize ? (
                    <div style={layoutStyles.attachmentFileSize}>
                      {this._formatFileSize(comprobante.fileSize)}
                    </div>
                  ) : null}
                </div>
                <div style={layoutStyles.attachmentActions}>
                  <button
                    type="button"
                    style={layoutStyles.attachmentActionButton}
                    onClick={() => this._abrirComprobante(comprobante.serverRelativeUrl)}
                    disabled={!comprobante.serverRelativeUrl || this.state.guardando}
                  >
                    Abrir
                  </button>
                  {!soloLectura && (
                    <button
                      type="button"
                      style={layoutStyles.attachmentActionButton}
                      onClick={() => {
                        void this._eliminarComprobante(comprobante.fileName);
                      }}
                      disabled={this.state.guardando || this.state.comprobantesSubiendo}
                    >
                      Eliminar
                    </button>
                  )}
                </div>
              </div>
            );
          })}

          {!soloLectura &&
            this.state.comprobantesPendientes.map((archivo: File, index: number) => {
              const esUltimo = index === this.state.comprobantesPendientes.length - 1;
              return (
                <div
                  key={archivo.name + '-' + archivo.size + '-' + index}
                  style={
                    esUltimo
                      ? { ...layoutStyles.attachmentRow, ...layoutStyles.attachmentRowLast }
                      : layoutStyles.attachmentRow
                  }
                >
                  <div style={layoutStyles.attachmentMeta}>
                    <div style={layoutStyles.attachmentFileName}>{archivo.name}</div>
                    <div style={layoutStyles.attachmentFileSize}>
                      {this._formatFileSize(archivo.size)} · Pendiente de guardar
                    </div>
                  </div>
                  <div style={layoutStyles.attachmentActions}>
                    <button
                      type="button"
                      style={layoutStyles.attachmentActionButton}
                      onClick={() => this._eliminarComprobantePendiente(index)}
                      disabled={this.state.guardando || this.state.comprobantesSubiendo}
                    >
                      Eliminar
                    </button>
                  </div>
                </div>
              );
            })}
        </div>

        {!soloLectura && this.state.mostrarEditorComprobante && (
          <div style={{ marginTop: 12 }}>
            <FileInputEspanol
              inputId="comprobantes-archivos"
              multiple={true}
              disabled={this.state.guardando || this.state.comprobantesSubiendo}
              buttonLabel="Seleccionar archivos"
              onChange={this._onSeleccionarComprobantes}
            />
            <div style={{ marginTop: 8 }}>
              <button
                type="button"
                style={layoutStyles.defaultButton}
                onClick={this._cerrarEditorComprobante}
                disabled={this.state.guardando || this.state.comprobantesSubiendo}
              >
                Cerrar
              </button>
            </div>
          </div>
        )}
      </div>
    );
  }

  private _renderSelectorTipoPago(soloLectura: boolean): React.ReactNode {
    const tipoPago = this._normalizarTipoPago(this.state.tipoPago);

    return (
      <div style={layoutStyles.section}>
        <div style={layoutStyles.sectionHeader}>Tipo de movimiento</div>
        {soloLectura ? (
          <div>
            <span
              style={{
                ...layoutStyles.tipoPagoDisplayBadge,
                ...(tipoPago === 'Egreso'
                  ? layoutStyles.tipoPagoDisplayEgreso
                  : layoutStyles.tipoPagoDisplayIngreso)
              }}
            >
              {tipoPago || '—'}
            </span>
          </div>
        ) : (
          <>
            <div style={layoutStyles.tipoPagoGrid}>
              {TIPOS_PAGO.map((opcion: TipoPago) => {
                const activo = this.state.tipoPago === opcion;
                const esIngreso = opcion === 'Ingreso';
                return (
                  <button
                    key={opcion}
                    type="button"
                    style={{
                      ...layoutStyles.tipoPagoCard,
                      ...(activo
                        ? esIngreso
                          ? layoutStyles.tipoPagoCardIngresoActive
                          : layoutStyles.tipoPagoCardEgresoActive
                        : {}),
                      ...(this.state.guardando ? layoutStyles.tipoPagoCardDisabled : {})
                    }}
                    onClick={() => this._onCambiarTipoPago(opcion)}
                    disabled={this.state.guardando}
                    aria-pressed={activo}
                  >
                    <div style={layoutStyles.tipoPagoCardTitle}>{opcion}</div>
                    <div style={layoutStyles.tipoPagoCardHint}>
                      {esIngreso
                        ? 'Registrar un ingreso de dinero'
                        : 'Registrar un egreso de dinero'}
                    </div>
                  </button>
                );
              })}
            </div>
            {this._renderFieldError(this.state.fieldErrors.tipoPago)}
          </>
        )}
      </div>
    );
  }

  private _renderDialogoAprobacion(): React.ReactNode {
    if (!this.state.mostrarDialogoAprobacion) {
      return null;
    }

    return (
      <div style={layoutStyles.dialogOverlay} role="presentation" onClick={this._onCancelarDialogoAprobacion}>
        <div
          style={layoutStyles.dialogBox}
          role="dialog"
          aria-modal="true"
          aria-labelledby="dialogo-aprobacion-titulo"
          onClick={(event: React.MouseEvent<HTMLDivElement>) => event.stopPropagation()}
        >
          <div id="dialogo-aprobacion-titulo" style={layoutStyles.dialogTitle}>
            Confirmar aprobación
          </div>
          <div style={layoutStyles.dialogText}>
            Aprobar este ingreso marcará la transferencia como aprobada. ¿Desea continuar?
          </div>
          <div style={layoutStyles.dialogActions}>
            <button
              type="button"
              style={
                this.state.aprobando
                  ? { ...layoutStyles.defaultButton, ...layoutStyles.buttonDisabled }
                  : layoutStyles.defaultButton
              }
              onClick={this._onCancelarDialogoAprobacion}
              disabled={this.state.aprobando}
            >
              Cancelar
            </button>
            <button
              type="button"
              style={
                this.state.aprobando
                  ? { ...layoutStyles.primaryButton, ...layoutStyles.buttonDisabled }
                  : layoutStyles.primaryButton
              }
              onClick={() => {
                void this._approveTransfer();
              }}
              disabled={this.state.aprobando}
            >
              {this.state.aprobando ? 'Aprobando...' : 'Aprobar'}
            </button>
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IRegistroPagosFormProps> {
    const soloLectura = this._esSoloLectura();
    const requiereCotizacion = this._requiereCotizacion();
    const requiereCuentaBancaria = this._requiereCuentaBancaria();

    if (this.state.cargando) {
      return (
        <div className={styles.registroPagosForm} style={layoutStyles.page}>
          Cargando datos...
        </div>
      );
    }

    return (
      <div className={styles.registroPagosForm} style={layoutStyles.page}>
        {this._renderDialogoAprobacion()}
        <div style={layoutStyles.container}>
          <div style={layoutStyles.header}>
            <div style={layoutStyles.title}>Registro de Pago</div>
            <div style={layoutStyles.subtitle}>{this._getSubtitle()}</div>
          </div>

          {this._renderSeccionAprobacion()}

          {this.state.mensajeExito && <div style={layoutStyles.success}>{this.state.mensajeExito}</div>}
          {this.state.error && <div style={layoutStyles.error}>{this.state.error}</div>}

          {this._renderSelectorTipoPago(soloLectura)}

          <div style={layoutStyles.section}>
            <div style={layoutStyles.sectionHeader}>Datos del pago</div>

            <div style={layoutStyles.formGrid}>
              <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.fieldFullWidth }}>
                <label style={layoutStyles.label}>Tipo de ingreso</label>
                {soloLectura ? (
                  <div style={layoutStyles.readOnlyValue}>
                    {this._getEtiquetaTipoIngreso(this.state.tipoIngreso)}
                  </div>
                ) : (
                  <div style={layoutStyles.choiceGroup}>
                    {TIPOS_INGRESO.map((opcion) => (
                      <label
                        key={opcion.key}
                        style={{
                          ...layoutStyles.choiceOption,
                          ...(this.state.guardando ? layoutStyles.choiceOptionDisabled : {})
                        }}
                      >
                        <input
                          type="radio"
                          name="tipoIngreso"
                          value={opcion.key}
                          checked={this.state.tipoIngreso === opcion.key}
                          onChange={() => this._onCambiarTipoIngreso(opcion.key)}
                          disabled={this.state.guardando}
                        />
                        {opcion.label}
                      </label>
                    ))}
                  </div>
                )}
              </div>

              {this._esAsociadoAViaje() && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Viaje</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>{this.state.viajeTitulo || '—'}</div>
                  ) : (
                    <div style={{ position: 'relative' }}>
                      <input
                        type="text"
                        style={{
                          ...layoutStyles.fieldInput,
                          ...(this.state.fieldErrors.viaje ? layoutStyles.fieldInputError : {})
                        }}
                        value={this.state.viajeBusquedaTexto}
                        onChange={this._onCambiarBusquedaViaje}
                        placeholder="Buscar viaje por título..."
                        disabled={this.state.guardando}
                      />
                      {this.state.viajeBuscando && (
                        <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>Buscando viajes...</div>
                      )}
                      {!this.state.viajeBuscando &&
                        this.state.viajeBusquedaTexto &&
                        this.state.viajesResultados.length > 0 && (
                          <div style={layoutStyles.autocompleteList}>
                            {this.state.viajesResultados.map((viaje: IViajeLookupItem) => (
                              <div
                                key={viaje.id}
                                style={layoutStyles.autocompleteItem}
                                onClick={() => this._onSeleccionarViaje(viaje)}
                              >
                                {viaje.titulo}
                              </div>
                            ))}
                          </div>
                        )}
                      {this.state.viajeId && this.state.viajeTitulo && (
                        <div style={{ fontSize: 12, color: '#107c10', marginTop: 4 }}>
                          Viaje seleccionado: {this.state.viajeTitulo}
                        </div>
                      )}
                    </div>
                  )}
                  {this._renderFieldError(this.state.fieldErrors.viaje)}
                </div>
              )}

              {this._esAsociadoAViaje() && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Pasajero</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>{this._getPasajeroSeleccionadoTitulo() || '—'}</div>
                  ) : (
                    <>
                      <select
                        style={{
                          ...layoutStyles.select,
                          ...(this.state.fieldErrors.pasajero ? layoutStyles.fieldInputError : {})
                        }}
                        value={this.state.pasajeroId || ''}
                        onChange={this._onCambiarPasajeroViaje}
                        disabled={
                          this.state.guardando ||
                          !this.state.viajeId ||
                          this.state.pasajerosViajeCargando
                        }
                      >
                        <option value="">
                          {this.state.pasajerosViajeCargando
                            ? 'Cargando pasajeros...'
                            : this.state.viajeId
                              ? 'Seleccione...'
                              : 'Seleccione un viaje primero'}
                        </option>
                        {this.state.pasajerosViaje.map((pasajero: IPasajeroLookupItem) => (
                          <option key={pasajero.id} value={pasajero.id}>
                            {pasajero.nombreApellido}
                            {pasajero.dni ? ' — DNI ' + pasajero.dni : ''}
                          </option>
                        ))}
                      </select>
                      {this.state.viajeId &&
                        !this.state.pasajerosViajeCargando &&
                        this.state.pasajerosViaje.length === 0 && (
                          <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>
                            El viaje seleccionado no tiene pasajeros asociados.
                          </div>
                        )}
                    </>
                  )}
                  {this._renderFieldError(this.state.fieldErrors.pasajero)}
                </div>
              )}

              {this._esAsociadoAViaje() && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Servicio a abonar</label>
                  {soloLectura ? (
                    <>
                      <div style={layoutStyles.readOnlyValue}>{this.state.concepto || '—'}</div>
                      {this._renderSaldoPendienteServicio()}
                    </>
                  ) : (
                    <>
                      <select
                        style={{
                          ...layoutStyles.select,
                          ...(this.state.fieldErrors.concepto ? layoutStyles.fieldInputError : {})
                        }}
                        value={this.state.servicioAsociadoId || ''}
                        onChange={this._onCambiarServicioConcepto}
                        disabled={
                          this.state.guardando ||
                          !this.state.viajeId ||
                          this.state.serviciosViajeCargando
                        }
                      >
                        <option value="">
                          {this.state.serviciosViajeCargando
                            ? 'Cargando servicios...'
                            : this.state.viajeId
                              ? 'Seleccione...'
                              : 'Seleccione un viaje primero'}
                        </option>
                        {this.state.serviciosViaje.map((servicio: IServicioViajeItem) => (
                          <option key={servicio.id} value={servicio.id}>
                            {servicio.concepto}
                            {servicio.moneda ? ' — ' + servicio.moneda : ''}
                          </option>
                        ))}
                      </select>
                      {this.state.serviciosViajeCargando && (
                        <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>
                          Cargando servicios del viaje...
                        </div>
                      )}
                      {this.state.viajeId &&
                        !this.state.serviciosViajeCargando &&
                        this.state.serviciosViaje.length === 0 && (
                          <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>
                            No hay servicios disponibles para este viaje.
                          </div>
                        )}
                      {this._renderSaldoPendienteServicio()}
                    </>
                  )}
                  {this._renderFieldError(this.state.fieldErrors.concepto)}
                </div>
              )}

              {!this._esAsociadoAViaje() && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>{soloLectura ? 'Pasajero' : 'Nombre'}</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>{this._getPasajeroSeleccionadoTitulo() || '—'}</div>
                  ) : (
                    <input
                      type="text"
                      style={{
                        ...layoutStyles.fieldInput,
                        ...(this.state.fieldErrors.pasajeroNombre ? layoutStyles.fieldInputError : {})
                      }}
                      value={this.state.pasajeroNombre}
                      onChange={this._onCambiarPasajeroNombre}
                      placeholder="Nombre y apellido del pasajero"
                      disabled={this.state.guardando}
                    />
                  )}
                  {this._renderFieldError(this.state.fieldErrors.pasajeroNombre)}
                </div>
              )}

              {!this._esAsociadoAViaje() && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>DNI</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>{this.state.pasajeroDni || '—'}</div>
                  ) : (
                    <input
                      type="text"
                      style={{
                        ...layoutStyles.fieldInput,
                        ...(this.state.fieldErrors.pasajeroDni ? layoutStyles.fieldInputError : {})
                      }}
                      value={this.state.pasajeroDni}
                      onChange={this._onCambiarPasajeroDni}
                      placeholder="DNI del pasajero"
                      disabled={this.state.guardando}
                    />
                  )}
                  {this._renderFieldError(this.state.fieldErrors.pasajeroDni)}
                </div>
              )}

              <div style={layoutStyles.fieldGroup}>
                <label style={layoutStyles.label}>Medio de Pago</label>
                {soloLectura ? (
                  <div style={layoutStyles.readOnlyValue}>{this.state.medioPago || '—'}</div>
                ) : (
                  <select
                    style={{
                      ...layoutStyles.select,
                      ...(this.state.fieldErrors.medioPago ? layoutStyles.fieldInputError : {})
                    }}
                    value={this.state.medioPago}
                    onChange={this._onCambiarMedioPago}
                    disabled={this.state.guardando}
                  >
                    <option value="">Seleccione...</option>
                    {MEDIOS_PAGO.map((medio: MedioPago) => (
                      <option key={medio} value={medio}>
                        {medio}
                      </option>
                    ))}
                  </select>
                )}
                {this._renderFieldError(this.state.fieldErrors.medioPago)}
              </div>

              {requiereCuentaBancaria && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Cuenta Bancaria</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>{this.state.banco || '—'}</div>
                  ) : (
                    <select
                      style={{
                        ...layoutStyles.select,
                        ...(this.state.fieldErrors.banco ? layoutStyles.fieldInputError : {})
                      }}
                      value={this.state.banco}
                      onChange={this._onCambiarBanco}
                      disabled={this.state.guardando}
                    >
                      <option value="">Seleccione...</option>
                      {this.state.opcionesBanco.map((opcion: string) => (
                        <option key={opcion} value={opcion}>
                          {opcion}
                        </option>
                      ))}
                    </select>
                  )}
                  {this._renderFieldError(this.state.fieldErrors.banco)}
                </div>
              )}

              <div style={layoutStyles.fieldGroup}>
                <label style={layoutStyles.label}>Motivo de elección</label>
                {soloLectura || this._esAsociadoAViaje() ? (
                  <div style={layoutStyles.readOnlyValue}>{this.state.motivo || '—'}</div>
                ) : (
                  <select
                    style={layoutStyles.select}
                    value={this.state.motivo}
                    onChange={this._onCambiarMotivo}
                    disabled={this.state.guardando}
                  >
                    <option value="">Seleccione...</option>
                    {this.state.opcionesMotivo.map((opcion: string) => (
                      <option key={opcion} value={opcion}>
                        {opcion}
                      </option>
                    ))}
                  </select>
                )}
              </div>

              <div style={layoutStyles.fieldGroup}>
                <label style={layoutStyles.label}>Fecha</label>
                {soloLectura ? (
                  <div style={layoutStyles.readOnlyDateValue}>
                    {formatDateDisplay(this.state.fechaPago) || '—'}
                  </div>
                ) : (
                  <div style={layoutStyles.dateInputWrap}>
                    <input
                      type="date"
                      lang="es-AR"
                      style={{
                        ...layoutStyles.dateInput,
                        ...(this.state.fieldErrors.fechaPago ? layoutStyles.dateInputError : {})
                      }}
                      value={getDateOnlyFromSharePoint(this.state.fechaPago)}
                      onChange={this._onCambiarFechaPago}
                      disabled={this.state.guardando}
                    />
                    <div style={layoutStyles.dateInputOverlay}>
                      {formatDateDisplay(this.state.fechaPago)}
                    </div>
                  </div>
                )}
                {this._renderFieldError(this.state.fieldErrors.fechaPago)}
              </div>

              <div style={layoutStyles.fieldGroup}>
                <label style={layoutStyles.label}>Monto</label>
                {soloLectura ? (
                  <div style={layoutStyles.readOnlyValue}>
                    {this.state.monto ? Number(this.state.monto).toLocaleString('es-AR') : '—'}
                  </div>
                ) : (
                  <input
                    type="number"
                    min="0"
                    step="0.01"
                    style={{
                      ...layoutStyles.fieldInput,
                      ...(this.state.fieldErrors.monto ? layoutStyles.fieldInputError : {})
                    }}
                    value={this.state.monto}
                    onChange={this._onCambiarMonto}
                    placeholder="0.00"
                    disabled={this.state.guardando}
                  />
                )}
                {this._renderFieldError(this.state.fieldErrors.monto)}
              </div>

              <div style={layoutStyles.fieldGroup}>
                <label style={layoutStyles.label}>Moneda</label>
                {soloLectura ? (
                  <div style={layoutStyles.readOnlyValue}>{this.state.moneda || '—'}</div>
                ) : (
                  <select
                    style={{
                      ...layoutStyles.select,
                      ...(this.state.fieldErrors.moneda ? layoutStyles.fieldInputError : {})
                    }}
                    value={this.state.moneda}
                    onChange={this._onCambiarMoneda}
                    disabled={this.state.guardando}
                  >
                    <option value="">Seleccione...</option>
                    {MONEDAS.map((moneda: Moneda) => (
                      <option key={moneda} value={moneda}>
                        {moneda}
                      </option>
                    ))}
                  </select>
                )}
                {this._renderFieldError(this.state.fieldErrors.moneda)}
              </div>

              {requiereCotizacion && (
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Cotización</label>
                  {soloLectura ? (
                    <div style={layoutStyles.readOnlyValue}>
                      {this.state.cotizacion ? Number(this.state.cotizacion).toLocaleString('es-AR') : '—'}
                    </div>
                  ) : (
                    <input
                      type="number"
                      min="0"
                      step="0.01"
                      style={{
                        ...layoutStyles.fieldInput,
                        ...(this.state.fieldErrors.cotizacion ? layoutStyles.fieldInputError : {})
                      }}
                      value={this.state.cotizacion}
                      onChange={this._onCambiarCotizacion}
                      placeholder="0.00"
                      disabled={this.state.guardando}
                    />
                  )}
                  {this._renderFieldError(this.state.fieldErrors.cotizacion)}
                </div>
              )}
            </div>
          </div>

          {this._renderSeccionComprobantes(soloLectura)}

          <div style={layoutStyles.bottomActions}>
            {soloLectura ? (
              <button
                type="button"
                style={layoutStyles.primaryButton}
                onClick={this._navegarAFormularioEdicion}
                disabled={this.state.guardando || !this._getItemId()}
              >
                Editar
              </button>
            ) : (
              <button
                type="button"
                style={
                  this.state.guardando
                    ? { ...layoutStyles.primaryButton, ...layoutStyles.buttonDisabled }
                    : layoutStyles.primaryButton
                }
                onClick={() => {
                  void this._onGuardar();
                }}
                disabled={this.state.guardando}
              >
                {this.state.guardando ? 'Guardando...' : this.state.pagoId ? 'Guardar' : 'Crear pago'}
              </button>
            )}
            <button
              type="button"
              style={
                this.state.guardando
                  ? { ...layoutStyles.defaultButton, ...layoutStyles.buttonDisabled }
                  : layoutStyles.defaultButton
              }
              onClick={this._onCancelar}
              disabled={this.state.guardando}
            >
              Cancelar
            </button>
          </div>
        </div>
      </div>
    );
  }
}
