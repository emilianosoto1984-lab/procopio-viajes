import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import SharePointViajesService, {
  IDestinoGeneralItem,
  IDestinoItem,
  ILiquidacionData,
  ILiquidacionItem,
  IOperadorItem,
  IPagoData,
  IPasajeroDraft,
  IPasajeroDocumentoItem,
  IPasajeroItem,
  IServicioViajeData,
  IServicioViajeItem,
  IFacturaItem,
  IPresupuestoItem,
  IVoucherItem,
  IViajeData
} from '../services/SharePointViajesService';
import styles from './ProcopioForms.module.scss';

export interface IProcopioFormsProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'ProcopioForms';
const COLORES_VIAJE = ['Azul', 'Rojo', 'Amarillo', 'Verde', 'Violeta', 'Naranja', 'Turquesa', 'Rosa', 'Gris', 'Marrón'] as const;
const ESTADOS_VIAJE = ['Activo', 'En Viaje', 'Finalizado'] as const;
type EstadoViaje = typeof ESTADOS_VIAJE[number];
const MONEDAS_SERVICIO = ['Dólares', 'Pesos'] as const;
type MonedaServicio = typeof MONEDAS_SERVICIO[number];
const MONEDAS_LIQUIDACION = ['Dólares', 'Pesos'] as const;
type MonedaLiquidacion = typeof MONEDAS_LIQUIDACION[number];
const MONEDAS_PAGO = ['Dólares', 'Pesos'] as const;
type MonedaPago = typeof MONEDAS_PAGO[number];

interface IServicio {
  id: number;
  concepto: string;
  precioCliente: number;
  moneda: MonedaServicio;
  operadorId?: number;
  operadorNombre?: string;
}

interface IMovimiento {
  id: number;
  movimiento: string;
  medioPago: string;
  fecha: string;
  moneda: string;
  monto: number;
  observaciones?: string;
  cotizacion?: number;
  tipo: 'Ingreso' | 'Egreso';
  liquidacionOperadorId?: number;
  servicioAsociadoId?: number;
  liquidacionOperadorNombre?: string;
}

interface IPasajero extends IPasajeroItem {}
interface IPasajeroPendiente extends IPasajeroDraft {
  tempId: number;
}

interface IProcopioFormsState {
  viajeId: number | null;
  nombreViaje: string;
  destinoGeneralId: number;
  destinoId: number;
  fechaSalida: string;
  fechaRegreso: string;
  estado: EstadoViaje;
  colorViaje: string;
  pasajeros: IPasajero[];
  pasajerosIds: number[];
  nuevosPasajeros: IPasajeroPendiente[];
  servicios: IServicio[];
  serviciosEliminadosIds: number[];
  movimientos: IMovimiento[];
  vouchers: IVoucherItem[];
  facturas: IFacturaItem[];
  presupuesto: IPresupuestoItem | null;
  liquidacionesOperador: ILiquidacionItem[];
  observaciones: string;
  servicioEnEdicion: {
    concepto: string;
    precioCliente: string;
    moneda: MonedaServicio;
    operadorId: string;
  };
  servicioEnEdicionId: number | null;
  mostrarEditorServicio: boolean;
  movimientoEnEdicion: {
    tipo: 'Ingreso' | 'Egreso';
    movimiento: string;
    medioPago: string;
    fecha: string;
    moneda: MonedaPago;
    monto: string;
    observaciones: string;
    cotizacion: string;
    liquidacionOperadorId: string;
    servicioAsociadoId: string;
  };
  nuevoPasajeroDraft: IPasajeroDraft;
  pasajeroBusquedaTexto: string;
  mostrarEditorPasajero: boolean;
  pasajeroEnEdicionId: number | null;
  pasajeroCamposHabilitados: boolean;
  pasajeroDocumentosExpandidoId: number | null;
  pasajeroDocumentosPorId: { [pasajeroId: number]: IPasajeroDocumentoItem[] };
  pasajeroDocumentosCargandoId: number | null;
  movimientoEnEdicionId: number | null;
  mostrarEditorMovimiento: boolean;
  mostrarEditorVoucher: boolean;
  vouchersSubiendo: boolean;
  voucherEnEdicion: {
    archivos: File[];
  };
  mostrarEditorFactura: boolean;
  facturasSubiendo: boolean;
  facturaEnEdicion: {
    archivos: File[];
  };
  mostrarEditorPresupuesto: boolean;
  presupuestoSubiendo: boolean;
  presupuestoArchivoPendiente: File | null;
  mostrarEditorLiquidacion: boolean;
  liquidacionEnEdicion: {
    codigoReferencia: string;
    operadorId: string;
    monto: string;
    moneda: MonedaLiquidacion;
    archivoFile: File | null;
  };
  liquidacionEnEdicionId: number | null;
  cargando: boolean;
  guardando: boolean;
  error: string;
  sectionErrors: {
    pasajeros: string;
    servicios: string;
    liquidaciones: string;
    movimientos: string;
    vouchers: string;
    facturas: string;
    presupuesto: string;
  };
  destinos: IDestinoItem[];
  destinosGenerales: IDestinoGeneralItem[];
  operadores: IOperadorItem[];
}

const layoutStyles: { [key: string]: React.CSSProperties } = {
  page: { fontFamily: 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif', fontSize: 14, backgroundColor: '#f3f2f1', padding: 16, boxSizing: 'border-box' },
  container: { maxWidth: 1100, margin: '0 auto' },
  header: { marginBottom: 16 },
  title: { fontSize: 24, fontWeight: 600, marginBottom: 4 },
  subtitle: { fontSize: 14, color: '#605e5c' },
  section: { backgroundColor: '#ffffff', borderRadius: 6, border: '1px solid #e1dfdd', padding: 16, marginBottom: 16, boxShadow: '0 1px 2px rgba(0,0,0,0.04)' },
  sectionHeader: { fontSize: 16, fontWeight: 600, marginBottom: 12 },
  fieldGroup: { display: 'flex', flexDirection: 'column', marginBottom: 12 },
  label: { fontWeight: 600, marginBottom: 4 },
  input: { padding: '6px 8px', borderRadius: 4, border: '1px solid #c8c6c4', fontSize: 14 },
  tableHeaderRow: { backgroundColor: '#f7f8fb' },
  table: { width: '100%', borderCollapse: 'collapse', marginTop: 6, marginBottom: 10 },
  th: {
    textAlign: 'left',
    padding: '11px 10px',
    borderBottom: '1px solid #d9dce3',
    fontWeight: 700,
    fontSize: 12.5,
    textTransform: 'uppercase',
    letterSpacing: '0.08em',
    color: '#5f6472',
    lineHeight: 1.2
  },
  td: {
    padding: '11px 10px',
    borderBottom: '1px solid #eceef3',
    fontSize: 14,
    color: '#2b2f38',
    lineHeight: 1.35
  },
  actionsCell: { textAlign: 'right', whiteSpace: 'nowrap' },
  emptyTableCell: {
    padding: '22px 16px',
    textAlign: 'center',
    color: '#7b8190',
    fontSize: 13.5,
    borderBottom: 'none'
  },
  smallButton: { padding: '4px 8px', borderRadius: 4, border: '1px solid #c8c6c4', backgroundColor: '#ffffff', cursor: 'pointer', fontSize: 12, marginLeft: 4 },
  primaryButton: { padding: '6px 14px', borderRadius: 4, border: '1px solid #0078d4', backgroundColor: '#0078d4', color: '#ffffff', cursor: 'pointer', fontSize: 14, fontWeight: 600, marginRight: 8 },
  defaultButton: { padding: '6px 14px', borderRadius: 4, border: '1px solid #c8c6c4', backgroundColor: '#ffffff', cursor: 'pointer', fontSize: 14, fontWeight: 400, marginRight: 8 },
  buttonDisabled: { opacity: 0.5, cursor: 'default' },
  sectionToolbar: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 },
  toolbarTitle: { fontSize: 16, fontWeight: 600 },
  inlineEditor: { backgroundColor: '#faf9f8', borderRadius: 4, border: '1px dashed #c8c6c4', padding: 12, marginTop: 8 },
  inlineEditorRow: { display: 'flex', gap: 12, marginBottom: 8 },
  inlineEditorField: { flex: 1 },
  bottomActions: { display: 'flex', justifyContent: 'flex-end', marginTop: 16 },
  info: { marginBottom: 10, color: '#605e5c' },
  error: { color: '#a4262c', marginBottom: 10 },
  voucherAttachmentZone: {
    border: '1px solid #e1dfdd',
    borderRadius: 4,
    backgroundColor: '#faf9f8',
    padding: 0,
    overflow: 'hidden'
  },
  voucherAttachmentEmpty: { padding: '22px 16px', color: '#7b8190', fontSize: 13.5, textAlign: 'center' },
  voucherAttachmentRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: 12,
    padding: '12px 16px',
    borderBottom: '1px solid #eceef3',
    backgroundColor: '#ffffff'
  },
  voucherAttachmentRowLast: { borderBottom: 'none' },
  voucherAttachmentIcon: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 36,
    height: 44,
    flexShrink: 0,
    borderRadius: 2,
    border: '1px solid #c8c6c4',
    backgroundColor: '#f3f2f1',
    boxSizing: 'border-box',
    color: '#605e5c'
  },
  voucherAttachmentMeta: { flex: 1, minWidth: 0 },
  voucherAttachmentFileName: { fontSize: 14, fontWeight: 600, color: '#2b2f38', wordBreak: 'break-word' },
  voucherAttachmentActions: { flexShrink: 0, paddingTop: 2 },
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
  fileInputHint: { fontSize: 12, color: '#605e5c' }
};

/** Border-radius tokens used only by the Pasajeros section (card > surface > controls). */
const pasajerosRadius = {
  /** Main section card */
  card: 8,
  /** Table shell, autocomplete panel */
  surface: 6,
  /** Inputs and all buttons */
  control: 4
} as const;

const pasajerosStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: pasajerosRadius.card },
  tableOuter: {
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    marginBottom: 8,
    borderRadius: pasajerosRadius.surface,
    overflow: 'hidden' as const
  },
  table: { ...layoutStyles.table, width: '100%', marginTop: 0, marginBottom: 0 },
  input: {
    ...layoutStyles.input,
    width: '100%',
    height: 38,
    padding: '0 14px',
    borderRadius: 8,
    border: '1.4px solid #CDD0D7',
    background: '#FAFAFC',
    fontSize: 15,
    color: '#232529',
    outline: 'none',
    transition: 'border-color 0.18s',
    boxSizing: 'border-box' as const,
    lineHeight: 1.2
  },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: pasajerosRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: pasajerosRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: pasajerosRadius.control },
  btnSmall: { ...layoutStyles.smallButton, borderRadius: pasajerosRadius.control },
  autocompleteList: {
    border: '1px solid #c8c6c4',
    borderRadius: pasajerosRadius.surface,
    backgroundColor: '#ffffff',
    maxHeight: 140,
    overflowY: 'auto' as const,
    marginTop: 4
  },
  /** Spacing only; no panel border — radii apply to inputs/buttons/list above */
  inlineFormRoot: { marginTop: 10, paddingTop: 6 },
  dateInputShell: { position: 'relative' as const, width: '100%' },
  dateInputOverlay: {
    position: 'absolute' as const,
    left: 14,
    right: 42,
    top: 0,
    height: 38,
    display: 'flex',
    alignItems: 'center',
    fontSize: 15,
    pointerEvents: 'none' as const,
    whiteSpace: 'nowrap' as const,
    overflow: 'hidden',
    textOverflow: 'ellipsis'
  },
  dateInputOverlayValue: { color: '#232529' },
  dateInputOverlayPlaceholder: { color: '#8a8a8a' },
  documentosToggle: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: 6,
    padding: '4px 10px',
    fontSize: 12,
    fontWeight: 600,
    color: '#005A9E',
    background: 'transparent',
    border: '1px solid #c8c6c4',
    borderRadius: pasajerosRadius.control,
    cursor: 'pointer'
  },
  documentosToggleActive: {
    backgroundColor: '#f3f2f1',
    borderColor: '#8a8886'
  },
  documentosExpandCell: {
    padding: '10px 16px 14px',
    backgroundColor: '#faf9f8',
    borderTop: '1px solid #edebe9'
  },
  documentosZone: {
    border: '1px solid #edebe9',
    borderRadius: pasajerosRadius.surface,
    backgroundColor: '#ffffff',
    overflow: 'hidden' as const
  },
  documentosEmpty: { padding: '12px 14px', color: '#7b8190', fontSize: 13 },
  documentosInfo: { padding: '12px 14px', color: '#605e5c', fontSize: 13, lineHeight: 1.45 },
  documentosRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '10px 14px',
    borderBottom: '1px solid #f3f2f1'
  },
  documentosRowLast: { borderBottom: 'none' },
  documentosFileName: { flex: 1, minWidth: 0, fontSize: 13, fontWeight: 600, color: '#2b2f38', wordBreak: 'break-word' as const },
  documentosUploadRow: { marginTop: 10, display: 'flex', flexWrap: 'wrap' as const, alignItems: 'center', gap: 8 }
};

/** Border-radius tokens for Servicios section only (aligned with Pasajeros: card > surface > controls). */
const serviciosRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const serviciosStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: serviciosRadius.card },
  tableOuter: {
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    marginBottom: 8,
    borderRadius: serviciosRadius.surface,
    overflow: 'hidden' as const
  },
  table: { ...layoutStyles.table, width: '100%', marginTop: 0, marginBottom: 0 },
  input: { ...layoutStyles.input, borderRadius: serviciosRadius.control },
  select: { ...layoutStyles.input, borderRadius: serviciosRadius.control },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: serviciosRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: serviciosRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: serviciosRadius.control },
  btnSmall: { ...layoutStyles.smallButton, borderRadius: serviciosRadius.control },
  inlineFormRoot: { marginTop: 10, paddingTop: 6 }
};

/** Border-radius tokens for Liquidaciones Operador only (aligned with Pasajeros / Servicios). */
const liquidacionesRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const liquidacionesStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: liquidacionesRadius.card },
  tableOuter: {
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    marginBottom: 8,
    borderRadius: liquidacionesRadius.surface,
    overflow: 'hidden' as const
  },
  table: { ...layoutStyles.table, width: '100%', marginTop: 0, marginBottom: 0 },
  input: { ...layoutStyles.input, borderRadius: liquidacionesRadius.control },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: liquidacionesRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: liquidacionesRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: liquidacionesRadius.control },
  btnSmall: { ...layoutStyles.smallButton, borderRadius: liquidacionesRadius.control },
  inlineFormRoot: { marginTop: 10, paddingTop: 6 },
  archivoNombreEnTabla: { wordBreak: 'break-word' as const, fontSize: 13, color: '#201f1e' },
  archivoLink: {
    wordBreak: 'break-word' as const,
    fontSize: 13,
    color: '#005A9E',
    textDecoration: 'none',
    fontWeight: 600,
    cursor: 'pointer'
  }
};

/** Border-radius tokens for Registro de Ingresos y Egresos only (aligned with other sections). */
const movimientosRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const movimientosStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: movimientosRadius.card },
  tableOuter: {
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    marginBottom: 8,
    borderRadius: movimientosRadius.surface,
    overflow: 'hidden' as const
  },
  table: { ...layoutStyles.table, width: '100%', marginTop: 0, marginBottom: 0 },
  input: { ...layoutStyles.input, borderRadius: movimientosRadius.control },
  select: { ...layoutStyles.input, borderRadius: movimientosRadius.control },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: movimientosRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: movimientosRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: movimientosRadius.control },
  btnSmall: { ...layoutStyles.smallButton, borderRadius: movimientosRadius.control },
  inlineFormRoot: { marginTop: 10, paddingTop: 6 }
};

/** Border-radius tokens for Vouchers only (aligned with other sections). */
const vouchersRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const vouchersStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: vouchersRadius.card },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: vouchersRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: vouchersRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: vouchersRadius.control },
  btnSmall: { ...layoutStyles.smallButton, borderRadius: vouchersRadius.control },
  input: { ...layoutStyles.input, borderRadius: vouchersRadius.control },
  subheading: { fontSize: 16, fontWeight: 600, marginBottom: 10, marginTop: 2 },
  /** Full-width list shell inside the section card */
  attachmentZone: {
    ...layoutStyles.voucherAttachmentZone,
    borderRadius: vouchersRadius.surface,
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    border: '1px solid #e1dfdd',
    overflow: 'hidden' as const,
    backgroundColor: '#faf9f8'
  },
  attachmentEmpty: { ...layoutStyles.voucherAttachmentEmpty, padding: '14px 16px' },
  attachmentRow: { ...layoutStyles.voucherAttachmentRow, padding: '12px 16px' },
  attachmentRowLast: layoutStyles.voucherAttachmentRowLast,
  attachmentIcon: { ...layoutStyles.voucherAttachmentIcon, borderRadius: vouchersRadius.control },
  attachmentMeta: layoutStyles.voucherAttachmentMeta,
  attachmentFileName: { ...layoutStyles.voucherAttachmentFileName, fontSize: 14, fontWeight: 600, wordBreak: 'break-word' as const },
  attachmentActions: layoutStyles.voucherAttachmentActions,
  /** Inline upload block — continuation of section, no dashed panel */
  inlineUploadRoot: { marginTop: 10, paddingTop: 8 },
  uploadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '12px 16px',
    fontSize: 13,
    color: '#605e5c',
    backgroundColor: '#ffffff',
    borderBottom: '1px solid #eceef3'
  }
};

/** Border-radius tokens for Facturas only (aligned with Vouchers). */
const facturasRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const facturasStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: facturasRadius.card },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: facturasRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: facturasRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: facturasRadius.control },
  input: { ...layoutStyles.input, borderRadius: facturasRadius.control },
  subheading: { fontSize: 16, fontWeight: 600, marginBottom: 10, marginTop: 2 },
  attachmentZone: {
    ...layoutStyles.voucherAttachmentZone,
    borderRadius: facturasRadius.surface,
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    border: '1px solid #e1dfdd',
    overflow: 'hidden' as const,
    backgroundColor: '#faf9f8'
  },
  attachmentEmpty: { ...layoutStyles.voucherAttachmentEmpty, padding: '14px 16px' },
  attachmentRow: { ...layoutStyles.voucherAttachmentRow, padding: '12px 16px' },
  attachmentRowLast: layoutStyles.voucherAttachmentRowLast,
  attachmentIcon: { ...layoutStyles.voucherAttachmentIcon, borderRadius: facturasRadius.control },
  attachmentMeta: layoutStyles.voucherAttachmentMeta,
  attachmentFileName: { ...layoutStyles.voucherAttachmentFileName, fontSize: 14, fontWeight: 600, wordBreak: 'break-word' as const },
  attachmentActions: layoutStyles.voucherAttachmentActions,
  inlineUploadRoot: { marginTop: 10, paddingTop: 8 },
  uploadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '12px 16px',
    fontSize: 13,
    color: '#605e5c',
    backgroundColor: '#ffffff',
    borderBottom: '1px solid #eceef3'
  }
};

/** Border-radius tokens for Presupuesto only (aligned with Vouchers). */
const presupuestoRadius = {
  card: 8,
  surface: 6,
  control: 4
} as const;

const presupuestoStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: presupuestoRadius.card },
  btnAdd: { ...layoutStyles.defaultButton, borderRadius: presupuestoRadius.control, padding: '6px 12px' },
  btnDefault: { ...layoutStyles.defaultButton, borderRadius: presupuestoRadius.control },
  btnPrimary: { ...layoutStyles.primaryButton, borderRadius: presupuestoRadius.control },
  input: { ...layoutStyles.input, borderRadius: presupuestoRadius.control },
  attachmentZone: {
    ...layoutStyles.voucherAttachmentZone,
    borderRadius: presupuestoRadius.surface,
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 4,
    border: '1px solid #e1dfdd',
    overflow: 'hidden' as const,
    backgroundColor: '#faf9f8'
  },
  attachmentEmpty: { ...layoutStyles.voucherAttachmentEmpty, padding: '14px 16px' },
  attachmentRow: { ...layoutStyles.voucherAttachmentRow, padding: '12px 16px', ...layoutStyles.voucherAttachmentRowLast },
  attachmentIcon: { ...layoutStyles.voucherAttachmentIcon, borderRadius: presupuestoRadius.control },
  attachmentMeta: layoutStyles.voucherAttachmentMeta,
  attachmentFileName: { ...layoutStyles.voucherAttachmentFileName, fontSize: 14, fontWeight: 600, wordBreak: 'break-word' as const },
  attachmentActions: layoutStyles.voucherAttachmentActions,
  inlineUploadRoot: { marginTop: 10, paddingTop: 8 },
  uploadingRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '12px 16px',
    fontSize: 13,
    color: '#605e5c',
    backgroundColor: '#ffffff',
    borderBottom: '1px solid #eceef3'
  }
};

/** Border-radius tokens for Observaciones only (aligned with other sections). */
const observacionesRadius = {
  card: 8,
  /** Inner field shell — matches list/table surfaces elsewhere */
  surface: 6,
  control: 4
} as const;

const observacionesStyles = {
  sectionCard: { ...layoutStyles.section, borderRadius: observacionesRadius.card },
  title: { fontSize: 16, fontWeight: 600, marginBottom: 10, color: '#201f1e' },
  /** Full-bleed wrapper so the field aligns with tables/lists inside other cards */
  textareaOuter: {
    width: 'calc(100% + 32px)',
    marginLeft: -16,
    marginRight: -16,
    marginTop: 2
  },
  /** Framed surface: same border/shadow language as voucher list & inputs, not a raw control */
  textareaShell: {
    borderRadius: observacionesRadius.surface,
    border: '1px solid #e1dfdd',
    backgroundColor: '#ffffff',
    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
    overflow: 'hidden' as const
  },
  textareaShellDisabled: {
    backgroundColor: '#f3f2f1',
    boxShadow: 'none'
  },
  textarea: {
    width: '100%',
    maxWidth: '100%',
    minHeight: 132,
    margin: 0,
    padding: '12px 14px',
    border: 'none',
    outline: 'none',
    resize: 'vertical' as const,
    boxSizing: 'border-box' as const,
    display: 'block' as const,
    fontSize: 14,
    lineHeight: 1.5,
    fontFamily: 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif',
    color: '#201f1e',
    backgroundColor: 'transparent'
  },
  textareaDisabled: {
    color: '#605e5c',
    cursor: 'not-allowed' as const
  }
};

/** Inline SVGs for grid row actions (consistent set, no extra packages). */
const GridIconEdit: React.FC = () => (
  <svg width={16} height={16} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
  </svg>
);

const GridIconTrash: React.FC = () => (
  <svg width={16} height={16} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <polyline points="3 6 5 6 21 6" />
    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
  </svg>
);

const GridIconDownload: React.FC = () => (
  <svg width={16} height={16} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

const AttachmentFileIcon: React.FC = () => (
  <svg width={22} height={22} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
    <line x1="16" y1="13" x2="8" y2="13" />
    <line x1="16" y1="17" x2="8" y2="17" />
  </svg>
);

interface IAttachmentIconBoxProps {
  title: string;
  boxStyle?: React.CSSProperties;
}

const InlineLoadingSpinner: React.FC = () => (
  <svg width={18} height={18} viewBox="0 0 24 24" aria-hidden="true">
    <circle cx="12" cy="12" r="10" stroke="#c8c6c4" strokeWidth="3" fill="none" />
    <path d="M12 2a10 10 0 0 1 10 10" stroke="#0078d4" strokeWidth="3" fill="none" strokeLinecap="round">
      <animateTransform attributeName="transform" type="rotate" from="0 12 12" to="360 12 12" dur="0.8s" repeatCount="indefinite" />
    </path>
  </svg>
);

const AttachmentIconBox: React.FC<IAttachmentIconBoxProps> = (props: IAttachmentIconBoxProps) => (
  <div style={{ ...layoutStyles.voucherAttachmentIcon, ...props.boxStyle }} title={props.title} aria-hidden="true">
    <AttachmentFileIcon />
  </div>
);

const GridIconPaperclip: React.FC = () => (
  <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M21.44 11.05l-9.19 9.19a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48" />
  </svg>
);

const GridIconChevronDown: React.FC = () => (
  <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <polyline points="6 9 12 15 18 9" />
  </svg>
);

const GridIconChevronUp: React.FC = () => (
  <svg width={14} height={14} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <polyline points="18 15 12 9 6 15" />
  </svg>
);

const gridActionBarStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
  alignItems: 'center',
  gap: 4,
  flexWrap: 'nowrap' as const
};

interface IGridIconActionButtonProps {
  title: string;
  onClick: () => void;
  disabled?: boolean;
  children: React.ReactNode;
}

interface IFileInputEspanolProps {
  inputId: string;
  multiple?: boolean;
  disabled?: boolean;
  accept?: string;
  buttonLabel?: string;
  hint?: string;
  onChange: (ev: React.ChangeEvent<HTMLInputElement>) => void;
  buttonStyle?: React.CSSProperties;
}

/** Input file nativo oculto + etiqueta en español (el navegador muestra "Choose files" en inglés). */
const FileInputEspanol: React.FC<IFileInputEspanolProps> = (props: IFileInputEspanolProps) => {
  const {
    inputId,
    multiple,
    disabled,
    accept,
    buttonLabel,
    hint,
    onChange,
    buttonStyle
  } = props;
  const etiqueta = buttonLabel || (multiple ? 'Seleccionar archivos' : 'Seleccionar archivo');
  return (
    <div style={layoutStyles.fileInputWrap}>
      <input
        id={inputId}
        type="file"
        multiple={multiple}
        disabled={disabled}
        accept={accept}
        onChange={onChange}
        style={layoutStyles.fileInputHidden}
        aria-label={etiqueta}
      />
      <label
        htmlFor={inputId}
        style={
          disabled
            ? { ...layoutStyles.fileInputButton, ...layoutStyles.buttonDisabled, pointerEvents: 'none' as const, ...buttonStyle }
            : { ...layoutStyles.fileInputButton, ...buttonStyle }
        }
      >
        {etiqueta}
      </label>
      {hint ? <span style={layoutStyles.fileInputHint}>{hint}</span> : null}
    </div>
  );
};

const GridIconActionButton: React.FC<IGridIconActionButtonProps> = (props: IGridIconActionButtonProps) => {
  const { title, onClick, disabled, children } = props;
  const isDisabled = !!disabled;
  return (
    <button
      type="button"
      title={title}
      aria-label={title}
      disabled={isDisabled}
      onClick={onClick}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 28,
        height: 28,
        padding: 0,
        border: 'none',
        borderRadius: 4,
        backgroundColor: 'transparent',
        color: '#323130',
        cursor: isDisabled ? 'not-allowed' : 'pointer',
        opacity: isDisabled ? 0.45 : 1,
        boxSizing: 'border-box'
      }}
      onMouseEnter={(e: React.MouseEvent<HTMLButtonElement>) => {
        if (!isDisabled) {
          e.currentTarget.style.backgroundColor = '#edebe9';
        }
      }}
      onMouseLeave={(e: React.MouseEvent<HTMLButtonElement>) => {
        e.currentTarget.style.backgroundColor = 'transparent';
      }}
    >
      {children}
    </button>
  );
};

export default class ProcopioForms extends React.Component<IProcopioFormsProps, IProcopioFormsState> {
  private _service: SharePointViajesService;

  public constructor(props: IProcopioFormsProps) {
    super(props);
    this._service = new SharePointViajesService(props.context);
    this.state = {
      viajeId: null,
      nombreViaje: '',
      destinoGeneralId: 0,
      destinoId: 0,
      fechaSalida: '',
      fechaRegreso: '',
      estado: 'Activo',
      colorViaje: this._getRandomColorViaje(),
      pasajeros: [],
      pasajerosIds: [],
      nuevosPasajeros: [],
      servicios: [],
      serviciosEliminadosIds: [],
      movimientos: [],
      vouchers: [],
      facturas: [],
      presupuesto: null,
      liquidacionesOperador: [],
      observaciones: '',
      servicioEnEdicion: { concepto: '', precioCliente: '', moneda: 'Pesos', operadorId: '' },
      servicioEnEdicionId: null,
      mostrarEditorServicio: false,
      nuevoPasajeroDraft: { nombreApellido: '', dni: '', pasaporte: '', telefono: '', email: '', observaciones: '', fechaNacimiento: '' },
      pasajeroBusquedaTexto: '',
      mostrarEditorPasajero: false,
      pasajeroEnEdicionId: null,
      pasajeroCamposHabilitados: false,
      pasajeroDocumentosExpandidoId: null,
      pasajeroDocumentosPorId: {},
      pasajeroDocumentosCargandoId: null,
      movimientoEnEdicion: { tipo: 'Ingreso', movimiento: '', medioPago: '', fecha: this._getTodayDateInput(), moneda: 'Pesos', monto: '', observaciones: '', cotizacion: '', liquidacionOperadorId: '', servicioAsociadoId: '' },
      movimientoEnEdicionId: null,
      mostrarEditorMovimiento: false,
      mostrarEditorVoucher: false,
      vouchersSubiendo: false,
      voucherEnEdicion: { archivos: [] },
      mostrarEditorFactura: false,
      facturasSubiendo: false,
      facturaEnEdicion: { archivos: [] },
      mostrarEditorPresupuesto: false,
      presupuestoSubiendo: false,
      presupuestoArchivoPendiente: null,
      mostrarEditorLiquidacion: false,
      liquidacionEnEdicion: { codigoReferencia: '', operadorId: '', monto: '', moneda: 'Pesos', archivoFile: null },
      liquidacionEnEdicionId: null,
      cargando: true,
      guardando: false,
      error: '',
      sectionErrors: {
        pasajeros: '',
        servicios: '',
        liquidaciones: '',
        movimientos: '',
        vouchers: '',
        facturas: '',
        presupuesto: ''
      },
      destinos: [],
      destinosGenerales: [],
      operadores: []
    };
  }

  public async componentDidMount(): Promise<void> {
    Log.info(LOG_SOURCE, 'React Element: ProcopioForms mounted');
    await this._cargarDatosIniciales();
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: ProcopioForms unmounted');
  }

  private async _cargarDatosIniciales(): Promise<void> {
    try {
      const [pasajeros, destinos, destinosGenerales, operadores] = await Promise.all([
        this._service.getPasajeros(),
        this._service.getDestinos(),
        this._service.getDestinosGenerales(),
        this._service.getOperadores()
      ]);
      const itemId = this._getItemId();
      if ((this.props.displayMode === FormDisplayMode.Edit || this.props.displayMode === FormDisplayMode.Display) && itemId) {
        const viaje = await this._service.getViajeById(itemId);
        const [pagos, liquidaciones, vouchers, facturas, presupuesto, serviciosViaje] = await Promise.all([
          this._service.getPagosByViaje(itemId),
          this._service.getLiquidacionesByViaje(itemId),
          this._service.getVouchersByViaje(itemId),
          this._service.getFacturasByViaje(itemId),
          this._service.getPresupuestoByViaje(itemId),
          this._service.getServiciosViajeByViaje(itemId)
        ]);
        const serviciosDetalle: IServicio[] = serviciosViaje.map((servicio: IServicioViajeItem) => ({
          id: servicio.id,
          concepto: servicio.concepto,
          precioCliente: servicio.precioCliente,
          moneda: this._normalizarMonedaServicio(servicio.moneda),
          operadorId: servicio.operadorId,
          operadorNombre: servicio.operadorNombre
        }));
        const destinoSeleccionado = destinos.filter((d: IDestinoItem) => d.id === viaje.destinoId)[0];
        const destinoGeneralId = viaje.destinoGeneralId && viaje.destinoGeneralId > 0
          ? viaje.destinoGeneralId
          : destinoSeleccionado && destinoSeleccionado.destinoGeneralId
            ? destinoSeleccionado.destinoGeneralId
            : 0;
        this.setState({
          viajeId: itemId,
          nombreViaje: viaje.nombre,
          destinoGeneralId,
          destinoId: viaje.destinoId,
          fechaSalida: this._getDateOnlyFromSharePoint(viaje.fechaSalida),
          fechaRegreso: this._getDateOnlyFromSharePoint(viaje.fechaLlegada),
          estado: this._normalizarEstado(viaje.estado),
          colorViaje: viaje.colorViaje || this.state.colorViaje,
          observaciones: viaje.observaciones,
          destinos,
          destinosGenerales,
          operadores,
          pasajeros,
          pasajerosIds: viaje.pasajerosIds,
          servicios: serviciosDetalle,
          serviciosEliminadosIds: [],
          movimientos: pagos.map(p => ({
            id: p.id,
            movimiento: p.concepto,
            medioPago: p.medioPago,
            fecha: p.fechaPago,
            moneda: this._normalizarMonedaPago(p.moneda),
            monto: p.importe,
            observaciones: p.observaciones,
            cotizacion: p.cotizacion,
            tipo:
              p.tipoPago === 'Ingreso' || p.tipoPago === 'Egreso'
                ? p.tipoPago
                : p.liquidacionOperadorId && p.liquidacionOperadorId > 0
                  ? 'Egreso'
                  : 'Ingreso',
            liquidacionOperadorId: p.liquidacionOperadorId,
            servicioAsociadoId: p.servicioAsociadoId,
            liquidacionOperadorNombre: p.liquidacionOperadorNombre
          })),
          liquidacionesOperador: liquidaciones,
          vouchers,
          facturas,
          presupuesto,
          cargando: false
        });
        return;
      }
      this.setState({ pasajeros, destinos, destinosGenerales, operadores, cargando: false });
    } catch (error) {
      this.setState({ cargando: false, error: 'No se pudieron cargar los datos de SharePoint.' });
    }
  }

  private _getItemId(): number | null {
    const fromContext = this.props.context.itemId;
    if (fromContext !== undefined && fromContext > 0) {
      return fromContext;
    }
    const possibleId = (this.props.context as any).itemId;
    const parsed = Number(possibleId);
    return parsed > 0 ? parsed : null;
  }

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
    window.location.assign(url);
  };

  private _getPasajeroDisplayUrl(pasajeroId: number): string {
    const webUrl = this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    return webUrl + '/Lists/Pasajeros/DispForm.aspx?ID=' + pasajeroId;
  }

  private _esSoloLectura(): boolean {
    return this.props.displayMode === FormDisplayMode.Display;
  }

  private _getSubtitle(): string {
    switch (this.props.displayMode) {
      case FormDisplayMode.New: return 'Nuevo viaje';
      case FormDisplayMode.Edit: return 'Editar viaje';
      default: return 'Ver viaje';
    }
  }

  private _formatDateDisplay(value: string): string {
    if (!value) {
      return '';
    }
    const datePart = this._getDateOnlyFromSharePoint(value); // YYYY-MM-DD
    const parts = datePart.split('-');
    if (parts.length !== 3) {
      return value;
    }
    const year = parts[0];
    const month = parts[1];
    const day = parts[2];
    return day + '/' + month + '/' + year;
  }

  private _getDateOnlyFromSharePoint(value: string): string {
    if (!value) {
      return '';
    }
    return value.split('T')[0];
  }

  private _getRandomColorViaje(): string {
    const index = Math.floor(Math.random() * COLORES_VIAJE.length);
    return COLORES_VIAJE[index];
  }

  private _getTodayDateInput(): string {
    return new Date().toISOString().substr(0, 10);
  }

  private _setSectionError = (section: keyof IProcopioFormsState['sectionErrors'], message: string): void => {
    this.setState(prev => ({
      sectionErrors: {
        ...prev.sectionErrors,
        [section]: message
      }
    }));
  };

  private _clearSectionError = (section: keyof IProcopioFormsState['sectionErrors']): void => {
    this.setState(prev => ({
      sectionErrors: {
        ...prev.sectionErrors,
        [section]: ''
      }
    }));
  };

  private _onCambiarNombre = (ev: React.ChangeEvent<HTMLInputElement>): void => this.setState({ nombreViaje: ev.target.value });
  private _onCambiarDestinoGeneral = (ev: React.ChangeEvent<HTMLSelectElement>): void => {
    const destinoGeneralId = Number(ev.target.value) || 0;
    this.setState(prev => {
      const destinoActual = prev.destinos.filter((d: IDestinoItem) => d.id === prev.destinoId)[0];
      const mantenerDestino =
        destinoActual &&
        destinoGeneralId > 0 &&
        destinoActual.destinoGeneralId === destinoGeneralId;
      return {
        destinoGeneralId,
        destinoId: mantenerDestino ? prev.destinoId : 0
      };
    });
  };

  private _onCambiarDestino = (ev: React.ChangeEvent<HTMLSelectElement>): void => {
    const destinoId = Number(ev.target.value) || 0;
    this.setState(prev => {
      const destino = prev.destinos.filter((d: IDestinoItem) => d.id === destinoId)[0];
      return {
        destinoId,
        destinoGeneralId: destino && destino.destinoGeneralId ? destino.destinoGeneralId : prev.destinoGeneralId
      };
    });
  };
  private _onCambiarFechaSalida = (ev: React.ChangeEvent<HTMLInputElement>): void =>
    this.setState({ fechaSalida: this._getDateOnlyFromSharePoint(ev.target.value) });
  private _onCambiarFechaRegreso = (ev: React.ChangeEvent<HTMLInputElement>): void =>
    this.setState({ fechaRegreso: this._getDateOnlyFromSharePoint(ev.target.value) });
  private _onCambiarEstado = (ev: React.ChangeEvent<HTMLSelectElement>): void => this.setState({ estado: this._normalizarEstado(ev.target.value) });
  private _onCambiarObservaciones = (ev: React.ChangeEvent<HTMLTextAreaElement>): void => this.setState({ observaciones: ev.target.value });

  private _onCambiarNuevoPasajero = (campo: keyof IPasajeroDraft, value: string): void => {
    this.setState(prev => ({ nuevoPasajeroDraft: { ...prev.nuevoPasajeroDraft, [campo]: value } }));
  };

  private _onCambiarFechaNacimientoPasajero = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    this._onCambiarNuevoPasajero('fechaNacimiento', this._getDateOnlyFromSharePoint(ev.target.value));
  };

  private _onCambiarBusquedaPasajero = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ pasajeroBusquedaTexto: ev.target.value, pasajeroCamposHabilitados: false });
  };

  private _onSeleccionarPasajeroDesdeBusqueda = (pasajeroId: number): void => {
    this.setState(prev => ({
      pasajerosIds: prev.pasajerosIds.indexOf(pasajeroId) >= 0 ? prev.pasajerosIds : prev.pasajerosIds.concat([pasajeroId]),
      pasajeroBusquedaTexto: '',
      mostrarEditorPasajero: false,
      pasajeroEnEdicionId: null,
      pasajeroCamposHabilitados: false
    }));
  };

  private _abrirEditorPasajero = (): void => {
    if (this._esSoloLectura()) { return; }
    this._clearSectionError('pasajeros');
    this.setState({
      mostrarEditorPasajero: true,
      pasajeroEnEdicionId: null,
      pasajeroBusquedaTexto: '',
      pasajeroCamposHabilitados: false,
      nuevoPasajeroDraft: { nombreApellido: '', dni: '', pasaporte: '', telefono: '', email: '', observaciones: '', fechaNacimiento: '' }
    });
  };

  private _editarPasajero = (id: number): void => {
    if (this._esSoloLectura()) { return; }
    this._clearSectionError('pasajeros');
    const pasajero = this.state.pasajeros.filter((p: IPasajero) => p.id === id)[0];
    if (!pasajero) { return; }
    this.setState({
      mostrarEditorPasajero: true,
      pasajeroEnEdicionId: id,
      pasajeroBusquedaTexto: '',
      pasajeroCamposHabilitados: true,
      nuevoPasajeroDraft: {
        nombreApellido: pasajero.nombreApellido,
        dni: pasajero.dni,
        pasaporte: pasajero.pasaporte || '',
        telefono: pasajero.telefono || '',
        email: pasajero.email || '',
        observaciones: pasajero.observaciones || '',
        fechaNacimiento: pasajero.fechaNacimiento || ''
      }
    });
  };

  private _habilitarNuevoPasajero = (): void => {
    this.setState({
      pasajeroCamposHabilitados: true,
      pasajeroBusquedaTexto: '',
      nuevoPasajeroDraft: { nombreApellido: '', dni: '', pasaporte: '', telefono: '', email: '', observaciones: '', fechaNacimiento: '' }
    });
  };

  private _resetEditorPasajero = (): IPasajeroDraft => ({
    nombreApellido: '',
    dni: '',
    pasaporte: '',
    telefono: '',
    email: '',
    observaciones: '',
    fechaNacimiento: ''
  });

  private _guardarPasajeroDraft = async (): Promise<void> => {
    const nombre = (this.state.nuevoPasajeroDraft.nombreApellido || '').trim();
    const dni = (this.state.nuevoPasajeroDraft.dni || '').trim();
    if (!nombre) {
      return;
    }
    if (!dni) {
      this._setSectionError('pasajeros', 'Indica el DNI del pasajero.');
      return;
    }
    const dniNormalizado = dni.toLowerCase();
    const editId = this.state.pasajeroEnEdicionId;
    const existeDni = this.state.pasajeros.some(
      (p: IPasajero) => p.id !== editId && ((p.dni || '').trim().toLowerCase() === dniNormalizado)
    );
    if (existeDni) {
      this._setSectionError('pasajeros', 'Ya existe un pasajero con ese DNI. Usa el buscador de pasajeros para agregarlo.');
      return;
    }

    const draftData: IPasajeroDraft = {
      ...this.state.nuevoPasajeroDraft,
      nombreApellido: nombre,
      dni,
      pasaporte: this.state.nuevoPasajeroDraft.pasaporte || '',
      telefono: this.state.nuevoPasajeroDraft.telefono || '',
      email: this.state.nuevoPasajeroDraft.email || '',
      observaciones: this.state.nuevoPasajeroDraft.observaciones || '',
      fechaNacimiento: this.state.nuevoPasajeroDraft.fechaNacimiento || ''
    };

    if (editId !== null) {
      const pasajeroActualizado: IPasajero = {
        id: editId,
        nombreApellido: draftData.nombreApellido,
        dni: draftData.dni,
        pasaporte: draftData.pasaporte || '',
        telefono: draftData.telefono || '',
        email: draftData.email || '',
        observaciones: draftData.observaciones || '',
        fechaNacimiento: draftData.fechaNacimiento || ''
      };

      this.setState({ guardando: true });
      try {
        if (editId > 0) {
          await this._service.updatePasajero(editId, draftData);
        }
        this.setState(prev => ({
          guardando: false,
          sectionErrors: { ...prev.sectionErrors, pasajeros: '' },
          pasajeros: prev.pasajeros.map((p: IPasajero) => (p.id === editId ? pasajeroActualizado : p)),
          nuevosPasajeros: editId < 0
            ? prev.nuevosPasajeros.map((d: IPasajeroPendiente) =>
              d.tempId === editId ? { ...draftData, tempId: editId } : d
            )
            : prev.nuevosPasajeros,
          nuevoPasajeroDraft: this._resetEditorPasajero(),
          mostrarEditorPasajero: false,
          pasajeroEnEdicionId: null,
          pasajeroCamposHabilitados: false
        }));
      } catch (error) {
        this.setState({
          guardando: false,
          sectionErrors: { ...this.state.sectionErrors, pasajeros: (error as Error).message || 'No se pudo actualizar el pasajero.' }
        });
      }
      return;
    }

    const tempId = this._nextTempPasajeroId();
    const draft: IPasajeroPendiente = {
      ...draftData,
      tempId
    };
    const pasajeroTemporal: IPasajero = {
      id: tempId,
      nombreApellido: draft.nombreApellido,
      dni: draft.dni,
      pasaporte: draft.pasaporte || '',
      telefono: draft.telefono || '',
      email: draft.email || '',
      observaciones: draft.observaciones || '',
      fechaNacimiento: draft.fechaNacimiento || ''
    };
    this.setState(prev => ({
      sectionErrors: { ...prev.sectionErrors, pasajeros: '' },
      nuevosPasajeros: prev.nuevosPasajeros.concat([draft]),
      pasajeros: prev.pasajeros.concat([pasajeroTemporal]),
      pasajerosIds: this._uniqueNumbers(prev.pasajerosIds.concat([tempId])),
      nuevoPasajeroDraft: this._resetEditorPasajero(),
      mostrarEditorPasajero: false,
      pasajeroEnEdicionId: null,
      pasajeroCamposHabilitados: false
    }));
  };

  private _cancelarPasajero = (): void => {
    this.setState({
      mostrarEditorPasajero: false,
      pasajeroEnEdicionId: null,
      pasajeroBusquedaTexto: '',
      pasajeroCamposHabilitados: false,
      nuevoPasajeroDraft: this._resetEditorPasajero()
    });
  };

  private _quitarPasajeroSeleccionado = (pasajeroId: number): void => {
    this.setState(prev => {
      const documentosPorId = { ...prev.pasajeroDocumentosPorId };
      delete documentosPorId[pasajeroId];
      return {
        pasajerosIds: prev.pasajerosIds.filter((id: number) => id !== pasajeroId),
        pasajeros: pasajeroId < 0 ? prev.pasajeros.filter((p: IPasajero) => p.id !== pasajeroId) : prev.pasajeros,
        nuevosPasajeros: pasajeroId < 0 ? prev.nuevosPasajeros.filter((p: IPasajeroPendiente) => p.tempId !== pasajeroId) : prev.nuevosPasajeros,
        pasajeroDocumentosExpandidoId:
          prev.pasajeroDocumentosExpandidoId === pasajeroId ? null : prev.pasajeroDocumentosExpandidoId,
        pasajeroDocumentosPorId: documentosPorId,
        pasajeroDocumentosCargandoId:
          prev.pasajeroDocumentosCargandoId === pasajeroId ? null : prev.pasajeroDocumentosCargandoId
      };
    });
  };

  private _esPasajeroPersistidoEnSharePoint(pasajeroId: number): boolean {
    return pasajeroId > 0;
  }

  private _getDocumentosPasajeroCache(pasajeroId: number): IPasajeroDocumentoItem[] {
    return this.state.pasajeroDocumentosPorId[pasajeroId] || [];
  }

  private async _cargarDocumentosPasajero(pasajeroId: number, forzarRecarga?: boolean): Promise<void> {
    if (!this._esPasajeroPersistidoEnSharePoint(pasajeroId)) {
      return;
    }
    if (!forzarRecarga && this.state.pasajeroDocumentosPorId[pasajeroId]) {
      return;
    }
    this.setState({ pasajeroDocumentosCargandoId: pasajeroId });
    try {
      const documentos = await this._service.getDocumentosPasajero(pasajeroId);
      this.setState(prev => ({
        pasajeroDocumentosPorId: { ...prev.pasajeroDocumentosPorId, [pasajeroId]: documentos },
        pasajeroDocumentosCargandoId:
          prev.pasajeroDocumentosCargandoId === pasajeroId ? null : prev.pasajeroDocumentosCargandoId
      }));
    } catch (error) {
      this.setState(prev => ({
        pasajeroDocumentosCargandoId:
          prev.pasajeroDocumentosCargandoId === pasajeroId ? null : prev.pasajeroDocumentosCargandoId,
        sectionErrors: {
          ...prev.sectionErrors,
          pasajeros: 'No se pudieron cargar los documentos del pasajero.'
        }
      }));
    }
  }

  private _toggleDocumentosPasajero = (pasajeroId: number): void => {
    if (this.state.pasajeroDocumentosExpandidoId === pasajeroId) {
      this.setState({ pasajeroDocumentosExpandidoId: null });
      return;
    }
    this.setState({ pasajeroDocumentosExpandidoId: pasajeroId });
    if (this._esPasajeroPersistidoEnSharePoint(pasajeroId)) {
      void this._cargarDocumentosPasajero(pasajeroId);
    }
  };

  private _abrirDocumentoPasajero = (serverRelativeUrl: string): void => {
    if (!serverRelativeUrl) {
      return;
    }
    window.open(serverRelativeUrl, '_blank');
  };

  private _onSeleccionarArchivosDocumentosPasajero = (pasajeroId: number, ev: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura() || !this._esPasajeroPersistidoEnSharePoint(pasajeroId)) {
      return;
    }
    const archivos =
      ev.target.files && ev.target.files.length > 0
        ? (Array.prototype.slice.call(ev.target.files) as File[])
        : [];
    ev.target.value = '';
    if (archivos.length === 0) {
      return;
    }
    void this._subirDocumentosPasajero(pasajeroId, archivos);
  };

  private _subirDocumentosPasajero = async (pasajeroId: number, archivos: File[]): Promise<void> => {
    if (this._esSoloLectura() || !this._esPasajeroPersistidoEnSharePoint(pasajeroId) || archivos.length === 0) {
      return;
    }
    this.setState({ guardando: true });
    this._clearSectionError('pasajeros');
    const fallidos: string[] = [];
    let subidos = 0;
    for (const archivo of archivos) {
      try {
        await this._service.uploadDocumentoPasajero(pasajeroId, archivo);
        subidos++;
      } catch (error) {
        const detalle = error instanceof Error && error.message ? error.message : '';
        fallidos.push(detalle ? archivo.name + ' (' + detalle + ')' : archivo.name);
      }
    }
    await this._cargarDocumentosPasajero(pasajeroId, true);
    if (fallidos.length > 0) {
      const mensaje =
        subidos > 0
          ? 'Se adjuntaron ' + subidos + ' de ' + archivos.length + ' documento(s). No se pudieron guardar: ' + fallidos.join(', ') + '.'
          : 'No se pudieron adjuntar los documentos: ' + fallidos.join(', ') + '.';
      this._setSectionError('pasajeros', mensaje);
    }
    this.setState({ guardando: false });
  };

  private _eliminarDocumentoPasajero = async (pasajeroId: number, fileName: string): Promise<void> => {
    if (this._esSoloLectura() || !this._esPasajeroPersistidoEnSharePoint(pasajeroId)) {
      return;
    }
    try {
      this.setState({ guardando: true });
      this._clearSectionError('pasajeros');
      await this._service.deleteDocumentoPasajero(pasajeroId, fileName);
      await this._cargarDocumentosPasajero(pasajeroId, true);
    } catch (error) {
      this._setSectionError('pasajeros', 'No se pudo eliminar el documento del pasajero.');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _renderDocumentosPasajeroExpandido(p: IPasajero, soloLectura: boolean): React.ReactNode {
    const expandido = this.state.pasajeroDocumentosExpandidoId === p.id;
    if (!expandido) {
      return null;
    }
    const columnas = soloLectura ? 4 : 5;
    if (!this._esPasajeroPersistidoEnSharePoint(p.id)) {
      return (
        <tr key={'docs-' + p.id}>
          <td colSpan={columnas} style={pasajerosStyles.documentosExpandCell}>
            <div style={pasajerosStyles.documentosZone}>
              <div style={pasajerosStyles.documentosInfo}>
                Guardá primero el pasajero para poder adjuntar documentos.
              </div>
            </div>
          </td>
        </tr>
      );
    }
    const cargando = this.state.pasajeroDocumentosCargandoId === p.id;
    const documentos = this._getDocumentosPasajeroCache(p.id);
    return (
      <tr key={'docs-' + p.id}>
        <td colSpan={columnas} style={pasajerosStyles.documentosExpandCell}>
          <div style={{ fontSize: 12, fontWeight: 600, color: '#605e5c', marginBottom: 8 }}>Documentos del pasajero</div>
          <div style={pasajerosStyles.documentosZone}>
            {cargando ? (
              <div style={pasajerosStyles.documentosEmpty}>Cargando documentos...</div>
            ) : documentos.length === 0 ? (
              <div style={pasajerosStyles.documentosEmpty}>No hay documentos adjuntos.</div>
            ) : (
              documentos.map((doc: IPasajeroDocumentoItem, index: number) => {
                const esUltimo = index === documentos.length - 1;
                return (
                  <div
                    key={doc.fileName}
                    style={esUltimo ? { ...pasajerosStyles.documentosRow, ...pasajerosStyles.documentosRowLast } : pasajerosStyles.documentosRow}
                  >
                    <GridIconPaperclip />
                    <div style={pasajerosStyles.documentosFileName}>{doc.fileName}</div>
                    <div style={gridActionBarStyle}>
                      <GridIconActionButton
                        title="Abrir / Descargar"
                        onClick={() => this._abrirDocumentoPasajero(doc.serverRelativeUrl)}
                        disabled={this.state.guardando || !doc.serverRelativeUrl}
                      >
                        <GridIconDownload />
                      </GridIconActionButton>
                      {!soloLectura && (
                        <GridIconActionButton
                          title="Eliminar"
                          onClick={() => { void this._eliminarDocumentoPasajero(p.id, doc.fileName); }}
                          disabled={this.state.guardando}
                        >
                          <GridIconTrash />
                        </GridIconActionButton>
                      )}
                    </div>
                  </div>
                );
              })
            )}
          </div>
          {!soloLectura && (
            <div style={pasajerosStyles.documentosUploadRow}>
              <span style={{ ...layoutStyles.label, marginBottom: 0 }}>Adjuntar archivos</span>
              <FileInputEspanol
                inputId={'pasajero-documentos-' + p.id}
                multiple={true}
                disabled={this.state.guardando}
                buttonLabel="Seleccionar archivos"
                hint="Podés elegir varios archivos."
                onChange={ev => this._onSeleccionarArchivosDocumentosPasajero(p.id, ev)}
                buttonStyle={{ borderRadius: pasajerosRadius.control }}
              />
            </div>
          )}
        </td>
      </tr>
    );
  };

  private _abrirEditorMovimiento = (): void => {
    if (this._esSoloLectura()) { return; }
    this._clearSectionError('movimientos');
    this.setState({
      mostrarEditorMovimiento: true,
      movimientoEnEdicionId: null,
      movimientoEnEdicion: { tipo: 'Ingreso', movimiento: '', medioPago: '', fecha: this._getTodayDateInput(), moneda: 'Pesos', monto: '', observaciones: '', cotizacion: '', liquidacionOperadorId: '', servicioAsociadoId: '' }
    });
  };

  private _onSeleccionarServicioIngreso = (servicioIdRaw: string): void => {
    const servicioId = Number(servicioIdRaw);
    const servicio = this.state.servicios.filter((s: IServicio) => s.id === servicioId)[0];
    this._clearSectionError('movimientos');
    this.setState(prev => ({
      movimientoEnEdicion: {
        ...prev.movimientoEnEdicion,
        servicioAsociadoId: servicioIdRaw,
        movimiento: servicio ? servicio.concepto : ''
      }
    }));
  };

  private _onSeleccionarLiquidacionEgreso = (liquidacionIdRaw: string): void => {
    const liquidacionId = Number(liquidacionIdRaw);
    const liquidacion = this.state.liquidacionesOperador.filter((l: ILiquidacionItem) => l.id === liquidacionId)[0];
    this.setState(prev => ({
      movimientoEnEdicion: {
        ...prev.movimientoEnEdicion,
        liquidacionOperadorId: liquidacionIdRaw,
        movimiento: liquidacion ? liquidacion.codigoReferencia : ''
      }
    }));
  };

  private _getMonedaRelacionadaMovimiento(): string {
    if (this.state.movimientoEnEdicion.tipo === 'Ingreso') {
      const servicioId = Number(this.state.movimientoEnEdicion.servicioAsociadoId);
      const servicio = this.state.servicios.filter((s: IServicio) => s.id === servicioId)[0];
      return servicio ? this._normalizarMonedaServicio(servicio.moneda) : '';
    }
    const liquidacionId = Number(this.state.movimientoEnEdicion.liquidacionOperadorId);
    const liquidacion = this.state.liquidacionesOperador.filter((l: ILiquidacionItem) => l.id === liquidacionId)[0];
    return liquidacion ? this._normalizarMonedaLiquidacion(liquidacion.moneda) : '';
  }

  private _requiereCotizacionMovimiento(): boolean {
    const monedaRelacionada = this._getMonedaRelacionadaMovimiento();
    if (!monedaRelacionada) {
      return false;
    }
    return monedaRelacionada !== this._normalizarMonedaPago(this.state.movimientoEnEdicion.moneda);
  }

  private _abrirEditorServicio = (): void => {
    if (this._esSoloLectura()) { return; }
    this._clearSectionError('servicios');
    this.setState({
      mostrarEditorServicio: true,
      servicioEnEdicionId: null,
      servicioEnEdicion: { concepto: '', precioCliente: '', moneda: 'Pesos', operadorId: '' }
    });
  };

  private _actualizarCampoServicio = (campo: keyof IProcopioFormsState['servicioEnEdicion'], valor: string): void => {
    this._clearSectionError('servicios');
    this.setState(prev => ({ servicioEnEdicion: { ...prev.servicioEnEdicion, [campo]: valor } }));
  };

  private _getOperadorNombrePorId(operadorId: number): string {
    if (operadorId <= 0) {
      return '';
    }
    const operador = this.state.operadores.filter((o: IOperadorItem) => o.id === operadorId)[0];
    return operador ? operador.titulo : '';
  }

  private _getServicioOperadorDisplay(servicio: IServicio): string {
    if (servicio.operadorNombre) {
      return servicio.operadorNombre;
    }
    if (servicio.operadorId) {
      return this._getOperadorNombrePorId(servicio.operadorId);
    }
    return '';
  }

  private _guardarServicioMock = (): void => {
    if (this._esSoloLectura()) { return; }
    const concepto = (this.state.servicioEnEdicion.concepto || '').trim();
    const operadorId = Number(this.state.servicioEnEdicion.operadorId);
    if (!concepto) {
      this._setSectionError('servicios', 'Selecciona un concepto para el servicio.');
      return;
    }
    if (operadorId <= 0) {
      this._setSectionError('servicios', 'Selecciona un operador para el servicio.');
      return;
    }
    this.setState(prev => {
      const esEdicion = prev.servicioEnEdicionId !== null;
      const servicio: IServicio = {
        id: esEdicion
          ? (prev.servicioEnEdicionId as number)
          : this._nextTempServicioId(prev.servicios),
        concepto,
        precioCliente: Number(prev.servicioEnEdicion.precioCliente) || 0,
        moneda: this._normalizarMonedaServicio(prev.servicioEnEdicion.moneda),
        operadorId,
        operadorNombre: (() => {
          const operadorItem = prev.operadores.filter((o: IOperadorItem) => o.id === operadorId)[0];
          return operadorItem ? operadorItem.titulo : '';
        })()
      };

      const servicios = esEdicion
        ? prev.servicios.map((s: IServicio) => (s.id === prev.servicioEnEdicionId ? servicio : s))
        : prev.servicios.concat([servicio]);

      return {
        servicios,
        sectionErrors: { ...prev.sectionErrors, servicios: '' },
        mostrarEditorServicio: false,
        servicioEnEdicionId: null,
        servicioEnEdicion: { concepto: '', precioCliente: '', moneda: 'Pesos', operadorId: '' }
      };
    });
  };

  private _cancelarServicioMock = (): void => {
    this.setState({
      mostrarEditorServicio: false,
      servicioEnEdicionId: null,
      servicioEnEdicion: { concepto: '', precioCliente: '', moneda: 'Pesos', operadorId: '' }
    });
  };

  private _editarServicio = (id: number): void => {
    if (this._esSoloLectura()) { return; }
    this.setState(prev => {
      const servicio = prev.servicios.filter((s: IServicio) => s.id === id)[0];
      if (!servicio) { return prev; }
      return {
        ...prev,
        mostrarEditorServicio: true,
        servicioEnEdicionId: id,
        servicioEnEdicion: {
          concepto: servicio.concepto,
          precioCliente: String(servicio.precioCliente),
          moneda: this._normalizarMonedaServicio(servicio.moneda),
          operadorId: servicio.operadorId ? String(servicio.operadorId) : ''
        }
      };
    });
  };

  private _eliminarServicio = (id: number): void => {
    if (this._esSoloLectura()) { return; }
    this.setState(prev => ({
      servicios: prev.servicios.filter((s: IServicio) => s.id !== id),
      serviciosEliminadosIds: id > 0 ? prev.serviciosEliminadosIds.concat([id]) : prev.serviciosEliminadosIds
    }));
  };

  private _actualizarCampoMovimiento = (
    campo: Exclude<keyof IProcopioFormsState['movimientoEnEdicion'], 'tipo'>,
    valor: string
  ): void => {
    this._clearSectionError('movimientos');
    this.setState(prev => ({ movimientoEnEdicion: { ...prev.movimientoEnEdicion, [campo]: valor } }));
  };

  private _onCambiarTipoMovimiento = (ev: React.ChangeEvent<HTMLSelectElement>): void => {
    const tipo = ev.target.value === 'Egreso' ? 'Egreso' : 'Ingreso';
    this.setState(prev => ({
      movimientoEnEdicion: {
        ...prev.movimientoEnEdicion,
        tipo,
        movimiento: '',
        liquidacionOperadorId: '',
        servicioAsociadoId: ''
      }
    }));
  };

  private _editarMovimiento = (id: number): void => {
    if (this._esSoloLectura()) { return; }
    this.setState(prev => {
      const mov = prev.movimientos.filter((m: IMovimiento) => m.id === id)[0];
      if (!mov) { return prev; }
      return {
        ...prev,
        mostrarEditorMovimiento: true,
        movimientoEnEdicionId: id,
        movimientoEnEdicion: {
          tipo: mov.tipo,
          movimiento: mov.movimiento,
          medioPago: mov.medioPago,
          fecha: mov.fecha,
          moneda: this._normalizarMonedaPago(mov.moneda),
          monto: String(mov.monto),
          observaciones: mov.observaciones || '',
          cotizacion: mov.cotizacion !== undefined ? String(mov.cotizacion) : '',
          liquidacionOperadorId: mov.liquidacionOperadorId ? String(mov.liquidacionOperadorId) : '',
          servicioAsociadoId: mov.servicioAsociadoId ? String(mov.servicioAsociadoId) : ''
        }
      };
    });
  };

  private _cancelarMovimiento = (): void => {
    this.setState({ mostrarEditorMovimiento: false, movimientoEnEdicionId: null });
  };

  private _abrirEditorVoucher = (): void => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    this._clearSectionError('vouchers');
    this.setState({ mostrarEditorVoucher: true, voucherEnEdicion: { archivos: [] } });
  };

  private async _refrescarVouchers(viajeId: number): Promise<void> {
    const vouchers = await this._service.getVouchersByViaje(viajeId);
    this.setState({ vouchers });
  }

  private _onSeleccionarArchivosVoucher = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const archivos =
      ev.target.files && ev.target.files.length > 0
        ? (Array.prototype.slice.call(ev.target.files) as File[])
        : [];
    ev.target.value = '';
    if (archivos.length === 0) {
      return;
    }
    void this._subirArchivosVoucher(archivos);
  };

  private _subirArchivosVoucher = async (archivos: File[]): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId || archivos.length === 0) { return; }
    const viajeId = this.state.viajeId;
    this.setState({ guardando: true, vouchersSubiendo: true });
    this._clearSectionError('vouchers');
    const fallidos: string[] = [];
    let subidos = 0;
    try {
      for (const archivo of archivos) {
        try {
          await this._service.uploadVoucher(viajeId, archivo);
          subidos++;
        } catch (error) {
          const detalle = error instanceof Error && error.message ? error.message : '';
          fallidos.push(detalle ? archivo.name + ' (' + detalle + ')' : archivo.name);
        }
      }
      await this._refrescarVouchers(viajeId);
      if (fallidos.length === 0) {
        this.setState({ voucherEnEdicion: { archivos: [] } });
        return;
      }
      const mensaje = subidos > 0
        ? 'Se adjuntaron ' + subidos + ' de ' + archivos.length + ' voucher(s). No se pudieron guardar: ' + fallidos.join(', ') + '.'
        : 'No se pudieron adjuntar los vouchers: ' + fallidos.join(', ') + '.';
      this.setState(prev => ({
        sectionErrors: { ...prev.sectionErrors, vouchers: mensaje }
      }));
    } finally {
      this.setState({ guardando: false, vouchersSubiendo: false });
    }
  };

  private _cancelarVoucher = (): void => {
    this.setState({ mostrarEditorVoucher: false, voucherEnEdicion: { archivos: [] } });
  };

  private _abrirVoucher = (url: string): void => {
    if (!url) { return; }
    window.open(url, '_blank');
  };

  private _eliminarVoucher = async (fileName: string): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const viajeId = this.state.viajeId;
    try {
      this.setState({ guardando: true });
      this._clearSectionError('vouchers');
      await this._service.deleteVoucher(viajeId, fileName);
      await this._refrescarVouchers(viajeId);
    } catch (error) {
      this._setSectionError('vouchers', 'No se pudo eliminar el voucher.');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _getPresupuestoNombreVisible(fileName: string): string {
    const prefix = 'Presupuesto_';
    return fileName.indexOf(prefix) === 0 ? fileName.substring(prefix.length) : fileName;
  }

  private _getFacturaNombreVisible(fileName: string): string {
    const prefix = 'Factura_';
    return fileName.indexOf(prefix) === 0 ? fileName.substring(prefix.length) : fileName;
  }

  private _abrirEditorFactura = (): void => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    this._clearSectionError('facturas');
    this.setState({ mostrarEditorFactura: true, facturaEnEdicion: { archivos: [] } });
  };

  private async _refrescarFacturas(viajeId: number): Promise<void> {
    const facturas = await this._service.getFacturasByViaje(viajeId);
    this.setState({ facturas });
  }

  private _onSeleccionarArchivosFactura = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const archivos =
      ev.target.files && ev.target.files.length > 0
        ? (Array.prototype.slice.call(ev.target.files) as File[])
        : [];
    ev.target.value = '';
    if (archivos.length === 0) {
      return;
    }
    void this._subirArchivosFactura(archivos);
  };

  private _subirArchivosFactura = async (archivos: File[]): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId || archivos.length === 0) { return; }
    const viajeId = this.state.viajeId;
    this.setState({ guardando: true, facturasSubiendo: true });
    this._clearSectionError('facturas');
    const fallidos: string[] = [];
    let subidos = 0;
    try {
      for (const archivo of archivos) {
        try {
          await this._service.uploadFactura(viajeId, archivo);
          subidos++;
        } catch (error) {
          const detalle = error instanceof Error && error.message ? error.message : '';
          fallidos.push(detalle ? archivo.name + ' (' + detalle + ')' : archivo.name);
        }
      }
      await this._refrescarFacturas(viajeId);
      if (fallidos.length === 0) {
        this.setState({ facturaEnEdicion: { archivos: [] } });
        return;
      }
      const mensaje = subidos > 0
        ? 'Se adjuntaron ' + subidos + ' de ' + archivos.length + ' factura(s). No se pudieron guardar: ' + fallidos.join(', ') + '.'
        : 'No se pudieron adjuntar las facturas: ' + fallidos.join(', ') + '.';
      this.setState(prev => ({
        sectionErrors: { ...prev.sectionErrors, facturas: mensaje }
      }));
    } finally {
      this.setState({ guardando: false, facturasSubiendo: false });
    }
  };

  private _cancelarFactura = (): void => {
    this.setState({ mostrarEditorFactura: false, facturaEnEdicion: { archivos: [] } });
  };

  private _abrirFactura = (url: string): void => {
    if (!url) { return; }
    window.open(url, '_blank');
  };

  private _eliminarFactura = async (fileName: string): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const viajeId = this.state.viajeId;
    try {
      this.setState({ guardando: true });
      this._clearSectionError('facturas');
      await this._service.deleteFactura(viajeId, fileName);
      await this._refrescarFacturas(viajeId);
    } catch (error) {
      this._setSectionError('facturas', 'No se pudo eliminar la factura.');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _abrirEditorPresupuesto = (): void => {
    if (this._esSoloLectura()) { return; }
    this._clearSectionError('presupuesto');
    this.setState({ mostrarEditorPresupuesto: true });
  };

  private async _refrescarPresupuesto(viajeId: number): Promise<void> {
    const presupuesto = await this._service.getPresupuestoByViaje(viajeId);
    this.setState({ presupuesto });
  }

  private _onSeleccionarArchivoPresupuesto = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    if (this._esSoloLectura()) { return; }
    const archivo = ev.target.files && ev.target.files.length > 0 ? ev.target.files[0] : null;
    ev.target.value = '';
    if (!archivo) {
      return;
    }
    void this._subirArchivoPresupuesto(archivo);
  };

  private _subirArchivoPresupuesto = async (archivo: File): Promise<void> => {
    if (this._esSoloLectura()) { return; }
    const viajeId = this.state.viajeId;
    const fileName = this._service.buildPresupuestoAttachmentFileName(archivo.name);

    if (!viajeId) {
      this._clearSectionError('presupuesto');
      this.setState({
        presupuesto: { fileName, serverRelativeUrl: '' },
        presupuestoArchivoPendiente: archivo
      });
      return;
    }

    this.setState({ guardando: true, presupuestoSubiendo: true });
    this._clearSectionError('presupuesto');
    try {
      const presupuesto = await this._service.uploadPresupuesto(viajeId, archivo);
      this.setState({ presupuesto, presupuestoArchivoPendiente: null });
    } catch (error) {
      this.setState(prev => ({
        sectionErrors: { ...prev.sectionErrors, presupuesto: 'No se pudo adjuntar el presupuesto.' }
      }));
    } finally {
      this.setState({ guardando: false, presupuestoSubiendo: false });
    }
  };

  private _cancelarPresupuesto = (): void => {
    this.setState({ mostrarEditorPresupuesto: false });
  };

  private _abrirPresupuesto = (url: string): void => {
    if (url) {
      window.open(url, '_blank');
      return;
    }
    const archivo = this.state.presupuestoArchivoPendiente;
    if (!archivo) { return; }
    const blobUrl = URL.createObjectURL(archivo);
    window.open(blobUrl, '_blank');
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
  };

  private _eliminarPresupuesto = async (): Promise<void> => {
    if (this._esSoloLectura() || !this.state.presupuesto) { return; }
    if (!this.state.viajeId) {
      this.setState({
        presupuesto: null,
        presupuestoArchivoPendiente: null,
        mostrarEditorPresupuesto: false
      });
      return;
    }
    const viajeId = this.state.viajeId;
    const fileName = this.state.presupuesto.fileName;
    try {
      this.setState({ guardando: true });
      this._clearSectionError('presupuesto');
      await this._service.deletePresupuesto(viajeId, fileName);
      await this._refrescarPresupuesto(viajeId);
    } catch (error) {
      this._setSectionError('presupuesto', 'No se pudo eliminar el presupuesto.');
    } finally {
      this.setState({
        guardando: false,
        mostrarEditorPresupuesto: false,
        presupuestoArchivoPendiente: null
      });
    }
  };

  private _abrirLiquidacionArchivo = (url: string): void => {
    if (!url) { return; }
    window.open(url, '_blank');
  };

  private _abrirEditorLiquidacion = (): void => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    this._clearSectionError('liquidaciones');
    this.setState({
      mostrarEditorLiquidacion: true,
      liquidacionEnEdicionId: null,
      liquidacionEnEdicion: { codigoReferencia: '', operadorId: '', monto: '', moneda: 'Pesos', archivoFile: null }
    });
  };

  private _editarLiquidacion = (id: number): void => {
    if (this._esSoloLectura()) { return; }
    this.setState(prev => {
      const liquidacion = prev.liquidacionesOperador.filter((l: ILiquidacionItem) => l.id === id)[0];
      if (!liquidacion) { return prev; }
      return {
        ...prev,
        mostrarEditorLiquidacion: true,
        liquidacionEnEdicionId: id,
        liquidacionEnEdicion: {
          codigoReferencia: liquidacion.codigoReferencia,
          operadorId: liquidacion.operadorId ? String(liquidacion.operadorId) : '',
          monto: String(liquidacion.monto),
          moneda: this._normalizarMonedaLiquidacion(liquidacion.moneda),
          archivoFile: null
        }
      };
    });
  };

  private _actualizarCampoLiquidacion = (campo: 'codigoReferencia' | 'operadorId' | 'monto' | 'moneda', valor: string): void => {
    const valorNormalizado = campo === 'moneda' ? this._normalizarMonedaLiquidacion(valor) : valor;
    this._clearSectionError('liquidaciones');
    this.setState(prev => ({
      liquidacionEnEdicion: {
        ...prev.liquidacionEnEdicion,
        [campo]: valorNormalizado
      }
    }));
  };

  private _onArchivoLiquidacionChange = (ev: React.ChangeEvent<HTMLInputElement>): void => {
    const file = ev.target.files && ev.target.files.length > 0 ? ev.target.files[0] : null;
    this.setState(prev => ({ liquidacionEnEdicion: { ...prev.liquidacionEnEdicion, archivoFile: file } }));
  };

  private async _refrescarLiquidaciones(viajeId: number): Promise<void> {
    const liquidaciones = await this._service.getLiquidacionesByViaje(viajeId);
    this.setState({ liquidacionesOperador: liquidaciones });
  }

  private _guardarLiquidacion = async (): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const viajeId = this.state.viajeId;
    const codigo = this.state.liquidacionEnEdicion.codigoReferencia.trim();
    if (!codigo) {
      this._setSectionError('liquidaciones', 'Indica el código de referencia de la liquidación.');
      return;
    }
    const operadorId = Number(this.state.liquidacionEnEdicion.operadorId);
    if (operadorId <= 0) {
      this._setSectionError('liquidaciones', 'Selecciona un operador para la liquidación.');
      return;
    }
    const monto = Number(this.state.liquidacionEnEdicion.monto);
    if (isNaN(monto)) {
      this._setSectionError('liquidaciones', 'El monto de la liquidación no es válido.');
      return;
    }
    try {
      this.setState({ guardando: true });
      this._clearSectionError('liquidaciones');
      const data: ILiquidacionData = {
        viajeId,
        codigoReferencia: codigo,
        operadorId,
        monto,
        moneda: this._normalizarMonedaLiquidacion(this.state.liquidacionEnEdicion.moneda),
        file: this.state.liquidacionEnEdicion.archivoFile || undefined
      };
      if (this.state.liquidacionEnEdicionId) {
        await this._service.updateLiquidacion(this.state.liquidacionEnEdicionId, data);
      } else {
        await this._service.createLiquidacion(data);
      }
      await this._refrescarLiquidaciones(viajeId);
      this.setState({
        mostrarEditorLiquidacion: false,
        liquidacionEnEdicion: { codigoReferencia: '', operadorId: '', monto: '', moneda: 'Pesos', archivoFile: null },
        liquidacionEnEdicionId: null,
        guardando: false
      });
    } catch (error) {
      this.setState({ guardando: false });
      this._setSectionError('liquidaciones', 'No se pudo guardar la liquidación en SharePoint.');
    }
  };

  private _cancelarLiquidacion = (): void => {
    this.setState({
      mostrarEditorLiquidacion: false,
      liquidacionEnEdicion: { codigoReferencia: '', operadorId: '', monto: '', moneda: 'Pesos', archivoFile: null },
      liquidacionEnEdicionId: null
    });
  };

  private _eliminarLiquidacion = async (id: number): Promise<void> => {
    if (this._esSoloLectura() || !this.state.viajeId) { return; }
    const viajeId = this.state.viajeId;
    try {
      this.setState({ guardando: true });
      this._clearSectionError('liquidaciones');
      await this._service.deleteLiquidacion(id);
      await this._refrescarLiquidaciones(viajeId);
    } catch (error) {
      this._setSectionError('liquidaciones', 'No se pudo eliminar la liquidación.');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _getTotalesServiciosPorMoneda(servicios: IServicio[] = this.state.servicios): { usd: number; ars: number } {
    return servicios.reduce(
      (acc: { usd: number; ars: number }, servicio: IServicio) => {
        const monto = Number(servicio.precioCliente) || 0;
        if (this._normalizarMonedaServicio(servicio.moneda) === 'Dólares') {
          acc.usd += monto;
        } else {
          acc.ars += monto;
        }
        return acc;
      },
      { usd: 0, ars: 0 }
    );
  }

  /**
   * Convierte un ingreso a la moneda del servicio asociado.
   * Cotizacion se interpreta como ARS por 1 USD.
   */
  private _convertirIngresoAMonedaServicio(movimiento: IMovimiento, servicio: IServicio): number {
    const monto = Number(movimiento.monto) || 0;
    if (monto <= 0) {
      return 0;
    }

    const monedaPago = this._normalizarMonedaPago(movimiento.moneda);
    const monedaServicio = this._normalizarMonedaServicio(servicio.moneda);
    if (monedaPago === monedaServicio) {
      return monto;
    }

    const cotizacion = Number(movimiento.cotizacion);
    if (!isFinite(cotizacion) || cotizacion <= 0) {
      return 0;
    }

    if (monedaPago === 'Pesos' && monedaServicio === 'Dólares') {
      return monto / cotizacion;
    }
    if (monedaPago === 'Dólares' && monedaServicio === 'Pesos') {
      return monto * cotizacion;
    }

    return monto;
  }

  private _getTotalesIngresosPorMoneda(
    movimientos: IMovimiento[] = this.state.movimientos,
    servicios: IServicio[] = this.state.servicios
  ): { usd: number; ars: number } {
    const serviciosById: { [id: number]: IServicio } = {};
    servicios.forEach((servicio: IServicio) => {
      serviciosById[servicio.id] = servicio;
    });

    return movimientos
      .filter((m: IMovimiento) => m.tipo === 'Ingreso' && m.monto > 0 && !!m.servicioAsociadoId)
      .reduce(
        (acc: { usd: number; ars: number }, m: IMovimiento) => {
          const servicio = serviciosById[m.servicioAsociadoId as number];
          if (!servicio) {
            return acc;
          }

          const montoConvertido = this._convertirIngresoAMonedaServicio(m, servicio);
          if (this._normalizarMonedaServicio(servicio.moneda) === 'Dólares') {
            acc.usd += montoConvertido;
          } else {
            acc.ars += montoConvertido;
          }
          return acc;
        },
        { usd: 0, ars: 0 }
      );
  }

  private _getTotalIngresosPorServicioEnMonedaServicio(servicio: IServicio, movimientoIdAExcluir?: number): number {
    return this.state.movimientos
      .filter(
        (m: IMovimiento) =>
          m.tipo === 'Ingreso' &&
          m.monto > 0 &&
          m.servicioAsociadoId === servicio.id &&
          (movimientoIdAExcluir === undefined || m.id !== movimientoIdAExcluir)
      )
      .reduce(
        (acc: number, m: IMovimiento) => acc + this._convertirIngresoAMonedaServicio(m, servicio),
        0
      );
  }

  private _renderResumenMonedas(totales: { usd: number; ars: number }): React.ReactNode {
    const rowStyle: React.CSSProperties = {
      display: 'flex',
      alignItems: 'baseline',
      justifyContent: 'space-between',
      gap: 8
    };
    const labelStyle: React.CSSProperties = {
      fontSize: 11,
      fontWeight: 600,
      color: '#808493',
      letterSpacing: '0.04em',
      textTransform: 'uppercase'
    };
    const valueStyle: React.CSSProperties = {
      fontSize: 20,
      fontWeight: 700,
      color: '#232529',
      lineHeight: 1.15
    };

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        <div style={rowStyle}>
          <span style={labelStyle}>USD</span>
          <span style={valueStyle}>{totales.usd.toLocaleString('en-US', { style: 'currency', currency: 'USD' })}</span>
        </div>
        <div style={rowStyle}>
          <span style={labelStyle}>ARS</span>
          <span style={valueStyle}>{totales.ars.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}</span>
        </div>
      </div>
    );
  }

  private _getConceptosServiciosDerivados(servicios: IServicio[]): string[] {
    const conceptosUnicos: string[] = [];
    const vistos: { [key: string]: boolean } = {};
    servicios.forEach((servicio: IServicio) => {
      const concepto = (servicio.concepto || '').trim();
      if (!concepto) {
        return;
      }
      if (!vistos[concepto]) {
        vistos[concepto] = true;
        conceptosUnicos.push(concepto);
      }
    });
    return conceptosUnicos;
  }

  private _esEstadoValido(estado: string): estado is EstadoViaje {
    return ESTADOS_VIAJE.indexOf(estado as EstadoViaje) >= 0;
  }

  private _normalizarEstado(estado: string): EstadoViaje {
    return this._esEstadoValido(estado) ? estado : 'Activo';
  }

  private _esMonedaServicioValida(moneda: string): moneda is MonedaServicio {
    return MONEDAS_SERVICIO.indexOf(moneda as MonedaServicio) >= 0;
  }

  private _normalizarMonedaServicio(moneda: string): MonedaServicio {
    return this._esMonedaServicioValida(moneda) ? moneda : 'Pesos';
  }

  private _esMonedaLiquidacionValida(moneda: string): moneda is MonedaLiquidacion {
    return MONEDAS_LIQUIDACION.indexOf(moneda as MonedaLiquidacion) >= 0;
  }

  private _normalizarMonedaLiquidacion(moneda: string): MonedaLiquidacion {
    return this._esMonedaLiquidacionValida(moneda) ? moneda : 'Pesos';
  }

  private _esMonedaPagoValida(moneda: string): moneda is MonedaPago {
    return MONEDAS_PAGO.indexOf(moneda as MonedaPago) >= 0;
  }

  private _normalizarMonedaPago(moneda: string): MonedaPago {
    return this._esMonedaPagoValida(moneda) ? moneda : 'Pesos';
  }

  private _getEtiquetaMonedaPago(moneda: string): 'USD' | 'ARS' {
    return this._normalizarMonedaPago(moneda) === 'Dólares' ? 'USD' : 'ARS';
  }

  /** Texto de Concepto solo para la grilla de pagos (egresos con prefijo fijo). */
  private _getConceptoMostradoGrillaMovimiento(m: IMovimiento): string {
    if (m.tipo !== 'Egreso') {
      return m.movimiento || '—';
    }
    const ref =
      (m.liquidacionOperadorNombre && m.liquidacionOperadorNombre.trim()) ||
      (m.liquidacionOperadorId
        ? (this.state.liquidacionesOperador.filter((l: ILiquidacionItem) => l.id === m.liquidacionOperadorId)[0]?.codigoReferencia || '').trim()
        : '') ||
      (m.movimiento || '').trim();
    return ref ? 'Liquidacion Operador - ' + ref : 'Liquidacion Operador';
  }

  private _mapViajeData(pasajerosIds: number[], serviciosFuente: IServicio[] = this.state.servicios): IViajeData {
    const estado = this._normalizarEstado(this.state.estado);
    return {
      nombre: this.state.nombreViaje,
      destinoGeneralId: this.state.destinoGeneralId > 0 ? this.state.destinoGeneralId : undefined,
      destinoId: this.state.destinoId,
      fechaSalida: this.state.fechaSalida,
      fechaLlegada: this.state.fechaRegreso,
      estado,
      colorViaje: this.state.colorViaje,
      pasajerosIds,
      servicios: this._getConceptosServiciosDerivados(serviciosFuente),
      observaciones: this.state.observaciones
    };
  }

  private async _resolverPasajerosIds(): Promise<number[]> {
    if (this.state.nuevosPasajeros.length === 0) {
      return this._uniqueNumbers(this.state.pasajerosIds.filter((id: number) => id > 0));
    }

    const created = await Promise.all(
      this.state.nuevosPasajeros.map((draft: IPasajeroPendiente) =>
        this._service.createPasajero({
          nombreApellido: draft.nombreApellido,
          dni: draft.dni,
          pasaporte: draft.pasaporte,
          telefono: draft.telefono,
          email: draft.email,
          observaciones: draft.observaciones,
          fechaNacimiento: draft.fechaNacimiento
        })
      )
    );

    const tempToReal: { [tempId: number]: IPasajeroItem } = {};
    this.state.nuevosPasajeros.forEach((draft: IPasajeroPendiente, index: number) => {
      tempToReal[draft.tempId] = created[index];
    });

    const pasajerosActualizados = this.state.pasajeros.map((p: IPasajero) => tempToReal[p.id] || p);
    const pasajerosIdsResueltos = this._uniqueNumbers(
      this.state.pasajerosIds
        .map((id: number) => (tempToReal[id] ? tempToReal[id].id : id))
        .filter((id: number) => id > 0)
    );

    this.setState({ pasajeros: pasajerosActualizados, pasajerosIds: pasajerosIdsResueltos, nuevosPasajeros: [] });
    return pasajerosIdsResueltos;
  }

  private _nextTempPasajeroId(): number {
    const used = this.state.pasajeros.map((p: IPasajero) => p.id);
    let next = -1;
    while (used.indexOf(next) >= 0) {
      next -= 1;
    }
    return next;
  }

  private _nextTempServicioId(servicios: IServicio[]): number {
    const usados = servicios.map((servicio: IServicio) => servicio.id);
    let next = -1;
    while (usados.indexOf(next) >= 0) {
      next -= 1;
    }
    return next;
  }

  private _mapServicioToServicioViajeData(viajeId: number, servicio: IServicio): IServicioViajeData {
    const operadorId = servicio.operadorId || 0;
    return {
      viajeId,
      concepto: servicio.concepto,
      precioCliente: Number(servicio.precioCliente) || 0,
      moneda: this._normalizarMonedaServicio(servicio.moneda),
      operadorId,
      operadorNombre: servicio.operadorNombre || this._getOperadorNombrePorId(operadorId)
    };
  }

  private async _sincronizarServiciosViaje(viajeId: number): Promise<IServicio[]> {
    const eliminados = this._uniqueNumbers(this.state.serviciosEliminadosIds.filter((id: number) => id > 0));
    const serviciosActuales = this.state.servicios;

    if (eliminados.length > 0) {
      await Promise.all(eliminados.map((id: number) => this._service.deleteServicioViaje(id)));
    }

    await Promise.all(serviciosActuales.map(async (servicio: IServicio) => {
      const data = this._mapServicioToServicioViajeData(viajeId, servicio);
      if (servicio.id > 0) {
        await this._service.updateServicioViaje(servicio.id, data);
        return;
      }
      await this._service.createServicioViaje(data);
    }));

    const serviciosViaje = await this._service.getServiciosViajeByViaje(viajeId);
    const serviciosSincronizados = serviciosViaje.map((servicio: IServicioViajeItem) => ({
      id: servicio.id,
      concepto: servicio.concepto,
      precioCliente: servicio.precioCliente,
      moneda: this._normalizarMonedaServicio(servicio.moneda),
      operadorId: servicio.operadorId,
      operadorNombre: servicio.operadorNombre
    }));
    this.setState({
      servicios: serviciosSincronizados,
      serviciosEliminadosIds: []
    });
    return serviciosSincronizados;
  }

  private _uniqueNumbers(values: number[]): number[] {
    const map: { [key: number]: boolean } = {};
    const result: number[] = [];
    values.forEach((value: number) => {
      if (!map[value]) {
        map[value] = true;
        result.push(value);
      }
    });
    return result;
  }

  private async _guardarPagoEnSharePoint(viajeId: number): Promise<void> {
    const esEgreso = this.state.movimientoEnEdicion.tipo === 'Egreso';
    const esIngreso = this.state.movimientoEnEdicion.tipo === 'Ingreso';
    const selLiq = Number(this.state.movimientoEnEdicion.liquidacionOperadorId);
    const selServicio = Number(this.state.movimientoEnEdicion.servicioAsociadoId);
    const esEdicion = !!this.state.movimientoEnEdicionId;

    let liquidacionOperadorId: number | undefined;
    if (esEgreso) {
      liquidacionOperadorId = selLiq > 0 ? selLiq : (esEdicion ? 0 : undefined);
    } else {
      liquidacionOperadorId = esEdicion ? 0 : undefined;
    }

    let servicioAsociadoId: number | undefined;
    if (esIngreso) {
      servicioAsociadoId = selServicio > 0 ? selServicio : (esEdicion ? 0 : undefined);
    } else {
      servicioAsociadoId = esEdicion ? 0 : undefined;
    }

    const pagoData: IPagoData = {
      viajeId,
      concepto: this.state.movimientoEnEdicion.movimiento,
      fechaPago: this.state.movimientoEnEdicion.fecha,
      importe: Number(this.state.movimientoEnEdicion.monto) || 0,
      medioPago: this.state.movimientoEnEdicion.medioPago,
      moneda: this._normalizarMonedaPago(this.state.movimientoEnEdicion.moneda),
      tipoPago: this.state.movimientoEnEdicion.tipo === 'Egreso' ? 'Egreso' : 'Ingreso',
      observaciones: (this.state.movimientoEnEdicion.observaciones || '').trim(),
      cotizacion: this._requiereCotizacionMovimiento() ? (Number(this.state.movimientoEnEdicion.cotizacion) || 0) : undefined,
      liquidacionOperadorId,
      servicioAsociadoId
    };

    const nombreLiquidacion =
      esEgreso && selLiq > 0
        ? this.state.liquidacionesOperador.filter((l: ILiquidacionItem) => l.id === selLiq)[0]?.codigoReferencia
        : undefined;

    if (this.state.movimientoEnEdicionId) {
      await this._service.updatePago(this.state.movimientoEnEdicionId, pagoData);
      this.setState(prev => ({
        movimientos: prev.movimientos.map(m =>
          m.id === prev.movimientoEnEdicionId
            ? {
                ...m,
                movimiento: pagoData.concepto,
                fecha: pagoData.fechaPago,
                monto: pagoData.importe,
                observaciones: pagoData.observaciones,
                cotizacion: pagoData.cotizacion,
                medioPago: pagoData.medioPago,
                moneda: pagoData.moneda,
                tipo: this.state.movimientoEnEdicion.tipo,
                liquidacionOperadorId: pagoData.liquidacionOperadorId && pagoData.liquidacionOperadorId > 0 ? pagoData.liquidacionOperadorId : undefined,
                servicioAsociadoId: pagoData.servicioAsociadoId && pagoData.servicioAsociadoId > 0 ? pagoData.servicioAsociadoId : undefined,
                liquidacionOperadorNombre: nombreLiquidacion
              }
            : m
        ),
        mostrarEditorMovimiento: false,
        movimientoEnEdicionId: null
      }));
      return;
    }

    const created = await this._service.createPago(pagoData);
    this.setState(prev => ({
      movimientos: prev.movimientos.concat([
        {
          id: created.id,
          movimiento: created.concepto,
          fecha: created.fechaPago,
          monto: created.importe,
          observaciones: created.observaciones,
          cotizacion: created.cotizacion,
          medioPago: created.medioPago,
          moneda: this._normalizarMonedaPago(created.moneda),
          tipo: this.state.movimientoEnEdicion.tipo,
          liquidacionOperadorId: pagoData.liquidacionOperadorId && pagoData.liquidacionOperadorId > 0 ? pagoData.liquidacionOperadorId : undefined,
          servicioAsociadoId: created.servicioAsociadoId && created.servicioAsociadoId > 0 ? created.servicioAsociadoId : undefined,
          liquidacionOperadorNombre: nombreLiquidacion
        }
      ]),
      mostrarEditorMovimiento: false,
      movimientoEnEdicionId: null
    }));
  }

  private _guardarMovimiento = async (): Promise<void> => {
    if (this._esSoloLectura()) { return; }
    if (!this.state.viajeId) {
      this._setSectionError('movimientos', 'Primero guarda el viaje para habilitar pagos.');
      return;
    }
    const esIngreso = this.state.movimientoEnEdicion.tipo === 'Ingreso';
    let servicioSeleccionadoIngreso: IServicio | undefined;
    if (esIngreso) {
      const serviciosDisponiblesIngreso = this.state.servicios.filter((servicio: IServicio) => servicio.id > 0 && !!(servicio.concepto || '').trim());
      const servicioSeleccionadoId = Number(this.state.movimientoEnEdicion.servicioAsociadoId);
      if (serviciosDisponiblesIngreso.length === 0) {
        this._setSectionError('movimientos', 'Primero debes agregar al menos un servicio para poder asociar el concepto de un ingreso.');
        return;
      }
      if (servicioSeleccionadoId <= 0) {
        this._setSectionError('movimientos', 'Selecciona un servicio para asociar el concepto del ingreso.');
        return;
      }
      servicioSeleccionadoIngreso = serviciosDisponiblesIngreso.filter((servicio: IServicio) => servicio.id === servicioSeleccionadoId)[0];
      if (!servicioSeleccionadoIngreso) {
        this._setSectionError('movimientos', 'Selecciona un servicio válido para el ingreso.');
        return;
      }
    } else {
      const liquidacionesDisponiblesEgreso = this.state.liquidacionesOperador.filter(
        (l: ILiquidacionItem) => l.id > 0 && !!(l.codigoReferencia || '').trim()
      );
      const liquidacionSeleccionadaId = Number(this.state.movimientoEnEdicion.liquidacionOperadorId);
      if (liquidacionesDisponiblesEgreso.length === 0) {
        this._setSectionError('movimientos', 'Primero debes agregar al menos una liquidación de operador para poder asociar el concepto de un egreso.');
        return;
      }
      if (liquidacionSeleccionadaId <= 0) {
        this._setSectionError('movimientos', 'Selecciona una liquidación para asociar el concepto del egreso.');
        return;
      }
      const liquidacionSeleccionada = liquidacionesDisponiblesEgreso.filter((l: ILiquidacionItem) => l.id === liquidacionSeleccionadaId)[0];
      if (!liquidacionSeleccionada) {
        this._setSectionError('movimientos', 'Selecciona una liquidación válida para el egreso.');
        return;
      }
    }
    const concepto = (this.state.movimientoEnEdicion.movimiento || '').trim();
    if (!concepto) {
      this._setSectionError('movimientos', 'Indica el concepto del movimiento.');
      return;
    }
    if (this._requiereCotizacionMovimiento()) {
      const cotizacion = Number(this.state.movimientoEnEdicion.cotizacion);
      if (isNaN(cotizacion) || cotizacion <= 0) {
        this._setSectionError('movimientos', 'Indica una cotización válida mayor a 0.');
        return;
      }
    }
    if (esIngreso && servicioSeleccionadoIngreso) {
      const movimientoEditadoId = this.state.movimientoEnEdicionId || undefined;
      const totalYaIngresado = this._getTotalIngresosPorServicioEnMonedaServicio(servicioSeleccionadoIngreso, movimientoEditadoId);
      const ingresoActual: IMovimiento = {
        id: movimientoEditadoId || 0,
        movimiento: this.state.movimientoEnEdicion.movimiento,
        medioPago: this.state.movimientoEnEdicion.medioPago,
        fecha: this.state.movimientoEnEdicion.fecha,
        moneda: this._normalizarMonedaPago(this.state.movimientoEnEdicion.moneda),
        monto: Number(this.state.movimientoEnEdicion.monto) || 0,
        observaciones: this.state.movimientoEnEdicion.observaciones,
        cotizacion: this._requiereCotizacionMovimiento() ? (Number(this.state.movimientoEnEdicion.cotizacion) || 0) : undefined,
        tipo: 'Ingreso',
        servicioAsociadoId: servicioSeleccionadoIngreso.id
      };
      const ingresoActualConvertido = this._convertirIngresoAMonedaServicio(ingresoActual, servicioSeleccionadoIngreso);
      const presupuestoServicio = Number(servicioSeleccionadoIngreso.precioCliente) || 0;
      if (totalYaIngresado + ingresoActualConvertido > presupuestoServicio) {
        this._setSectionError('movimientos', 'El monto ingresado supera el servicio presupuestado.');
        return;
      }
    }
    try {
      this.setState({ guardando: true });
      this._clearSectionError('movimientos');
      await this._guardarPagoEnSharePoint(this.state.viajeId);
    } catch (error) {
      this._setSectionError('movimientos', 'No se pudo guardar el pago en "Registro de Pagos".');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _eliminarMovimiento = async (id: number): Promise<void> => {
    if (this._esSoloLectura()) { return; }
    try {
      this.setState({ guardando: true });
      this._clearSectionError('movimientos');
      await this._service.deletePago(id);
      this.setState(prev => ({ movimientos: prev.movimientos.filter(m => m.id !== id) }));
    } catch (error) {
      this._setSectionError('movimientos', 'No se pudo eliminar el pago.');
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _onGuardarViaje = async (): Promise<void> => {
    if (this._esSoloLectura()) { return; }
    try {
      this.setState({ guardando: true, error: '' });
      const pasajerosIdsResueltos = await this._resolverPasajerosIds();

      if (this.state.viajeId) {
        const serviciosSincronizados = await this._sincronizarServiciosViaje(this.state.viajeId);
        const data = this._mapViajeData(pasajerosIdsResueltos, serviciosSincronizados);
        await this._service.updateViaje(this.state.viajeId, data);
        this.props.onSave();
        return;
      }

      const dataInicial = this._mapViajeData(pasajerosIdsResueltos, []);
      const creado = await this._service.createViaje(dataInicial);
      const serviciosSincronizados = await this._sincronizarServiciosViaje(creado.id);
      const dataDerivada = this._mapViajeData(pasajerosIdsResueltos, serviciosSincronizados);
      await this._service.updateViaje(creado.id, dataDerivada);
      if (this.state.presupuestoArchivoPendiente) {
        await this._service.uploadPresupuesto(creado.id, this.state.presupuestoArchivoPendiente);
      }
      this.props.onSave();
    } catch (error) {
      this.setState({ error: 'No se pudo guardar el viaje en "Registro de Viajes". Verifica lookup IDs.' });
    } finally {
      this.setState({ guardando: false });
    }
  };

  private _onCancelarViaje = (): void => this.props.onClose();

  public render(): React.ReactElement<IProcopioFormsProps> {
    const soloLectura = this._esSoloLectura();
    const isNewMode = this.props.displayMode === FormDisplayMode.New;
    const pagosHabilitados = soloLectura || !!this.state.viajeId;
    const totalesViaje = this._getTotalesServiciosPorMoneda();
    const totalesIngresos = this._getTotalesIngresosPorMoneda();
    const saldosPendientes = {
      usd: totalesViaje.usd - totalesIngresos.usd,
      ars: totalesViaje.ars - totalesIngresos.ars
    };
    const pasajerosFiltrados: IPasajero[] = this.state.pasajeroBusquedaTexto
      ? this.state.pasajeros.filter((p: IPasajero) =>
          p.nombreApellido.toLowerCase().indexOf(this.state.pasajeroBusquedaTexto.toLowerCase()) >= 0 ||
          p.dni.toLowerCase().indexOf(this.state.pasajeroBusquedaTexto.toLowerCase()) >= 0
        )
      : [];
    const pasajerosSeleccionados: IPasajero[] = this.state.pasajeros.filter(
      (p: IPasajero) => this.state.pasajerosIds.indexOf(p.id) >= 0
    );
    const serviciosDisponiblesIngreso = this.state.servicios.filter(
      (servicio: IServicio) => servicio.id > 0 && !!(servicio.concepto || '').trim()
    );
    const liquidacionesDisponiblesEgreso = this.state.liquidacionesOperador.filter(
      (l: ILiquidacionItem) => l.id > 0 && !!(l.codigoReferencia || '').trim()
    );
    const destinosParticularesFiltrados = this.state.destinoGeneralId > 0
      ? this.state.destinos.filter((d: IDestinoItem) => d.destinoGeneralId === this.state.destinoGeneralId)
      : [];
    const requiereCotizacionMovimiento = this._requiereCotizacionMovimiento();

    if (this.state.cargando) {
      return <div style={layoutStyles.page}>Cargando datos...</div>;
    }

    return (
      <div style={layoutStyles.page}>
        <div style={layoutStyles.container}>
          <div style={layoutStyles.header}>
            <div style={layoutStyles.title}>Viaje</div>
            <div style={layoutStyles.subtitle}>{this._getSubtitle()}</div>
          </div>

          <div style={layoutStyles.section}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: 16 }}>
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 14,
                    background: '#FAFAFC',
                    border: '1px solid #E3E4E8',
                    borderRadius: 12,
                    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
                    padding: '16px 18px',
                    minHeight: 104
                  }}
                >
                  <div style={{ width: 52, height: 52, borderRadius: '50%', background: 'rgba(0,120,212,0.10)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#0078D4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <line x1="12" y1="1" x2="12" y2="23" />
                      <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7H14.5a3.5 3.5 0 0 1 0 7H6" />
                    </svg>
                  </div>
                  <div style={{ minWidth: 0, flex: 1, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <div style={{ fontSize: 16, color: '#666B78', lineHeight: 1.25 }}>Total del viaje</div>
                    {this._renderResumenMonedas(totalesViaje)}
                  </div>
                </div>
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 14,
                    background: '#FAFAFC',
                    border: '1px solid #E3E4E8',
                    borderRadius: 12,
                    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
                    padding: '16px 18px',
                    minHeight: 104
                  }}
                >
                  <div style={{ width: 52, height: 52, borderRadius: '50%', background: 'rgba(16,124,16,0.10)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#107C10" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <polyline points="22 7 13.5 15.5 8.5 10.5 2 17" />
                      <polyline points="16 7 22 7 22 13" />
                    </svg>
                  </div>
                  <div style={{ minWidth: 0, flex: 1, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <div style={{ fontSize: 16, color: '#666B78', lineHeight: 1.25 }}>Ingresos</div>
                    {this._renderResumenMonedas(totalesIngresos)}
                  </div>
                </div>
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 14,
                    background: '#FAFAFC',
                    border: '1px solid #E3E4E8',
                    borderRadius: 12,
                    boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
                    padding: '16px 18px',
                    minHeight: 104
                  }}
                >
                  <div style={{ width: 52, height: 52, borderRadius: '50%', background: 'rgba(242,107,28,0.12)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#F26B1C" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <circle cx="12" cy="12" r="9" />
                      <polyline points="12 7 12 12 15 14" />
                    </svg>
                  </div>
                  <div style={{ minWidth: 0, flex: 1, display: 'flex', flexDirection: 'column', gap: 6 }}>
                    <div style={{ fontSize: 16, color: '#666B78', lineHeight: 1.25 }}>Saldo pendiente</div>
                    {this._renderResumenMonedas(saldosPendientes)}
                  </div>
                </div>
              </div>
            </div>

          <div style={presupuestoStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Presupuesto</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={presupuestoStyles.btnAdd}
                  disabled={this.state.guardando}
                  onClick={this._abrirEditorPresupuesto}
                >
                  {this.state.presupuesto ? 'Reemplazar presupuesto' : 'Adjuntar presupuesto'}
                </button>
              )}
            </div>
            {this.state.sectionErrors.presupuesto && <div style={layoutStyles.error}>{this.state.sectionErrors.presupuesto}</div>}
            <div style={presupuestoStyles.attachmentZone}>
              {this.state.presupuestoSubiendo && (
                <div style={presupuestoStyles.uploadingRow}>
                  <InlineLoadingSpinner />
                  <span>Adjuntando archivo...</span>
                </div>
              )}
              {!this.state.presupuestoSubiendo && !this.state.presupuesto ? (
                <div style={presupuestoStyles.attachmentEmpty}>No hay presupuesto adjunto.</div>
              ) : this.state.presupuesto ? (
                <div style={presupuestoStyles.attachmentRow}>
                  <AttachmentIconBox title="Presupuesto" boxStyle={{ borderRadius: presupuestoRadius.control }} />
                  <div style={presupuestoStyles.attachmentMeta}>
                    <div style={presupuestoStyles.attachmentFileName}>
                      {this._getPresupuestoNombreVisible(this.state.presupuesto.fileName)}
                    </div>
                  </div>
                  <div style={{ ...presupuestoStyles.attachmentActions, ...gridActionBarStyle }}>
                    <GridIconActionButton
                      title="Abrir / Descargar"
                      onClick={() =>
                        this._abrirPresupuesto(
                          this.state.presupuesto ? this.state.presupuesto.serverRelativeUrl : ''
                        )
                      }
                      disabled={
                        this.state.guardando ||
                        (!this.state.presupuesto.serverRelativeUrl && !this.state.presupuestoArchivoPendiente)
                      }
                    >
                      <GridIconDownload />
                    </GridIconActionButton>
                    {!soloLectura && (
                      <GridIconActionButton title="Eliminar" onClick={() => { void this._eliminarPresupuesto(); }} disabled={this.state.guardando}>
                        <GridIconTrash />
                      </GridIconActionButton>
                    )}
                  </div>
                </div>
              ) : null}
            </div>
            {!soloLectura && this.state.mostrarEditorPresupuesto && (
              <div style={presupuestoStyles.inlineUploadRoot}>
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Archivo de presupuesto</label>
                  <FileInputEspanol
                    inputId="presupuesto-archivo"
                    disabled={this.state.guardando}
                    buttonLabel="Seleccionar archivo"
                    onChange={this._onSeleccionarArchivoPresupuesto}
                  />
                </div>
                <div>
                  <button type="button" style={presupuestoStyles.btnDefault} onClick={this._cancelarPresupuesto} disabled={this.state.guardando}>
                    Cerrar
                  </button>
                </div>
              </div>
            )}
          </div>
          {this.state.error && <div style={layoutStyles.error}>{this.state.error}</div>}

          <div style={{ ...layoutStyles.section, padding: '38px 32px 32px 32px' }}>
            <div
              style={{
                fontSize: 18,
                fontWeight: 650,
                color: '#242933',
                marginBottom: 32,
                letterSpacing: '-0.2px',
              }}
            >
              Datos del viaje
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
              <div style={{ marginBottom: 0 }}>
                <label
                  style={{
                    display: 'block',
                    fontWeight: 550,
                    fontSize: 13.8,
                    color: '#565868',
                    marginBottom: 7,
                    letterSpacing: '-0.1px',
                  }}
                >
                  Nombre
                </label>
                <input
                  type="text"
                  style={{
                    width: '100%',
                    height: 38,
                    padding: '0 14px',
                    borderRadius: 8,
                    border: '1.4px solid #CDD0D7',
                    background: '#FAFAFC',
                    fontSize: 15,
                    color: '#232529',
                    outline: 'none',
                    transition: 'border-color 0.18s',
                    boxSizing: 'border-box',
                  }}
                  value={this.state.nombreViaje}
                  onChange={this._onCambiarNombre}
                  disabled={soloLectura || this.state.guardando}
                  placeholder="Nombre del viaje"
                />
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                <div style={{ marginBottom: 0 }}>
                  <label
                    style={{
                      display: 'block',
                      fontWeight: 550,
                      fontSize: 13.8,
                      color: '#565868',
                      marginBottom: 7,
                      letterSpacing: '-0.1px',
                    }}
                  >
                    Destino general
                  </label>
                  <select
                    style={{
                      width: '100%',
                      height: 38,
                      padding: '0 14px',
                      borderRadius: 8,
                      border: '1.4px solid #CDD0D7',
                      background: '#FAFAFC',
                      fontSize: 15,
                      color: '#232529',
                      outline: 'none',
                      appearance: 'none',
                      WebkitAppearance: 'none',
                      MozAppearance: 'none',
                      backgroundImage:
                        'url("data:image/svg+xml;charset=UTF-8,<svg fill=\'%23838495\' height=\'20\' viewBox=\'0 0 20 20\' width=\'20\' xmlns=\'http://www.w3.org/2000/svg\'><path d=\'M7.293 8.293a1 1 0 0 1 1.414 0L10 9.586l1.293-1.293a1 1 0 1 1 1.414 1.414l-2 2a1 1 0 0 1-1.414 0l-2-2a1 1 0 0 1 0-1.414z\'/></svg>")',
                      backgroundRepeat: 'no-repeat',
                      backgroundPosition: 'right 10px center',
                      backgroundSize: 20,
                      transition: 'border-color 0.18s',
                      boxSizing: 'border-box',
                    }}
                    value={this.state.destinoGeneralId}
                    onChange={this._onCambiarDestinoGeneral}
                    disabled={soloLectura || this.state.guardando}
                  >
                    <option value={0}>Seleccione destino general...</option>
                    {this.state.destinosGenerales.map((d: IDestinoGeneralItem) => (
                      <option key={d.id} value={d.id}>
                        {d.titulo}
                      </option>
                    ))}
                  </select>
                </div>
                <div style={{ marginBottom: 0 }}>
                  <label
                    style={{
                      display: 'block',
                      fontWeight: 550,
                      fontSize: 13.8,
                      color: '#565868',
                      marginBottom: 7,
                      letterSpacing: '-0.1px',
                    }}
                  >
                    Destino particular
                  </label>
                  <select
                    style={{
                      width: '100%',
                      height: 38,
                      padding: '0 14px',
                      borderRadius: 8,
                      border: '1.4px solid #CDD0D7',
                      background: '#FAFAFC',
                      fontSize: 15,
                      color: '#232529',
                      outline: 'none',
                      appearance: 'none',
                      WebkitAppearance: 'none',
                      MozAppearance: 'none',
                      backgroundImage:
                        'url("data:image/svg+xml;charset=UTF-8,<svg fill=\'%23838495\' height=\'20\' viewBox=\'0 0 20 20\' width=\'20\' xmlns=\'http://www.w3.org/2000/svg\'><path d=\'M7.293 8.293a1 1 0 0 1 1.414 0L10 9.586l1.293-1.293a1 1 0 1 1 1.414 1.414l-2 2a1 1 0 0 1-1.414 0l-2-2a1 1 0 0 1 0-1.414z\'/></svg>")',
                      backgroundRepeat: 'no-repeat',
                      backgroundPosition: 'right 10px center',
                      backgroundSize: 20,
                      transition: 'border-color 0.18s',
                      boxSizing: 'border-box',
                    }}
                    value={this.state.destinoId}
                    onChange={this._onCambiarDestino}
                    disabled={soloLectura || this.state.guardando || this.state.destinoGeneralId <= 0}
                  >
                    <option value={0}>Seleccione destino particular...</option>
                    {destinosParticularesFiltrados.map((d: IDestinoItem) => (
                      <option key={d.id} value={d.id}>
                        {d.titulo}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              <div
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  gap: 14,
                  marginTop: 10,
                }}
              >
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
                  <div>
                    <label
                      style={{
                        display: 'block',
                        fontWeight: 550,
                        fontSize: 13.8,
                        color: '#565868',
                        marginBottom: 7,
                        letterSpacing: '-0.1px',
                      }}
                    >
                      Fecha de salida
                    </label>
                    {soloLectura ? (
                      <div
                        style={{
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
                          alignItems: 'center',
                        }}
                      >
                        {this._formatDateDisplay(this.state.fechaSalida)}
                      </div>
                    ) : (
                    <div style={{ position: 'relative' }}>
                      <input
                        type="date"
                        lang="es-AR"
                        style={{
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
                          boxSizing: 'border-box',
                        }}
                        value={this._getDateOnlyFromSharePoint(this.state.fechaSalida)}
                        onChange={this._onCambiarFechaSalida}
                        disabled={this.state.guardando}
                      />
                      <div
                        style={{
                          position: 'absolute',
                          left: 14,
                          right: 42,
                          top: 0,
                          height: 38,
                          display: 'flex',
                          alignItems: 'center',
                          fontSize: 15,
                          color: '#232529',
                          pointerEvents: 'none',
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis'
                        }}
                      >
                        {this._formatDateDisplay(this.state.fechaSalida)}
                      </div>
                    </div>
                    )}
                  </div>
                  <div>
                    <label
                      style={{
                        display: 'block',
                        fontWeight: 550,
                        fontSize: 13.8,
                        color: '#565868',
                        marginBottom: 7,
                        letterSpacing: '-0.1px',
                      }}
                    >
                      Fecha de regreso
                    </label>
                    {soloLectura ? (
                      <div
                        style={{
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
                          alignItems: 'center',
                        }}
                      >
                        {this._formatDateDisplay(this.state.fechaRegreso)}
                      </div>
                    ) : (
                    <div style={{ position: 'relative' }}>
                      <input
                        type="date"
                        lang="es-AR"
                        style={{
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
                          boxSizing: 'border-box',
                        }}
                        value={this._getDateOnlyFromSharePoint(this.state.fechaRegreso)}
                        onChange={this._onCambiarFechaRegreso}
                        disabled={this.state.guardando}
                      />
                      <div
                        style={{
                          position: 'absolute',
                          left: 14,
                          right: 42,
                          top: 0,
                          height: 38,
                          display: 'flex',
                          alignItems: 'center',
                          fontSize: 15,
                          color: '#232529',
                          pointerEvents: 'none',
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis'
                        }}
                      >
                        {this._formatDateDisplay(this.state.fechaRegreso)}
                      </div>
                    </div>
                    )}
                  </div>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 16 }}>
                  <div>
                    <label
                      style={{
                        display: 'block',
                        fontWeight: 550,
                        fontSize: 13.8,
                        color: '#565868',
                        marginBottom: 7,
                        letterSpacing: '-0.1px',
                      }}
                    >
                      Estado
                    </label>
                    {soloLectura ? (
                      <div
                        style={{
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
                          alignItems: 'center',
                        }}
                      >
                        {this.state.estado}
                      </div>
                    ) : (
                      <select
                        style={{
                          width: '100%',
                          height: 38,
                          padding: '0 14px',
                          borderRadius: 8,
                          border: '1.4px solid #CDD0D7',
                          background: '#FAFAFC',
                          fontSize: 15,
                          color: '#232529',
                          outline: 'none',
                          appearance: 'none',
                          WebkitAppearance: 'none',
                          MozAppearance: 'none',
                          backgroundImage:
                            'url("data:image/svg+xml;charset=UTF-8,<svg fill=\'%23838495\' height=\'20\' viewBox=\'0 0 20 20\' width=\'20\' xmlns=\'http://www.w3.org/2000/svg\'><path d=\'M7.293 8.293a1 1 0 0 1 1.414 0L10 9.586l1.293-1.293a1 1 0 1 1 1.414 1.414l-2 2a1 1 0 0 1-1.414 0l-2-2a1 1 0 0 1 0-1.414z\'/></svg>")',
                          backgroundRepeat: 'no-repeat',
                          backgroundPosition: 'right 10px center',
                          backgroundSize: 20,
                          transition: 'border-color 0.18s',
                          boxSizing: 'border-box',
                        }}
                        value={this.state.estado}
                        onChange={this._onCambiarEstado}
                        disabled={this.state.guardando}
                      >
                        {ESTADOS_VIAJE.map((estado: EstadoViaje) => (
                          <option key={estado} value={estado}>
                            {estado}
                          </option>
                        ))}
                      </select>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
 

          <div style={pasajerosStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Pasajeros</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={pasajerosStyles.btnAdd}
                  onClick={this._abrirEditorPasajero}
                  disabled={this.state.guardando}
                >
                  Agregar pasajero
                </button>
              )}
            </div>
            {this.state.sectionErrors.pasajeros && <div style={layoutStyles.error}>{this.state.sectionErrors.pasajeros}</div>}
            <div style={pasajerosStyles.tableOuter}>
              <table style={pasajerosStyles.table}>
                <thead>
                  <tr style={layoutStyles.tableHeaderRow}>
                    <th style={{ ...layoutStyles.th, paddingLeft: 16 }}>Nombre y Apellido</th>
                    <th style={layoutStyles.th}>DNI</th>
                    <th style={layoutStyles.th}>Fecha Nacimiento</th>
                    <th style={layoutStyles.th}>Documentos</th>
                    {!soloLectura && <th style={{ ...layoutStyles.th, textAlign: 'right', paddingRight: 16 }}>Acciones</th>}
                  </tr>
                </thead>
                <tbody>
                  {pasajerosSeleccionados.map((p: IPasajero) => {
                    const documentosExpandidos = this.state.pasajeroDocumentosExpandidoId === p.id;
                    return (
                      <React.Fragment key={p.id}>
                        <tr>
                          <td style={{ ...layoutStyles.td, paddingLeft: 16 }}>
                            {soloLectura && this._esPasajeroPersistidoEnSharePoint(p.id) ? (
                              <a
                                href={this._getPasajeroDisplayUrl(p.id)}
                                target="_blank"
                                rel="noopener noreferrer"
                                style={{ color: '#005A9E', textDecoration: 'none', fontWeight: 600 }}
                                title="Abrir pasajero"
                              >
                                {p.nombreApellido}
                              </a>
                            ) : (
                              p.nombreApellido
                            )}
                          </td>
                          <td style={layoutStyles.td}>{p.dni}</td>
                          <td style={layoutStyles.td}>{this._formatDateDisplay(p.fechaNacimiento || '')}</td>
                          <td style={layoutStyles.td}>
                            <button
                              type="button"
                              title={documentosExpandidos ? 'Ocultar documentos' : 'Ver documentos'}
                              style={
                                documentosExpandidos
                                  ? { ...pasajerosStyles.documentosToggle, ...pasajerosStyles.documentosToggleActive }
                                  : pasajerosStyles.documentosToggle
                              }
                              onClick={() => this._toggleDocumentosPasajero(p.id)}
                              disabled={this.state.guardando}
                            >
                              <GridIconPaperclip />
                              <span>Documentos</span>
                              {documentosExpandidos ? <GridIconChevronUp /> : <GridIconChevronDown />}
                            </button>
                          </td>
                          {!soloLectura && (
                            <td style={{ ...layoutStyles.td, ...layoutStyles.actionsCell, paddingRight: 16 }}>
                              <div style={gridActionBarStyle}>
                                <GridIconActionButton title="Editar" onClick={() => this._editarPasajero(p.id)} disabled={this.state.guardando}>
                                  <GridIconEdit />
                                </GridIconActionButton>
                                <GridIconActionButton title="Quitar" onClick={() => this._quitarPasajeroSeleccionado(p.id)} disabled={this.state.guardando}>
                                  <GridIconTrash />
                                </GridIconActionButton>
                              </div>
                            </td>
                          )}
                        </tr>
                        {this._renderDocumentosPasajeroExpandido(p, soloLectura)}
                      </React.Fragment>
                    );
                  })}
                  {pasajerosSeleccionados.length === 0 && (
                    <tr>
                      <td style={layoutStyles.emptyTableCell} colSpan={soloLectura ? 4 : 5}>
                        No hay pasajeros agregados.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {!soloLectura && this.state.mostrarEditorPasajero && (
              <div style={pasajerosStyles.inlineFormRoot}>
                {this.state.pasajeroEnEdicionId === null && (
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Buscar pasajero existente</label>
                    <input
                      type="text"
                      style={pasajerosStyles.input}
                      value={this.state.pasajeroBusquedaTexto}
                      onChange={this._onCambiarBusquedaPasajero}
                      placeholder="Escriba nombre o DNI..."
                    />
                    {this.state.pasajeroBusquedaTexto && pasajerosFiltrados.length > 0 && (
                      <div style={pasajerosStyles.autocompleteList}>
                        {pasajerosFiltrados.map((p: IPasajero) => (
                          <div
                            key={p.id}
                            style={{ padding: '4px 8px', fontSize: 12, cursor: 'pointer', borderBottom: '1px solid #f3f2f1' }}
                            onClick={() => this._onSeleccionarPasajeroDesdeBusqueda(p.id)}
                          >
                            {p.nombreApellido} ({p.dni})
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, width: 120, alignSelf: 'flex-end' }}>
                    <label style={{ ...layoutStyles.label, visibility: 'hidden' }}>Nuevo</label>
                    <button type="button" style={pasajerosStyles.btnDefault} onClick={this._habilitarNuevoPasajero}>
                      Nuevo
                    </button>
                  </div>
                </div>
                )}
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>
                      {this.state.pasajeroEnEdicionId !== null ? 'Nombre y Apellido' : 'Nuevo pasajero: Nombre y Apellido'}
                    </label>
                    <input type="text" style={pasajerosStyles.input} value={this.state.nuevoPasajeroDraft.nombreApellido} onChange={e => this._onCambiarNuevoPasajero('nombreApellido', e.target.value)} disabled={!this.state.pasajeroCamposHabilitados} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>DNI</label>
                    <input type="text" style={pasajerosStyles.input} value={this.state.nuevoPasajeroDraft.dni} onChange={e => this._onCambiarNuevoPasajero('dni', e.target.value)} disabled={!this.state.pasajeroCamposHabilitados} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Pasaporte</label>
                    <input type="text" style={pasajerosStyles.input} value={this.state.nuevoPasajeroDraft.pasaporte || ''} onChange={e => this._onCambiarNuevoPasajero('pasaporte', e.target.value)} disabled={!this.state.pasajeroCamposHabilitados} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Teléfono</label>
                    <input type="text" style={pasajerosStyles.input} value={this.state.nuevoPasajeroDraft.telefono || ''} onChange={e => this._onCambiarNuevoPasajero('telefono', e.target.value)} disabled={!this.state.pasajeroCamposHabilitados} />
                  </div>
                </div>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Email</label>
                    <input type="text" style={pasajerosStyles.input} value={this.state.nuevoPasajeroDraft.email} onChange={e => this._onCambiarNuevoPasajero('email', e.target.value)} disabled={!this.state.pasajeroCamposHabilitados} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Fecha de nacimiento</label>
                    <div style={pasajerosStyles.dateInputShell}>
                      <input
                        type="date"
                        lang="es-AR"
                        className={styles.pasajeroDateInput}
                        value={this._getDateOnlyFromSharePoint(this.state.nuevoPasajeroDraft.fechaNacimiento || '')}
                        onChange={this._onCambiarFechaNacimientoPasajero}
                        disabled={!this.state.pasajeroCamposHabilitados || this.state.guardando}
                      />
                      <div
                        style={{
                          ...pasajerosStyles.dateInputOverlay,
                          ...(this.state.nuevoPasajeroDraft.fechaNacimiento
                            ? pasajerosStyles.dateInputOverlayValue
                            : pasajerosStyles.dateInputOverlayPlaceholder)
                        }}
                      >
                        {this._formatDateDisplay(this.state.nuevoPasajeroDraft.fechaNacimiento || '') || 'dd/mm/aaaa'}
                      </div>
                    </div>
                  </div>
                </div>
                <button type="button" style={pasajerosStyles.btnPrimary} onClick={() => { void this._guardarPasajeroDraft(); }} disabled={this.state.guardando || !this.state.pasajeroCamposHabilitados}>Guardar</button>
                <button type="button" style={pasajerosStyles.btnDefault} onClick={this._cancelarPasajero} disabled={this.state.guardando}>Cancelar</button>
              </div>
            )}
          </div>

          <div style={serviciosStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Servicios</div>
              {!soloLectura && (
                <button type="button" style={serviciosStyles.btnAdd} onClick={this._abrirEditorServicio} disabled={this.state.guardando}>
                  Agregar servicio
                </button>
              )}
            </div>
            {this.state.sectionErrors.servicios && <div style={layoutStyles.error}>{this.state.sectionErrors.servicios}</div>}
            <div style={serviciosStyles.tableOuter}>
              <table style={serviciosStyles.table}>
                <thead>
                  <tr style={layoutStyles.tableHeaderRow}>
                    <th style={{ ...layoutStyles.th, paddingLeft: 16 }}>Concepto</th>
                    <th style={layoutStyles.th}>Moneda</th>
                    <th style={{ ...layoutStyles.th, textAlign: 'right' }}>Precio cliente</th>
                    <th style={layoutStyles.th}>Operador</th>
                    {!soloLectura && <th style={{ ...layoutStyles.th, textAlign: 'right', paddingRight: 16 }}>Acciones</th>}
                  </tr>
                </thead>
                <tbody>
                  {this.state.servicios.map(s => (
                    <tr key={s.id}>
                      <td style={{ ...layoutStyles.td, paddingLeft: 16 }}>{s.concepto}</td>
                      <td style={layoutStyles.td}>{s.moneda}</td>
                      <td style={{ ...layoutStyles.td, textAlign: 'right' }}>{s.precioCliente.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}</td>
                      <td style={layoutStyles.td}>{this._getServicioOperadorDisplay(s)}</td>
                      {!soloLectura && (
                        <td style={{ ...layoutStyles.td, ...layoutStyles.actionsCell, paddingRight: 16 }}>
                          <div style={gridActionBarStyle}>
                            <GridIconActionButton title="Editar" onClick={() => this._editarServicio(s.id)} disabled={this.state.guardando}>
                              <GridIconEdit />
                            </GridIconActionButton>
                            <GridIconActionButton title="Eliminar" onClick={() => this._eliminarServicio(s.id)} disabled={this.state.guardando}>
                              <GridIconTrash />
                            </GridIconActionButton>
                          </div>
                        </td>
                      )}
                    </tr>
                  ))}
                  {this.state.servicios.length === 0 && (
                    <tr>
                      <td style={layoutStyles.emptyTableCell} colSpan={soloLectura ? 4 : 5}>
                        No hay servicios registrados.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {!soloLectura && this.state.mostrarEditorServicio && (
              <div style={serviciosStyles.inlineFormRoot}>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Concepto</label>
                    <select style={serviciosStyles.select} value={this.state.servicioEnEdicion.concepto} onChange={e => this._actualizarCampoServicio('concepto', e.target.value)}>
                      <option value="">Seleccione...</option>
                      <option value="Aéreos">Aéreos</option>
                      <option value="Traslados">Traslados</option>
                      <option value="Alojamiento">Alojamiento</option>
                      <option value="Crucero">Crucero</option>
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Moneda</label>
                    <select style={serviciosStyles.select} value={this.state.servicioEnEdicion.moneda} onChange={e => this._actualizarCampoServicio('moneda', e.target.value)}>
                      {MONEDAS_SERVICIO.map((moneda: MonedaServicio) => (
                        <option key={moneda} value={moneda}>
                          {moneda}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Precio cliente</label>
                    <input type="number" style={serviciosStyles.input} value={this.state.servicioEnEdicion.precioCliente} onChange={e => this._actualizarCampoServicio('precioCliente', e.target.value)} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Operador</label>
                    <select
                      style={serviciosStyles.select}
                      value={this.state.servicioEnEdicion.operadorId}
                      onChange={e => this._actualizarCampoServicio('operadorId', e.target.value)}
                      disabled={this.state.guardando}
                    >
                      <option value="">Seleccione...</option>
                      {this.state.operadores.map((operador: IOperadorItem) => (
                        <option key={operador.id} value={String(operador.id)}>
                          {operador.titulo}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div>
                  <button type="button" style={serviciosStyles.btnPrimary} onClick={this._guardarServicioMock} disabled={this.state.guardando}>Guardar</button>
                  <button type="button" style={serviciosStyles.btnDefault} onClick={this._cancelarServicioMock} disabled={this.state.guardando}>Cancelar</button>
                </div>
              </div>
            )}
          </div>

          {!isNewMode && (
          <div style={liquidacionesStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Liquidaciones Operador</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={pagosHabilitados ? liquidacionesStyles.btnAdd : { ...liquidacionesStyles.btnAdd, ...layoutStyles.buttonDisabled }}
                  onClick={this._abrirEditorLiquidacion}
                  disabled={!pagosHabilitados || this.state.guardando}
                >
                  Agregar liquidación
                </button>
              )}
            </div>
            {this.state.sectionErrors.liquidaciones && <div style={layoutStyles.error}>{this.state.sectionErrors.liquidaciones}</div>}
            <div style={liquidacionesStyles.tableOuter}>
              <table style={liquidacionesStyles.table}>
                <thead>
                  <tr style={layoutStyles.tableHeaderRow}>
                    <th style={{ ...layoutStyles.th, paddingLeft: 16 }}>Código Referencia</th>
                    <th style={layoutStyles.th}>Operador</th>
                    <th style={layoutStyles.th}>Moneda</th>
                    <th style={{ ...layoutStyles.th, textAlign: 'right' }}>Monto</th>
                    <th style={layoutStyles.th}>Archivo</th>
                    {!soloLectura && <th style={{ ...layoutStyles.th, textAlign: 'right', paddingRight: 16 }}>Acciones</th>}
                  </tr>
                </thead>
                <tbody>
                  {this.state.liquidacionesOperador.map((l: ILiquidacionItem) => (
                    <tr key={l.id}>
                      <td style={{ ...layoutStyles.td, paddingLeft: 16 }}>{l.codigoReferencia}</td>
                      <td style={layoutStyles.td}>{(l.operadorNombre || '').trim() || '—'}</td>
                      <td style={layoutStyles.td}>{this._normalizarMonedaLiquidacion(l.moneda)}</td>
                      <td style={{ ...layoutStyles.td, textAlign: 'right' }}>{l.monto.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}</td>
                      <td style={layoutStyles.td}>
                        {l.archivoNombre && l.archivoUrl ? (
                          <a
                            href={l.archivoUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            style={liquidacionesStyles.archivoLink}
                            title="Abrir / Descargar"
                            onClick={e => {
                              e.preventDefault();
                              this._abrirLiquidacionArchivo(l.archivoUrl || '');
                            }}
                          >
                            {l.archivoNombre}
                          </a>
                        ) : (
                          <span style={{ ...liquidacionesStyles.archivoNombreEnTabla, color: '#605e5c' }}>—</span>
                        )}
                      </td>
                      {!soloLectura && (
                        <td style={{ ...layoutStyles.td, ...layoutStyles.actionsCell, paddingRight: 16 }}>
                          <div style={gridActionBarStyle}>
                            <GridIconActionButton
                              title="Editar"
                              onClick={() => this._editarLiquidacion(l.id)}
                              disabled={!pagosHabilitados || this.state.guardando}
                            >
                              <GridIconEdit />
                            </GridIconActionButton>
                            <GridIconActionButton
                              title="Eliminar"
                              onClick={() => this._eliminarLiquidacion(l.id)}
                              disabled={!pagosHabilitados || this.state.guardando}
                            >
                              <GridIconTrash />
                            </GridIconActionButton>
                          </div>
                        </td>
                      )}
                    </tr>
                  ))}
                  {this.state.liquidacionesOperador.length === 0 && (
                    <tr>
                      <td style={layoutStyles.emptyTableCell} colSpan={soloLectura ? 5 : 6}>
                        No hay liquidaciones registradas.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {!soloLectura && this.state.mostrarEditorLiquidacion && (
              <div style={liquidacionesStyles.inlineFormRoot}>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Código Referencia</label>
                    <input
                      type="text"
                      style={liquidacionesStyles.input}
                      value={this.state.liquidacionEnEdicion.codigoReferencia}
                      onChange={e => this._actualizarCampoLiquidacion('codigoReferencia', e.target.value)}
                      disabled={this.state.guardando}
                    />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Operador</label>
                    <select
                      style={liquidacionesStyles.input}
                      value={this.state.liquidacionEnEdicion.operadorId}
                      onChange={e => this._actualizarCampoLiquidacion('operadorId', e.target.value)}
                      disabled={this.state.guardando}
                    >
                      <option value="">Seleccione...</option>
                      {this.state.operadores.map((operador: IOperadorItem) => (
                        <option key={operador.id} value={String(operador.id)}>
                          {operador.titulo}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Moneda</label>
                    <select
                      style={liquidacionesStyles.input}
                      value={this.state.liquidacionEnEdicion.moneda}
                      onChange={e => this._actualizarCampoLiquidacion('moneda', e.target.value)}
                      disabled={this.state.guardando}
                    >
                      {MONEDAS_LIQUIDACION.map((moneda: MonedaLiquidacion) => (
                        <option key={moneda} value={moneda}>
                          {moneda}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Monto</label>
                    <input
                      type="number"
                      style={liquidacionesStyles.input}
                      value={this.state.liquidacionEnEdicion.monto}
                      onChange={e => this._actualizarCampoLiquidacion('monto', e.target.value)}
                      disabled={this.state.guardando}
                    />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Archivo</label>
                    <FileInputEspanol
                      inputId="liquidacion-archivo"
                      disabled={this.state.guardando}
                      buttonLabel="Seleccionar archivo"
                      onChange={this._onArchivoLiquidacionChange}
                    />
                    {this.state.liquidacionEnEdicion.archivoFile && (
                      <div style={{ ...layoutStyles.fileInputHint, marginTop: 6 }}>
                        Archivo a cargar: {this.state.liquidacionEnEdicion.archivoFile.name}
                      </div>
                    )}
                  </div>
                </div>
                <div>
                  <button type="button" style={liquidacionesStyles.btnPrimary} onClick={this._guardarLiquidacion} disabled={this.state.guardando}>
                    {this.state.liquidacionEnEdicionId ? 'Guardar cambios' : 'Guardar'}
                  </button>
                  <button type="button" style={liquidacionesStyles.btnDefault} onClick={this._cancelarLiquidacion} disabled={this.state.guardando}>
                    Cancelar
                  </button>
                </div>
              </div>
            )}
          </div>
          )}

          {!isNewMode && (
          <div style={movimientosStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Registro de Ingresos y Egresos</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={pagosHabilitados ? movimientosStyles.btnAdd : { ...movimientosStyles.btnAdd, ...layoutStyles.buttonDisabled }}
                  onClick={this._abrirEditorMovimiento}
                  disabled={!pagosHabilitados || this.state.guardando}
                >
                  Agregar pago
                </button>
              )}
            </div>
            {this.state.sectionErrors.movimientos && <div style={layoutStyles.error}>{this.state.sectionErrors.movimientos}</div>}
            <div style={movimientosStyles.tableOuter}>
              <table style={movimientosStyles.table}>
                <thead>
                  <tr style={layoutStyles.tableHeaderRow}>
                    <th style={{ ...layoutStyles.th, paddingLeft: 16 }}>Tipo</th>
                    <th style={layoutStyles.th}>Concepto</th>
                    <th style={layoutStyles.th}>Moneda</th>
                    <th style={layoutStyles.th}>Medio de pago</th>
                    <th style={layoutStyles.th}>Fecha de pago</th>
                    <th style={{ ...layoutStyles.th, textAlign: 'right' }}>Monto</th>
                    {!soloLectura && <th style={{ ...layoutStyles.th, textAlign: 'right', paddingRight: 16 }}>Acciones</th>}
                  </tr>
                </thead>
                <tbody>
                  {this.state.movimientos.map(m => (
                    <tr key={m.id}>
                      <td style={{ ...layoutStyles.td, paddingLeft: 16 }}>{m.tipo}</td>
                      <td style={layoutStyles.td}>{this._getConceptoMostradoGrillaMovimiento(m)}</td>
                      <td style={layoutStyles.td}>{this._getEtiquetaMonedaPago(m.moneda)}</td>
                      <td style={layoutStyles.td}>{m.medioPago}</td>
                      <td style={layoutStyles.td}>{m.fecha}</td>
                      <td style={{ ...layoutStyles.td, textAlign: 'right' }}>{m.monto.toLocaleString('es-MX', { style: 'currency', currency: 'MXN' })}</td>
                      {!soloLectura && (
                        <td style={{ ...layoutStyles.td, ...layoutStyles.actionsCell, paddingRight: 16 }}>
                          <div style={gridActionBarStyle}>
                            <GridIconActionButton title="Editar" onClick={() => this._editarMovimiento(m.id)} disabled={!pagosHabilitados || this.state.guardando}>
                              <GridIconEdit />
                            </GridIconActionButton>
                            <GridIconActionButton title="Eliminar" onClick={() => this._eliminarMovimiento(m.id)} disabled={!pagosHabilitados || this.state.guardando}>
                              <GridIconTrash />
                            </GridIconActionButton>
                          </div>
                        </td>
                      )}
                    </tr>
                  ))}
                  {this.state.movimientos.length === 0 && (
                    <tr>
                      <td style={layoutStyles.emptyTableCell} colSpan={soloLectura ? 6 : 7}>
                        No hay ingresos o egresos registrados.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {!soloLectura && this.state.mostrarEditorMovimiento && (
              <div style={movimientosStyles.inlineFormRoot}>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Tipo</label>
                    <select
                      style={movimientosStyles.select}
                      value={this.state.movimientoEnEdicion.tipo}
                      onChange={this._onCambiarTipoMovimiento}
                      disabled={this.state.guardando}
                    >
                      <option value="Ingreso">Ingreso</option>
                      <option value="Egreso">Egreso</option>
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Concepto</label>
                    {this.state.movimientoEnEdicion.tipo === 'Ingreso' ? (
                      <select
                        style={movimientosStyles.select}
                        value={this.state.movimientoEnEdicion.servicioAsociadoId}
                        onChange={e => this._onSeleccionarServicioIngreso(e.target.value)}
                        disabled={this.state.guardando || serviciosDisponiblesIngreso.length === 0}
                      >
                        <option value="">Seleccione...</option>
                        {serviciosDisponiblesIngreso.map((servicio: IServicio) => (
                          <option key={servicio.id} value={String(servicio.id)}>
                            {servicio.concepto}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <select
                        style={movimientosStyles.select}
                        value={this.state.movimientoEnEdicion.liquidacionOperadorId}
                        onChange={e => this._onSeleccionarLiquidacionEgreso(e.target.value)}
                        disabled={this.state.guardando || liquidacionesDisponiblesEgreso.length === 0}
                      >
                        <option value="">Seleccione...</option>
                        {liquidacionesDisponiblesEgreso.map((l: ILiquidacionItem) => (
                          <option key={l.id} value={String(l.id)}>
                            {l.codigoReferencia}
                          </option>
                        ))}
                      </select>
                    )}
                    {this.state.movimientoEnEdicion.tipo === 'Ingreso' && serviciosDisponiblesIngreso.length === 0 && (
                      <div style={{ ...layoutStyles.info, marginBottom: 0, marginTop: 6 }}>
                        Primero debes agregar un servicio para poder asociar el concepto del ingreso.
                      </div>
                    )}
                    {this.state.movimientoEnEdicion.tipo === 'Egreso' && liquidacionesDisponiblesEgreso.length === 0 && (
                      <div style={{ ...layoutStyles.info, marginBottom: 0, marginTop: 6 }}>
                        Primero debes agregar una liquidación de operador para poder asociar el concepto del egreso.
                      </div>
                    )}
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Medio de pago</label>
                    <select style={movimientosStyles.select} value={this.state.movimientoEnEdicion.medioPago} onChange={e => this._actualizarCampoMovimiento('medioPago', e.target.value)} disabled={this.state.guardando}>
                      <option value="">Seleccione...</option>
                      <option value="Transferencia">Transferencia</option>
                      <option value="Tarjeta de Crédito">Tarjeta de Crédito</option>
                      <option value="Efectivo">Efectivo</option>
                    </select>
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Moneda</label>
                    <select
                      style={movimientosStyles.select}
                      value={this.state.movimientoEnEdicion.moneda}
                      onChange={e => this._actualizarCampoMovimiento('moneda', this._normalizarMonedaPago(e.target.value))}
                      disabled={this.state.guardando}
                    >
                      {MONEDAS_PAGO.map((moneda: MonedaPago) => (
                        <option key={moneda} value={moneda}>
                          {moneda}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Fecha de pago</label>
                    <input type="date" style={movimientosStyles.input} value={this.state.movimientoEnEdicion.fecha} onChange={e => this._actualizarCampoMovimiento('fecha', e.target.value)} disabled={this.state.guardando} />
                  </div>
                  <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                    <label style={layoutStyles.label}>Monto</label>
                    <input type="number" style={movimientosStyles.input} value={this.state.movimientoEnEdicion.monto} onChange={e => this._actualizarCampoMovimiento('monto', e.target.value)} disabled={this.state.guardando} />
                  </div>
                  {requiereCotizacionMovimiento && (
                    <div style={{ ...layoutStyles.fieldGroup, ...layoutStyles.inlineEditorField }}>
                      <label style={layoutStyles.label}>Cotización</label>
                      <input
                        type="number"
                        style={movimientosStyles.input}
                        value={this.state.movimientoEnEdicion.cotizacion}
                        onChange={e => this._actualizarCampoMovimiento('cotizacion', e.target.value)}
                        disabled={this.state.guardando}
                      />
                    </div>
                  )}
                </div>
                <div style={layoutStyles.inlineEditorRow}>
                  <div style={{ ...layoutStyles.fieldGroup, flex: '1 1 100%' }}>
                    <label style={layoutStyles.label}>Observaciones</label>
                    <textarea
                      style={{ ...movimientosStyles.input, minHeight: 72, resize: 'vertical', padding: '10px 14px' }}
                      value={this.state.movimientoEnEdicion.observaciones}
                      onChange={e => this._actualizarCampoMovimiento('observaciones', e.target.value)}
                      disabled={this.state.guardando}
                    />
                  </div>
                </div>
                <div>
                  <button type="button" style={movimientosStyles.btnPrimary} onClick={this._guardarMovimiento} disabled={this.state.guardando}>Guardar</button>
                  <button type="button" style={movimientosStyles.btnDefault} onClick={this._cancelarMovimiento} disabled={this.state.guardando}>Cancelar</button>
                </div>
              </div>
            )}
          </div>
          )}

          {!isNewMode && (
          <div style={vouchersStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Vouchers</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={pagosHabilitados ? vouchersStyles.btnAdd : { ...vouchersStyles.btnAdd, ...layoutStyles.buttonDisabled }}
                  disabled={!pagosHabilitados || this.state.guardando}
                  onClick={this._abrirEditorVoucher}
                >
                  Adjuntar documentos
                </button>
              )}
            </div>
            {this.state.sectionErrors.vouchers && <div style={layoutStyles.error}>{this.state.sectionErrors.vouchers}</div>}
            {!soloLectura && !this.state.viajeId && (
              <div style={layoutStyles.info}>Guardá primero el viaje para poder adjuntar vouchers.</div>
            )}
            <div style={vouchersStyles.subheading}>Documentos adjuntos</div>
            <div style={vouchersStyles.attachmentZone}>
              {this.state.vouchersSubiendo && (
                <div style={vouchersStyles.uploadingRow}>
                  <InlineLoadingSpinner />
                  <span>Adjuntando archivos...</span>
                </div>
              )}
              {!this.state.vouchersSubiendo && this.state.vouchers.length === 0 ? (
                <div style={vouchersStyles.attachmentEmpty}>No hay documentos adjuntos.</div>
              ) : (
                this.state.vouchers.map((v: IVoucherItem, index: number) => {
                  const esUltimo = index === this.state.vouchers.length - 1;
                  return (
                    <div
                      key={v.fileName}
                      style={esUltimo ? { ...vouchersStyles.attachmentRow, ...vouchersStyles.attachmentRowLast } : vouchersStyles.attachmentRow}
                    >
                      <AttachmentIconBox title="Documento" boxStyle={{ borderRadius: vouchersRadius.control }} />
                      <div style={vouchersStyles.attachmentMeta}>
                        <div style={vouchersStyles.attachmentFileName}>{v.fileName}</div>
                      </div>
                      <div style={{ ...vouchersStyles.attachmentActions, ...gridActionBarStyle }}>
                        <GridIconActionButton
                          title="Abrir / Descargar"
                          onClick={() => this._abrirVoucher(v.serverRelativeUrl)}
                          disabled={this.state.guardando || !v.serverRelativeUrl}
                        >
                          <GridIconDownload />
                        </GridIconActionButton>
                        {!soloLectura && (
                          <GridIconActionButton title="Eliminar" onClick={() => this._eliminarVoucher(v.fileName)} disabled={this.state.guardando}>
                            <GridIconTrash />
                          </GridIconActionButton>
                        )}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
            {!soloLectura && this.state.mostrarEditorVoucher && (
              <div style={vouchersStyles.inlineUploadRoot}>
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Archivos</label>
                  <FileInputEspanol
                    inputId="voucher-archivos"
                    multiple={true}
                    disabled={this.state.guardando}
                    buttonLabel="Seleccionar archivos"
                    onChange={this._onSeleccionarArchivosVoucher}
                  />
                </div>
                <div>
                  <button type="button" style={vouchersStyles.btnDefault} onClick={this._cancelarVoucher} disabled={this.state.guardando}>
                    Cerrar
                  </button>
                </div>
              </div>
            )}
          </div>
          )}

          {!isNewMode && (
          <div style={facturasStyles.sectionCard}>
            <div style={layoutStyles.sectionToolbar}>
              <div style={layoutStyles.toolbarTitle}>Facturas</div>
              {!soloLectura && (
                <button
                  type="button"
                  style={pagosHabilitados ? facturasStyles.btnAdd : { ...facturasStyles.btnAdd, ...layoutStyles.buttonDisabled }}
                  disabled={!pagosHabilitados || this.state.guardando}
                  onClick={this._abrirEditorFactura}
                >
                  Adjuntar facturas
                </button>
              )}
            </div>
            {this.state.sectionErrors.facturas && <div style={layoutStyles.error}>{this.state.sectionErrors.facturas}</div>}
            {!pagosHabilitados && (
              <div style={layoutStyles.info}>Guardá primero el viaje para poder adjuntar facturas.</div>
            )}
            <div style={facturasStyles.subheading}>Documentos adjuntos</div>
            <div style={facturasStyles.attachmentZone}>
              {this.state.facturasSubiendo && (
                <div style={facturasStyles.uploadingRow}>
                  <InlineLoadingSpinner />
                  <span>Adjuntando archivos...</span>
                </div>
              )}
              {!this.state.facturasSubiendo && this.state.facturas.length === 0 ? (
                <div style={facturasStyles.attachmentEmpty}>No hay facturas adjuntas.</div>
              ) : (
                this.state.facturas.map((f: IFacturaItem, index: number) => {
                  const esUltimo = index === this.state.facturas.length - 1;
                  return (
                    <div
                      key={f.fileName}
                      style={esUltimo ? { ...facturasStyles.attachmentRow, ...facturasStyles.attachmentRowLast } : facturasStyles.attachmentRow}
                    >
                      <AttachmentIconBox title="Factura" boxStyle={{ borderRadius: facturasRadius.control }} />
                      <div style={facturasStyles.attachmentMeta}>
                        <div style={facturasStyles.attachmentFileName}>{this._getFacturaNombreVisible(f.fileName)}</div>
                      </div>
                      <div style={{ ...facturasStyles.attachmentActions, ...gridActionBarStyle }}>
                        <GridIconActionButton
                          title="Abrir / Descargar"
                          onClick={() => this._abrirFactura(f.serverRelativeUrl)}
                          disabled={this.state.guardando || !f.serverRelativeUrl}
                        >
                          <GridIconDownload />
                        </GridIconActionButton>
                        {!soloLectura && (
                          <GridIconActionButton title="Eliminar" onClick={() => { void this._eliminarFactura(f.fileName); }} disabled={this.state.guardando}>
                            <GridIconTrash />
                          </GridIconActionButton>
                        )}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
            {!soloLectura && this.state.mostrarEditorFactura && pagosHabilitados && (
              <div style={facturasStyles.inlineUploadRoot}>
                <div style={layoutStyles.fieldGroup}>
                  <label style={layoutStyles.label}>Archivos</label>
                  <FileInputEspanol
                    inputId="factura-archivos"
                    multiple={true}
                    disabled={this.state.guardando}
                    buttonLabel="Seleccionar archivos"
                    onChange={this._onSeleccionarArchivosFactura}
                  />
                </div>
                <div>
                  <button type="button" style={facturasStyles.btnDefault} onClick={this._cancelarFactura} disabled={this.state.guardando}>
                    Cerrar
                  </button>
                </div>
              </div>
            )}
          </div>
          )}

          <div style={observacionesStyles.sectionCard}>
            <div style={observacionesStyles.title}>Observaciones</div>
            <div style={observacionesStyles.textareaOuter}>
              <div
                style={{
                  ...observacionesStyles.textareaShell,
                  ...(soloLectura || this.state.guardando ? observacionesStyles.textareaShellDisabled : {})
                }}
              >
                <textarea
                  id="procopio-observaciones"
                  style={{
                    ...observacionesStyles.textarea,
                    ...(soloLectura || this.state.guardando ? observacionesStyles.textareaDisabled : {})
                  }}
                  value={this.state.observaciones}
                  onChange={this._onCambiarObservaciones}
                  disabled={soloLectura || this.state.guardando}
                  aria-label="Observaciones"
                />
              </div>
            </div>
          </div>

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
              <button type="button" style={layoutStyles.primaryButton} onClick={this._onGuardarViaje} disabled={this.state.guardando}>
                {this.state.guardando ? 'Guardando...' : this.state.viajeId ? 'Guardar viaje' : 'Crear viaje'}
              </button>
            )}
            <button type="button" style={layoutStyles.defaultButton} onClick={this._onCancelarViaje} disabled={this.state.guardando}>Cancelar</button>
          </div>
        </div>
      </div>
    );
  }
}

