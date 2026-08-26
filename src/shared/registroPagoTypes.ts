/**
 * Modelo compartido para ítems de la lista SharePoint «Registro de Pagos».
 * Usado por Registro de Pagos y Registro de Viajes.
 */

/** Datos para crear/actualizar. `undefined` = no incluir en el payload (no pisar en MERGE). */
export interface IRegistroPagoPayload {
  /** Título descriptivo del ítem (columna Title). No sustituye a Concepto. */
  title?: string;
  concepto?: string | null;
  monto?: number;
  tipoPago?: string;
  medioPago?: string;
  /** Fecha solo-día en formato yyyy-MM-dd */
  fechaPago?: string;
  moneda?: string;
  cotizacion?: number | null;
  viajeAsociadoId?: number | null;
  servicioViajeId?: number | null;
  pasajeroId?: number | null;
  estado?: string;
  /**
   * Lookup a CuentasBancarias. `null` limpia el valor.
   * No usar el Choice Banco como identificador.
   */
  cuentaBancariaId?: number | null;
  /**
   * Choice/texto legacy. Solo lectura histórica / Viajes.
   * RegistroPagosForm ya no lo escribe.
   */
  banco?: string | null;
  /** Neto aplicado al viaje/servicio (Monto ya incluye recupero). */
  montoAplicadoViaje?: number | null;
  /** Porción del Monto correspondiente a recupero bancario. */
  montoGastosBancarios?: number | null;
  /** % histórico aplicado al crear/editar el pago. */
  porcentajeRecupero?: number | null;
  motivo?: string | null;
  liquidacionOperadorId?: number | null;
  observaciones?: string | null;
}

export interface IRegistroPagoItem {
  id: number;
  title: string;
  concepto: string;
  monto: number;
  tipoPago: string;
  medioPago: string;
  fechaPago: string;
  moneda: string;
  cotizacion?: number;
  viajeAsociadoId?: number | null;
  servicioViajeId?: number;
  pasajeroId?: number | null;
  /** Nombre del lookup Pasajero (solo lectura / UI). */
  pasajeroNombre?: string;
  estado?: string;
  /** Lookup CuentasBancarias (fuente de verdad). */
  cuentaBancariaId?: number | null;
  /** Título del lookup (solo lectura / UI). */
  cuentaBancariaTitulo?: string;
  /** Choice/texto legacy. Solo fallback histórico. */
  banco?: string;
  montoAplicadoViaje?: number | null;
  montoGastosBancarios?: number | null;
  porcentajeRecupero?: number | null;
  motivo?: string;
  liquidacionOperadorId?: number;
  liquidacionOperadorNombre?: string;
  observaciones?: string;
}

/** Claves lógicas esperadas en la lista Registro de Pagos. */
export const REGISTRO_PAGO_EXPECTED_FIELDS = [
  'Title',
  'Concepto',
  'Monto',
  'Importe',
  'TipoPago',
  'MedioPago',
  'FechaPago',
  'Moneda',
  'Cotizacion',
  'ViajeAsociado',
  'ServicioViaje',
  'Pasajero',
  'Estado',
  'Banco',
  'CuentaBancaria',
  'MontoAplicadoViaje',
  'MontoGastosBancarios',
  'PorcentajeRecupero',
  'Motivo',
  'LiquidacionOperador',
  'Observaciones'
] as const;

export type RegistroPagoExpectedField = (typeof REGISTRO_PAGO_EXPECTED_FIELDS)[number];
