import { getDateOnlyFromSharePoint } from './sharePointDateUtils';
import { normalizarMonedaPago } from './pagoMonedaUtils';
import { formatPorcentajeRecuperoDisplay } from './pagoRecuperoUtils';
import { requiereCuentaBancaria, normalizarMedioPago } from './pagoMedioUtils';

export type MotivoGeneracionRecibo = 'creacion' | 'aprobacion' | 'edicion';

/**
 * Reglas de cuándo generar o regenerar el recibo PDF (Registro de Pagos / Viajes).
 * - Egreso: nunca
 * - Creación: solo Ingreso con Estado Aprobado
 * - Aprobación: solo Ingreso
 * - Edición: solo Ingreso con Estado final Aprobado (regenera PDF existente)
 */
export function debeGenerarReciboPago(params: {
  tipoPago: string;
  estado: string;
  motivo: MotivoGeneracionRecibo;
}): boolean {
  const tipo = (params.tipoPago || '').trim();
  if (tipo !== 'Ingreso') {
    return false;
  }
  if (params.motivo === 'aprobacion') {
    return true;
  }
  // creacion y edicion: solo si el estado final es Aprobado
  return (params.estado || '').trim().toLowerCase() === 'aprobado';
}

export function buildNumeroRecibo(itemId: number, fechaPago: string): string {
  const year = _extractYearFromFechaPago(fechaPago);
  return 'RP-' + year + '-' + itemId;
}

export function buildReciboFileName(numeroRecibo: string): string {
  return 'Recibo-' + numeroRecibo + '.pdf';
}

export function formatMontoRecibo(monto: number, moneda: string): string {
  const valor = isNaN(monto) ? 0 : monto;
  const formatted = valor.toLocaleString('es-AR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return normalizarMonedaPago(moneda) === 'Dólares' ? 'USD ' + formatted : '$ ' + formatted;
}

/**
 * Label de cuenta para el PDF.
 * Preferir el label del formulario; Title Lookup; fallback Choice Banco histórico.
 */
export function resolverCuentaBancariaParaRecibo(params: {
  medioPago: string;
  cuentaLabel?: string | null;
  cuentaBancariaTitulo?: string | null;
  bancoHistorico?: string | null;
}): string {
  if (!requiereCuentaBancaria(normalizarMedioPago(params.medioPago))) {
    return '';
  }
  const label = (params.cuentaLabel || '').trim();
  if (label) {
    return label;
  }
  const titulo = (params.cuentaBancariaTitulo || '').trim();
  if (titulo) {
    return titulo;
  }
  return (params.bancoHistorico || '').trim();
}

/** True solo con recupero persistido (no recalcula). */
export function tieneDesgloseRecuperoRecibo(
  montoGastosBancarios?: number | null
): boolean {
  const gastos = Number(montoGastosBancarios);
  return isFinite(gastos) && gastos > 0;
}

function _escapeHtmlRecibo(value: string): string {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function _buildFilaRecibo(label: string, valueHtml: string, valueClass?: string): string {
  const cls = valueClass ? ' recibo-value ' + valueClass : ' recibo-value';
  return (
    '<tr>' +
    '<td class="recibo-label">' +
    _escapeHtmlRecibo(label) +
    '</td>' +
    '<td class="' +
    cls.trim() +
    '">' +
    valueHtml +
    '</td>' +
    '</tr>'
  );
}

/** Fila opcional de cuenta bancaria; vacío si no aplica (sin fila vacía). */
export function buildFilaCuentaBancariaReciboHtml(cuentaBancaria: string): string {
  const label = (cuentaBancaria || '').trim();
  if (!label) {
    return '';
  }
  return _buildFilaRecibo('Cuenta bancaria', _escapeHtmlRecibo(label));
}

/**
 * Cotización en pesos (ARS por 1 USD). Solo se muestra si fue ingresada (> 0).
 */
export function formatCotizacionRecibo(cotizacion?: number | null): string {
  const valor = Number(cotizacion);
  if (!isFinite(valor) || valor <= 0) {
    return '';
  }
  return (
    '$ ' +
    valor.toLocaleString('es-AR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }) +
    ' por USD'
  );
}

/** Fila opcional de cotización; vacío si no hay valor. */
export function buildFilaCotizacionReciboHtml(cotizacion?: number | null): string {
  const texto = formatCotizacionRecibo(cotizacion);
  if (!texto) {
    return '';
  }
  return _buildFilaRecibo('Cotización', _escapeHtmlRecibo(texto));
}

/**
 * Filas de importe: Monto simple, o desglose con valores ya persistidos.
 * No recalcula recupero.
 */
export function buildFilasImporteReciboHtml(params: {
  monto: number;
  moneda: string;
  montoAplicadoViaje?: number | null;
  montoGastosBancarios?: number | null;
  porcentajeRecupero?: number | null;
}): string {
  const moneda = params.moneda || '';
  const montoRecibido = Number(params.monto) || 0;

  if (!tieneDesgloseRecuperoRecibo(params.montoGastosBancarios)) {
    return _buildFilaRecibo(
      'Monto',
      _escapeHtmlRecibo(formatMontoRecibo(montoRecibido, moneda)),
      'monto'
    );
  }

  const aplicado = Number(params.montoAplicadoViaje);
  const gastos = Number(params.montoGastosBancarios);
  const porcentaje = Number(params.porcentajeRecupero) || 0;
  const aplicadoSeguro = isFinite(aplicado) ? aplicado : montoRecibido;
  const gastosSeguros = isFinite(gastos) ? gastos : 0;

  return (
    _buildFilaRecibo(
      'Importe recibido',
      _escapeHtmlRecibo(formatMontoRecibo(montoRecibido, moneda)),
      'monto'
    ) +
    _buildFilaRecibo(
      'Aplicado al viaje',
      _escapeHtmlRecibo(formatMontoRecibo(aplicadoSeguro, moneda))
    ) +
    _buildFilaRecibo(
      'Recupero de gastos bancarios (' +
        formatPorcentajeRecuperoDisplay(porcentaje) +
        '%)',
      _escapeHtmlRecibo(formatMontoRecibo(gastosSeguros, moneda))
    )
  );
}

function _extractYearFromFechaPago(fechaPago: string): number {
  const datePart = getDateOnlyFromSharePoint(fechaPago);
  if (datePart.length >= 4) {
    const year = parseInt(datePart.substring(0, 4), 10);
    if (!isNaN(year)) {
      return year;
    }
  }
  return new Date().getFullYear();
}
