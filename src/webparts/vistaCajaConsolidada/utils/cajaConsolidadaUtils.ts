import { EstadoCajaConsolidada } from '../models/ICajaConsolidadaItem';
import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';
import { IImportesMoneda } from '../models/IImportesMoneda';

export function createEmptyFiltrosCajaConsolidada(): IFiltrosCajaConsolidada {
  return {
    textoBusqueda: '',
    fechaSalidaDesde: '',
    fechaSalidaHasta: '',
    vendedor: '',
    estadoFinanciero: '',
    mostrarSaldosEnCero: false
  };
}

export function isSaldoPendienteEnCero(saldo: IImportesMoneda): boolean {
  return (saldo.ars || 0) <= 0.009 && (saldo.usd || 0) <= 0.009;
}

/** True si hay importe distinto de cero en alguna moneda. */
export function tieneImporteMoneda(importes?: IImportesMoneda): boolean {
  if (!importes) {
    return false;
  }
  return Math.abs(importes.ars || 0) > 0.009 || Math.abs(importes.usd || 0) > 0.009;
}

export function getEstadoCajaLabel(estado: EstadoCajaConsolidada): string {
  switch (estado) {
    case 'AlDia':
      return 'Al día';
    case 'Parcial':
      return 'Parcial';
    case 'Pendiente':
      return 'Pendiente';
    case 'SinMovimientos':
      return 'Sin movimientos';
    default:
      return 'Desconocido';
  }
}

function _tieneImporte(importes: IImportesMoneda): boolean {
  return Math.abs(importes.ars || 0) > 0.009 || Math.abs(importes.usd || 0) > 0.009;
}

/**
 * Deriva el estado financiero a partir de totales y saldo.
 */
export function resolveEstadoCajaFinanciero(
  totalViaje: IImportesMoneda,
  totalAbonado: IImportesMoneda,
  saldoPendiente: IImportesMoneda
): EstadoCajaConsolidada {
  if (!_tieneImporte(totalViaje) && !_tieneImporte(totalAbonado)) {
    return 'SinMovimientos';
  }
  if (isSaldoPendienteEnCero(saldoPendiente) && _tieneImporte(totalAbonado)) {
    return 'AlDia';
  }
  if (_tieneImporte(totalAbonado) && !isSaldoPendienteEnCero(saldoPendiente)) {
    return 'Parcial';
  }
  if (!_tieneImporte(totalAbonado) && _tieneImporte(totalViaje)) {
    return 'Pendiente';
  }
  return 'Desconocido';
}
