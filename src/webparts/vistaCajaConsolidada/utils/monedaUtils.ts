import { IImportesMoneda } from '../models/IImportesMoneda';
import { ICajaConsolidadaItem } from '../models/ICajaConsolidadaItem';

export function createEmptyImportesMoneda(): IImportesMoneda {
  return { ars: 0, usd: 0 };
}

function _sumCampoImportes(
  items: ICajaConsolidadaItem[],
  campo: 'saldoPendiente' | 'totalAbonado' | 'totalRecibido' | 'totalRecupero'
): IImportesMoneda {
  return (items || []).reduce(
    (acc: IImportesMoneda, item: ICajaConsolidadaItem) => {
      const importes = item[campo];
      acc.ars += importes ? importes.ars || 0 : 0;
      acc.usd += importes ? importes.usd || 0 : 0;
      return acc;
    },
    createEmptyImportesMoneda()
  );
}

/**
 * Suma de saldos pendientes de los viajes visibles (por moneda de servicio).
 * No es el Total Caja: el Total Caja usa sumTotalesAbonado (ingresos).
 */
export function sumSaldosPendientes(items: ICajaConsolidadaItem[]): IImportesMoneda {
  return _sumCampoImportes(items, 'saldoPendiente');
}

/**
 * Total Caja / ingresos = suma de total abonado (aplicado a viajes) de las filas visibles.
 */
export function sumTotalesAbonado(items: ICajaConsolidadaItem[]): IImportesMoneda {
  return _sumCampoImportes(items, 'totalAbonado');
}

/**
 * Suma de dinero recibido (Monto) de las filas visibles.
 * No sumar totalRecupero encima: el recupero ya está dentro de Monto.
 */
export function sumTotalesRecibido(items: ICajaConsolidadaItem[]): IImportesMoneda {
  return _sumCampoImportes(items, 'totalRecibido');
}

/** Suma informativa de recupero bancario (no se suma al recibido). */
export function sumTotalesRecupero(items: ICajaConsolidadaItem[]): IImportesMoneda {
  return _sumCampoImportes(items, 'totalRecupero');
}

export function formatImporteArs(valor: number): string {
  const monto = isNaN(valor) ? 0 : valor;
  return (
    '$ ' +
    monto.toLocaleString('es-AR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })
  );
}

export function formatImporteUsd(valor: number): string {
  const monto = isNaN(valor) ? 0 : valor;
  return (
    'USD ' +
    monto.toLocaleString('es-AR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })
  );
}
