import { toDateInput, toSharePointDateOnlyPayload } from './sharePointDateUtils';
import { IRegistroPagoFieldMap } from './registroPagoFieldMap';
import { IRegistroPagoItem, IRegistroPagoPayload } from './registroPagoTypes';

function getString(source: any, key: string): string {
  if (!key) {
    return '';
  }
  const value = source[key];
  return value === null || value === undefined ? '' : String(value);
}

function toNumber(value: any): number {
  const parsed = Number(value);
  return isNaN(parsed) ? 0 : parsed;
}

function extractLookupId(item: any, fieldName: string): number {
  if (!fieldName) {
    return 0;
  }
  const lookupIdKey = fieldName + 'Id';
  const lookupIdValue = item[lookupIdKey];
  if (lookupIdValue !== undefined && lookupIdValue !== null) {
    return Number(lookupIdValue) || 0;
  }
  const lookupValue = item[fieldName] as { Id?: number } | undefined;
  if (lookupValue && lookupValue.Id) {
    return Number(lookupValue.Id) || 0;
  }
  return 0;
}

function extractLookupTitle(item: any, fieldName: string): string {
  if (!fieldName || !item[fieldName]) {
    return '';
  }
  return getString(item[fieldName], 'Title');
}

export function normalizarTipoPagoRegistro(value: string): string {
  return (value || '').trim() === 'Egreso' ? 'Egreso' : 'Ingreso';
}

/**
 * Importe canónico: columna Importe; fallback legacy Monto si el canónico está vacío/0.
 */
export function leerMontoRegistroPago(item: any, fieldMap: IRegistroPagoFieldMap): number {
  const principal = toNumber(item[fieldMap.montoField]);
  if (principal > 0) {
    return principal;
  }
  if (fieldMap.montoLegacyField && fieldMap.montoLegacyField !== fieldMap.montoField) {
    const legacy = toNumber(item[fieldMap.montoLegacyField]);
    if (legacy > 0) {
      return legacy;
    }
  }
  return principal;
}

/**
 * Concepto canónico: Concepto; fallback legacy Title si Concepto vacío o no existe.
 */
export function leerConceptoRegistroPago(item: any, fieldMap: IRegistroPagoFieldMap): string {
  if (fieldMap.hasConceptoField) {
    const concepto = getString(item, fieldMap.Concepto).trim();
    if (concepto) {
      return concepto;
    }
  }
  return getString(item, fieldMap.Title).trim();
}

function setIfFieldExists(
  payload: { [key: string]: string | number | null },
  fieldInternalName: string,
  exists: boolean,
  value: string | number | null
): void {
  if (!exists || !fieldInternalName) {
    return;
  }
  payload[fieldInternalName] = value;
}

function resolveLookupPayloadValue(id: number | null | undefined): number | null | undefined {
  if (id === undefined) {
    return undefined;
  }
  if (id === null) {
    return null;
  }
  return id > 0 ? id : null;
}

/**
 * Construye el body REST para create/update.
 * - `undefined` en data → no se incluye la propiedad (MERGE no pisa el valor existente).
 * - `null` o string vacío (según campo) → limpia explícitamente.
 * - Solo escribe columnas que existen en el field map.
 * - Importe: una sola columna canónica (no escribe Importe y Monto a la vez).
 * - Concepto: columna Concepto si existe; Title queda para título descriptivo.
 * - Fecha: siempre `toSharePointDateOnlyPayload` (T12:00:00Z).
 */
export function buildRegistroPagoPayload(
  data: IRegistroPagoPayload,
  fieldMap: IRegistroPagoFieldMap
): { [key: string]: string | number | null } {
  const payload: { [key: string]: string | number | null } = {};

  if (data.title !== undefined) {
    const title = (data.title || '').trim() || 'Registro de Pago';
    payload[fieldMap.Title] = title;
  }

  if (data.concepto !== undefined) {
    if (fieldMap.hasConceptoField) {
      const concepto = (data.concepto || '').trim();
      // Si Concepto y Title son la misma columna (Title renombrado), el concepto gana.
      payload[fieldMap.Concepto] = concepto || null;
    } else if (data.title === undefined) {
      // Sin columna Concepto: solo entonces Title puede llevar el concepto (legacy).
      const concepto = (data.concepto || '').trim();
      if (concepto) {
        payload[fieldMap.Title] = concepto;
      }
    }
  }

  if (data.monto !== undefined) {
    payload[fieldMap.montoField] = data.monto;
  }

  if (data.tipoPago !== undefined && fieldMap.fieldExists.TipoPago) {
    payload[fieldMap.TipoPago] = normalizarTipoPagoRegistro(data.tipoPago);
  }

  if (data.medioPago !== undefined && fieldMap.fieldExists.MedioPago) {
    payload[fieldMap.MedioPago] = data.medioPago;
  }

  if (data.fechaPago !== undefined && fieldMap.fieldExists.FechaPago) {
    payload[fieldMap.FechaPago] = toSharePointDateOnlyPayload(data.fechaPago);
  }

  if (data.moneda !== undefined && fieldMap.fieldExists.Moneda) {
    payload[fieldMap.Moneda] = data.moneda;
  }

  if (data.cotizacion !== undefined && fieldMap.fieldExists.Cotizacion) {
    if (data.cotizacion !== null && data.cotizacion > 0) {
      payload[fieldMap.Cotizacion] = data.cotizacion;
    } else {
      payload[fieldMap.Cotizacion] = null;
    }
  }

  if (data.viajeAsociadoId !== undefined && fieldMap.fieldExists.ViajeAsociado) {
    const viajeId = resolveLookupPayloadValue(data.viajeAsociadoId);
    if (viajeId !== undefined) {
      payload[fieldMap.ViajeAsociado + 'Id'] = viajeId;
    }
  }

  if (data.servicioViajeId !== undefined && fieldMap.fieldExists.ServicioViaje) {
    const servicioId = resolveLookupPayloadValue(data.servicioViajeId);
    if (servicioId !== undefined) {
      payload[fieldMap.ServicioViaje + 'Id'] = servicioId;
    }
  }

  if (data.pasajeroId !== undefined && fieldMap.fieldExists.Pasajero) {
    const pasajeroId = resolveLookupPayloadValue(data.pasajeroId);
    if (pasajeroId !== undefined) {
      payload[fieldMap.Pasajero + 'Id'] = pasajeroId;
    }
  }

  if (data.estado !== undefined) {
    setIfFieldExists(payload, fieldMap.Estado, fieldMap.fieldExists.Estado, data.estado);
  }

  if (data.banco !== undefined && fieldMap.fieldExists.Banco) {
    const banco = (data.banco || '').trim();
    payload[fieldMap.Banco] = banco || null;
  }

  if (data.motivo !== undefined && fieldMap.fieldExists.Motivo) {
    const motivo = (data.motivo || '').trim();
    payload[fieldMap.Motivo] = motivo || null;
  }

  if (data.liquidacionOperadorId !== undefined && fieldMap.fieldExists.LiquidacionOperador) {
    const liquidacionId = resolveLookupPayloadValue(data.liquidacionOperadorId);
    if (liquidacionId !== undefined) {
      payload[fieldMap.LiquidacionOperador + 'Id'] = liquidacionId;
    }
  }

  if (data.observaciones !== undefined && fieldMap.fieldExists.Observaciones) {
    payload[fieldMap.Observaciones] = data.observaciones || '';
  }

  // TEMP diagnóstico Concepto — quitar cuando se confirme la causa
  console.log('[Concepto pago] fieldMap.hasConceptoField:', fieldMap.hasConceptoField);
  console.log('[Concepto pago] fieldMap.Concepto (internal):', fieldMap.Concepto || '(vacío)');
  console.log('[Concepto pago] data.concepto:', data.concepto);
  console.log('[Concepto pago] data.title:', data.title);
  console.log('[Concepto pago] payload final:', payload);
  if (fieldMap.hasConceptoField) {
    console.log(
      '[Concepto pago] valor escrito en columna Concepto:',
      payload[fieldMap.Concepto]
    );
  } else {
    console.warn(
      '[Concepto pago] No hay columna Concepto en el field map; Title=',
      payload[fieldMap.Title]
    );
  }

  return payload;
}

/**
 * Mapea un ítem REST de SharePoint al modelo compartido de lectura.
 */
export function mapSharePointItemToRegistroPago(
  item: any,
  fieldMap: IRegistroPagoFieldMap
): IRegistroPagoItem {
  const viajeAsociadoId = extractLookupId(item, fieldMap.ViajeAsociado);
  const pasajeroId = fieldMap.fieldExists.Pasajero
    ? extractLookupId(item, fieldMap.Pasajero)
    : 0;
  const pasajeroNombre = fieldMap.fieldExists.Pasajero
    ? extractLookupTitle(item, fieldMap.Pasajero)
    : '';
  const servicioViajeId = fieldMap.fieldExists.ServicioViaje
    ? extractLookupId(item, fieldMap.ServicioViaje)
    : 0;
  const liquidacionOperadorId = fieldMap.fieldExists.LiquidacionOperador
    ? extractLookupId(item, fieldMap.LiquidacionOperador)
    : 0;
  const liquidacionOperadorNombre = fieldMap.fieldExists.LiquidacionOperador
    ? extractLookupTitle(item, fieldMap.LiquidacionOperador)
    : '';

  const tipoPagoLeido = getString(item, fieldMap.TipoPago);
  const tipoPago =
    tipoPagoLeido === 'Egreso' || tipoPagoLeido === 'Ingreso'
      ? tipoPagoLeido
      : liquidacionOperadorId > 0
        ? 'Egreso'
        : 'Ingreso';

  const cotizacionRaw = item[fieldMap.Cotizacion];

  return {
    id: item.Id,
    title: getString(item, fieldMap.Title),
    concepto: leerConceptoRegistroPago(item, fieldMap),
    monto: leerMontoRegistroPago(item, fieldMap),
    tipoPago: normalizarTipoPagoRegistro(tipoPago),
    medioPago: getString(item, fieldMap.MedioPago),
    fechaPago: toDateInput(getString(item, fieldMap.FechaPago)),
    moneda: getString(item, fieldMap.Moneda),
    cotizacion:
      cotizacionRaw !== undefined && cotizacionRaw !== null
        ? toNumber(cotizacionRaw)
        : undefined,
    viajeAsociadoId: viajeAsociadoId > 0 ? viajeAsociadoId : null,
    servicioViajeId: servicioViajeId > 0 ? servicioViajeId : undefined,
    pasajeroId: pasajeroId > 0 ? pasajeroId : null,
    pasajeroNombre: pasajeroNombre || undefined,
    estado: fieldMap.fieldExists.Estado ? getString(item, fieldMap.Estado) : undefined,
    banco: fieldMap.fieldExists.Banco ? getString(item, fieldMap.Banco) : undefined,
    motivo: fieldMap.fieldExists.Motivo ? getString(item, fieldMap.Motivo) : undefined,
    liquidacionOperadorId:
      liquidacionOperadorId > 0 ? liquidacionOperadorId : undefined,
    liquidacionOperadorNombre: liquidacionOperadorNombre || undefined,
    observaciones: fieldMap.fieldExists.Observaciones
      ? getString(item, fieldMap.Observaciones)
      : undefined
  };
}
