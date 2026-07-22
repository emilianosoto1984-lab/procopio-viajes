import {
  REGISTRO_PAGO_EXPECTED_FIELDS,
  RegistroPagoExpectedField
} from './registroPagoTypes';

export interface ISharePointListFieldMeta {
  Title: string;
  InternalName: string;
  TypeAsString?: string;
  Hidden?: boolean;
}

/**
 * Mapa resuelto de columnas de «Registro de Pagos».
 * `montoField` es el canónico de escritura (Importe en esta lista).
 */
export interface IRegistroPagoFieldMap {
  Title: string;
  /** Internal name de Concepto si existe; si no, cadena vacía. */
  Concepto: string;
  hasConceptoField: boolean;
  /**
   * Columna canónica del importe numérico.
   * Preferencia: Importe; solo si no existe, Monto.
   */
  montoField: string;
  montoCanonicalKey: 'Importe' | 'Monto';
  /**
   * Columna alternativa de monto (p. ej. Monto) solo para lectura legacy
   * cuando difiere del canónico.
   */
  montoLegacyField: string;
  hasImporteField: boolean;
  TipoPago: string;
  MedioPago: string;
  FechaPago: string;
  Moneda: string;
  Cotizacion: string;
  ViajeAsociado: string;
  ServicioViaje: string;
  Pasajero: string;
  Estado: string;
  Banco: string;
  Motivo: string;
  LiquidacionOperador: string;
  Observaciones: string;
  /** Existencia real en SharePoint por clave lógica esperada. */
  fieldExists: { [key in RegistroPagoExpectedField]: boolean };
}

const DISPLAY_NAME_ALIASES: { [key: string]: string[] } = {
  Banco: ['Banco', 'Cuenta Bancaria'],
  Motivo: ['Motivo', 'Motivo de elección'],
  Concepto: ['Concepto'],
  Monto: ['Monto'],
  Importe: ['Importe'],
  TipoPago: ['TipoPago', 'Tipo de pago', 'Tipo de Pago'],
  MedioPago: ['MedioPago', 'Medio de pago', 'Medio de Pago'],
  FechaPago: ['FechaPago', 'Fecha de pago', 'Fecha de Pago'],
  Moneda: ['Moneda'],
  Cotizacion: ['Cotizacion', 'Cotización'],
  ViajeAsociado: ['ViajeAsociado', 'Viaje Asociado'],
  ServicioViaje: ['ServicioViaje', 'Servicio Viaje', 'Servicio a abonar'],
  Pasajero: ['Pasajero'],
  Estado: ['Estado'],
  LiquidacionOperador: ['LiquidacionOperador', 'Liquidación Operador', 'Liquidacion Operador'],
  Observaciones: ['Observaciones'],
  Title: ['Title', 'Título']
};

/** Campos calculados / de UI que no aceptan escritura REST real. */
const NON_WRITABLE_FIELD_INTERNALS: { [name: string]: boolean } = {
  LinkTitle: true,
  LinkTitleNoMenu: true,
  LinkFilename: true,
  LinkFilenameNoMenu: true,
  Edit: true,
  DocIcon: true,
  FileLeafRef: true
};

function isWritableTextLikeField(field: ISharePointListFieldMeta): boolean {
  const internal = field.InternalName || '';
  if (NON_WRITABLE_FIELD_INTERNALS[internal]) {
    return false;
  }
  const type = (field.TypeAsString || '').toLowerCase();
  if (
    type === 'computed' ||
    type === 'attachments' ||
    type === 'file' ||
    type.indexOf('lookup') >= 0 ||
    type === 'user' ||
    type === 'taxonomyfieldtype'
  ) {
    return false;
  }
  return true;
}

function findField(
  fields: ISharePointListFieldMeta[],
  logicalKey: string
): ISharePointListFieldMeta | undefined {
  const byInternal = fields.filter(
    (field: ISharePointListFieldMeta) => field.InternalName === logicalKey
  )[0];
  if (byInternal) {
    return byInternal;
  }

  const aliases = DISPLAY_NAME_ALIASES[logicalKey] || [logicalKey];
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    const byTitle = fields.filter(
      (field: ISharePointListFieldMeta) => field.Title === alias
    )[0];
    if (byTitle) {
      return byTitle;
    }
  }

  return undefined;
}

/**
 * Resuelve la columna real de Concepto.
 * En muchas listas Title se renombra a "Concepto" y aparece también LinkTitle
 * con el mismo display name; LinkTitle no persiste valores por REST.
 */
function findConceptoField(
  fields: ISharePointListFieldMeta[],
  titleInternalName: string
): ISharePointListFieldMeta | undefined {
  const exactInternal = fields.filter(
    (field: ISharePointListFieldMeta) =>
      field.InternalName === 'Concepto' && isWritableTextLikeField(field)
  )[0];
  if (exactInternal) {
    return exactInternal;
  }

  const byDisplayName = fields.filter(
    (field: ISharePointListFieldMeta) =>
      field.Title === 'Concepto' && isWritableTextLikeField(field)
  );

  const dedicada = byDisplayName.filter(
    (field: ISharePointListFieldMeta) =>
      field.InternalName !== 'Title' && field.InternalName !== titleInternalName
  )[0];
  if (dedicada) {
    return dedicada;
  }

  // Title renombrado a "Concepto" en la UI de SharePoint.
  const titleComoConcepto = byDisplayName.filter(
    (field: ISharePointListFieldMeta) =>
      field.InternalName === 'Title' || field.InternalName === titleInternalName
  )[0];
  if (titleComoConcepto) {
    console.warn(
      '[Concepto pago] La columna visible "Concepto" es Title renombrado (' +
        titleComoConcepto.InternalName +
        '). Se escribirá el concepto ahí.'
    );
    return titleComoConcepto;
  }

  const rejected = fields.filter((field: ISharePointListFieldMeta) => field.Title === 'Concepto');
  if (rejected.length > 0) {
    console.warn(
      '[Concepto pago] Se ignoró match no escribible de Concepto:',
      rejected.map(
        (field: ISharePointListFieldMeta) =>
          field.InternalName + '/' + (field.TypeAsString || '')
      )
    );
  }

  return undefined;
}

export function isRegistroPagoFieldMapDebugEnabled(): boolean {
  if (typeof window === 'undefined' || !window.location) {
    return false;
  }
  const href = (window.location.href || '').toLowerCase();
  const hostname = (window.location.hostname || '').toLowerCase();
  return (
    hostname === 'localhost' ||
    hostname === '127.0.0.1' ||
    href.indexOf('/workbench') >= 0 ||
    href.indexOf('workbench.aspx') >= 0
  );
}

function logFieldResolution(
  logicalKey: string,
  field: ISharePointListFieldMeta | undefined,
  note?: string
): void {
  if (field) {
    console.log(
      '[RegistroPagoFieldMap] ' +
        logicalKey +
        ' → visible="' +
        field.Title +
        '", internal="' +
        field.InternalName +
        '", type="' +
        (field.TypeAsString || '(n/d)') +
        '", existe=sí' +
        (note ? ' (' + note + ')' : '')
    );
  } else {
    console.warn(
      '[RegistroPagoFieldMap] ' +
        logicalKey +
        ' → existe=no' +
        (note ? ' (' + note + ')' : '')
    );
  }
}

/**
 * Resuelve internal names a partir de la metadata de campos de la lista.
 * Preferencia de importe: Importe (canónico) → Monto (legacy, solo si Importe no existe).
 * Preferencia de concepto: Concepto (canónico) → Title solo como fallback de lectura.
 */
export function resolveRegistroPagoFieldMap(
  fields: ISharePointListFieldMeta[],
  options?: { log?: boolean }
): IRegistroPagoFieldMap {
  const shouldLog =
    options && options.log !== undefined
      ? options.log
      : isRegistroPagoFieldMapDebugEnabled();

  const visibleFields = fields.filter(
    (field: ISharePointListFieldMeta) => field.Hidden !== true
  );

  const titleField =
    fields.filter(
      (field: ISharePointListFieldMeta) => field.InternalName === 'Title' && field.Hidden !== true
    )[0] ||
    visibleFields.filter(
      (field: ISharePointListFieldMeta) => field.InternalName === 'Title'
    )[0] ||
    findField(visibleFields, 'Title') || {
      Title: 'Title',
      InternalName: 'Title',
      TypeAsString: 'Text'
    };
  const titleInternalName = titleField.InternalName || 'Title';
  const conceptoField = findConceptoField(visibleFields, titleInternalName);
  // TEMP diagnóstico Concepto
  if (conceptoField) {
    console.log(
      '[Concepto pago] metadata Concepto → Title="' +
        conceptoField.Title +
        '", InternalName="' +
        conceptoField.InternalName +
        '", TypeAsString="' +
        (conceptoField.TypeAsString || '') +
        '"'
    );
  } else {
    console.warn('[Concepto pago] metadata: no se encontró columna Concepto escribible');
  }
  const montoFieldMeta = findField(visibleFields, 'Monto');
  const importeFieldMeta = findField(visibleFields, 'Importe');
  const tipoPagoField = findField(visibleFields, 'TipoPago');
  const medioPagoField = findField(visibleFields, 'MedioPago');
  const fechaPagoField = findField(visibleFields, 'FechaPago');
  const monedaField = findField(visibleFields, 'Moneda');
  const cotizacionField = findField(visibleFields, 'Cotizacion');
  const viajeField = findField(visibleFields, 'ViajeAsociado');
  const servicioField = findField(visibleFields, 'ServicioViaje');
  const pasajeroField = findField(visibleFields, 'Pasajero');
  const estadoField = findField(visibleFields, 'Estado');
  const bancoField = findField(visibleFields, 'Banco');
  const motivoField = findField(visibleFields, 'Motivo');
  const liquidacionField = findField(visibleFields, 'LiquidacionOperador');
  const observacionesField = findField(visibleFields, 'Observaciones');

  let montoCanonicalKey: 'Importe' | 'Monto' = 'Importe';
  let montoField = 'Importe';
  if (importeFieldMeta) {
    montoCanonicalKey = 'Importe';
    montoField = importeFieldMeta.InternalName;
  } else if (montoFieldMeta) {
    montoCanonicalKey = 'Monto';
    montoField = montoFieldMeta.InternalName;
  }

  const montoInternal = montoFieldMeta ? montoFieldMeta.InternalName : '';
  const hasMontoDistinct = !!montoFieldMeta && montoInternal !== montoField;

  const fieldExists = {} as { [key in RegistroPagoExpectedField]: boolean };
  REGISTRO_PAGO_EXPECTED_FIELDS.forEach((key: RegistroPagoExpectedField) => {
    if (key === 'Title') {
      fieldExists[key] = true;
      return;
    }
    if (key === 'Monto') {
      fieldExists[key] = !!montoFieldMeta;
      return;
    }
    if (key === 'Importe') {
      fieldExists[key] = !!importeFieldMeta;
      return;
    }
    if (key === 'Concepto') {
      fieldExists[key] = !!conceptoField;
      return;
    }
    const resolved = findField(visibleFields, key);
    fieldExists[key] = !!resolved;
  });

  if (shouldLog) {
    console.log('[RegistroPagoFieldMap] Resolviendo columnas de «Registro de Pagos»...');
    logFieldResolution('Title', titleField);
    logFieldResolution(
      'Concepto',
      conceptoField,
      conceptoField
        ? 'canónico para concepto'
        : 'lectura usará Title como fallback legacy'
    );
    logFieldResolution(
      'Importe',
      importeFieldMeta,
      importeFieldMeta ? 'canónico de importe' : undefined
    );
    logFieldResolution(
      'Monto',
      montoFieldMeta,
      !importeFieldMeta && montoFieldMeta
        ? 'usado como canónico porque Importe no existe'
        : hasMontoDistinct
          ? 'solo fallback de lectura legacy'
          : undefined
    );
    logFieldResolution('TipoPago', tipoPagoField);
    logFieldResolution('MedioPago', medioPagoField);
    logFieldResolution('FechaPago', fechaPagoField);
    logFieldResolution('Moneda', monedaField);
    logFieldResolution('Cotizacion', cotizacionField);
    logFieldResolution('ViajeAsociado', viajeField);
    logFieldResolution('ServicioViaje', servicioField);
    logFieldResolution('Pasajero', pasajeroField);
    logFieldResolution('Estado', estadoField);
    logFieldResolution('Banco', bancoField);
    logFieldResolution('Motivo', motivoField);
    logFieldResolution('LiquidacionOperador', liquidacionField);
    logFieldResolution('Observaciones', observacionesField);
    console.log(
      '[RegistroPagoFieldMap] Canónicos → importe=' +
        montoCanonicalKey +
        ' ("' +
        montoField +
        '"), concepto=' +
        (conceptoField ? 'Concepto ("' + conceptoField.InternalName + '")' : 'Title (fallback)')
    );
  }

  return {
    Title: titleField.InternalName || 'Title',
    Concepto: conceptoField ? conceptoField.InternalName : '',
    hasConceptoField: !!conceptoField,
    montoField,
    montoCanonicalKey,
    montoLegacyField: hasMontoDistinct ? montoInternal : '',
    hasImporteField: !!importeFieldMeta,
    TipoPago: tipoPagoField ? tipoPagoField.InternalName : 'TipoPago',
    MedioPago: medioPagoField ? medioPagoField.InternalName : 'MedioPago',
    FechaPago: fechaPagoField ? fechaPagoField.InternalName : 'FechaPago',
    Moneda: monedaField ? monedaField.InternalName : 'Moneda',
    Cotizacion: cotizacionField ? cotizacionField.InternalName : 'Cotizacion',
    ViajeAsociado: viajeField ? viajeField.InternalName : 'ViajeAsociado',
    ServicioViaje: servicioField ? servicioField.InternalName : 'ServicioViaje',
    Pasajero: pasajeroField ? pasajeroField.InternalName : 'Pasajero',
    Estado: estadoField ? estadoField.InternalName : 'Estado',
    Banco: bancoField ? bancoField.InternalName : 'Banco',
    Motivo: motivoField ? motivoField.InternalName : 'Motivo',
    LiquidacionOperador: liquidacionField
      ? liquidacionField.InternalName
      : 'LiquidacionOperador',
    Observaciones: observacionesField
      ? observacionesField.InternalName
      : 'Observaciones',
    fieldExists
  };
}

/** Campos $select útiles para leer un pago completo (incluye fallbacks legacy). */
export function buildRegistroPagoSelectFields(fieldMap: IRegistroPagoFieldMap): string[] {
  const selectFields: string[] = ['Id', fieldMap.Title];

  const pushUnique = (name: string): void => {
    if (name && selectFields.indexOf(name) < 0) {
      selectFields.push(name);
    }
  };

  if (fieldMap.hasConceptoField) {
    pushUnique(fieldMap.Concepto);
  }
  pushUnique(fieldMap.montoField);
  if (fieldMap.montoLegacyField) {
    pushUnique(fieldMap.montoLegacyField);
  }
  pushUnique(fieldMap.TipoPago);
  pushUnique(fieldMap.MedioPago);
  pushUnique(fieldMap.FechaPago);
  pushUnique(fieldMap.Moneda);
  pushUnique(fieldMap.Cotizacion);

  if (fieldMap.fieldExists.Observaciones) {
    pushUnique(fieldMap.Observaciones);
  }
  if (fieldMap.fieldExists.Estado) {
    pushUnique(fieldMap.Estado);
  }
  if (fieldMap.fieldExists.Banco) {
    pushUnique(fieldMap.Banco);
  }
  if (fieldMap.fieldExists.Motivo) {
    pushUnique(fieldMap.Motivo);
  }

  const viaje = fieldMap.ViajeAsociado;
  pushUnique(viaje + '/Id');
  pushUnique(viaje + '/Title');

  if (fieldMap.fieldExists.Pasajero) {
    pushUnique(fieldMap.Pasajero + '/Id');
    pushUnique(fieldMap.Pasajero + '/Title');
  }
  if (fieldMap.fieldExists.ServicioViaje) {
    pushUnique(fieldMap.ServicioViaje + '/Id');
    pushUnique(fieldMap.ServicioViaje + '/Title');
  }
  if (fieldMap.fieldExists.LiquidacionOperador) {
    pushUnique(fieldMap.LiquidacionOperador + '/Id');
    pushUnique(fieldMap.LiquidacionOperador + '/Title');
  }

  return selectFields;
}

export function buildRegistroPagoExpandFields(fieldMap: IRegistroPagoFieldMap): string[] {
  const expand: string[] = [fieldMap.ViajeAsociado];
  if (fieldMap.fieldExists.Pasajero) {
    expand.push(fieldMap.Pasajero);
  }
  if (fieldMap.fieldExists.ServicioViaje) {
    expand.push(fieldMap.ServicioViaje);
  }
  if (fieldMap.fieldExists.LiquidacionOperador) {
    expand.push(fieldMap.LiquidacionOperador);
  }
  return expand.filter(
    (field: string, index: number, fields: string[]) => fields.indexOf(field) === index
  );
}
