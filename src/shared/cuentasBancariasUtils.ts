import { ISharePointListFieldMeta } from './registroPagoFieldMap';

/**
 * Modelo de un ítem de la lista SharePoint «CuentasBancarias».
 * Los nombres de columnas se resuelven en runtime (InternalName / Title).
 */
export interface ICuentaBancaria {
  id: number;
  title: string;
  banco: string;
  moneda: string;
  cuenta: string;
  cbu: string;
  alias: string;
  aplicaRecupero: boolean;
  porcentajeRecupero: number;
  activo: boolean;
}

export interface ICuentasBancariasFieldMap {
  Title: string;
  Banco: string;
  Moneda: string;
  Cuenta: string;
  CBU: string;
  Alias: string;
  AplicaRecupero: string;
  PorcentajeRecupero: string;
  Activo: string;
}

const FIELD_ALIASES: { [key: string]: string[] } = {
  Title: ['Title', 'Título'],
  Banco: ['Banco'],
  Moneda: ['Moneda'],
  Cuenta: ['Cuenta', 'Número de cuenta', 'Numero de cuenta', 'Nro Cuenta', 'Nro. Cuenta'],
  CBU: ['CBU', 'Cbu'],
  Alias: ['Alias'],
  AplicaRecupero: ['AplicaRecupero', 'Aplica Recupero', 'Aplica recupero'],
  PorcentajeRecupero: [
    'PorcentajeRecupero',
    'Porcentaje Recupero',
    'Porcentaje de recupero',
    '% Recupero'
  ],
  Activo: ['Activo']
};

function findFieldByAliases(
  fields: ISharePointListFieldMeta[],
  logicalKey: string
): ISharePointListFieldMeta | undefined {
  const aliases = FIELD_ALIASES[logicalKey] || [logicalKey];
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    const byInternal = fields.filter(
      (field: ISharePointListFieldMeta) => field.InternalName === alias
    )[0];
    if (byInternal) {
      return byInternal;
    }
  }
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

export function resolveCuentasBancariasFieldMap(
  fields: ISharePointListFieldMeta[]
): ICuentasBancariasFieldMap {
  const visible = (fields || []).filter((field: ISharePointListFieldMeta) => !field.Hidden);
  const keys: (keyof ICuentasBancariasFieldMap)[] = [
    'Title',
    'Banco',
    'Moneda',
    'Cuenta',
    'CBU',
    'Alias',
    'AplicaRecupero',
    'PorcentajeRecupero',
    'Activo'
  ];
  const map: ICuentasBancariasFieldMap = {
    Title: 'Title',
    Banco: '',
    Moneda: '',
    Cuenta: '',
    CBU: '',
    Alias: '',
    AplicaRecupero: '',
    PorcentajeRecupero: '',
    Activo: ''
  };

  console.log('[CuentasBancarias] Resolviendo columnas de la lista...');
  keys.forEach((key: keyof ICuentasBancariasFieldMap) => {
    const resolved = findFieldByAliases(visible, key);
    map[key] = resolved ? resolved.InternalName : key === 'Title' ? 'Title' : '';
    if (resolved) {
      console.log(
        '[CuentasBancarias] ' +
          key +
          ' → "' +
          resolved.InternalName +
          '" (Title="' +
          resolved.Title +
          '"' +
          (resolved.TypeAsString ? ', Type=' + resolved.TypeAsString : '') +
          ')'
      );
    } else if (key !== 'Title') {
      console.warn('[CuentasBancarias] No se encontró columna para ' + key);
    }
  });

  return map;
}

export function buildCuentasBancariasSelectFields(map: ICuentasBancariasFieldMap): string[] {
  const select = ['Id', map.Title || 'Title'];
  const optional: (keyof ICuentasBancariasFieldMap)[] = [
    'Banco',
    'Moneda',
    'Cuenta',
    'CBU',
    'Alias',
    'AplicaRecupero',
    'PorcentajeRecupero',
    'Activo'
  ];
  optional.forEach((key: keyof ICuentasBancariasFieldMap) => {
    if (map[key] && select.indexOf(map[key]) < 0) {
      select.push(map[key]);
    }
  });
  return select;
}

function toBooleanCuenta(value: any): boolean {
  if (value === true || value === 1 || value === '1') {
    return true;
  }
  const text = String(value || '')
    .trim()
    .toLowerCase();
  return text === 'true' || text === 'sí' || text === 'si' || text === 'yes';
}

function toNumberCuenta(value: any): number {
  const parsed = Number(value);
  return isNaN(parsed) ? 0 : parsed;
}

function toStringCuenta(item: any, fieldInternalName: string): string {
  if (!fieldInternalName) {
    return '';
  }
  const value = item[fieldInternalName];
  return value === null || value === undefined ? '' : String(value);
}

export function mapSharePointItemToCuentaBancaria(
  item: any,
  map: ICuentasBancariasFieldMap
): ICuentaBancaria {
  return {
    id: Number(item.Id) || 0,
    title: toStringCuenta(item, map.Title || 'Title'),
    banco: toStringCuenta(item, map.Banco),
    moneda: toStringCuenta(item, map.Moneda),
    cuenta: toStringCuenta(item, map.Cuenta),
    cbu: toStringCuenta(item, map.CBU),
    alias: toStringCuenta(item, map.Alias),
    aplicaRecupero: map.AplicaRecupero ? toBooleanCuenta(item[map.AplicaRecupero]) : false,
    porcentajeRecupero: map.PorcentajeRecupero
      ? toNumberCuenta(item[map.PorcentajeRecupero])
      : 0,
    activo: map.Activo ? toBooleanCuenta(item[map.Activo]) : true
  };
}

/**
 * Label visible del selector. No usar como identificador persistido.
 */
export function getCuentaBancariaLabel(cuenta: ICuentaBancaria): string {
  const title = (cuenta.title || '').trim();
  if (title) {
    return title;
  }
  const banco = (cuenta.banco || '').trim();
  const moneda = (cuenta.moneda || '').trim();
  const nroCuenta = (cuenta.cuenta || '').trim();
  if (banco && moneda) {
    return banco + ' - ' + moneda;
  }
  if (banco && nroCuenta) {
    return banco + ' - ' + nroCuenta;
  }
  return banco || nroCuenta || (cuenta.id > 0 ? 'Cuenta #' + cuenta.id : '');
}

export function findCuentaBancariaById(
  cuentas: ICuentaBancaria[],
  id: number
): ICuentaBancaria | undefined {
  if (!id || id <= 0 || !cuentas || cuentas.length === 0) {
    return undefined;
  }
  return cuentas.filter((cuenta: ICuentaBancaria) => cuenta.id === id)[0];
}

function normalizarTextoCuenta(value: string): string {
  return (value || '').trim().toLowerCase();
}

export type CoincidenciaHistoricaBanco = 'unica' | 'sin_coincidencia' | 'ambiguo';

export interface ICoincidenciaHistoricaCuentaBancaria {
  cuenta?: ICuentaBancaria;
  coincidencias: number;
  estado: CoincidenciaHistoricaBanco;
}

/**
 * Fallback temporal: resuelve una cuenta a partir del Choice/texto histórico `Banco`.
 * Solo acepta coincidencia única. No usar como lógica normal (el Lookup por Id es la fuente de verdad).
 */
export function resolverCuentaBancariaHistorica(
  cuentas: ICuentaBancaria[],
  valorGuardado: string
): ICoincidenciaHistoricaCuentaBancaria {
  const valor = normalizarTextoCuenta(valorGuardado);
  if (!valor || !cuentas || cuentas.length === 0) {
    return { coincidencias: 0, estado: 'sin_coincidencia' };
  }

  const encontradas: ICuentaBancaria[] = [];
  const seen: { [id: number]: boolean } = {};
  const agregar = (cuenta: ICuentaBancaria | undefined): void => {
    if (!cuenta || seen[cuenta.id]) {
      return;
    }
    seen[cuenta.id] = true;
    encontradas.push(cuenta);
  };

  cuentas.forEach((cuenta: ICuentaBancaria) => {
    if (normalizarTextoCuenta(cuenta.title) === valor) {
      agregar(cuenta);
    }
  });
  cuentas.forEach((cuenta: ICuentaBancaria) => {
    if (normalizarTextoCuenta(getCuentaBancariaLabel(cuenta)) === valor) {
      agregar(cuenta);
    }
  });
  cuentas.forEach((cuenta: ICuentaBancaria) => {
    if (normalizarTextoCuenta(cuenta.banco) === valor) {
      agregar(cuenta);
    }
  });

  if (encontradas.length === 1) {
    return { cuenta: encontradas[0], coincidencias: 1, estado: 'unica' };
  }
  if (encontradas.length > 1) {
    return { coincidencias: encontradas.length, estado: 'ambiguo' };
  }
  return { coincidencias: 0, estado: 'sin_coincidencia' };
}

/** @deprecated Usar resolverCuentaBancariaHistorica. Solo coincidencia única. */
export function findCuentaBancariaByValorGuardado(
  cuentas: ICuentaBancaria[],
  valorGuardado: string
): ICuentaBancaria | undefined {
  return resolverCuentaBancariaHistorica(cuentas, valorGuardado).cuenta;
}

export interface IPlanMigracionCuentaBancariaItem {
  pagoId: number;
  bancoHistorico: string;
  cuentaBancariaId?: number;
  estado: 'migrable' | 'sin_coincidencia' | 'ambiguo' | 'ya_migrado';
}

/**
 * Estrategia de migración (NO ejecutar escritura).
 * Solo propone Id cuando hay exactamente una coincidencia histórica.
 */
export function planificarMigracionCuentaBancaria(
  pagos: { id: number; cuentaBancariaId?: number | null; banco?: string }[],
  cuentas: ICuentaBancaria[]
): IPlanMigracionCuentaBancariaItem[] {
  return (pagos || []).map((pago) => {
    if (pago.cuentaBancariaId && pago.cuentaBancariaId > 0) {
      return {
        pagoId: pago.id,
        bancoHistorico: (pago.banco || '').trim(),
        cuentaBancariaId: pago.cuentaBancariaId,
        estado: 'ya_migrado' as const
      };
    }
    const bancoHistorico = (pago.banco || '').trim();
    if (!bancoHistorico) {
      return {
        pagoId: pago.id,
        bancoHistorico: '',
        estado: 'sin_coincidencia' as const
      };
    }
    const match = resolverCuentaBancariaHistorica(cuentas, bancoHistorico);
    if (match.estado === 'unica' && match.cuenta) {
      return {
        pagoId: pago.id,
        bancoHistorico,
        cuentaBancariaId: match.cuenta.id,
        estado: 'migrable' as const
      };
    }
    return {
      pagoId: pago.id,
      bancoHistorico,
      estado: match.estado === 'ambiguo' ? 'ambiguo' : 'sin_coincidencia'
    };
  });
}
