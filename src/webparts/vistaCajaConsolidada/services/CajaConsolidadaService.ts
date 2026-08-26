import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  buildRegistroPagoExpandFields,
  buildRegistroPagoSelectFields,
  IRegistroPagoFieldMap,
  ISharePointListFieldMeta,
  resolveRegistroPagoFieldMap
} from '../../../shared/registroPagoFieldMap';
import { mapSharePointItemToRegistroPago } from '../../../shared/registroPagoPayload';
import { toDateInput } from '../../../shared/sharePointDateUtils';
import {
  getSaldosPendientesPorMoneda,
  getTotalesIngresosPorMoneda,
  getTotalesServiciosPorMoneda
} from '../../../shared/pagoTotalesUtils';
import { ICajaConsolidadaItem } from '../models/ICajaConsolidadaItem';
import { ICajaConsolidadaVistaData } from '../models/ICajaConsolidadaVistaData';
import { IFiltrosCajaConsolidada } from '../models/IFiltrosCajaConsolidada';
import { IImportesMoneda } from '../models/IImportesMoneda';
import { isSaldoPendienteEnCero, resolveEstadoCajaFinanciero, tieneImporteMoneda } from '../utils/cajaConsolidadaUtils';
import {
  calcularTotalesGeneralesCaja,
  IPagoParaCaja,
  listarEgresosSinViaje,
  sumEgresosAsociadosPorMonedaPago,
  sumTotalRecibidoPorMonedaPago,
  sumTotalRecuperoPorMonedaPago
} from '../utils/cajaMovimientosUtils';
import { createEmptyImportesMoneda } from '../utils/monedaUtils';
import { ICajaConsolidadaService } from './ICajaConsolidadaService';

const LISTA_VIAJES = 'Registro de Viajes';
const LISTA_PAGOS = 'Registro de Pagos';
const LISTA_SERVICIOS_VIAJE = 'ServiciosViaje';

interface IViajeCajaRow {
  id: number;
  nombre: string;
  destino: string;
  fechaSalida: string;
}

interface IServicioCajaRow {
  id: number;
  viajeId: number;
  precioCliente: number;
  moneda: string;
}

type IPagoCajaRow = IPagoParaCaja;

type IStringMap = { [key: string]: string };

/**
 * Servicio de datos reales para Vista Caja Consolidada.
 * Lee Viajes, ServiciosViaje y Pagos; calcula totales con helpers compartidos.
 */
export default class CajaConsolidadaService implements ICajaConsolidadaService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webUrl: string;
  private _registroPagoFieldMap: IRegistroPagoFieldMap | undefined;
  private _viajesFieldMap: IStringMap | undefined;
  private _serviciosFieldMap: IStringMap | undefined;

  public constructor(context: WebPartContext) {
    this._spHttpClient = context.spHttpClient;
    this._webUrl = (context.pageContext.web.absoluteUrl || '').replace(/\/$/, '');
  }

  public async getVistaData(filtros?: IFiltrosCajaConsolidada): Promise<ICajaConsolidadaVistaData> {
    const [viajes, servicios, pagos] = await Promise.all([
      this._getViajes(),
      this._getServicios(),
      this._getPagos()
    ]);

    const serviciosPorViaje: { [viajeId: number]: IServicioCajaRow[] } = {};
    servicios.forEach((servicio: IServicioCajaRow) => {
      if (!serviciosPorViaje[servicio.viajeId]) {
        serviciosPorViaje[servicio.viajeId] = [];
      }
      serviciosPorViaje[servicio.viajeId].push(servicio);
    });

    const pagosPorViaje: { [viajeId: number]: IPagoCajaRow[] } = {};
    pagos.forEach((pago: IPagoCajaRow) => {
      const viajeId = pago.viajeId && pago.viajeId > 0 ? pago.viajeId : 0;
      if (!(viajeId > 0)) {
        return;
      }
      if (!pagosPorViaje[viajeId]) {
        pagosPorViaje[viajeId] = [];
      }
      pagosPorViaje[viajeId].push(pago);
    });

    let items = viajes.map((viaje: IViajeCajaRow) => {
      const serviciosViaje = serviciosPorViaje[viaje.id] || [];
      const pagosViaje = pagosPorViaje[viaje.id] || [];
      // Ingresos aplicados: MontoAplicadoViaje ?? Monto (moneda del servicio).
      const totalViaje = getTotalesServiciosPorMoneda(serviciosViaje);
      const ingresosAplicados = getTotalesIngresosPorMoneda(pagosViaje, serviciosViaje);
      // Cobro pendiente del cliente (NO resta egresos ni recupero).
      const saldoPendiente = getSaldosPendientesPorMoneda(totalViaje, ingresosAplicados);
      // Caja asociada al viaje (moneda del pago). Recupero informativo; no sumar encima de Monto.
      const totalRecibido = sumTotalRecibidoPorMonedaPago(pagosViaje);
      const totalRecupero = sumTotalRecuperoPorMonedaPago(pagosViaje);
      const egresosAsociados = sumEgresosAsociadosPorMonedaPago(pagosViaje);
      // Resultado económico del viaje: ingresos aplicados - egresos (por bucket USD/ARS).
      const resultadoViaje: IImportesMoneda = {
        usd: (ingresosAplicados.usd || 0) - (egresosAsociados.usd || 0),
        ars: (ingresosAplicados.ars || 0) - (egresosAsociados.ars || 0)
      };

      return {
        viajeId: viaje.id,
        nombreViaje: viaje.nombre,
        destino: viaje.destino,
        fechaSalida: viaje.fechaSalida,
        vendedor: '',
        totalViaje: this._toImportesMoneda(totalViaje),
        totalAbonado: this._toImportesMoneda(ingresosAplicados),
        egresosAsociados: this._toImportesMoneda(egresosAsociados),
        totalRecibido: this._toImportesMoneda(totalRecibido),
        totalRecupero: this._toImportesMoneda(totalRecupero),
        saldoPendiente: this._toImportesMoneda(saldoPendiente),
        resultadoViaje: this._toImportesMoneda(resultadoViaje),
        estadoFinanciero: resolveEstadoCajaFinanciero(
          this._toImportesMoneda(totalViaje),
          this._toImportesMoneda(ingresosAplicados),
          this._toImportesMoneda(saldoPendiente)
        )
      } as ICajaConsolidadaItem;
    });

    items = this._aplicarFiltrosLocales(items, filtros);
    items.sort((a: ICajaConsolidadaItem, b: ICajaConsolidadaItem) => {
      const fa = a.fechaSalida || '';
      const fb = b.fechaSalida || '';
      if (fa === fb) {
        return (a.nombreViaje || '').localeCompare(b.nombreViaje || '', 'es');
      }
      return fa < fb ? 1 : -1;
    });

    // Totales generales desde TODOS los pagos (incluye egresos sin viaje).
    // Anti doble-suma: totalRecibido = sum(Monto).
    const totalesGenerales = calcularTotalesGeneralesCaja(pagos);
    const egresosSinViaje = listarEgresosSinViaje(pagos);

    return {
      items,
      egresosSinViaje,
      totalesGenerales,
      totalIngresosAplicados: this._sumCampoItems(items, 'totalAbonado'),
      totalEgresosAsociadosViajes: this._sumCampoItems(items, 'egresosAsociados'),
      totalSaldoPendienteCobro: this._sumCampoItems(items, 'saldoPendiente')
    };
  }

  private _sumCampoItems(
    items: ICajaConsolidadaItem[],
    campo: 'totalAbonado' | 'egresosAsociados' | 'saldoPendiente'
  ): IImportesMoneda {
    return (items || []).reduce(
      (acc: IImportesMoneda, item: ICajaConsolidadaItem) => {
        const importes = item[campo];
        acc.usd += importes ? importes.usd || 0 : 0;
        acc.ars += importes ? importes.ars || 0 : 0;
        return acc;
      },
      createEmptyImportesMoneda()
    );
  }

  private _toImportesMoneda(totales: { usd: number; ars: number }): IImportesMoneda {
    return { usd: totales.usd || 0, ars: totales.ars || 0 };
  }

  private _aplicarFiltrosLocales(
    items: ICajaConsolidadaItem[],
    filtros?: IFiltrosCajaConsolidada
  ): ICajaConsolidadaItem[] {
    const texto = ((filtros && filtros.textoBusqueda) || '').trim().toLowerCase();
    const desde = ((filtros && filtros.fechaSalidaDesde) || '').trim();
    const hasta = ((filtros && filtros.fechaSalidaHasta) || '').trim();
    const estado = filtros ? filtros.estadoFinanciero : '';
    const mostrarSaldosEnCero = !!(filtros && filtros.mostrarSaldosEnCero);

    return items.filter((item: ICajaConsolidadaItem) => {
      // Por defecto se ocultan saldos en 0, pero se conservan viajes con recupero o egresos.
      if (
        !mostrarSaldosEnCero &&
        isSaldoPendienteEnCero(item.saldoPendiente) &&
        !tieneImporteMoneda(item.totalRecupero) &&
        !tieneImporteMoneda(item.egresosAsociados)
      ) {
        return false;
      }
      if (texto) {
        const hay =
          (item.nombreViaje || '').toLowerCase().indexOf(texto) >= 0 ||
          (item.destino || '').toLowerCase().indexOf(texto) >= 0;
        if (!hay) {
          return false;
        }
      }
      if (desde && (item.fechaSalida || '') < desde) {
        return false;
      }
      if (hasta && (item.fechaSalida || '') > hasta) {
        return false;
      }
      if (estado && item.estadoFinanciero !== estado) {
        return false;
      }
      return true;
    });
  }

  private async _getViajes(): Promise<IViajeCajaRow[]> {
    const map = await this._getViajesFieldMap();
    const destinoLookup = map.Destino || 'Destino';
    const selectFields = [
      'Id',
      map.Nombre || 'Title',
      map.FechaSalida || 'FechaSalida',
      destinoLookup + '/Id',
      destinoLookup + '/Title'
    ];
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      LISTA_VIAJES +
      "')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$expand=' +
      encodeURIComponent(destinoLookup) +
      '&$top=5000';

    const json: any = await this._get(url);
    return this._getResults(json).map((item: any) => ({
      id: Number(item.Id) || 0,
      nombre: this._getString(item, map.Nombre || 'Title'),
      destino: item[destinoLookup] ? this._getString(item[destinoLookup], 'Title') : '',
      fechaSalida: toDateInput(this._getString(item, map.FechaSalida || 'FechaSalida'))
    })).filter((viaje: IViajeCajaRow) => viaje.id > 0);
  }

  private async _getServicios(): Promise<IServicioCajaRow[]> {
    const map = await this._getServiciosFieldMap();
    const lookupField = map.ViajeAsociado || 'ViajeAsociado';
    const selectFields = [
      'Id',
      map.PrecioCliente || 'PrecioCliente',
      map.Moneda || 'Moneda',
      lookupField + '/Id'
    ];
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      LISTA_SERVICIOS_VIAJE +
      "')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$expand=' +
      encodeURIComponent(lookupField) +
      '&$top=5000';

    const json: any = await this._get(url);
    return this._getResults(json)
      .map((item: any) => ({
        id: Number(item.Id) || 0,
        viajeId: this._extractLookupId(item, lookupField),
        precioCliente: this._toNumber(item[map.PrecioCliente || 'PrecioCliente']),
        moneda: this._getString(item, map.Moneda || 'Moneda')
      }))
      .filter((servicio: IServicioCajaRow) => servicio.id > 0 && servicio.viajeId > 0);
  }

  private async _getPagos(): Promise<IPagoCajaRow[]> {
    const map = await this._getPagosFieldMap();
    const selectFields = buildRegistroPagoSelectFields(map);
    // Asegurar columnas de recupero aunque el field map no las haya marcado.
    this._ensureSelectField(selectFields, map.MontoAplicadoViaje || 'MontoAplicadoViaje');
    this._ensureSelectField(selectFields, map.MontoGastosBancarios || 'RecuperoGastosBancarios');
    this._ensureSelectField(selectFields, 'RecuperoGastosBancarios');
    this._ensureSelectField(selectFields, map.PorcentajeRecupero || 'PorcentajeRecupero');
    const expandFields = buildRegistroPagoExpandFields(map);
    let url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      LISTA_PAGOS +
      "')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$top=5000';
    if (expandFields.length > 0) {
      url += '&$expand=' + encodeURIComponent(expandFields.join(','));
    }

    let json: any;
    try {
      json = await this._get(url);
    } catch (error) {
      // Si alguna columna de recupero no existe, reintentar con el select canónico.
      console.warn(
        '[CajaConsolidada] Falló $select con columnas de recupero; reintento con select base.',
        error
      );
      const baseSelect = buildRegistroPagoSelectFields(map);
      let fallbackUrl =
        this._webUrl +
        "/_api/web/lists/getByTitle('" +
        LISTA_PAGOS +
        "')/items?$select=" +
        encodeURIComponent(baseSelect.join(',')) +
        '&$top=5000';
      if (expandFields.length > 0) {
        fallbackUrl += '&$expand=' + encodeURIComponent(expandFields.join(','));
      }
      json = await this._get(fallbackUrl);
    }

    return this._getResults(json)
      .map((item: any) => {
        const shared = mapSharePointItemToRegistroPago(item, map);
        let montoGastosBancarios = this._leerNumeroRecupero(
          item,
          map.MontoGastosBancarios || 'RecuperoGastosBancarios',
          shared.montoGastosBancarios
        );
        if (montoGastosBancarios === null || montoGastosBancarios === undefined) {
          montoGastosBancarios = this._leerNumeroRecupero(
            item,
            'RecuperoGastosBancarios',
            shared.montoGastosBancarios
          );
        }
        const montoAplicadoViaje = this._leerNumeroRecupero(
          item,
          map.MontoAplicadoViaje || 'MontoAplicadoViaje',
          shared.montoAplicadoViaje
        );
        return {
          id: shared.id,
          viajeId: shared.viajeAsociadoId && shared.viajeAsociadoId > 0 ? shared.viajeAsociadoId : 0,
          tipoPago: shared.tipoPago,
          monto: shared.monto,
          montoAplicadoViaje:
            montoAplicadoViaje !== undefined ? montoAplicadoViaje : shared.montoAplicadoViaje,
          montoGastosBancarios:
            montoGastosBancarios !== undefined
              ? montoGastosBancarios
              : shared.montoGastosBancarios,
          porcentajeRecupero: shared.porcentajeRecupero,
          moneda: shared.moneda,
          cotizacion: shared.cotizacion,
          estado: shared.estado,
          servicioAsociadoId: shared.servicioViajeId,
          concepto: shared.concepto,
          cuentaBancariaId: shared.cuentaBancariaId,
          cuentaBancariaTitulo: shared.cuentaBancariaTitulo,
          banco: shared.banco,
          fechaPago: shared.fechaPago,
          medioPago: shared.medioPago,
          observaciones: shared.observaciones
        } as IPagoCajaRow;
      })
      .filter((pago: IPagoCajaRow) => pago.id > 0);
  }

  private _ensureSelectField(selectFields: string[], fieldName: string): void {
    const name = (fieldName || '').trim();
    if (!name) {
      return;
    }
    const already = selectFields.filter(
      (field: string) => (field || '').toLowerCase() === name.toLowerCase()
    )[0];
    if (!already) {
      selectFields.push(name);
    }
  }

  /**
   * Lee un número de recupero desde el ítem raw si el mapper no lo resolvió
   * (p. ej. fieldExists=false pero la columna sí vino en $select).
   */
  private _leerNumeroRecupero(
    item: any,
    fieldInternalName: string,
    mappedValue: number | null | undefined
  ): number | null | undefined {
    if (mappedValue !== null && mappedValue !== undefined && isFinite(Number(mappedValue))) {
      return Number(mappedValue);
    }
    if (!item || !fieldInternalName) {
      return mappedValue;
    }
    const raw = item[fieldInternalName];
    if (raw === undefined || raw === null || raw === '') {
      return mappedValue;
    }
    const parsed = Number(raw);
    return isFinite(parsed) ? parsed : mappedValue;
  }

  private async _getPagosFieldMap(): Promise<IRegistroPagoFieldMap> {
    if (this._registroPagoFieldMap) {
      return this._registroPagoFieldMap;
    }
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      LISTA_PAGOS +
      "')/fields?$select=Title,InternalName,TypeAsString,Hidden";
    const json: any = await this._get(url);
    const fields: ISharePointListFieldMeta[] = this._getResults(json);
    this._registroPagoFieldMap = resolveRegistroPagoFieldMap(fields, { log: false });
    return this._registroPagoFieldMap;
  }

  private async _getViajesFieldMap(): Promise<IStringMap> {
    if (this._viajesFieldMap) {
      return this._viajesFieldMap;
    }
    this._viajesFieldMap = await this._resolveSimpleFieldMap(LISTA_VIAJES, {
      Nombre: 'Title',
      Destino: 'Destino',
      FechaSalida: 'FechaSalida'
    });
    this._viajesFieldMap.Nombre = 'Title';
    return this._viajesFieldMap;
  }

  private async _getServiciosFieldMap(): Promise<IStringMap> {
    if (this._serviciosFieldMap) {
      return this._serviciosFieldMap;
    }
    this._serviciosFieldMap = await this._resolveSimpleFieldMap(LISTA_SERVICIOS_VIAJE, {
      ViajeAsociado: 'ViajeAsociado',
      PrecioCliente: 'PrecioCliente',
      Moneda: 'Moneda'
    });
    return this._serviciosFieldMap;
  }

  private async _resolveSimpleFieldMap(
    listTitle: string,
    fallbackByKey: IStringMap
  ): Promise<IStringMap> {
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/fields?$select=Title,InternalName&$filter=Hidden eq false";
    const json: any = await this._get(url);
    const fields = this._getResults(json);
    const byInternal: IStringMap = {};
    const byTitle: IStringMap = {};
    fields.forEach((field: any) => {
      byInternal[field.InternalName] = field.InternalName;
      byTitle[field.Title] = field.InternalName;
    });

    const map: IStringMap = {};
    Object.keys(fallbackByKey).forEach((key: string) => {
      const fallback = fallbackByKey[key];
      map[key] = byInternal[fallback] || byTitle[fallback] || fallback;
    });
    return map;
  }

  private _extractLookupId(item: any, fieldName: string): number {
    const idKey = fieldName + 'Id';
    if (item[idKey] !== undefined && item[idKey] !== null) {
      return Number(item[idKey]) || 0;
    }
    if (item[fieldName] && item[fieldName].Id) {
      return Number(item[fieldName].Id) || 0;
    }
    return 0;
  }

  private _getString(source: any, key: string): string {
    if (!key || !source) {
      return '';
    }
    const value = source[key];
    return value === null || value === undefined ? '' : String(value);
  }

  private _toNumber(value: any): number {
    const parsed = Number(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  private _getResults(json: any): any[] {
    if (!json) {
      return [];
    }
    if (Array.isArray(json.value)) {
      return json.value;
    }
    if (json.d && Array.isArray(json.d.results)) {
      return json.d.results;
    }
    return [];
  }

  private async _get(url: string): Promise<any> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );
    if (!response.ok) {
      throw new Error('Error HTTP ' + response.status + ' al consultar SharePoint.');
    }
    return response.json();
  }
}
