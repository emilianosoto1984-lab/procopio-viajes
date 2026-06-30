import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { toDateInput, toSharePointDateOnlyPayload } from '../../../shared/sharePointDateUtils';
import { IPagoSaldoItem } from '../../../shared/pagoSaldoUtils';

export { IPagoSaldoItem };

export interface IRegistroPagoData {
  tipoPago: string;
  viajeId?: number | null;
  pasajeroId: number;
  medioPago: string;
  monto: number;
  moneda: string;
  fechaPago: string;
  estado: string;
  banco?: string;
  motivo?: string;
  cotizacion?: number;
  viajeTitulo?: string;
  pasajeroNombre?: string;
  concepto?: string;
  servicioAsociadoId?: number;
}

export function getEstadoByMedioPago(medioPago: string): string {
  switch ((medioPago || '').trim()) {
    case 'Efectivo':
      return 'Aprobado';
    case 'Transferencia':
    case 'Tarjeta de Credito':
      return 'Pendiente';
    default:
      return 'Pendiente';
  }
}

export interface IRegistroPagoItem extends IRegistroPagoData {
  id: number;
}

export interface IViajeLookupItem {
  id: number;
  titulo: string;
}

export interface IPasajeroLookupItem {
  id: number;
  nombreApellido: string;
  dni: string;
}

export interface IServicioViajeItem {
  id: number;
  viajeId: number;
  concepto: string;
  precioCliente: number;
  moneda: string;
  operadorId: number;
  operadorNombre: string;
}

export interface IComprobanteItem {
  fileName: string;
  serverRelativeUrl: string;
  fileSize?: number;
}

type IStringMap = { [key: string]: string };
type ISPRequestHeaders = { [key: string]: string };

const GET_HEADERS: ISPRequestHeaders = {
  Accept: 'application/json;odata.metadata=minimal'
};

const WRITE_HEADERS: ISPRequestHeaders = {
  Accept: 'application/json;odata.metadata=minimal',
  'Content-Type': 'application/json;odata.metadata=minimal'
};

const LISTA_PAGOS = 'Registro de Pagos';
const LISTA_VIAJES = 'Registro de Viajes';
const LISTA_PASAJEROS = 'Pasajeros';
const LISTA_SERVICIOS_VIAJE = 'ServiciosViaje';

export default class SharePointPagosService {
  private readonly _context: FormCustomizerContext;
  private readonly _webUrl: string;
  private _fieldMaps: { [listTitle: string]: IStringMap } = {};

  public constructor(context: FormCustomizerContext) {
    this._context = context;
    this._webUrl = context.pageContext.web.absoluteUrl;
  }

  public async getPagoById(id: number): Promise<IRegistroPagoItem> {
    const map = await this._getPagosFieldMap();
    const lookupField = map.ViajeAsociado;
    const pasajeroField = map.Pasajero;
    const montoField = map.Monto;
    const servicioLookup = map.ServicioViaje;
    const selectFields = [
      'Id',
      montoField,
      map.MedioPago,
      map.Moneda,
      map.Estado,
      map.FechaPago,
      map.Cotizacion,
      map.Banco,
      map.Motivo,
      map.TipoPago,
      map.Concepto,
      lookupField + '/Id',
      lookupField + '/Title',
      pasajeroField + '/Id',
      pasajeroField + '/Title',
      servicioLookup + '/Id'
    ];
    const expandFields = [lookupField, pasajeroField, servicioLookup].filter(
      (field: string, index: number, fields: string[]) => fields.indexOf(field) === index
    );
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('Registro de Pagos')/items(" +
        id +
        ')?$select=' +
        encodeURIComponent(selectFields.join(',')) +
        '&$expand=' +
        encodeURIComponent(expandFields.join(','))
    );
    const item: any = await this._get(url);
    const viajeId = this._extractLookupId(item, lookupField);
    const viajeTitulo = item[lookupField] ? this._getString(item[lookupField], 'Title') : '';
    const pasajeroId = this._extractLookupId(item, pasajeroField);
    const servicioAsociadoId = this._extractLookupId(item, servicioLookup);

    return {
      id: item.Id,
      viajeId: viajeId > 0 ? viajeId : null,
      viajeTitulo,
      pasajeroId,
      tipoPago: this._normalizarTipoPago(this._getString(item, map.TipoPago)),
      servicioAsociadoId: servicioAsociadoId > 0 ? servicioAsociadoId : undefined,
      concepto: this._getConceptoPago(item, map),
      medioPago: this._getString(item, map.MedioPago),
      monto: this._leerMontoPago(item, montoField, map),
      moneda: this._getString(item, map.Moneda),
      estado: this._getString(item, map.Estado),
      fechaPago: toDateInput(this._getString(item, map.FechaPago)),
      banco: this._getString(item, map.Banco),
      motivo: this._getString(item, map.Motivo),
      cotizacion:
        item[map.Cotizacion] !== undefined && item[map.Cotizacion] !== null
          ? this._toNumber(item[map.Cotizacion])
          : undefined
    };
  }

  public async createPago(data: IRegistroPagoData): Promise<IRegistroPagoItem> {
    const map = await this._getPagosFieldMap();
    const payload = this._buildPagoPayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items");
    const created = await this._post(url, payload);
    return {
      id: created.Id,
      viajeId: data.viajeId,
      viajeTitulo: data.viajeTitulo,
      pasajeroId: data.pasajeroId,
      tipoPago: this._normalizarTipoPago(data.tipoPago),
      concepto: data.concepto,
      medioPago: data.medioPago,
      monto: data.monto,
      moneda: data.moneda,
      fechaPago: data.fechaPago,
      estado: data.estado,
      banco: data.banco,
      motivo: data.motivo,
      cotizacion: data.cotizacion,
      servicioAsociadoId:
        data.servicioAsociadoId && data.servicioAsociadoId > 0 ? data.servicioAsociadoId : undefined
    };
  }

  public async updatePago(id: number, data: IRegistroPagoData): Promise<void> {
    const map = await this._getPagosFieldMap();
    const payload = this._buildPagoPayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async approveTransfer(id: number): Promise<void> {
    const map = await this._getPagosFieldMap();
    const payload: { [key: string]: string } = {};
    payload[map.Estado] = 'Aprobado';
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async getAttachments(itemId: number): Promise<IComprobanteItem[]> {
    const url = this._buildPagoAttachmentsUrl(itemId);
    const json: any = await this._get(url);
    const files = this._getResults(json);
    return files
      .map((file: any) => ({
        fileName: this._getString(file, 'FileName'),
        serverRelativeUrl: this._getString(file, 'ServerRelativeUrl'),
        fileSize: file.FileLength !== undefined ? this._toNumber(file.FileLength) : undefined
      }))
      .filter((file: IComprobanteItem) => file.fileName.length > 0);
  }

  public async uploadAttachment(itemId: number, file: File): Promise<IComprobanteItem> {
    const fileName = this._sanitizeAttachmentFileName(file.name);
    if (!fileName.trim()) {
      throw new Error('El archivo no tiene un nombre válido.');
    }

    const existentes = await this.getAttachments(itemId);
    const duplicado = existentes.some(
      (attachment: IComprobanteItem) => attachment.fileName.toLowerCase() === fileName.toLowerCase()
    );
    if (duplicado) {
      throw new Error("Ya existe un archivo con el nombre '" + fileName + "'.");
    }

    await this._uploadAttachmentFile(itemId, file, fileName);
    const actualizados = await this.getAttachments(itemId);
    const subido = actualizados.filter(
      (attachment: IComprobanteItem) => attachment.fileName.toLowerCase() === fileName.toLowerCase()
    )[0];
    return subido || { fileName, serverRelativeUrl: '', fileSize: file.size };
  }

  public async deleteAttachment(itemId: number, fileName: string): Promise<void> {
    const escapedName = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_PAGOS +
        "')/items(" +
        itemId +
        ")/AttachmentFiles('" +
        escapedName +
        "')"
    );
    await this._post(url, undefined, {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  }

  public async searchViajesByTitle(searchText: string, top: number = 20): Promise<IViajeLookupItem[]> {
    const trimmed = (searchText || '').trim();
    if (!trimmed) {
      return [];
    }

    const escaped = trimmed.replace(/'/g, "''");
    const filter = "substringof('" + escaped + "',Title)";
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('Registro de Viajes')/items?$select=Id,Title&$filter=" +
        encodeURIComponent(filter) +
        '&$top=' +
        top +
        '&$orderby=Title'
    );
    const json: any = await this._get(url);
    const values = this._getResults(json);
    return values.map((item: any) => ({
      id: item.Id,
      titulo: this._getString(item, 'Title')
    }));
  }

  public async getViajeById(id: number): Promise<IViajeLookupItem> {
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Viajes')/items(" + id + ')?$select=Id,Title');
    const item = await this._get(url);
    return {
      id: item.Id,
      titulo: this._getString(item, 'Title')
    };
  }

  public async getPasajerosByViaje(viajeId: number): Promise<IPasajeroLookupItem[]> {
    if (!viajeId || viajeId <= 0) {
      return [];
    }

    const viajesMap = await this._getViajesFieldMap();
    const pasajerosField = viajesMap.Pasajeros;
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_VIAJES +
        "')/items(" +
        viajeId +
        ')?$select=Id,' +
        encodeURIComponent(pasajerosField + 'Id')
    );
    const item: any = await this._get(url);
    const pasajeroIds = this._extractMultiLookupIds(item, pasajerosField);
    if (pasajeroIds.length === 0) {
      return [];
    }

    return this._getPasajerosByIds(pasajeroIds);
  }

  public async getPasajeroById(id: number): Promise<IPasajeroLookupItem> {
    const map = await this._getPasajerosFieldMap();
    const selectFields = ['Id', 'Title', map.NombreApellido, map.DNI];
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_PASAJEROS +
        "')/items(" +
        id +
        ')?$select=' +
        encodeURIComponent(selectFields.join(','))
    );
    const item: any = await this._get(url);
    return this._mapPasajeroItem(item, map);
  }

  public async findPasajeroByDni(dni: string): Promise<IPasajeroLookupItem | null> {
    const trimmed = (dni || '').trim();
    if (!trimmed) {
      return null;
    }

    const map = await this._getPasajerosFieldMap();
    const escaped = trimmed.replace(/'/g, "''");
    const filter = map.DNI + " eq '" + escaped + "'";
    const selectFields = ['Id', 'Title', map.NombreApellido, map.DNI];
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_PASAJEROS +
        "')/items?$select=" +
        encodeURIComponent(selectFields.join(',')) +
        '&$filter=' +
        encodeURIComponent(filter) +
        '&$top=1'
    );
    const json: any = await this._get(url);
    const values = this._getResults(json);
    if (!values.length) {
      return null;
    }
    return this._mapPasajeroItem(values[0], map);
  }

  public async createPasajero(nombreApellido: string, dni: string): Promise<IPasajeroLookupItem> {
    const map = await this._getPasajerosFieldMap();
    const nombre = (nombreApellido || '').trim();
    const documento = (dni || '').trim();
    const payload: { [key: string]: string } = {};
    payload.Title = nombre;
    payload[map.NombreApellido] = nombre;
    payload[map.DNI] = documento;
    payload[map.Pasaporte] = '';
    payload[map.Telefono] = '';
    payload[map.Email] = '';
    payload[map.Observaciones] = '';

    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_PASAJEROS + "')/items");
    const created = await this._post(url, payload);
    return {
      id: created.Id,
      nombreApellido: nombre,
      dni: documento
    };
  }

  public async findOrCreatePasajero(nombreApellido: string, dni: string): Promise<IPasajeroLookupItem> {
    const existente = await this.findPasajeroByDni(dni);
    if (existente) {
      return existente;
    }
    return this.createPasajero(nombreApellido, dni);
  }

  public async getServiciosViajeByViaje(viajeId: number): Promise<IServicioViajeItem[]> {
    const map = await this._getServiciosViajeFieldMap();
    const lookupField = map.ViajeAsociado;
    const operadorLookup = map.Operador;
    const selectFields = [
      'Id',
      map.Concepto,
      map.PrecioCliente,
      map.Moneda,
      lookupField + '/Id',
      operadorLookup + '/Id',
      operadorLookup + '/Title'
    ];
    const baseUrl =
      "/_api/web/lists/getByTitle('" +
      LISTA_SERVICIOS_VIAJE +
      "')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$expand=' +
      encodeURIComponent(lookupField + ',' + operadorLookup);

    const lookupFilterUrl = this._buildUrl(baseUrl + '&$filter=' + encodeURIComponent(lookupField + '/Id eq ' + viajeId));
    let json: any;
    try {
      json = await this._get(lookupFilterUrl);
    } catch (error) {
      const simpleFilterUrl = this._buildUrl(baseUrl + '&$filter=' + encodeURIComponent(lookupField + ' eq ' + viajeId));
      json = await this._get(simpleFilterUrl);
    }

    const values = this._getResults(json);
    return values.map((item: any) => ({
      id: item.Id,
      viajeId: this._extractViajeId(item, lookupField),
      concepto: this._getString(item, map.Concepto),
      precioCliente: this._toNumber(item[map.PrecioCliente]),
      moneda: this._getString(item, map.Moneda),
      operadorId: this._extractLookupId(item, operadorLookup) || 0,
      operadorNombre: item[operadorLookup] ? this._getString(item[operadorLookup], 'Title') : ''
    }));
  }

  public async getPagosSaldoByViaje(viajeId: number): Promise<IPagoSaldoItem[]> {
    const map = await this._getPagosFieldMap();
    const lookupField = map.ViajeAsociado;
    const montoField = map.Monto;
    const servicioLookup = map.ServicioViaje;
    const selectFields = [
      'Id',
      'Title',
      map.Concepto,
      montoField,
      map.Moneda,
      map.TipoPago,
      map.Cotizacion,
      lookupField + '/Id',
      servicioLookup + '/Id'
    ];
    if (map.Importe && map.Importe !== montoField) {
      selectFields.push(map.Importe);
    }
    const expandFields = [lookupField, servicioLookup].filter(
      (field: string, index: number, fields: string[]) => fields.indexOf(field) === index
    );
    const baseUrl =
      "/_api/web/lists/getByTitle('" +
      LISTA_PAGOS +
      "')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$expand=' +
      encodeURIComponent(expandFields.join(','));

    const lookupFilterUrl = this._buildUrl(
      baseUrl + '&$filter=' + encodeURIComponent(lookupField + '/Id eq ' + viajeId)
    );
    let json: any;
    try {
      json = await this._get(lookupFilterUrl);
    } catch (error) {
      const simpleFilterUrl = this._buildUrl(
        baseUrl + '&$filter=' + encodeURIComponent(lookupField + ' eq ' + viajeId)
      );
      json = await this._get(simpleFilterUrl);
    }

    const values = this._getResults(json);
    return values.map((item: any) => {
      const servicioAsociadoId = this._extractLookupId(item, servicioLookup);
      const tipoPagoLeido = this._getString(item, map.TipoPago);
      return {
        id: item.Id,
        tipoPago: this._normalizarTipoPago(tipoPagoLeido),
        monto: this._leerMontoPago(item, montoField, map),
        moneda: this._getString(item, map.Moneda),
        cotizacion:
          item[map.Cotizacion] !== undefined && item[map.Cotizacion] !== null
            ? this._toNumber(item[map.Cotizacion])
            : undefined,
        servicioAsociadoId: servicioAsociadoId > 0 ? servicioAsociadoId : undefined,
        concepto: this._getConceptoPago(item, map)
      };
    });
  }

  public async getChoiceFieldOptions(listTitle: string, fieldInternalName: string): Promise<string[]> {
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        listTitle +
        "')/fields/getByInternalNameOrTitle('" +
        fieldInternalName +
        "')?$select=Choices"
    );
    const json: any = await this._get(url);
    const choices = json.Choices;
    if (Array.isArray(choices)) {
      return choices.map((choice: string) => String(choice));
    }
    if (choices && Array.isArray(choices.results)) {
      return choices.results.map((choice: string) => String(choice));
    }
    return [];
  }

  public async getBancoChoices(): Promise<string[]> {
    const map = await this._getPagosFieldMap();
    return this.getChoiceFieldOptions(LISTA_PAGOS, map.Banco || 'Banco');
  }

  public async getMotivoChoices(): Promise<string[]> {
    const map = await this._getPagosFieldMap();
    return this.getChoiceFieldOptions(LISTA_PAGOS, map.Motivo || 'Motivo');
  }

  private _buildPagoPayload(map: IStringMap, data: IRegistroPagoData): { [key: string]: string | number | null } {
    const payload: { [key: string]: string | number | null } = {};
    const tituloViaje = (data.viajeTitulo || '').trim();
    const tituloPasajero = (data.pasajeroNombre || '').trim();
    const titulo = tituloViaje || tituloPasajero || 'Registro de Pago';
    payload.Title = titulo;
    if (data.concepto && data.concepto.trim()) {
      payload[map.Concepto] = data.concepto.trim();
    } else {
      payload[map.Concepto] = null;
    }
    if (data.viajeId && data.viajeId > 0) {
      payload[map.ViajeAsociado + 'Id'] = data.viajeId;
    } else {
      payload[map.ViajeAsociado + 'Id'] = null;
    }
    payload[map.Pasajero + 'Id'] = data.pasajeroId > 0 ? data.pasajeroId : null;
    payload[map.TipoPago] = this._normalizarTipoPago(data.tipoPago);
    payload[map.Monto] = data.monto;
    payload[map.MedioPago] = data.medioPago;
    payload[map.Moneda] = data.moneda;
    payload[map.Estado] = data.estado;
    const fechaPagoPayload = toSharePointDateOnlyPayload(data.fechaPago);
    console.log('FechaPago preparada para guardar:', fechaPagoPayload);
    payload[map.FechaPago] = fechaPagoPayload;
    if (data.banco && data.banco.trim()) {
      payload[map.Banco] = data.banco;
    } else {
      payload[map.Banco] = null;
    }
    if (data.motivo && data.motivo.trim()) {
      payload[map.Motivo] = data.motivo;
    } else {
      payload[map.Motivo] = null;
    }
    if (data.cotizacion !== undefined && data.cotizacion !== null) {
      payload[map.Cotizacion] = data.cotizacion > 0 ? data.cotizacion : null;
    } else {
      payload[map.Cotizacion] = null;
    }
    const servicioKey = map.ServicioViaje + 'Id';
    if (data.servicioAsociadoId !== undefined) {
      if (data.servicioAsociadoId > 0) {
        payload[servicioKey] = data.servicioAsociadoId;
      } else {
        payload[servicioKey] = null;
      }
    }
    return payload;
  }

  private async _getPagosFieldMap(): Promise<IStringMap> {
    const map = await this._getFieldMap(LISTA_PAGOS, {
      ViajeAsociado: 'ViajeAsociado',
      // Lookup a lista Pasajeros. Crear el campo en SharePoint si aún no existe (InternalName: Pasajero).
      Pasajero: 'Pasajero',
      Monto: 'Monto',
      MedioPago: 'MedioPago',
      Moneda: 'Moneda',
      Estado: 'Estado',
      FechaPago: 'FechaPago',
      Cotizacion: 'Cotizacion',
      Banco: 'Banco',
      Motivo: 'Motivo',
      TipoPago: 'TipoPago',
      Concepto: 'Concepto',
      ServicioViaje: 'ServicioViaje'
    }, {
      Banco: 'Cuenta Bancaria',
      Motivo: 'Motivo de elección'
    });

    const fieldsUrl = this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_PAGOS + "')/fields?$select=Title,InternalName&$filter=Hidden eq false"
    );
    const json: any = await this._get(fieldsUrl);
    const fields: any[] = this._getResults(json);
    const importeField = fields.filter((field) => field.Title === 'Importe' || field.InternalName === 'Importe')[0];
    if (importeField) {
      map.Importe = importeField.InternalName;
      if (!map.Monto || map.Monto === 'Monto') {
        const montoField = fields.filter((field) => field.InternalName === 'Monto' || field.Title === 'Monto')[0];
        if (!montoField) {
          map.Monto = importeField.InternalName;
        }
      }
    }

    return map;
  }

  private async _getViajesFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_VIAJES, {
      Pasajeros: 'Pasajeros'
    });
  }

  private async _getServiciosViajeFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_SERVICIOS_VIAJE, {
      ViajeAsociado: 'ViajeAsociado',
      Concepto: 'Concepto',
      PrecioCliente: 'PrecioCliente',
      Moneda: 'Moneda',
      Operador: 'Operador'
    });
  }

  private async _getPasajerosFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(
      LISTA_PASAJEROS,
      {
        NombreApellido: 'Title',
        DNI: 'DNI',
        Pasaporte: 'Pasaporte',
        Telefono: 'Telefono',
        Email: 'Email',
        Observaciones: 'Observaciones'
      },
      {
        NombreApellido: 'Nombre y Apellido'
      }
    );
  }

  private async _getPasajerosByIds(ids: number[]): Promise<IPasajeroLookupItem[]> {
    const uniqueIds = ids.filter((id: number) => id > 0).filter(
      (id: number, index: number, source: number[]) => source.indexOf(id) === index
    );
    if (uniqueIds.length === 0) {
      return [];
    }

    const map = await this._getPasajerosFieldMap();
    const selectFields = ['Id', 'Title', map.NombreApellido, map.DNI];
    const filter = uniqueIds.map((id: number) => 'Id eq ' + id).join(' or ');
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_PASAJEROS +
        "')/items?$select=" +
        encodeURIComponent(selectFields.join(',')) +
        '&$filter=' +
        encodeURIComponent(filter) +
        '&$top=' +
        uniqueIds.length
    );
    const json: any = await this._get(url);
    const values = this._getResults(json);
    return values
      .map((item: any) => this._mapPasajeroItem(item, map))
      .sort((a: IPasajeroLookupItem, b: IPasajeroLookupItem) =>
        a.nombreApellido.localeCompare(b.nombreApellido, 'es')
      );
  }

  private _mapPasajeroItem(item: any, map: IStringMap): IPasajeroLookupItem {
    return {
      id: item.Id,
      nombreApellido: this._getString(item, map.NombreApellido) || this._getString(item, 'Title'),
      dni: this._getString(item, map.DNI)
    };
  }

  private async _getFieldMap(
    listTitle: string,
    fallbackByKey: IStringMap,
    displayNameByKey?: IStringMap
  ): Promise<IStringMap> {
    if (this._fieldMaps[listTitle]) {
      return this._fieldMaps[listTitle];
    }

    const fieldsUrl = this._buildUrl(
      "/_api/web/lists/getByTitle('" + encodeURIComponent(listTitle).replace(/%20/g, ' ') + "')/fields?$select=Title,InternalName&$filter=Hidden eq false"
    );
    const json: any = await this._get(fieldsUrl);
    const fields: any[] = this._getResults(json);
    const byTitle: IStringMap = {};

    fields.forEach((field) => {
      byTitle[field.Title] = field.InternalName;
    });

    const map: IStringMap = {};
    Object.keys(fallbackByKey).forEach((key: string) => {
      if (key === 'Concepto') {
        const conceptoField = fields.filter(
          (field) => field.InternalName === 'Concepto' || field.Title === 'Concepto'
        )[0];
        map[key] = conceptoField ? conceptoField.InternalName : 'Title';
        return;
      }
      const displayName = displayNameByKey && displayNameByKey[key] ? displayNameByKey[key] : key;
      map[key] = byTitle[displayName] || byTitle[key] || fallbackByKey[key];
    });

    this._fieldMaps[listTitle] = map;
    return map;
  }

  private _extractViajeId(item: any, lookupField: string): number {
    const lookupValue = item[lookupField];
    if (lookupValue && lookupValue.Id) {
      return Number(lookupValue.Id) || 0;
    }
    return this._extractLookupId(item, lookupField);
  }

  private _extractLookupId(item: any, fieldName: string): number {
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

  private _extractMultiLookupIds(item: any, fieldName: string): number[] {
    const value = item[fieldName + 'Id'];
    if (Array.isArray(value)) {
      return value.map((id: any) => Number(id)).filter((id: number) => id > 0);
    }
    if (value && Array.isArray(value.results)) {
      return value.results.map((id: any) => Number(id)).filter((id: number) => id > 0);
    }
    return [];
  }

  private _buildUrl(relativeUrl: string): string {
    return this._webUrl + relativeUrl;
  }

  private async _get(url: string): Promise<any> {
    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: GET_HEADERS
    });
    return this._parseJsonResponse(response, 'GET', url);
  }

  private async _post(
    url: string,
    body?: { [key: string]: string | number | null },
    extraHeaders?: ISPRequestHeaders
  ): Promise<any> {
    const headers: ISPRequestHeaders = {
      ...WRITE_HEADERS,
      ...(extraHeaders || {})
    };
    const response = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    return this._parseJsonResponse(response, 'POST', url);
  }

  private async _parseJsonResponse(
    response: SPHttpClientResponse,
    method: string,
    url: string
  ): Promise<any> {
    if (response.ok) {
      if (response.status === 204) {
        return {};
      }
      const text = await response.text();
      return text ? JSON.parse(text) : {};
    }

    const errorText = await response.text();
    throw new Error(method + ' ' + url + ' failed: ' + response.status + ' ' + errorText);
  }

  private _getResults(json: any): any[] {
    if (Array.isArray(json?.value)) {
      return json.value;
    }
    if (json?.d && Array.isArray(json.d.results)) {
      return json.d.results;
    }
    if (Array.isArray(json?.d)) {
      return json.d;
    }
    return [];
  }

  private _toNumber(value: any): number {
    const parsed = Number(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  private _getConceptoPago(item: any, map: IStringMap): string {
    const concepto = this._getString(item, map.Concepto).trim();
    if (concepto) {
      return concepto;
    }
    return this._getString(item, 'Title').trim();
  }

  private _leerMontoPago(item: any, montoField: string, map: IStringMap): number {
    const montoPrincipal = this._toNumber(item[montoField]);
    if (montoPrincipal > 0) {
      return montoPrincipal;
    }
    const importeField = map.Importe || 'Importe';
    if (importeField !== montoField) {
      const importe = this._toNumber(item[importeField]);
      if (importe > 0) {
        return importe;
      }
    }
    return montoPrincipal;
  }

  private _normalizarTipoPago(value: string): string {
    return (value || '').trim() === 'Egreso' ? 'Egreso' : 'Ingreso';
  }

  private _getString(source: any, key: string): string {
    const value = source[key];
    return value === null || value === undefined ? '' : String(value);
  }

  private _buildPagoAttachmentsUrl(itemId: number): string {
    return this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_PAGOS + "')/items(" + itemId + ')/AttachmentFiles'
    );
  }

  private _escapeODataFileName(fileName: string): string {
    return fileName.replace(/'/g, "''");
  }

  private _sanitizeAttachmentFileName(fileName: string): string {
    const trimmed = (fileName || 'documento').trim();
    const dot = trimmed.lastIndexOf('.');
    let base = dot >= 0 ? trimmed.substring(0, dot) : trimmed;
    const ext = dot >= 0 ? trimmed.substring(dot) : '';
    base = base.replace(/[#%?&'"<>\\\/:*|]/g, '-');
    base = base.replace(/[\x00-\x1f]/g, '-');
    base = base.replace(/-+/g, '-');
    base = base.replace(/^[\s-]+|[\s-]+$/g, '');
    if (!base) {
      base = 'documento';
    }
    return base + ext;
  }

  private async _uploadAttachmentFile(itemId: number, file: File, fileName: string): Promise<void> {
    const escaped = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_PAGOS +
        "')/items(" +
        itemId +
        ")/AttachmentFiles/add(FileName='" +
        escaped +
        "')"
    );
    const buffer = await file.arrayBuffer();
    const headers: ISPRequestHeaders = {
      Accept: 'application/json;odata.metadata=minimal',
      'Content-Type': 'application/octet-stream'
    };
    const response = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers,
      body: buffer
    });
    await this._parseJsonResponse(response, 'POST', url);
  }
}
