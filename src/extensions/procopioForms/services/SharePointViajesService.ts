import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { toDateInput, toSharePointDateOnlyPayload } from '../../../shared/sharePointDateUtils';

export interface IViajeData {
  nombre: string;
  destinoGeneralId?: number;
  destinoId: number;
  fechaSalida: string;
  fechaLlegada: string;
  estado: string;
  colorViaje: string;
  pasajerosIds: number[];
  servicios: string[];
  observaciones: string;
}

export interface IViajeItem extends IViajeData {
  id: number;
}

export interface IPagoData {
  viajeId: number;
  concepto: string;
  fechaPago: string;
  importe: number;
  medioPago: string;
  moneda: string;
  tipoPago: string;
  observaciones?: string;
  cotizacion?: number;
  liquidacionOperadorId?: number;
  servicioAsociadoId?: number;
}

export interface IPagoItem extends IPagoData {
  id: number;
  liquidacionOperadorNombre?: string;
}

export interface IPasajeroItem {
  id: number;
  nombreApellido: string;
  dni: string;
  pasaporte: string;
  telefono: string;
  email: string;
  observaciones: string;
  fechaNacimiento?: string;
}

export interface IPasajeroDraft {
  nombreApellido: string;
  dni: string;
  pasaporte: string;
  telefono: string;
  email: string;
  observaciones?: string;
  fechaNacimiento?: string;
}

export interface IDestinoItem {
  id: number;
  titulo: string;
  destinoGeneralId?: number;
  destinoGeneralNombre?: string;
}

export interface IDestinoGeneralItem {
  id: number;
  titulo: string;
}

export interface IOperadorItem {
  id: number;
  titulo: string;
}

export interface ILiquidacionItem {
  id: number;
  codigoReferencia: string;
  monto: number;
  moneda: string;
  operadorId?: number;
  operadorNombre?: string;
  archivoNombre?: string;
  archivoUrl?: string;
}

export interface ILiquidacionData {
  viajeId: number;
  codigoReferencia: string;
  monto: number;
  moneda: string;
  operadorId: number;
  file?: File;
}

export interface IVoucherItem {
  fileName: string;
  serverRelativeUrl: string;
}

export interface IPresupuestoItem {
  fileName: string;
  serverRelativeUrl: string;
}

export interface IFacturaItem {
  fileName: string;
  serverRelativeUrl: string;
}

export interface IPasajeroDocumentoItem {
  fileName: string;
  serverRelativeUrl: string;
}

const PRESUPUESTO_FILE_PREFIX = 'Presupuesto_';
const FACTURA_FILE_PREFIX = 'Factura_';

export interface IServicioViajeData {
  viajeId: number;
  concepto: string;
  precioCliente: number;
  moneda: string;
  operadorId: number;
  operadorNombre?: string;
}

export interface IServicioViajeItem extends IServicioViajeData {
  id: number;
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

const LISTA_VIAJES = 'Registro de Viajes';
const LISTA_PASAJEROS = 'Pasajeros';
const LISTA_PAGOS = 'Registro de Pagos';
const LISTA_DESTINOS = 'Destinos';
const LISTA_DESTINOS_GENERALES = 'DestinosGenerales';
const LISTA_LIQUIDACIONES = 'Liquidaciones Operador';
const LISTA_SERVICIOS_VIAJE = 'ServiciosViaje';
const LISTA_OPERADORES = 'Operadores';

export default class SharePointViajesService {
  private readonly _context: FormCustomizerContext;
  private readonly _webUrl: string;
  private _fieldMaps: { [listTitle: string]: IStringMap } = {};

  public constructor(context: FormCustomizerContext) {
    this._context = context;
    this._webUrl = context.pageContext.web.absoluteUrl;
  }

  public async getViajeById(id: number): Promise<IViajeItem> {
    const map = await this._getViajesFieldMap();
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Viajes')/items(" + id + ')');
    const item = await this._get(url);

    console.log('[fecha][rest raw] FechaSalida:', item[map.FechaSalida]);
    console.log('[fecha][rest raw] FechaLlegada:', item[map.FechaLlegada]);

    const mapped: IViajeItem = {
      id: item.Id,
      nombre: this._getString(item, map.Nombre),
      destinoGeneralId: this._extractLookupId(item, map.DestinoGeneral),
      destinoId: this._extractLookupId(item, map.Destino),
      fechaSalida: this._toDateInput(this._getString(item, map.FechaSalida)),
      fechaLlegada: this._toDateInput(this._getString(item, map.FechaLlegada)),
      estado: this._getString(item, map.Estado),
      colorViaje: this._getString(item, map.ColorViaje),
      pasajerosIds: this._extractMultiLookupIds(item, map.Pasajeros),
      servicios: this._extractMultiChoiceValues(item, map.Servicios),
      observaciones: this._getString(item, map.Observaciones)
    };
    console.log('[fecha][mapped] FechaSalida:', mapped.fechaSalida);
    console.log('[fecha][mapped] FechaLlegada:', mapped.fechaLlegada);
    return mapped;
  }

  public async createViaje(data: IViajeData): Promise<IViajeItem> {
    const map = await this._getViajesFieldMap();
    const payload = this._buildViajePayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Viajes')/items");
    let created: any;
    try {
      created = await this._post(url, payload);
    } catch (error) {
      if (!this._isInvalidFieldError(error as Error, map.DestinoGeneral + 'Id')) {
        throw error;
      }
      const payloadSinDestinoGeneral = { ...payload };
      delete payloadSinDestinoGeneral[map.DestinoGeneral + 'Id'];
      created = await this._post(url, payloadSinDestinoGeneral);
    }
    return this.getViajeById(created.Id);
  }

  public async updateViaje(id: number, data: IViajeData): Promise<void> {
    const map = await this._getViajesFieldMap();
    const payload = this._buildViajePayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Viajes')/items(" + id + ')');
    try {
      await this._post(url, payload, {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      });
    } catch (error) {
      if (!this._isInvalidFieldError(error as Error, map.DestinoGeneral + 'Id')) {
        throw error;
      }
      const payloadSinDestinoGeneral = { ...payload };
      delete payloadSinDestinoGeneral[map.DestinoGeneral + 'Id'];
      await this._post(url, payloadSinDestinoGeneral, {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      });
    }
  }

  public async getDestinos(): Promise<IDestinoItem[]> {
    const map = await this._getDestinosFieldMap();
    const destinoGeneralLookup = map.DestinoGeneral;
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
        LISTA_DESTINOS +
        "')/items?$select=" +
        encodeURIComponent('Id,Title,' + destinoGeneralLookup + '/Id,' + destinoGeneralLookup + '/Title') +
        '&$expand=' +
        encodeURIComponent(destinoGeneralLookup) +
        '&$top=5000'
    );
    const json = await this._get(url);
    const values = this._getResults(json);
    const destinos = values
      .map((item: any) => ({
        id: item.Id as number,
        titulo: this._getString(item, 'Title'),
        destinoGeneralId: this._extractLookupId(item, destinoGeneralLookup) || undefined,
        destinoGeneralNombre: item[destinoGeneralLookup] ? this._getString(item[destinoGeneralLookup], 'Title') : undefined
      }))
      .filter((d: IDestinoItem) => d.titulo.length > 0)
      .sort((a: IDestinoItem, b: IDestinoItem) => a.titulo.localeCompare(b.titulo, 'es'));
    return destinos;
  }

  public async getDestinosGenerales(): Promise<IDestinoGeneralItem[]> {
    const listNames = [LISTA_DESTINOS_GENERALES, 'Destino General'];
    let json: any = null;
    let lastError: Error | undefined;
    for (let i = 0; i < listNames.length; i++) {
      const listName = listNames[i];
      const url = this._buildUrl(
        "/_api/web/lists/getByTitle('" + listName + "')/items?$select=Id,Title&$top=5000"
      );
      try {
        json = await this._get(url);
        break;
      } catch (error) {
        lastError = error as Error;
      }
    }
    if (!json) {
      throw lastError || new Error('No se pudo cargar la lista de destinos generales.');
    }
    const values = this._getResults(json);
    return values
      .map((item: any) => ({
        id: item.Id as number,
        titulo: this._getString(item, 'Title')
      }))
      .filter((d: IDestinoGeneralItem) => d.id > 0 && d.titulo.length > 0)
      .sort((a: IDestinoGeneralItem, b: IDestinoGeneralItem) => a.titulo.localeCompare(b.titulo, 'es'));
  }

  public async getOperadores(): Promise<IOperadorItem[]> {
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_OPERADORES + "')/items?$select=Id,Title&$top=5000"
    );
    const json = await this._get(url);
    const values = this._getResults(json);
    return values
      .map((item: any) => ({
        id: item.Id as number,
        titulo: this._getString(item, 'Title')
      }))
      .filter((o: IOperadorItem) => o.id > 0 && o.titulo.length > 0);
  }

  public async getPasajeros(): Promise<IPasajeroItem[]> {
    const map = await this._getPasajerosFieldMap();
    const selectFields = ['Id', 'Title', map.NombreApellido, map.DNI, map.Pasaporte, map.Telefono, map.Email, map.Observaciones, map.FechaNacimiento];
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('Pasajeros')/items?$select=" + encodeURIComponent(selectFields.join(',')) + '&$top=5000'
    );
    const json = await this._get(url);
    const values = this._getResults(json);

    return values.map((item: any) => ({
      id: item.Id,
      nombreApellido: this._getString(item, map.NombreApellido) || this._getString(item, 'Title'),
      dni: this._getString(item, map.DNI),
      pasaporte: this._getString(item, map.Pasaporte),
      telefono: this._getString(item, map.Telefono),
      email: this._getString(item, map.Email),
      observaciones: this._getString(item, map.Observaciones),
      fechaNacimiento: this._toDateInput(this._getString(item, map.FechaNacimiento))
    }));
  }

  public async createPasajero(data: IPasajeroDraft): Promise<IPasajeroItem> {
    const map = await this._getPasajerosFieldMap();
    const payload: any = {};
    // Always persist passenger name in Title (source of truth in this list).
    payload.Title = data.nombreApellido;
    payload[map.NombreApellido] = data.nombreApellido;
    payload[map.DNI] = data.dni;
    payload[map.Pasaporte] = data.pasaporte;
    payload[map.Telefono] = data.telefono;
    payload[map.Email] = data.email;
    payload[map.Observaciones] = data.observaciones || '';
    if (data.fechaNacimiento) {
      payload[map.FechaNacimiento] = this._toSharePointDateOnlyPayload(data.fechaNacimiento);
    }

    const url = this._buildUrl("/_api/web/lists/getByTitle('Pasajeros')/items");
    const created = await this._post(url, payload);

    return {
      id: created.Id,
      nombreApellido: data.nombreApellido,
      dni: data.dni,
      pasaporte: data.pasaporte,
      telefono: data.telefono,
      email: data.email,
      observaciones: data.observaciones || '',
      fechaNacimiento: data.fechaNacimiento || ''
    };
  }

  public async updatePasajero(id: number, data: IPasajeroDraft): Promise<void> {
    const map = await this._getPasajerosFieldMap();
    const payload: any = {};
    payload.Title = data.nombreApellido;
    payload[map.NombreApellido] = data.nombreApellido;
    payload[map.DNI] = data.dni;
    payload[map.Pasaporte] = data.pasaporte;
    payload[map.Telefono] = data.telefono;
    payload[map.Email] = data.email;
    payload[map.Observaciones] = data.observaciones || '';
    if (data.fechaNacimiento) {
      payload[map.FechaNacimiento] = this._toSharePointDateOnlyPayload(data.fechaNacimiento);
    }

    const url = this._buildUrl("/_api/web/lists/getByTitle('Pasajeros')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async getDocumentosPasajero(pasajeroId: number): Promise<IPasajeroDocumentoItem[]> {
    const json = await this._get(this._buildPasajeroAttachmentsUrl(pasajeroId));
    const files = this._getResults(json);
    return files
      .map((file: any) => ({
        fileName: this._getString(file, 'FileName'),
        serverRelativeUrl: this._getString(file, 'ServerRelativeUrl')
      }))
      .filter((file: IPasajeroDocumentoItem) => file.fileName.length > 0);
  }

  public async uploadDocumentoPasajero(pasajeroId: number, file: File): Promise<IPasajeroDocumentoItem> {
    const storedFileName = this._sanitizeAttachmentFileName(file.name);
    await this._uploadAttachmentToPasajero(pasajeroId, file);
    const documentos = await this.getDocumentosPasajero(pasajeroId);
    const uploaded = documentos.filter((d: IPasajeroDocumentoItem) => d.fileName === storedFileName)[0];
    return uploaded || { fileName: storedFileName, serverRelativeUrl: '' };
  }

  public async deleteDocumentoPasajero(pasajeroId: number, fileName: string): Promise<void> {
    const escapedName = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_PASAJEROS +
      "')/items(" +
      pasajeroId +
      ")/AttachmentFiles('" +
      escapedName +
      "')"
    );
    await this._post(url, undefined, {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  }

  public async getPagosByViaje(viajeId: number): Promise<IPagoItem[]> {
    const map = await this._getPagosFieldMap();
    const lookupField = map.ViajeAsociado;
    const liqLookup = map.LiquidacionOperador;
    const servicioLookup = map.ServicioViaje;

    const selectFields = [
      'Id',
      map.Concepto,
      map.FechaPago,
      map.Importe,
      map.MedioPago,
      map.Moneda,
      map.TipoPago,
      map.Observaciones,
      map.Cotizacion,
      lookupField + '/Id',
      liqLookup + '/Id',
      liqLookup + '/Title',
      servicioLookup + '/Id',
      servicioLookup + '/Title'
    ];
    const expandFields = lookupField + ',' + liqLookup + ',' + servicioLookup;
    const baseUrl = "/_api/web/lists/getByTitle('Registro de Pagos')/items?$select=" +
      encodeURIComponent(selectFields.join(',')) +
      '&$expand=' + encodeURIComponent(expandFields);

    const lookupFilterUrl = this._buildUrl(baseUrl + '&$filter=' + encodeURIComponent(lookupField + '/Id eq ' + viajeId));
    let json: any;
    try {
      json = await this._get(lookupFilterUrl);
    } catch (error) {
      const simpleFilterUrl = this._buildUrl(baseUrl + '&$filter=' + encodeURIComponent(lookupField + ' eq ' + viajeId));
      json = await this._get(simpleFilterUrl);
    }

    const values = this._getResults(json);
    return values.map((item: any) => {
      const liquidacionOperadorId = this._extractLookupId(item, liqLookup);
      const servicioAsociadoId = this._extractLookupId(item, servicioLookup);
      const liquidacionOperadorNombre = item[liqLookup] && item[liqLookup].Title
        ? this._getString(item[liqLookup], 'Title')
        : '';
      const tipoPagoLeido = this._getString(item, map.TipoPago);
      const tipoPago =
        tipoPagoLeido === 'Egreso' || tipoPagoLeido === 'Ingreso'
          ? tipoPagoLeido
          : liquidacionOperadorId > 0
            ? 'Egreso'
            : 'Ingreso';
      return {
        id: item.Id,
        viajeId: this._extractViajeId(item, lookupField),
        concepto: this._getString(item, map.Concepto),
        fechaPago: this._toDateInput(this._getString(item, map.FechaPago)),
        importe: this._toNumber(item[map.Importe]),
        medioPago: this._getString(item, map.MedioPago),
        moneda: this._getString(item, map.Moneda),
        tipoPago,
        observaciones: this._getString(item, map.Observaciones),
        cotizacion: item[map.Cotizacion] !== undefined && item[map.Cotizacion] !== null ? this._toNumber(item[map.Cotizacion]) : undefined,
        liquidacionOperadorId: liquidacionOperadorId > 0 ? liquidacionOperadorId : undefined,
        servicioAsociadoId: servicioAsociadoId > 0 ? servicioAsociadoId : undefined,
        liquidacionOperadorNombre: liquidacionOperadorNombre || undefined
      };
    });
  }

  public async createPago(data: IPagoData): Promise<IPagoItem> {
    const map = await this._getPagosFieldMap();
    const payload = this._buildPagoPayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items");
    const created = await this._post(url, payload);
    return {
      id: created.Id,
      viajeId: data.viajeId,
      concepto: data.concepto,
      fechaPago: data.fechaPago,
      importe: data.importe,
      medioPago: data.medioPago,
      moneda: data.moneda,
      tipoPago: data.tipoPago,
      observaciones: data.observaciones,
      cotizacion: data.cotizacion,
      liquidacionOperadorId: data.liquidacionOperadorId,
      servicioAsociadoId: data.servicioAsociadoId,
      liquidacionOperadorNombre: undefined
    };
  }

  public async updatePago(id: number, data: IPagoData): Promise<void> {
    const map = await this._getPagosFieldMap();
    const payload = this._buildPagoPayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async deletePago(id: number): Promise<void> {
    const url = this._buildUrl("/_api/web/lists/getByTitle('Registro de Pagos')/items(" + id + ')');
    await this._post(url, undefined, {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  }

  public async getLiquidacionesByViaje(viajeId: number): Promise<ILiquidacionItem[]> {
    const map = await this._getLiquidacionesFieldMap();
    const lookupField = map.ViajeAsociado;
    const operadorLookup = map.Operador;
    const selectFields = ['Id', 'Title', map.Monto, map.Moneda, lookupField + '/Id', operadorLookup + '/Id', operadorLookup + '/Title'];
    const baseUrl =
      "/_api/web/lists/getByTitle('" +
      LISTA_LIQUIDACIONES +
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
    const items: ILiquidacionItem[] = [];
    for (let i = 0; i < values.length; i++) {
      const item = values[i];
      const id = item.Id as number;
      const attachments = await this._getLiquidacionAttachments(id);
      items.push({
        id,
        codigoReferencia: this._getString(item, 'Title'),
        monto: this._toNumber(item[map.Monto]),
        moneda: this._getString(item, map.Moneda),
        operadorId: this._extractLookupId(item, operadorLookup) || undefined,
        operadorNombre: item[operadorLookup] ? this._getString(item[operadorLookup], 'Title') : undefined,
        archivoNombre: attachments.length > 0 ? attachments[0].fileName : undefined,
        archivoUrl: attachments.length > 0 ? attachments[0].serverRelativeUrl : undefined
      });
    }
    return items;
  }

  public async createLiquidacion(data: ILiquidacionData): Promise<ILiquidacionItem> {
    const map = await this._getLiquidacionesFieldMap();
    const payload: any = {};
    payload.Title = data.codigoReferencia;
    payload[map.Monto] = data.monto;
    payload[map.Moneda] = data.moneda;
    payload[map.ViajeAsociado + 'Id'] = data.viajeId;
    payload[map.Operador + 'Id'] = data.operadorId;

    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_LIQUIDACIONES + "')/items");
    const created = await this._post(url, payload);
    const id = created.Id as number;

    let archivoNombre: string | undefined;
    if (data.file) {
      archivoNombre = await this._addAttachmentToListItem(id, data.file);
    }

    return {
      id,
      codigoReferencia: data.codigoReferencia,
      monto: data.monto,
      moneda: data.moneda,
      operadorId: data.operadorId,
      archivoNombre
    };
  }

  public async updateLiquidacion(id: number, data: ILiquidacionData): Promise<void> {
    const map = await this._getLiquidacionesFieldMap();
    const payload: any = {};
    payload.Title = data.codigoReferencia;
    payload[map.Monto] = data.monto;
    payload[map.Moneda] = data.moneda;
    payload[map.ViajeAsociado + 'Id'] = data.viajeId;
    payload[map.Operador + 'Id'] = data.operadorId;

    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_LIQUIDACIONES + "')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async deleteLiquidacion(id: number): Promise<void> {
    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_LIQUIDACIONES + "')/items(" + id + ')');
    await this._post(url, undefined, {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  }

  public async getVouchersByViaje(itemId: number): Promise<IVoucherItem[]> {
    const json = await this._get(this._buildViajeAttachmentsUrl(itemId));
    const files = this._getResults(json);
    return files.map((file: any) => ({
      fileName: this._getString(file, 'FileName'),
      serverRelativeUrl: this._getString(file, 'ServerRelativeUrl')
    })).filter((file: IVoucherItem) => file.fileName.length > 0 && !this._isReservedViajeAttachmentFileName(file.fileName));
  }

  public async getFacturasByViaje(itemId: number): Promise<IFacturaItem[]> {
    const json = await this._get(this._buildViajeAttachmentsUrl(itemId));
    const files = this._getResults(json);
    return files
      .map((file: any) => ({
        fileName: this._getString(file, 'FileName'),
        serverRelativeUrl: this._getString(file, 'ServerRelativeUrl')
      }))
      .filter((file: IFacturaItem) => file.fileName.length > 0 && this._isFacturaFileName(file.fileName));
  }

  public async uploadFactura(itemId: number, file: File): Promise<IFacturaItem> {
    const storedFileName = this._buildFacturaFileName(file.name);
    await this._uploadAttachmentToViaje(itemId, file, storedFileName);
    const facturas = await this.getFacturasByViaje(itemId);
    const uploaded = facturas.filter((f: IFacturaItem) => f.fileName === storedFileName)[0];
    return uploaded || { fileName: storedFileName, serverRelativeUrl: '' };
  }

  public async deleteFactura(itemId: number, fileName: string): Promise<void> {
    const escapedName = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_VIAJES +
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

  public async getPresupuestoByViaje(itemId: number): Promise<IPresupuestoItem | null> {
    const json = await this._get(this._buildViajeAttachmentsUrl(itemId));
    const files = this._getResults(json);
    const presupuestos = files
      .map((file: any) => ({
        fileName: this._getString(file, 'FileName'),
        serverRelativeUrl: this._getString(file, 'ServerRelativeUrl')
      }))
      .filter((file: IPresupuestoItem) => file.fileName.length > 0 && this._isPresupuestoFileName(file.fileName));
    return presupuestos.length > 0 ? presupuestos[0] : null;
  }

  public async uploadPresupuesto(itemId: number, file: File): Promise<IPresupuestoItem> {
    const json = await this._get(this._buildViajeAttachmentsUrl(itemId));
    const files = this._getResults(json);
    for (let i = 0; i < files.length; i++) {
      const fileName = this._getString(files[i], 'FileName');
      if (fileName && this._isPresupuestoFileName(fileName)) {
        await this.deletePresupuesto(itemId, fileName);
      }
    }
    const storedFileName = this._buildPresupuestoFileName(file.name);
    await this._uploadAttachmentToViaje(itemId, file, storedFileName);
    const uploaded = await this.getPresupuestoByViaje(itemId);
    if (!uploaded) {
      throw new Error('No se pudo confirmar el adjunto de presupuesto.');
    }
    return uploaded;
  }

  public async deletePresupuesto(itemId: number, fileName: string): Promise<void> {
    const escapedName = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_VIAJES +
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

  public async uploadVoucher(itemId: number, file: File): Promise<IVoucherItem> {
    const existentes = await this.getVouchersByViaje(itemId);
    const nombresExistentes = existentes.map((v: IVoucherItem) => v.fileName);
    const storedFileName = this._buildUniqueViajeAttachmentFileName(file.name, nombresExistentes);
    await this._uploadAttachmentToViaje(itemId, file, storedFileName);
    const vouchers = await this.getVouchersByViaje(itemId);
    const uploaded = vouchers.filter((v: IVoucherItem) => v.fileName === storedFileName)[0];
    return uploaded || { fileName: storedFileName, serverRelativeUrl: '' };
  }

  public async deleteVoucher(itemId: number, fileName: string): Promise<void> {
    const escapedName = this._escapeODataFileName(fileName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_VIAJES +
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

  public async createServicioViaje(data: IServicioViajeData): Promise<IServicioViajeItem> {
    const map = await this._getServiciosViajeFieldMap();
    const payload = this._buildServicioViajePayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_SERVICIOS_VIAJE + "')/items");
    const created = await this._post(url, payload);
    return {
      id: created.Id,
      viajeId: data.viajeId,
      concepto: data.concepto,
      precioCliente: data.precioCliente,
      moneda: data.moneda,
      operadorId: data.operadorId,
      operadorNombre: data.operadorNombre || ''
    };
  }

  public async updateServicioViaje(id: number, data: IServicioViajeData): Promise<void> {
    const map = await this._getServiciosViajeFieldMap();
    const payload = this._buildServicioViajePayload(map, data);
    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_SERVICIOS_VIAJE + "')/items(" + id + ')');
    await this._post(url, payload, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  public async deleteServicioViaje(id: number): Promise<void> {
    const url = this._buildUrl("/_api/web/lists/getByTitle('" + LISTA_SERVICIOS_VIAJE + "')/items(" + id + ')');
    await this._post(url, undefined, {
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  }

  private async _getLiquidacionesFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_LIQUIDACIONES, {
      ViajeAsociado: 'ViajeAsociado',
      Monto: 'Monto',
      Moneda: 'Moneda',
      Operador: 'Operador'
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

  private async _getLiquidacionAttachments(itemId: number): Promise<{ fileName: string; serverRelativeUrl: string }[]> {
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_LIQUIDACIONES + "')/items(" + itemId + ')/AttachmentFiles'
    );
    const json = await this._get(url);
    const files = this._getResults(json);
    const attachments: { fileName: string; serverRelativeUrl: string }[] = [];
    for (let i = 0; i < files.length; i++) {
      const fileName = this._getString(files[i], 'FileName');
      if (fileName) {
        attachments.push({
          fileName,
          serverRelativeUrl: this._getString(files[i], 'ServerRelativeUrl')
        });
      }
    }
    return attachments;
  }

  private _escapeODataFileName(fileName: string): string {
    return fileName.replace(/'/g, "''");
  }

  /** Reemplaza caracteres conflictivos en URLs REST/OData y en nombres de SharePoint. */
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

  public buildPresupuestoAttachmentFileName(originalFileName: string): string {
    return this._buildPresupuestoFileName(originalFileName);
  }

  private _buildUniqueViajeAttachmentFileName(originalName: string, existingNames: string[]): string {
    const safeName = this._sanitizeAttachmentFileName(originalName);
    const ocupados: { [key: string]: boolean } = {};
    existingNames.forEach((name: string) => {
      ocupados[name.toLowerCase()] = true;
    });
    if (!ocupados[safeName.toLowerCase()]) {
      return safeName;
    }
    const dot = safeName.lastIndexOf('.');
    const base = dot >= 0 ? safeName.substring(0, dot) : safeName;
    const ext = dot >= 0 ? safeName.substring(dot) : '';
    let counter = 1;
    let candidate = base + ' (' + counter + ')' + ext;
    while (ocupados[candidate.toLowerCase()]) {
      counter++;
      candidate = base + ' (' + counter + ')' + ext;
    }
    return candidate;
  }

  private _isPresupuestoFileName(fileName: string): boolean {
    return fileName.indexOf(PRESUPUESTO_FILE_PREFIX) === 0;
  }

  private _isFacturaFileName(fileName: string): boolean {
    return fileName.indexOf(FACTURA_FILE_PREFIX) === 0;
  }

  private _isReservedViajeAttachmentFileName(fileName: string): boolean {
    return this._isPresupuestoFileName(fileName) || this._isFacturaFileName(fileName);
  }

  private _buildFacturaFileName(originalFileName: string): string {
    const sanitized = this._sanitizeAttachmentFileName(originalFileName || 'documento');
    if (this._isFacturaFileName(sanitized)) {
      return sanitized;
    }
    return FACTURA_FILE_PREFIX + sanitized;
  }

  private _buildPresupuestoFileName(originalFileName: string): string {
    const sanitized = this._sanitizeAttachmentFileName(originalFileName || 'documento');
    if (this._isPresupuestoFileName(sanitized)) {
      return sanitized;
    }
    return PRESUPUESTO_FILE_PREFIX + sanitized;
  }

  private _buildViajeAttachmentsUrl(itemId: number): string {
    return this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_VIAJES + "')/items(" + itemId + ')/AttachmentFiles'
    );
  }

  private _buildPasajeroAttachmentsUrl(itemId: number): string {
    return this._buildUrl(
      "/_api/web/lists/getByTitle('" + LISTA_PASAJEROS + "')/items(" + itemId + ')/AttachmentFiles'
    );
  }

  private async _uploadAttachmentToPasajero(itemId: number, file: File): Promise<void> {
    const storedName = this._sanitizeAttachmentFileName(file.name);
    const escaped = this._escapeODataFileName(storedName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_PASAJEROS +
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

  private async _uploadAttachmentToViaje(itemId: number, file: File, fileName?: string): Promise<void> {
    const storedName = this._sanitizeAttachmentFileName(fileName || file.name);
    const escaped = this._escapeODataFileName(storedName);
    const url = this._buildUrl(
      "/_api/web/lists/getByTitle('" +
      LISTA_VIAJES +
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

  private async _addAttachmentToListItem(itemId: number, file: File): Promise<string> {
    const name = this._sanitizeAttachmentFileName(file.name);
    const escaped = this._escapeODataFileName(name);
    const relativePath =
      "/_api/web/lists/getByTitle('" +
      LISTA_LIQUIDACIONES +
      "')/items(" +
      itemId +
      ")/AttachmentFiles/add(FileName='" +
      escaped +
      "')";
    const url = this._buildUrl(relativePath);
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
    return name;
  }

  private async _getViajesFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_VIAJES, {
      Nombre: 'Title',
      DestinoGeneral: 'DestinoGeneral',
      Destino: 'Destino',
      FechaSalida: 'FechaSalida',
      FechaLlegada: 'FechaLlegada',
      Estado: 'Estado',
      ColorViaje: 'ColorViaje',
      Pasajeros: 'Pasajeros',
      Servicios: 'Servicios',
      Observaciones: 'Observaciones'
    }, {
      DestinoGeneral: 'Destino general'
    });
  }

  private async _getDestinosFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_DESTINOS, {
      DestinoGeneral: 'DestinoGeneral'
    }, {
      DestinoGeneral: 'Destino general'
    });
  }

  private async _getPagosFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_PAGOS, {
      ViajeAsociado: 'ViajeAsociado',
      Concepto: 'Title',
      FechaPago: 'FechaPago',
      Importe: 'Importe',
      MedioPago: 'MedioPago',
      Moneda: 'Moneda',
      TipoPago: 'TipoPago',
      Observaciones: 'Observaciones',
      Cotizacion: 'Cotizacion',
      LiquidacionOperador: 'LiquidacionOperador',
      ServicioViaje: 'ServicioViaje'
    });
  }

  private async _getPasajerosFieldMap(): Promise<IStringMap> {
    return this._getFieldMap(LISTA_PASAJEROS, {
      NombreApellido: 'Title',
      DNI: 'DNI',
      Pasaporte: 'Pasaporte',
      Telefono: 'Telefono',
      Email: 'Email',
      Observaciones: 'Observaciones',
      FechaNacimiento: 'FechaNacimiento'
    }, {
      NombreApellido: 'Nombre y Apellido'
    });
  }

  private async _getFieldMap(listTitle: string, fallbackByKey: IStringMap, displayNameByKey?: IStringMap): Promise<IStringMap> {
    if (this._fieldMaps[listTitle]) {
      return this._fieldMaps[listTitle];
    }

    const fieldsUrl = this._buildUrl(
      "/_api/web/lists/getByTitle('" + encodeURIComponent(listTitle).replace(/%20/g, ' ') + "')/fields?$select=Title,InternalName&$filter=Hidden eq false"
    );
    const json = await this._get(fieldsUrl);
    const fields: any[] = this._getResults(json);
    const byTitle: IStringMap = {};

    fields.forEach((field: any) => {
      byTitle[field.Title] = field.InternalName;
    });

    const map: IStringMap = {};
    Object.keys(fallbackByKey).forEach((key: string) => {
      if (key === 'Nombre') {
        map[key] = 'Title';
        return;
      }
      if (listTitle === LISTA_PAGOS && key === 'Concepto') {
        map[key] = 'Title';
        return;
      }
      const displayName = displayNameByKey && displayNameByKey[key] ? displayNameByKey[key] : key;
      map[key] = byTitle[displayName] || fallbackByKey[key];
    });

    this._fieldMaps[listTitle] = map;
    return map;
  }

  private _buildViajePayload(map: IStringMap, data: IViajeData): any {
    const payload: any = {};
    payload[map.Nombre] = data.nombre;
    // DestinoGeneral is only used by UI as parent filter; do not persist it in Registro de Viajes.
    payload[map.Destino + 'Id'] = data.destinoId;
    payload[map.FechaSalida] = this._toSharePointDateOnlyPayload(data.fechaSalida);
    payload[map.FechaLlegada] = this._toSharePointDateOnlyPayload(data.fechaLlegada);
    console.log('[fecha][input] FechaSalida:', data.fechaSalida);
    console.log('[fecha][input] FechaLlegada:', data.fechaLlegada);
    console.log('[fecha][payload] FechaSalida:', payload[map.FechaSalida]);
    console.log('[fecha][payload] FechaLlegada:', payload[map.FechaLlegada]);
    payload[map.Estado] = data.estado;
    payload[map.ColorViaje] = data.colorViaje;
    payload[map.Pasajeros + 'Id'] = data.pasajerosIds;
    payload[map.Servicios] = data.servicios;
    payload[map.Observaciones] = data.observaciones;
    return payload;
  }

  private _buildPagoPayload(map: IStringMap, data: IPagoData): any {
    const payload: any = {};
    payload[map.ViajeAsociado + 'Id'] = data.viajeId;
    payload[map.Concepto] = data.concepto;
    payload[map.FechaPago] = this._toIsoDateTime(data.fechaPago);
    payload[map.Importe] = data.importe;
    payload[map.MedioPago] = data.medioPago;
    payload[map.Moneda] = data.moneda;
    payload[map.TipoPago] = data.tipoPago === 'Egreso' ? 'Egreso' : 'Ingreso';
    payload[map.Observaciones] = data.observaciones || '';
    if (data.cotizacion !== undefined && data.cotizacion !== null) {
      payload[map.Cotizacion] = data.cotizacion > 0 ? data.cotizacion : null;
    } else {
      payload[map.Cotizacion] = null;
    }
    const liqKey = map.LiquidacionOperador + 'Id';
    if (data.liquidacionOperadorId !== undefined) {
      if (data.liquidacionOperadorId > 0) {
        payload[liqKey] = data.liquidacionOperadorId;
      } else {
        payload[liqKey] = null;
      }
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

  private _buildServicioViajePayload(map: IStringMap, data: IServicioViajeData): any {
    const payload: any = {};
    payload.Title = data.concepto;
    payload[map.ViajeAsociado + 'Id'] = data.viajeId;
    payload[map.Concepto] = data.concepto;
    payload[map.PrecioCliente] = data.precioCliente;
    payload[map.Moneda] = data.moneda;
    payload[map.Operador + 'Id'] = data.operadorId;
    return payload;
  }

  private _extractViajeId(item: any, lookupField: string): number {
    if (item[lookupField] && item[lookupField].Id) {
      return item[lookupField].Id;
    }
    if (item[lookupField + 'Id']) {
      return Number(item[lookupField + 'Id']);
    }
    return 0;
  }

  private _extractLookupId(item: any, fieldName: string): number {
    if (item[fieldName + 'Id'] !== undefined && item[fieldName + 'Id'] !== null) {
      return Number(item[fieldName + 'Id']) || 0;
    }
    if (item[fieldName] && item[fieldName].Id) {
      return Number(item[fieldName].Id) || 0;
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

  private _extractMultiChoiceValues(item: any, fieldName: string): string[] {
    const value = item[fieldName];
    if (Array.isArray(value)) {
      return value.map((entry: any) => String(entry));
    }
    if (value && Array.isArray(value.results)) {
      return value.results.map((entry: any) => String(entry));
    }
    if (value) {
      return [String(value)];
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

  private async _post(url: string, body?: any, extraHeaders?: ISPRequestHeaders): Promise<any> {
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

  private async _parseJsonResponse(response: SPHttpClientResponse, method: string, url: string): Promise<any> {
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
    if (Array.isArray(json?.d?.results)) {
      return json.d.results;
    }
    if (Array.isArray(json?.d)) {
      return json.d;
    }
    return [];
  }

  private _toDateInput(value: string): string {
    return toDateInput(value);
  }

  private _toIsoDateTime(value: string): string {
    if (!value) {
      return '';
    }
    return value + 'T00:00:00Z';
  }

  private _toSharePointDateOnlyPayload(value: string): string {
    return toSharePointDateOnlyPayload(value);
  }

  private _toNumber(value: any): number {
    const parsed = Number(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  private _getString(source: any, key: string): string {
    const value = source[key];
    return value === null || value === undefined ? '' : String(value);
  }

  private _isInvalidFieldError(error: Error, fieldInternalName: string): boolean {
    const message = (error && error.message ? error.message : '').toLowerCase();
    const field = (fieldInternalName || '').toLowerCase();
    return message.indexOf('does not exist') >= 0 && message.indexOf(field) >= 0;
  }
}
