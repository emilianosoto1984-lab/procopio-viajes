import { SPHttpClient } from '@microsoft/sp-http';
import DocumentTemplateService from './DocumentTemplateService';
import PdfGeneratorService from './PdfGeneratorService';
import SharePointAttachmentService from './SharePointAttachmentService';
import { formatDateDisplay } from './sharePointDateUtils';
import {
  buildNumeroRecibo,
  buildReciboFileName,
  formatMontoRecibo
} from './reciboPagoUtils';

const LISTA_PAGOS = 'Registro de Pagos';
const TEMPLATE_PATH = '/Documentos compartidos/Templates/ReciboPago.html';
const LOGO_PATH = '/Documentos compartidos/Templates/logo-procopio.png';

export interface IReciboPagoServiceContext {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  webServerRelativeUrl: string;
}

export interface IReciboPagoGeneracionData {
  itemId: number;
  fechaPago: string;
  nombreApellido: string;
  dni: string;
  concepto: string;
  monto: number;
  moneda: string;
  formaPago: string;
}

export interface IReciboPagoGeneracionResult {
  success: boolean;
  skipped: boolean;
  numeroRecibo: string;
  fileName: string;
}

/**
 * Generación y adjunto de recibo PDF sobre la lista Registro de Pagos.
 * Compartido entre Registro de Pagos y Registro de Viajes.
 */
export default class ReciboPagoService {
  private readonly _templateService: DocumentTemplateService;
  private readonly _pdfGenerator: PdfGeneratorService;
  private readonly _attachmentService: SharePointAttachmentService;
  private readonly _enProceso: { [itemId: number]: boolean } = {};

  public constructor(context: IReciboPagoServiceContext) {
    this._templateService = new DocumentTemplateService({
      spHttpClient: context.spHttpClient,
      webAbsoluteUrl: context.webAbsoluteUrl,
      webServerRelativeUrl: context.webServerRelativeUrl
    });
    this._pdfGenerator = new PdfGeneratorService();
    this._attachmentService = new SharePointAttachmentService({
      spHttpClient: context.spHttpClient,
      webAbsoluteUrl: context.webAbsoluteUrl
    });
  }

  public async generarYAdjuntarRecibo(data: IReciboPagoGeneracionData): Promise<IReciboPagoGeneracionResult> {
    const numeroRecibo = buildNumeroRecibo(data.itemId, data.fechaPago);
    const fileName = buildReciboFileName(numeroRecibo);

    if (this._enProceso[data.itemId]) {
      console.log('[ReciboPagoService] Generación ya en curso para el ítem', data.itemId, '- se omite');
      return {
        success: true,
        skipped: true,
        numeroRecibo,
        fileName
      };
    }

    this._enProceso[data.itemId] = true;
    try {
      const yaExiste = await this._attachmentService.attachmentExists(
        LISTA_PAGOS,
        data.itemId,
        fileName
      );
      if (yaExiste) {
        console.log('[ReciboPagoService] Recibo ya adjunto, se omite la generación:', fileName);
        return {
          success: true,
          skipped: true,
          numeroRecibo,
          fileName
        };
      }

      return await this._generarYAdjuntar(data, numeroRecibo, fileName);
    } finally {
      delete this._enProceso[data.itemId];
    }
  }

  public async regenerarRecibo(data: IReciboPagoGeneracionData): Promise<IReciboPagoGeneracionResult> {
    const numeroRecibo = buildNumeroRecibo(data.itemId, data.fechaPago);
    const fileName = buildReciboFileName(numeroRecibo);

    if (this._enProceso[data.itemId]) {
      console.log('[ReciboPagoService] Regeneración ya en curso para el ítem', data.itemId, '- se omite');
      return {
        success: true,
        skipped: true,
        numeroRecibo,
        fileName
      };
    }

    this._enProceso[data.itemId] = true;
    try {
      await this._eliminarRecibosPreviosDelItem(data.itemId, fileName);
      return await this._generarYAdjuntar(data, numeroRecibo, fileName);
    } finally {
      delete this._enProceso[data.itemId];
    }
  }

  /** Elimina el recibo esperado y huérfanos Recibo-RP-*-{id}.pdf (p. ej. si cambió el año). */
  private async _eliminarRecibosPreviosDelItem(
    itemId: number,
    fileNameEsperado: string
  ): Promise<void> {
    const attachments = await this._attachmentService.getAttachments(LISTA_PAGOS, itemId);
    const expectedLower = (fileNameEsperado || '').trim().toLowerCase();
    const reciboPattern = new RegExp('^recibo-rp-\\d+-' + itemId + '\\.pdf$', 'i');

    for (let i = 0; i < attachments.length; i++) {
      const attachment = attachments[i];
      const name = (attachment.fileName || '').trim();
      if (!name) {
        continue;
      }
      const nameLower = name.toLowerCase();
      if (nameLower === expectedLower || reciboPattern.test(name)) {
        await this._attachmentService.deleteAttachment(LISTA_PAGOS, itemId, name);
        console.log('[ReciboPagoService] Recibo anterior eliminado para regeneración:', name);
      }
    }
  }

  /** Devuelve la URL del recibo adjunto esperado, o cadena vacía si no existe. */
  public async getReciboAttachmentUrl(itemId: number, fechaPago: string): Promise<string> {
    const fileName = buildReciboFileName(buildNumeroRecibo(itemId, fechaPago));
    const attachments = await this._attachmentService.getAttachments(LISTA_PAGOS, itemId);
    const found = attachments.filter(
      (attachment) => attachment.fileName.toLowerCase() === fileName.toLowerCase()
    )[0];
    return found && found.serverRelativeUrl ? found.serverRelativeUrl : '';
  }

  private async _generarYAdjuntar(
    data: IReciboPagoGeneracionData,
    numeroRecibo: string,
    fileName: string
  ): Promise<IReciboPagoGeneracionResult> {
    const template = await this._templateService.getTemplate(TEMPLATE_PATH);
    const html = this._templateService.renderTemplate(template, {
      LogoEmpresa: this._templateService.getFileAbsoluteUrl(LOGO_PATH),
      NumeroRecibo: numeroRecibo,
      Fecha: formatDateDisplay(data.fechaPago),
      NombreApellido: data.nombreApellido,
      DNI: data.dni,
      Concepto: data.concepto,
      Monto: formatMontoRecibo(data.monto, data.moneda),
      FormaPago: data.formaPago
    });

    const pdfBlob = await this._pdfGenerator.generate(html);
    await this._attachmentService.addAttachment(LISTA_PAGOS, data.itemId, fileName, pdfBlob);

    console.log('[ReciboPagoService] Recibo generado y adjuntado:', fileName);
    return {
      success: true,
      skipped: false,
      numeroRecibo,
      fileName
    };
  }
}
