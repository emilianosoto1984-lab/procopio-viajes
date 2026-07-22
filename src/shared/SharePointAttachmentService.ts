import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISharePointAttachmentServiceContext {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
}

export interface ISharePointAttachment {
  fileName: string;
  serverRelativeUrl: string;
  fileSize?: number;
}

type ISPRequestHeaders = { [key: string]: string };

const GET_HEADERS: ISPRequestHeaders = {
  Accept: 'application/json;odata.metadata=minimal'
};

export default class SharePointAttachmentService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webUrl: string;

  public constructor(context: ISharePointAttachmentServiceContext) {
    this._spHttpClient = context.spHttpClient;
    this._webUrl = (context.webAbsoluteUrl || '').replace(/\/$/, '');
  }

  public async getAttachments(listTitle: string, itemId: number): Promise<ISharePointAttachment[]> {
    const url = this._buildAttachmentsUrl(listTitle, itemId);
    const json: any = await this._get(url);
    const files = this._getResults(json);
    return files
      .map((file: any) => ({
        fileName: this._getString(file, 'FileName'),
        serverRelativeUrl: this._getString(file, 'ServerRelativeUrl'),
        fileSize: file.FileLength !== undefined ? Number(file.FileLength) || 0 : undefined
      }))
      .filter((file: ISharePointAttachment) => file.fileName.length > 0);
  }

  public async attachmentExists(listTitle: string, itemId: number, fileName: string): Promise<boolean> {
    const normalizedName = (fileName || '').trim().toLowerCase();
    if (!normalizedName) {
      return false;
    }

    const attachments = await this.getAttachments(listTitle, itemId);
    return attachments.some(
      (attachment: ISharePointAttachment) => attachment.fileName.toLowerCase() === normalizedName
    );
  }

  public async addAttachment(
    listTitle: string,
    itemId: number,
    fileName: string,
    fileBlob: Blob
  ): Promise<void> {
    const normalizedName = (fileName || '').trim();
    if (!normalizedName) {
      throw new Error('El nombre del archivo adjunto es obligatorio.');
    }

    const exists = await this.attachmentExists(listTitle, itemId, normalizedName);
    if (exists) {
      throw new Error("Ya existe un archivo adjunto con el nombre '" + normalizedName + "'.");
    }

    const escapedName = this._escapeODataFileName(normalizedName);
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items(" +
      itemId +
      ")/AttachmentFiles/add(FileName='" +
      escapedName +
      "')";

    const buffer = await fileBlob.arrayBuffer();
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal',
          'Content-Type': 'application/octet-stream'
        },
        body: buffer
      }
    );

    if (!response.ok) {
      throw new Error(
        'No se pudo adjuntar el archivo a SharePoint. HTTP ' + response.status
      );
    }
  }

  public async deleteAttachment(listTitle: string, itemId: number, fileName: string): Promise<void> {
    const normalizedName = (fileName || '').trim();
    if (!normalizedName) {
      throw new Error('El nombre del archivo adjunto es obligatorio.');
    }

    const escapedName = this._escapeODataFileName(normalizedName);
    const url =
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items(" +
      itemId +
      ")/AttachmentFiles('" +
      escapedName +
      "')";

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal',
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*'
        }
      }
    );

    if (!response.ok) {
      throw new Error('No se pudo eliminar el adjunto en SharePoint. HTTP ' + response.status);
    }
  }

  private _buildAttachmentsUrl(listTitle: string, itemId: number): string {
    return (
      this._webUrl +
      "/_api/web/lists/getByTitle('" +
      listTitle +
      "')/items(" +
      itemId +
      ')/AttachmentFiles'
    );
  }

  private _escapeODataFileName(fileName: string): string {
    return fileName.replace(/'/g, "''");
  }

  private _getString(source: any, key: string): string {
    const value = source[key];
    return value === null || value === undefined ? '' : String(value);
  }

  private _getResults(json: any): any[] {
    if (Array.isArray(json.value)) {
      return json.value;
    }
    if (Array.isArray(json.results)) {
      return json.results;
    }
    return [];
  }

  private async _get(url: string): Promise<any> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: GET_HEADERS }
    );

    if (!response.ok) {
      throw new Error('Error al consultar adjuntos en SharePoint. HTTP ' + response.status);
    }

    return response.json();
  }
}
