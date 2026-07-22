import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { renderTemplate as renderTemplateHtml } from './htmlTemplateUtils';

export { renderTemplate } from './htmlTemplateUtils';

export interface IDocumentTemplateServiceContext {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  webServerRelativeUrl: string;
}

const TEXT_RESPONSE_HEADERS = {
  Accept: 'text/html,application/xhtml+xml,text/plain,*/*'
};

export default class DocumentTemplateService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webAbsoluteUrl: string;
  private readonly _webServerRelativeUrl: string;

  public constructor(context: IDocumentTemplateServiceContext) {
    this._spHttpClient = context.spHttpClient;
    this._webAbsoluteUrl = (context.webAbsoluteUrl || '').replace(/\/$/, '');
    this._webServerRelativeUrl = (context.webServerRelativeUrl || '').replace(/\/$/, '');
  }

  /**
   * Lee el contenido de una plantilla HTML almacenada en SharePoint.
   * templatePath puede ser URL absoluta, ruta server-relative completa o relativa al sitio actual.
   */
  public async getTemplate(templatePath: string): Promise<string> {
    const trimmedPath = (templatePath || '').trim();
    if (!trimmedPath) {
      throw new Error('La ruta de la plantilla HTML es obligatoria.');
    }

    const serverRelativeUrl = this._resolveServerRelativeUrl(trimmedPath);
    const apiUrl = this._buildGetFileByServerRelativePathUrl(serverRelativeUrl);

    try {
      return await this._getText(apiUrl);
    } catch (apiError) {
      if (this._isAbsoluteHttpUrl(trimmedPath)) {
        return await this._getText(trimmedPath);
      }
      throw apiError;
    }
  }

  /**
   * Devuelve la URL absoluta de un archivo en SharePoint (imágenes, logos, etc.).
   */
  public getFileAbsoluteUrl(filePath: string): string {
    const serverRelativeUrl = this._resolveServerRelativeUrl(filePath);
    const origin = new URL(this._webAbsoluteUrl).origin;
    return origin + serverRelativeUrl;
  }

  /**
   * Reemplaza placeholders {{Clave}} en la plantilla HTML.
   * Valores null/undefined se convierten a string vacío.
   */
  public renderTemplate(
    template: string,
    data: Record<string, string | null | undefined>
  ): string {
    return renderTemplateHtml(template, data);
  }

  private _resolveServerRelativeUrl(templatePath: string): string {
    if (this._isAbsoluteHttpUrl(templatePath)) {
      const pathname = decodeURIComponent(new URL(templatePath).pathname);
      if (pathname) {
        return pathname;
      }
    }

    const normalizedPath = templatePath.replace(/\\/g, '/');
    if (this._webServerRelativeUrl && normalizedPath.indexOf(this._webServerRelativeUrl) === 0) {
      return normalizedPath;
    }

    if (normalizedPath.indexOf('/') === 0) {
      if (this._webServerRelativeUrl) {
        return this._webServerRelativeUrl + normalizedPath;
      }
      return normalizedPath;
    }

    if (this._webServerRelativeUrl) {
      return this._webServerRelativeUrl + '/' + normalizedPath;
    }

    return '/' + normalizedPath;
  }

  private _buildGetFileByServerRelativePathUrl(serverRelativeUrl: string): string {
    const encodedPath = encodeURIComponent(serverRelativeUrl);
    return (
      this._webAbsoluteUrl +
      "/_api/web/GetFileByServerRelativePath(decodedurl='" +
      encodedPath +
      "')/$value"
    );
  }

  private _isAbsoluteHttpUrl(value: string): boolean {
    return /^https?:\/\//i.test(value);
  }

  private async _getText(url: string): Promise<string> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: TEXT_RESPONSE_HEADERS }
    );

    if (!response.ok) {
      throw new Error(
        'No se pudo leer la plantilla HTML desde SharePoint. HTTP ' + response.status
      );
    }

    const content = await response.text();
    if (!content || !content.trim()) {
      throw new Error('La plantilla HTML está vacía.');
    }

    return content;
  }
}
