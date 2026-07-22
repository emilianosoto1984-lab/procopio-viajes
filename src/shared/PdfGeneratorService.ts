import html2pdf from 'html2pdf.js';

type Html2PdfMargin = number | [number, number] | [number, number, number, number];

const A4_WIDTH_PX = 794;
const STYLE_SETTLE_MS = 300;
const LOG_SOURCE = '[PdfGeneratorService]';
const PDF_STYLE_ATTR = 'data-procopio-pdf-style';

export interface IPdfGeneratorHtml2CanvasOptions {
  scale?: number;
  useCORS?: boolean;
  allowTaint?: boolean;
  backgroundColor?: string;
  logging?: boolean;
  scrollX?: number;
  scrollY?: number;
  width?: number;
  height?: number;
}

export interface IPdfGeneratorJsPdfOptions {
  unit?: string;
  format?: string | [number, number];
  orientation?: 'portrait' | 'landscape';
}

export interface IPdfGeneratorImageOptions {
  type?: 'jpeg' | 'png' | 'webp';
  quality?: number;
}

export interface IPdfGeneratorOptions {
  margin?: Html2PdfMargin;
  filename?: string;
  image?: IPdfGeneratorImageOptions;
  html2canvas?: IPdfGeneratorHtml2CanvasOptions;
  jsPDF?: IPdfGeneratorJsPdfOptions;
}

interface IPdfRenderTarget {
  element: HTMLElement;
  host: HTMLDivElement;
  injectedStyles: HTMLStyleElement[];
  cleanup: () => void;
}

const DEFAULT_PDF_OPTIONS: IPdfGeneratorOptions = {
  margin: 10,
  filename: 'documento.pdf',
  image: {
    type: 'jpeg',
    quality: 0.98
  },
  html2canvas: {
    scale: 2,
    useCORS: true,
    allowTaint: true,
    backgroundColor: '#ffffff',
    logging: true,
    scrollX: 0,
    scrollY: 0
  },
  jsPDF: {
    unit: 'mm',
    format: 'a4',
    orientation: 'portrait'
  }
};

const OFFSCREEN_HOST_STYLES =
  'position:fixed;left:-10000px;top:0;width:' +
  A4_WIDTH_PX +
  'px;border:0;margin:0;padding:0;background:#ffffff;overflow:visible;';

export default class PdfGeneratorService {
  /**
   * Convierte HTML renderizado en un Blob PDF sin descargarlo al equipo del usuario.
   */
  public async generate(html: string, options?: IPdfGeneratorOptions): Promise<Blob> {
    const trimmedHtml = (html || '').trim();
    if (!trimmedHtml) {
      throw new Error('El HTML para generar el PDF está vacío.');
    }

    const target = this._createRenderTarget(trimmedHtml);

    try {
      await this._waitForImages(target.element);
      await this._settleLayout();

      this._logRenderDiagnostics(target);

      const canvasWidth = target.element.scrollWidth || target.element.offsetWidth || A4_WIDTH_PX;
      const canvasHeight = target.element.scrollHeight || target.element.offsetHeight;
      if (!canvasHeight) {
        throw new Error('El contenido HTML no tiene altura renderizable para generar el PDF.');
      }

      const pdfOptions = this._mergeOptions(options, canvasWidth, canvasHeight);
      const pdfBlob = await html2pdf()
        .set(pdfOptions)
        .from(target.element)
        .outputPdf('blob');

      if (!(pdfBlob instanceof Blob)) {
        throw new Error('No se pudo generar el PDF.');
      }

      return pdfBlob;
    } finally {
      target.cleanup();
    }
  }

  /**
   * html2canvas corre en el documento principal y no lee estilos del <head> de un iframe.
   * Se inyectan los <style> de la plantilla en document.head y el contenido va a un div fuera de pantalla.
   */
  private _createRenderTarget(html: string): IPdfRenderTarget {
    const fullHtml = /<html[\s>]/i.test(html)
      ? html
      : '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head><body>' + html + '</body></html>';

    const parsed = new DOMParser().parseFromString(fullHtml, 'text/html');
    const injectedStyles: HTMLStyleElement[] = [];
    const styleNodes = parsed.head.querySelectorAll('style');

    for (let i = 0; i < styleNodes.length; i++) {
      const sourceStyle = styleNodes[i];
      const styleEl = document.createElement('style');
      styleEl.setAttribute(PDF_STYLE_ATTR, 'true');
      styleEl.textContent = sourceStyle.textContent || '';
      document.head.appendChild(styleEl);
      injectedStyles.push(styleEl);
    }

    const host = document.createElement('div');
    host.className = 'procopio-pdf-render-host';
    host.setAttribute('aria-hidden', 'true');
    host.style.cssText = OFFSCREEN_HOST_STYLES;

    const receiptSource = parsed.querySelector('.receipt-page');
    if (receiptSource) {
      host.appendChild(receiptSource.cloneNode(true));
    } else if (parsed.body) {
      host.innerHTML = parsed.body.innerHTML;
    }

    document.body.appendChild(host);

    const receiptPage = host.querySelector('.receipt-page') as HTMLElement | null;
    const element = receiptPage || host;

    console.log(LOG_SOURCE, 'Contenedor temporal creado con HTML final (longitud):', fullHtml.length);
    console.log(LOG_SOURCE, 'Estilos inyectados en document.head:', injectedStyles.length);
    console.log(LOG_SOURCE, 'Clase .receipt-page encontrada:', !!receiptPage);

    return {
      element,
      host,
      injectedStyles,
      cleanup: () => {
        if (host.parentNode) {
          host.parentNode.removeChild(host);
        }
        injectedStyles.forEach((styleEl: HTMLStyleElement) => {
          if (styleEl.parentNode) {
            styleEl.parentNode.removeChild(styleEl);
          }
        });
      }
    };
  }

  private _logRenderDiagnostics(target: IPdfRenderTarget): void {
    const element = target.element;

    console.log(
      LOG_SOURCE,
      'HTML del contenedor temporal:',
      target.host.innerHTML.substring(0, 500)
    );

    const computedWidth = window.getComputedStyle(element).width;
    console.log(LOG_SOURCE, 'getComputedStyle(element).width:', computedWidth);

    const receiptBg = window.getComputedStyle(element).backgroundColor;
    console.log(LOG_SOURCE, 'getComputedStyle(element).backgroundColor:', receiptBg);

    const logoImg = element.querySelector('img') as HTMLImageElement | null;
    if (logoImg) {
      console.log(LOG_SOURCE, 'Logo naturalWidth:', logoImg.naturalWidth, 'naturalHeight:', logoImg.naturalHeight);
      if (!logoImg.naturalWidth) {
        console.warn(LOG_SOURCE, 'El logo no terminó de cargar antes de generar el PDF.');
      }
    } else {
      console.warn(LOG_SOURCE, 'No se encontró imagen de logo en el elemento a capturar.');
    }
  }

  private _mergeOptions(
    options: IPdfGeneratorOptions | undefined,
    canvasWidth: number,
    canvasHeight: number
  ): IPdfGeneratorOptions {
    const merged = options ? this._mergeBaseOptions(options) : DEFAULT_PDF_OPTIONS;
    const html2canvas = merged.html2canvas || {};

    return {
      ...merged,
      html2canvas: {
        scale: html2canvas.scale !== undefined ? html2canvas.scale : DEFAULT_PDF_OPTIONS.html2canvas!.scale,
        useCORS: html2canvas.useCORS !== undefined ? html2canvas.useCORS : DEFAULT_PDF_OPTIONS.html2canvas!.useCORS,
        allowTaint: html2canvas.allowTaint !== undefined ? html2canvas.allowTaint : DEFAULT_PDF_OPTIONS.html2canvas!.allowTaint,
        backgroundColor: html2canvas.backgroundColor || '#ffffff',
        logging: html2canvas.logging !== undefined ? html2canvas.logging : DEFAULT_PDF_OPTIONS.html2canvas!.logging,
        scrollX: html2canvas.scrollX !== undefined ? html2canvas.scrollX : 0,
        scrollY: html2canvas.scrollY !== undefined ? html2canvas.scrollY : 0,
        width: canvasWidth,
        height: canvasHeight
      }
    };
  }

  private _mergeBaseOptions(options: IPdfGeneratorOptions): IPdfGeneratorOptions {
    return {
      margin: options.margin !== undefined ? options.margin : DEFAULT_PDF_OPTIONS.margin,
      filename: options.filename || DEFAULT_PDF_OPTIONS.filename,
      image: {
        type: options.image && options.image.type ? options.image.type : DEFAULT_PDF_OPTIONS.image!.type,
        quality: options.image && options.image.quality !== undefined
          ? options.image.quality
          : DEFAULT_PDF_OPTIONS.image!.quality
      },
      html2canvas: {
        scale: options.html2canvas && options.html2canvas.scale !== undefined
          ? options.html2canvas.scale
          : DEFAULT_PDF_OPTIONS.html2canvas!.scale,
        useCORS: options.html2canvas && options.html2canvas.useCORS !== undefined
          ? options.html2canvas.useCORS
          : DEFAULT_PDF_OPTIONS.html2canvas!.useCORS,
        allowTaint: options.html2canvas && options.html2canvas.allowTaint !== undefined
          ? options.html2canvas.allowTaint
          : DEFAULT_PDF_OPTIONS.html2canvas!.allowTaint,
        logging: options.html2canvas && options.html2canvas.logging !== undefined
          ? options.html2canvas.logging
          : DEFAULT_PDF_OPTIONS.html2canvas!.logging,
        scrollX: options.html2canvas && options.html2canvas.scrollX !== undefined
          ? options.html2canvas.scrollX
          : DEFAULT_PDF_OPTIONS.html2canvas!.scrollX,
        scrollY: options.html2canvas && options.html2canvas.scrollY !== undefined
          ? options.html2canvas.scrollY
          : DEFAULT_PDF_OPTIONS.html2canvas!.scrollY,
        backgroundColor: options.html2canvas && options.html2canvas.backgroundColor
          ? options.html2canvas.backgroundColor
          : DEFAULT_PDF_OPTIONS.html2canvas!.backgroundColor
      },
      jsPDF: {
        unit: options.jsPDF && options.jsPDF.unit ? options.jsPDF.unit : DEFAULT_PDF_OPTIONS.jsPDF!.unit,
        format: options.jsPDF && options.jsPDF.format ? options.jsPDF.format : DEFAULT_PDF_OPTIONS.jsPDF!.format,
        orientation: options.jsPDF && options.jsPDF.orientation
          ? options.jsPDF.orientation
          : DEFAULT_PDF_OPTIONS.jsPDF!.orientation
      }
    };
  }

  private _settleLayout(): Promise<void> {
    return new Promise((resolve: () => void) => {
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          window.setTimeout(resolve, STYLE_SETTLE_MS);
        });
      });
    });
  }

  private _waitForImages(container: HTMLElement): Promise<void> {
    const images = Array.prototype.slice.call(container.querySelectorAll('img')) as HTMLImageElement[];

    if (!images.length) {
      return Promise.resolve();
    }

    const imagePromises = images.map((image: HTMLImageElement) => {
      if (image.complete && image.naturalWidth > 0) {
        return Promise.resolve();
      }

      return new Promise<void>((resolve: () => void) => {
        const finalize = (): void => {
          image.removeEventListener('load', finalize);
          image.removeEventListener('error', finalize);
          resolve();
        };

        image.addEventListener('load', finalize);
        image.addEventListener('error', finalize);
      });
    });

    return Promise.all(imagePromises).then(() => undefined);
  }
}
