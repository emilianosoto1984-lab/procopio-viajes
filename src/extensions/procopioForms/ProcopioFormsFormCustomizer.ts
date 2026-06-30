import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import ProcopioForms, { IProcopioFormsProps } from './components/ProcopioForms';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProcopioFormsFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'ProcopioFormsFormCustomizer';
const FORM_NAME = 'PROCOPIO REGISTRO DE VIAJES';

export default class ProcopioFormsFormCustomizer
  extends BaseFormCustomizer<IProcopioFormsFormCustomizerProperties> {

  private readonly VERSION = '1.0.58';

  public onInit(): Promise<void> {
    const cargadoEn = new Date().toLocaleString('es-AR');
    const urlActual = window.location.href;
    const modo = this._getModoEntorno(urlActual);

    console.log(
      `%c ${FORM_NAME} v${this.VERSION} | Cargado: ${cargadoEn} `,
      'background: green; color: white; font-size: 14px;'
    );
    console.log('[Procopio Forms] Contexto de carga', {
      formulario: FORM_NAME,
      version: this.VERSION,
      url: urlActual,
      modo,
      fechaHora: cargadoEn
    });

    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated ProcopioFormsFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    return Promise.resolve();
  }

  private _getModoEntorno(url: string): 'DEV' | 'PROD' {
    const urlLower = url.toLowerCase();
    if (
      urlLower.indexOf('localhost') >= 0 ||
      urlLower.indexOf('/workbench') >= 0 ||
      urlLower.indexOf('/sites/dev-') >= 0 ||
      urlLower.indexOf('dev-viajes') >= 0
    ) {
      return 'DEV';
    }
    return process.env.NODE_ENV === 'production' ? 'PROD' : 'DEV';
  }

  public render(): void {
    // Use this method to perform your custom rendering.

    const procopioForms: React.ReactElement<IProcopioFormsProps> =
      React.createElement(ProcopioForms, {
        context: this.context,
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose
       } as IProcopioFormsProps);

    ReactDOM.render(procopioForms, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
