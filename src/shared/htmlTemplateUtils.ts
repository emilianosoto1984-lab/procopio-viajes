export const RECIBO_PAGO_TEMPLATE_PLACEHOLDERS = [
  'LogoEmpresa',
  'NumeroRecibo',
  'Fecha',
  'NombreApellido',
  'DNI',
  'Concepto',
  'FormaPago',
  'FilaCuentaBancaria',
  'FilasImporte',
  'FilaCotizacion'
] as const;

export type ReciboPagoTemplatePlaceholder = typeof RECIBO_PAGO_TEMPLATE_PLACEHOLDERS[number];

export type ReciboPagoTemplateData = Partial<Record<ReciboPagoTemplatePlaceholder, string | null | undefined>>;

/**
 * Reemplaza placeholders {{Clave}} en la plantilla HTML.
 * Valores null/undefined se convierten a string vacío.
 */
export function renderTemplate(
  templateHtml: string,
  data: Record<string, string | null | undefined>
): string {
  if (!templateHtml) {
    return '';
  }

  let result = templateHtml;
  Object.keys(data).forEach((key: string) => {
    const rawValue = data[key];
    const value = rawValue === null || rawValue === undefined ? '' : String(rawValue);
    const placeholder = '{{' + key + '}}';
    result = result.split(placeholder).join(value);
  });
  return result;
}
