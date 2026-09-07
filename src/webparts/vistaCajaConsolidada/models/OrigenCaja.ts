/**
 * Origen de fondos de la vista Movimientos.
 * Efectivo es fijo; las cuentas bancarias se resuelven desde CuentasBancarias.
 */
export type OrigenCaja =
  | { tipo: 'efectivo'; id: 'efectivo'; nombre: 'Efectivo' }
  | { tipo: 'banco'; id: number; nombre: string };
