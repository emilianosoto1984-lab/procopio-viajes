/**
 * Recupero de gastos bancarios sobre transferencias.
 * El monto ingresado YA INCLUYE el recupero: neto = bruto / (1 + %/100).
 */

export interface IDesgloseRecuperoBancario {
  aplicaRecupero: boolean;
  porcentaje: number;
  montoRecibido: number;
  montoAplicadoViaje: number;
  montoGastosBancarios: number;
}

/** Redondeo monetario a 2 decimales (evita residuos de floating point). */
export function roundMoney(value: number): number {
  const n = Number(value);
  if (!isFinite(n)) {
    return 0;
  }
  return Math.round(n * 100) / 100;
}

export function cuentaAplicaRecuperoBancario(params: {
  medioPago: string;
  aplicaRecupero?: boolean;
  porcentajeRecupero?: number;
}): boolean {
  if ((params.medioPago || '').trim() !== 'Transferencia') {
    return false;
  }
  if (params.aplicaRecupero !== true) {
    return false;
  }
  const porcentaje = Number(params.porcentajeRecupero);
  return isFinite(porcentaje) && porcentaje > 0;
}

/**
 * Calcula el desglose a partir del monto bruto recibido.
 * Si no aplica recupero: aplicado = recibido, gastos = 0, porcentaje = 0.
 */
export function calcularDesgloseRecuperoBancario(
  montoRecibido: number,
  params: {
    medioPago: string;
    aplicaRecupero?: boolean;
    porcentajeRecupero?: number;
  }
): IDesgloseRecuperoBancario {
  const bruto = roundMoney(montoRecibido);
  const aplica = cuentaAplicaRecuperoBancario(params) && bruto > 0;
  if (!aplica) {
    return {
      aplicaRecupero: false,
      porcentaje: 0,
      montoRecibido: bruto,
      montoAplicadoViaje: bruto,
      montoGastosBancarios: 0
    };
  }

  const porcentaje = roundMoney(Number(params.porcentajeRecupero));
  const factor = 1 + porcentaje / 100;
  const montoAplicadoViaje = roundMoney(bruto / factor);
  const montoGastosBancarios = roundMoney(bruto - montoAplicadoViaje);

  return {
    aplicaRecupero: true,
    porcentaje,
    montoRecibido: bruto,
    montoAplicadoViaje,
    montoGastosBancarios
  };
}

/**
 * Monto que impacta saldo del servicio.
 * Histórico sin columna: usa Monto completo.
 */
export function getMontoQueAplicaAlViaje(pago: {
  monto: number;
  montoAplicadoViaje?: number | null;
}): number {
  const aplicado = pago.montoAplicadoViaje;
  if (aplicado === null || aplicado === undefined) {
    return Number(pago.monto) || 0;
  }
  const parsed = Number(aplicado);
  if (!isFinite(parsed)) {
    return Number(pago.monto) || 0;
  }
  return parsed;
}

export function formatPorcentajeRecuperoDisplay(porcentaje: number): string {
  const n = Number(porcentaje) || 0;
  return n.toLocaleString('es-AR', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
}
