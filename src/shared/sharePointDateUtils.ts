/**
 * Utilidades para campos Date/DateTime de SharePoint sin corrimiento de zona horaria.
 * Los campos solo-fecha se guardan con mediodía UTC (T12:00:00Z), igual que Fecha de Partida.
 */

export function getDateOnlyFromSharePoint(value: string): string {
  if (!value) {
    return '';
  }
  return value.split('T')[0];
}

export function formatDateDisplay(value: string): string {
  if (!value) {
    return '';
  }
  const datePart = getDateOnlyFromSharePoint(value);
  const parts = datePart.split('-');
  if (parts.length !== 3) {
    return value;
  }
  const year = parts[0];
  const month = parts[1];
  const day = parts[2];
  return day + '/' + month + '/' + year;
}

export function getTodayDateInput(): string {
  return new Date().toISOString().substr(0, 10);
}

export function toDateInput(value: string): string {
  if (!value) {
    return '';
  }
  return value.length >= 10 ? value.substr(0, 10) : value;
}

export function toSharePointDateOnlyPayload(value: string): string {
  if (!value) {
    return '';
  }
  return value + 'T12:00:00Z';
}
