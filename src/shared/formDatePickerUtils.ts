import { IDatePickerStrings } from '@fluentui/react';
import { getDateOnlyFromSharePoint } from './sharePointDateUtils';

export const FORM_DATE_PICKER_STRINGS: IDatePickerStrings = {
  months: [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ],
  shortMonths: ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'],
  days: ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'],
  shortDays: ['D', 'L', 'M', 'M', 'J', 'V', 'S'],
  goToToday: 'Ir a hoy',
  prevMonthAriaLabel: 'Mes anterior',
  nextMonthAriaLabel: 'Mes siguiente',
  prevYearAriaLabel: 'Año anterior',
  nextYearAriaLabel: 'Año siguiente',
  closeButtonAriaLabel: 'Cerrar',
  monthPickerHeaderAriaLabel: '{0}, seleccionar para cambiar el año',
  yearPickerHeaderAriaLabel: '{0}, seleccionar para cambiar el mes',
  isRequiredErrorMessage: 'La fecha es obligatoria.',
  invalidInputErrorMessage: 'Formato de fecha inválido. Use dd/mm/aaaa.'
};

export function dateToInputString(date: Date): string {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const monthStr = month < 10 ? '0' + month : String(month);
  const dayStr = day < 10 ? '0' + day : String(day);
  return year + '-' + monthStr + '-' + dayStr;
}

export function parseDateStringToDate(value: string): Date | undefined {
  const datePart = getDateOnlyFromSharePoint(value);
  const parts = datePart.split('-');
  if (parts.length !== 3) {
    return undefined;
  }
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!year || !month || !day) {
    return undefined;
  }
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return undefined;
  }
  return date;
}

export function formatDateForPicker(date?: Date): string {
  if (!date) {
    return '';
  }
  const day = date.getDate();
  const month = date.getMonth() + 1;
  const dayStr = day < 10 ? '0' + day : String(day);
  const monthStr = month < 10 ? '0' + month : String(month);
  return dayStr + '/' + monthStr + '/' + date.getFullYear();
}

export function fechaTextoDesdeValor(value: string): string {
  const date = parseDateStringToDate(value);
  return date ? formatDateForPicker(date) : '';
}

export function aplicarMascaraFecha(value: string): string {
  const digits = (value || '').replace(/\D/g, '').substring(0, 8);
  if (digits.length <= 2) {
    return digits;
  }
  if (digits.length <= 4) {
    return digits.substring(0, 2) + '/' + digits.substring(2);
  }
  return digits.substring(0, 2) + '/' + digits.substring(2, 4) + '/' + digits.substring(4);
}

export function parseDateFromPickerString(dateStr: string): Date | null {
  const trimmed = (dateStr || '').trim();
  if (!trimmed) {
    return null;
  }
  const parts = trimmed.split('/');
  if (parts.length !== 3) {
    return null;
  }
  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);
  if (!day || !month || !year || parts[2].length < 4) {
    return null;
  }
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}
