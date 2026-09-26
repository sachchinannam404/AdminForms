/**
 * Client-side validation helpers and required-field indicators
 */

import { IFormField } from '../config/requestTypeRegistry';

export interface IValidationResult {
  valid: boolean;
  errors: Record<string, string>;
  firstError?: string;
}

export function isEmpty(value: any): boolean {
  if (value === undefined || value === null) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  if (typeof value === 'number' && isNaN(value)) return true;
  return false;
}

export function validateEmail(email: string): boolean {
  if (!email) return true; // optional unless required elsewhere
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

export function validateRequiredFields(
  fields: IFormField[],
  getValue: (field: IFormField) => any
): IValidationResult {
  const errors: Record<string, string> = {};
  for (const field of fields) {
    if (!field.required) continue;
    const v = getValue(field);
    if (isEmpty(v)) {
      errors[field.key] = `${field.label} is required`;
    }
  }
  const keys = Object.keys(errors);
  return {
    valid: keys.length === 0,
    errors,
    firstError: keys.length ? errors[keys[0]] : undefined
  };
}

export function validateEntity(
  data: Record<string, any>,
  rules: { key: string; label: string; required?: boolean; email?: boolean; min?: number }[]
): IValidationResult {
  const errors: Record<string, string> = {};
  for (const rule of rules) {
    const v = data[rule.key];
    if (rule.required && isEmpty(v)) {
      errors[rule.key] = `${rule.label} is required`;
      continue;
    }
    if (rule.email && v && !validateEmail(String(v))) {
      errors[rule.key] = `${rule.label} must be a valid email`;
    }
    if (rule.min !== undefined && typeof v === 'number' && v < rule.min) {
      errors[rule.key] = `${rule.label} must be at least ${rule.min}`;
    }
  }
  const keys = Object.keys(errors);
  return {
    valid: keys.length === 0,
    errors,
    firstError: keys.length ? errors[keys[0]] : undefined
  };
}
