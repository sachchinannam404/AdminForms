/**
 * Lightweight unit tests for validation helpers (run with jest if configured).
 * These assert pure functions without SharePoint.
 */

import { isEmpty, validateEmail, validateEntity, validateRequiredFields } from '../utils/validation';
import { IFormField } from '../config/requestTypeRegistry';

describe('validation helpers', () => {
  it('isEmpty detects nullish and blank strings', () => {
    expect(isEmpty(undefined)).toBe(true);
    expect(isEmpty(null)).toBe(true);
    expect(isEmpty('')).toBe(true);
    expect(isEmpty('  ')).toBe(true);
    expect(isEmpty('ok')).toBe(false);
    expect(isEmpty(0)).toBe(false);
  });

  it('validateEmail accepts valid addresses', () => {
    expect(validateEmail('')).toBe(true);
    expect(validateEmail('a@b.com')).toBe(true);
    expect(validateEmail('bad')).toBe(false);
  });

  it('validateRequiredFields reports missing required', () => {
    const fields: IFormField[] = [
      { key: 'title', label: 'Title', type: 'text', required: true },
      { key: 'notes', label: 'Notes', type: 'text' }
    ];
    const result = validateRequiredFields(fields, (f) => (f.key === 'title' ? '' : 'x'));
    expect(result.valid).toBe(false);
    expect(result.errors.title).toContain('required');
  });

  it('validateEntity supports email and min rules', () => {
    const result = validateEntity(
      { email: 'bad', qty: 0 },
      [
        { key: 'email', label: 'Email', email: true },
        { key: 'qty', label: 'Qty', min: 1 }
      ]
    );
    expect(result.valid).toBe(false);
    expect(result.errors.email).toBeDefined();
    expect(result.errors.qty).toBeDefined();
  });
});
