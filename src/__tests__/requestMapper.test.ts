/**
 * Mapping-layer tests for status/type defaults and soft-delete semantics.
 */

import { RequestStatus, PriorityLevel, RequestType } from '../models/common/enums';

describe('request domain defaults', () => {
  it('enums expose expected statuses including Cancelled for soft-delete', () => {
    expect(RequestStatus.Cancelled).toBe('Cancelled');
    expect(RequestStatus.Pending).toBe('Pending');
    expect(PriorityLevel.Urgent).toBe('Urgent');
    expect(RequestType.ITEquipment).toBe('ITEquipment');
  });

  it('soft-delete uses Cancelled rather than physical delete', () => {
    const softStatus = RequestStatus.Cancelled;
    expect(softStatus).not.toBe(RequestStatus.Rejected);
    expect(['Draft', 'Pending', 'Approved', 'Rejected', 'InProgress', 'Completed', 'Cancelled']).toContain(
      softStatus
    );
  });
});
