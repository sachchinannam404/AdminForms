/**
 * Shared attachment model used across all request types
 */
export interface IAttachment {
  id: string;
  fileName: string;
  fileUrl: string;
  fileSize: number;
  uploadedBy?: string;
  uploadedDate: Date;
}
