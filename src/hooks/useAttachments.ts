import * as React from 'react';
import { IAttachment } from '../models/common/IAttachment';
import { AttachmentService } from '../services/attachmentService';

export function useAttachments(listTitle: string, itemId?: string) {
  const [attachments, setAttachments] = React.useState<IAttachment[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [pendingFiles, setPendingFiles] = React.useState<File[]>([]);
  const [error, setError] = React.useState<string | null>(null);

  const reload = React.useCallback(async () => {
    if (!itemId) {
      setAttachments([]);
      return;
    }
    setIsLoading(true);
    setError(null);
    try {
      const list = await AttachmentService.getAll(listTitle, itemId);
      setAttachments(list);
    } catch (e: any) {
      setError(e?.message || 'Failed to load attachments');
    } finally {
      setIsLoading(false);
    }
  }, [listTitle, itemId]);

  React.useEffect(() => {
    reload();
  }, [reload]);

  const uploadPending = async (targetItemId: string): Promise<IAttachment[]> => {
    const uploaded: IAttachment[] = [];
    for (const file of pendingFiles) {
      try {
        const att = await AttachmentService.upload(listTitle, targetItemId, file);
        uploaded.push(att);
      } catch (e) {
        console.error(e);
      }
    }
    setPendingFiles([]);
    await reload();
    return uploaded;
  };

  const remove = async (fileName: string) => {
    if (!itemId) return;
    await AttachmentService.delete(listTitle, itemId, fileName);
    await reload();
  };

  return {
    attachments,
    pendingFiles,
    setPendingFiles,
    isLoading,
    error,
    reload,
    uploadPending,
    remove
  };
}
