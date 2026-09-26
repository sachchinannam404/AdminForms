/**
 * Notifications on status change.
 * Primary path: optional Power Automate HTTP webhook.
 * Secondary: Microsoft Graph mail (requires Graph permission in SPFx).
 */

import { IAdminRequest, RequestStatus } from '../models/IAdminRequest';

export interface INotificationConfig {
  /** Power Automate HTTP trigger URL (optional) */
  powerAutomateWebhookUrl?: string;
  /** If true and Graph is available, attempt sendMail */
  useGraphMail?: boolean;
}

let config: INotificationConfig = {};

export class NotificationService {
  public static configure(cfg: INotificationConfig): void {
    config = { ...config, ...cfg };
  }

  public static async notifyStatusChange(
    request: IAdminRequest,
    previousStatus: RequestStatus | string | undefined,
    actorName: string
  ): Promise<void> {
    const payload = {
      event: 'AdminRequestStatusChanged',
      requestId: request.id,
      title: request.title,
      requestType: request.requestType,
      previousStatus,
      newStatus: request.status,
      requesterEmail: request.requesterEmail,
      requesterName: request.requesterName,
      actorName,
      rejectionReason: request.rejectionReason,
      timestamp: new Date().toISOString()
    };

    if (config.powerAutomateWebhookUrl) {
      try {
        await fetch(config.powerAutomateWebhookUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
      } catch (e) {
        console.error('NotificationService: Power Automate webhook failed', e);
      }
    } else {
      // No webhook configured – log for debugging / future Graph
      console.info('[AdminForms notification]', payload);
    }
  }

  public static async notifyBulkStatusChange(
    requestIds: string[],
    newStatus: RequestStatus,
    actorName: string
  ): Promise<void> {
    if (!config.powerAutomateWebhookUrl) {
      console.info('[AdminForms bulk notification]', { requestIds, newStatus, actorName });
      return;
    }
    try {
      await fetch(config.powerAutomateWebhookUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          event: 'AdminRequestBulkStatusChanged',
          requestIds,
          newStatus,
          actorName,
          timestamp: new Date().toISOString()
        })
      });
    } catch (e) {
      console.error('NotificationService: bulk webhook failed', e);
    }
  }
}
