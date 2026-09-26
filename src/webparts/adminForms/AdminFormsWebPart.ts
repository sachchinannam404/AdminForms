import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { AdminFormsApp } from '../../components/AdminFormsApp';
import { RequestService } from '../../services/RequestService';
import { SiteProvisioningService } from '../../services/SiteProvisioningService';

export interface IAdminFormsWebPartProps {
  description: string;
  powerAutomateWebhookUrl: string;
  runProvisioningOnInit: boolean;
}

export default class AdminFormsWebPart extends BaseClientSideWebPart<IAdminFormsWebPartProps> {
  private _provisioningDone = false;

  public async onInit(): Promise<void> {
    await super.onInit();
    RequestService.initialize(this.context);

    if (this.properties.runProvisioningOnInit && !this._provisioningDone) {
      try {
        const result = await SiteProvisioningService.ensureSiteArtifacts(this.context);
        console.info('[AdminForms] Provisioning:', result);
        this._provisioningDone = true;
      } catch (e) {
        console.warn('[AdminForms] Provisioning skipped or failed (may need site owner rights):', e);
      }
    }
  }

  public render(): void {
    const user = this.context.pageContext.user;

    const element = React.createElement(AdminFormsApp, {
      spfxContext: this.context,
      currentUserName: user?.displayName,
      currentUserEmail: user?.email,
      powerAutomateWebhookUrl: this.properties.powerAutomateWebhookUrl || undefined
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Admin Forms settings' },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                PropertyPaneTextField('powerAutomateWebhookUrl', {
                  label: 'Power Automate webhook URL (optional)',
                  description:
                    'HTTP trigger URL for status-change notifications',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneToggle('runProvisioningOnInit', {
                  label: 'Ensure lists & groups on load',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneLabel('provisioningHint', {
                  text:
                    'Provisioning creates Admin Requests + child lists, columns (RequestType, DetailsJson), versioning, and groups “Admin Forms Approvers” / “Admin Forms Admins”. Requires site owner or full control.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
