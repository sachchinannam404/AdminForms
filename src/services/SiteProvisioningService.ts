/**
 * Ensures SharePoint lists, columns, versioning, and security groups exist.
 * Call from web part onInit (site owner) or run scripts/Provision-AdminForms.ps1.
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/site-groups/web';

export interface IProvisionResult {
  listsCreated: string[];
  listsUpdated: string[];
  groupsCreated: string[];
  messages: string[];
}

const REQUEST_TYPES =
  'Stationery;ITEquipment;Travel;Leave;Facilities;Procurement;General';
const STATUSES =
  'Draft;Pending;Approved;Rejected;InProgress;Completed;Cancelled';
const PRIORITIES = 'Low;Medium;High;Urgent';
const ITEM_STATUSES = 'Pending;Ordered;InStock;Delivered;Cancelled';

const CHILD_LISTS = [
  'Stationery Items',
  'IT Equipment Items',
  'Travel Items',
  'Facilities Items',
  'Procurement Items'
];

const APPROVER_GROUP = 'Admin Forms Approvers';
const ADMIN_GROUP = 'Admin Forms Admins';

export class SiteProvisioningService {
  public static async ensureSiteArtifacts(context?: any): Promise<IProvisionResult> {
    if (context) {
      sp.setup({ spfxContext: context });
    }

    const result: IProvisionResult = {
      listsCreated: [],
      listsUpdated: [],
      groupsCreated: [],
      messages: []
    };

    await this.ensureParentList(result);
    for (const title of CHILD_LISTS) {
      await this.ensureChildList(title, result);
    }
    await this.ensureGroup(APPROVER_GROUP, 'Approvers for Admin Forms requests', result);
    await this.ensureGroup(ADMIN_GROUP, 'Administrators for Admin Forms', result);

    return result;
  }

  private static async listExists(title: string): Promise<boolean> {
    try {
      await sp.web.lists.getByTitle(title).select('Id').get();
      return true;
    } catch {
      return false;
    }
  }

  private static async ensureParentList(result: IProvisionResult): Promise<void> {
    const title = 'Admin Requests';
    const exists = await this.listExists(title);

    if (!exists) {
      await sp.web.lists.add(title, 'Parent list for all admin request types', 100, true, {
        EnableVersioning: true
      });
      result.listsCreated.push(title);
      result.messages.push(`Created list: ${title}`);
    } else {
      try {
        await sp.web.lists.getByTitle(title).update({ EnableVersioning: true });
        result.messages.push(`Versioning enabled on ${title}`);
      } catch (e) {
        result.messages.push(`Could not enable versioning on ${title}: ${e}`);
      }
      result.listsUpdated.push(title);
    }

    const list = sp.web.lists.getByTitle(title);

    await this.ensureTextField(list, 'Description', true);
    await this.ensureChoiceField(list, 'RequestType', REQUEST_TYPES.split(';'), false);
    await this.ensureTextField(list, 'RequesterName', false);
    await this.ensureTextField(list, 'RequesterEmail', false);
    await this.ensureTextField(list, 'Department', false);
    await this.ensureChoiceField(list, 'Status', STATUSES.split(';'), false);
    await this.ensureChoiceField(list, 'Priority', PRIORITIES.split(';'), false);
    await this.ensureDateField(list, 'TargetDeliveryDate');
    await this.ensureTextField(list, 'ApprovedBy', false);
    await this.ensureDateField(list, 'ApprovedDate');
    await this.ensureTextField(list, 'RejectionReason', true);
    await this.ensureCurrencyField(list, 'TotalBudget');
    await this.ensureTextField(list, 'Comments', true);
    await this.ensureTextField(list, 'DetailsJson', true);
  }

  private static async ensureChildList(title: string, result: IProvisionResult): Promise<void> {
    const exists = await this.listExists(title);
    if (!exists) {
      await sp.web.lists.add(title, `Line items for ${title}`, 100, true);
      result.listsCreated.push(title);
      result.messages.push(`Created list: ${title}`);
    } else {
      result.listsUpdated.push(title);
    }

    const list = sp.web.lists.getByTitle(title);
    await this.ensureTextField(list, 'RequestId', false);
    await this.ensureTextField(list, 'ItemName', false);
    await this.ensureTextField(list, 'Category', false);
    await this.ensureNumberField(list, 'Quantity');
    await this.ensureTextField(list, 'Unit', false);
    await this.ensureCurrencyField(list, 'UnitPrice');
    await this.ensureCurrencyField(list, 'TotalPrice');
    await this.ensureTextField(list, 'Description', true);
    await this.ensureChoiceField(list, 'Status', ITEM_STATUSES.split(';'), false);
    await this.ensureTextField(list, 'VendorName', false);
    await this.ensureTextField(list, 'VendorEmail', false);
    await this.ensureDateField(list, 'ExpectedDeliveryDate');
    await this.ensureDateField(list, 'ActualDeliveryDate');
    await this.ensureTextField(list, 'Remarks', true);
    await this.ensureTextField(list, 'ExtraJson', true);
  }

  private static async fieldExists(list: any, internalName: string): Promise<boolean> {
    try {
      await list.fields.getByInternalNameOrTitle(internalName).get();
      return true;
    } catch {
      return false;
    }
  }

  private static async ensureTextField(
    list: any,
    name: string,
    multiline: boolean
  ): Promise<void> {
    if (await this.fieldExists(list, name)) return;
    if (multiline) {
      await list.fields.addMultilineText(name, undefined, false);
    } else {
      await list.fields.addText(name);
    }
  }

  private static async ensureChoiceField(
    list: any,
    name: string,
    choices: string[],
    multi: boolean
  ): Promise<void> {
    if (await this.fieldExists(list, name)) return;
    await list.fields.addChoice(name, choices, multi ? 1 : 0, false);
  }

  private static async ensureDateField(list: any, name: string): Promise<void> {
    if (await this.fieldExists(list, name)) return;
    await list.fields.addDateTime(name);
  }

  private static async ensureNumberField(list: any, name: string): Promise<void> {
    if (await this.fieldExists(list, name)) return;
    await list.fields.addNumber(name);
  }

  private static async ensureCurrencyField(list: any, name: string): Promise<void> {
    if (await this.fieldExists(list, name)) return;
    await list.fields.addCurrency(name);
  }

  private static async ensureGroup(
    title: string,
    description: string,
    result: IProvisionResult
  ): Promise<void> {
    try {
      const groups = await sp.web.siteGroups.filter(`Title eq '${title.replace(/'/g, "''")}'`).get();
      if (groups && groups.length) {
        result.messages.push(`Group already exists: ${title}`);
        return;
      }
      await sp.web.siteGroups.add({
        Title: title,
        Description: description,
        AllowMembersEditMembership: false,
        OnlyAllowMembersViewMembership: false
      });
      result.groupsCreated.push(title);
      result.messages.push(`Created group: ${title}`);
    } catch (e) {
      result.messages.push(`Group ${title}: ${e}`);
    }
  }
}
