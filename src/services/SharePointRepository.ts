/**
 * Generic SharePoint list repository – reusable for any list + entity mapping
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export type EntityMapper<T> = (item: any) => T;
export type SharePointMapper<T> = (entity: Partial<T>) => Record<string, any>;

export class SharePointRepository<T> {
  constructor(
    private listTitle: string,
    private mapToEntity: EntityMapper<T>,
    private mapToSharePoint: SharePointMapper<T>,
    private defaultSelect?: string[]
  ) {}

  public static initialize(context: any): void {
    sp.setup({ spfxContext: context });
  }

  public async getById(id: string): Promise<T> {
    try {
      let query = sp.web.lists.getByTitle(this.listTitle).items.getById(parseInt(id, 10));
      if (this.defaultSelect && this.defaultSelect.length) {
        query = query.select(...this.defaultSelect) as any;
      }
      const item = await query.get();
      return this.mapToEntity(item);
    } catch (error) {
      console.error(`Error getting item ${id} from ${this.listTitle}:`, error);
      throw error;
    }
  }

  public async getAll(
    filter?: string,
    orderBy: string = 'Created',
    ascending: boolean = false,
    top: number = 500
  ): Promise<T[]> {
    try {
      let query = sp.web.lists.getByTitle(this.listTitle).items.top(top);
      if (this.defaultSelect && this.defaultSelect.length) {
        query = query.select(...this.defaultSelect) as any;
      }
      if (filter) {
        query = query.filter(filter) as any;
      }
      query = query.orderBy(orderBy, ascending) as any;
      const items = await query.get();
      return items.map((item: any) => this.mapToEntity(item));
    } catch (error) {
      console.error(`Error getting items from ${this.listTitle}:`, error);
      return [];
    }
  }

  public async create(entity: Partial<T>): Promise<any> {
    try {
      const data = this.mapToSharePoint(entity);
      const result = await sp.web.lists.getByTitle(this.listTitle).items.add(data);
      return result.data;
    } catch (error) {
      console.error(`Error creating item in ${this.listTitle}:`, error);
      throw error;
    }
  }

  public async update(id: string, entity: Partial<T>): Promise<void> {
    try {
      const data = this.mapToSharePoint(entity);
      // Remove undefined / empty keys that should not overwrite
      Object.keys(data).forEach((k) => {
        if (data[k] === undefined) {
          delete data[k];
        }
      });
      await sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(parseInt(id, 10))
        .update(data);
    } catch (error) {
      console.error(`Error updating item ${id} in ${this.listTitle}:`, error);
      throw error;
    }
  }

  public async delete(id: string): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(parseInt(id, 10))
        .delete();
    } catch (error) {
      console.error(`Error deleting item ${id} from ${this.listTitle}:`, error);
      throw error;
    }
  }

  public getListTitle(): string {
    return this.listTitle;
  }
}
