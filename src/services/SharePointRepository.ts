/**
 * Generic SharePoint list repository – list name + field maps, pagination, soft-delete support
 */

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export type EntityMapper<T> = (item: any) => T;
export type SharePointMapper<T> = (entity: Partial<T>) => Record<string, any>;

export interface IPagedResult<T> {
  items: T[];
  hasNext: boolean;
  /** Opaque skip token / next skip index for simple top+skip paging */
  nextSkip: number;
}

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

  public getListTitle(): string {
    return this.listTitle;
  }

  public async getById(id: string): Promise<T> {
    try {
      let query = sp.web.lists.getByTitle(this.listTitle).items.getById(parseInt(id, 10));
      if (this.defaultSelect?.length) {
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
      if (this.defaultSelect?.length) {
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

  /**
   * Simple top + skip pagination (SharePoint list items).
   * Prefer smaller page sizes (25–50) for large lists.
   */
  public async getPaged(
    options: {
      filter?: string;
      orderBy?: string;
      ascending?: boolean;
      top?: number;
      skip?: number;
    } = {}
  ): Promise<IPagedResult<T>> {
    const top = options.top ?? 50;
    const skip = options.skip ?? 0;
    const orderBy = options.orderBy ?? 'Created';
    const ascending = options.ascending ?? false;

    try {
      let query = sp.web.lists.getByTitle(this.listTitle).items.top(top + 1);
      if (skip > 0) {
        query = query.skip(skip) as any;
      }
      if (this.defaultSelect?.length) {
        query = query.select(...this.defaultSelect) as any;
      }
      if (options.filter) {
        query = query.filter(options.filter) as any;
      }
      query = query.orderBy(orderBy, ascending) as any;
      const raw = await query.get();
      const hasNext = raw.length > top;
      const page = hasNext ? raw.slice(0, top) : raw;
      return {
        items: page.map((item: any) => this.mapToEntity(item)),
        hasNext,
        nextSkip: skip + page.length
      };
    } catch (error) {
      console.error(`Error paging ${this.listTitle}:`, error);
      return { items: [], hasNext: false, nextSkip: skip };
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
      Object.keys(data).forEach((k) => {
        if (data[k] === undefined) delete data[k];
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

  /** Soft-delete: set Status = Cancelled (or provided status) instead of hard delete */
  public async softDelete(
    id: string,
    statusField: string = 'Status',
    statusValue: string = 'Cancelled'
  ): Promise<void> {
    try {
      await sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(parseInt(id, 10))
        .update({ [statusField]: statusValue });
    } catch (error) {
      console.error(`Error soft-deleting item ${id} in ${this.listTitle}:`, error);
      throw error;
    }
  }
}
