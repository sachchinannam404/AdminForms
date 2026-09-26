import * as React from 'react';
import { IAdminRequest, IAdminRequestFilter, RequestStatus } from '../models/IAdminRequest';
import { RequestService } from '../services/RequestService';
import { IPagedResult } from '../services/SharePointRepository';

export interface IUseAdminRequestsResult {
  items: IAdminRequest[];
  totalApprox: number;
  isLoading: boolean;
  error: string | null;
  hasNext: boolean;
  reload: () => Promise<void>;
  loadMore: () => Promise<void>;
  pageSize: number;
}

export function useAdminRequests(
  filter?: IAdminRequestFilter,
  pageSize: number = 50
): IUseAdminRequestsResult {
  const [items, setItems] = React.useState<IAdminRequest[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [skip, setSkip] = React.useState(0);
  const [hasNext, setHasNext] = React.useState(false);

  const filterKey = JSON.stringify(filter || {});

  const reload = React.useCallback(async () => {
    setIsLoading(true);
    setError(null);
    setSkip(0);
    try {
      // Prefer full filter path (includes client search/date); pagination on unfiltered bulk via repo when needed
      const list = await RequestService.getRequests(filter);
      setItems(list.slice(0, pageSize));
      setHasNext(list.length > pageSize);
      setSkip(pageSize);
    } catch (e: any) {
      setError(e?.message || 'Failed to load requests');
      setItems([]);
    } finally {
      setIsLoading(false);
    }
  }, [filterKey, pageSize]); // eslint-disable-line react-hooks/exhaustive-deps

  const loadMore = React.useCallback(async () => {
    if (!hasNext || isLoading) return;
    setIsLoading(true);
    try {
      const list = await RequestService.getRequests(filter);
      const next = list.slice(skip, skip + pageSize);
      setItems((prev) => [...prev, ...next]);
      setSkip((s) => s + pageSize);
      setHasNext(skip + pageSize < list.length);
    } catch (e: any) {
      setError(e?.message || 'Failed to load more');
    } finally {
      setIsLoading(false);
    }
  }, [filter, filterKey, hasNext, isLoading, pageSize, skip]);

  React.useEffect(() => {
    reload();
  }, [reload]);

  return {
    items,
    totalApprox: items.length,
    isLoading,
    error,
    hasNext,
    reload,
    loadMore,
    pageSize
  };
}

export function useRequestDetail(requestId?: string) {
  const [request, setRequest] = React.useState<IAdminRequest | null>(null);
  const [isLoading, setIsLoading] = React.useState(!!requestId);
  const [error, setError] = React.useState<string | null>(null);

  const reload = React.useCallback(async () => {
    if (!requestId) {
      setRequest(null);
      return;
    }
    setIsLoading(true);
    setError(null);
    try {
      const data = await RequestService.getRequest(requestId);
      setRequest(data);
    } catch (e: any) {
      setError(e?.message || 'Failed to load request');
    } finally {
      setIsLoading(false);
    }
  }, [requestId]);

  React.useEffect(() => {
    reload();
  }, [reload]);

  return { request, isLoading, error, reload, setRequest };
}

void RequestStatus;
void IPagedResult;
