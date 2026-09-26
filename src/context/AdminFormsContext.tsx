/**
 * Injectable context for services + current user (replaces pure static usage in UI)
 */

import * as React from 'react';
import { RequestService } from '../services/RequestService';
import { AttachmentService } from '../services/attachmentService';
import { RoleService, ICurrentUserContext } from '../services/RoleService';
import { NotificationService, INotificationConfig } from '../services/NotificationService';

export interface IAdminFormsServices {
  requestService: typeof RequestService;
  attachmentService: typeof AttachmentService;
  roleService: typeof RoleService;
  notificationService: typeof NotificationService;
}

export interface IAdminFormsContextValue {
  services: IAdminFormsServices;
  user: ICurrentUserContext | null;
  setUser: (u: ICurrentUserContext | null) => void;
  initialized: boolean;
}

const defaultServices: IAdminFormsServices = {
  requestService: RequestService,
  attachmentService: AttachmentService,
  roleService: RoleService,
  notificationService: NotificationService
};

const AdminFormsContext = React.createContext<IAdminFormsContextValue>({
  services: defaultServices,
  user: null,
  setUser: () => undefined,
  initialized: false
});

export interface IAdminFormsProviderProps {
  children: React.ReactNode;
  spfxContext?: any;
  displayName?: string;
  email?: string;
  notificationConfig?: INotificationConfig;
  /** Optional service overrides for tests */
  services?: Partial<IAdminFormsServices>;
}

export const AdminFormsProvider: React.FC<IAdminFormsProviderProps> = (props) => {
  const [user, setUser] = React.useState<ICurrentUserContext | null>(null);
  const [initialized, setInitialized] = React.useState(false);

  const services = React.useMemo(
    () => ({ ...defaultServices, ...props.services }),
    [props.services]
  );

  React.useEffect(() => {
    if (props.spfxContext) {
      RequestService.initialize(props.spfxContext);
    }
    if (props.notificationConfig) {
      NotificationService.configure(props.notificationConfig);
    }
    RoleService.getCurrentUser(props.displayName, props.email).then((u) => {
      setUser(u);
      setInitialized(true);
    });
  }, [props.spfxContext, props.displayName, props.email, props.notificationConfig]);

  const value = React.useMemo(
    () => ({ services, user, setUser, initialized }),
    [services, user, initialized]
  );

  return <AdminFormsContext.Provider value={value}>{props.children}</AdminFormsContext.Provider>;
};

export function useAdminForms(): IAdminFormsContextValue {
  return React.useContext(AdminFormsContext);
}
