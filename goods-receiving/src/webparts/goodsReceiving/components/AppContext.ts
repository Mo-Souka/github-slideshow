import * as React from 'react';
import { IAppConfig } from '../../../models/IFieldConfig';
import { IUserContext } from '../../../models/IUserContext';
import { ReceivingRecordService } from '../../../services/ReceivingRecordService';
import { AttachmentService } from '../../../services/AttachmentService';
import { UserService } from '../../../services/UserService';
import { Route } from './useHashRoute';

export interface IAppServices {
  records: ReceivingRecordService;
  attachments: AttachmentService;
  users: UserService;
}

export interface INotification {
  text: string;
  type: 'success' | 'info' | 'warning';
}

export interface IAppContext {
  config: IAppConfig;
  user: IUserContext;
  services: IAppServices;
  navigate: (route: Route) => void;
  /**
   * Shows a message (e.g. "Record saved"). With `then`, navigates to that screen
   * and shows the message there; otherwise it shows on the current screen.
   */
  notify: (notification: INotification, then?: Route) => void;
  /** Container width in pixels, used to switch between phone and desktop layouts. */
  width: number;
}

export const AppContext = React.createContext<IAppContext | undefined>(undefined);

export function useAppContext(): IAppContext {
  const context = React.useContext(AppContext);
  if (!context) {
    throw new Error('useAppContext must be used inside <AppContext.Provider>.');
  }
  return context;
}

/** Layout breakpoint: below this width the app uses the phone layout (cards, stacked buttons). */
export const NARROW_WIDTH = 640;
