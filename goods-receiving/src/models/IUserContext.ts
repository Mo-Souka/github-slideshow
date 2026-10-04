export interface IUserContext {
  /** SharePoint site user ID. */
  id: number;
  loginName: string;
  email: string;
  displayName: string;
  /** Member of the receivers group (or can add items to the list). */
  isReceiver: boolean;
  /** Member of the supervisors group (or has the Approve Items permission on the list). */
  isSupervisor: boolean;
}
