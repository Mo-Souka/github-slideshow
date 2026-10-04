import { SPFI } from '@pnp/sp';
import { PermissionKind } from '@pnp/sp/security';
import { IAppConfig } from '../models/IFieldConfig';
import { IPersonValue } from '../models/IReceivingRecord';
import { IUserContext } from '../models/IUserContext';
import { toAppError } from './errors';

export interface IPersonSuggestion {
  /** Claims login name, e.g. i:0#.f|membership|jane@contoso.com */
  key: string;
  displayName: string;
  email?: string;
  jobTitle?: string;
}

/** Current user, role detection and people search. */
export class UserService {
  private readonly _sp: SPFI;
  private readonly _config: IAppConfig;
  private readonly _ensuredUsers: Record<string, IPersonValue> = {};

  public constructor(sp: SPFI, config: IAppConfig) {
    this._sp = sp;
    this._config = config;
  }

  /**
   * Loads the signed-in user and works out their role.
   *
   * Supervisor: direct member of the supervisors group, OR has the
   *   "Approve Items" permission on the list (covers people who are members
   *   through a nested Microsoft 365 / security group, and site owners).
   * Receiver: direct member of the receivers group, OR can add items to the list.
   */
  public async getCurrentUser(): Promise<IUserContext> {
    try {
      const list = this._sp.web.lists.getByTitle(this._config.solution.list.title);
      const [user, groups, permissions] = await Promise.all([
        this._sp.web.currentUser.select('Id', 'Title', 'Email', 'LoginName')(),
        this._sp.web.currentUser.groups.select('Title')().catch(() => [] as { Title: string }[]),
        list.getCurrentUserEffectivePermissions().catch(() => undefined)
      ]);
      const groupTitles = groups.map((g) => g.Title.toLowerCase());
      const inGroup = (name: string): boolean => groupTitles.indexOf(name.toLowerCase()) >= 0;
      const canApprove = permissions ? list.hasPermissions(permissions, PermissionKind.ApproveItems) : false;
      const canAdd = permissions ? list.hasPermissions(permissions, PermissionKind.AddListItems) : false;
      const isSupervisor = inGroup(this._config.solution.groups.supervisors) || canApprove;
      const isReceiver = isSupervisor || inGroup(this._config.solution.groups.receivers) || canAdd;
      return {
        id: user.Id,
        loginName: user.LoginName,
        email: user.Email,
        displayName: user.Title,
        isReceiver,
        isSupervisor
      };
    } catch (error) {
      throw toAppError(error, { action: 'load your user profile', listTitle: this._config.solution.list.title });
    }
  }

  /** People picker search across the organisation. */
  public async searchPeople(text: string): Promise<IPersonSuggestion[]> {
    if (!text || text.trim().length < 2) return [];
    try {
      const results = await this._sp.profiles.clientPeoplePickerSearchUser({
        AllowEmailAddresses: false,
        AllowMultipleEntities: false,
        MaximumEntitySuggestions: 10,
        PrincipalSource: 15,
        PrincipalType: 1,
        QueryString: text.trim()
      });
      return results.map((entity) => {
        const data = (entity.EntityData || {}) as unknown as Record<string, string>;
        return {
          key: entity.Key,
          displayName: entity.DisplayText,
          email: data.Email || undefined,
          jobTitle: data.Title || undefined
        };
      });
    } catch (error) {
      throw toAppError(error, { action: 'search for people' });
    }
  }

  /** Makes sure the person exists in the site and returns their site user ID. */
  public async ensureUser(suggestion: IPersonSuggestion): Promise<IPersonValue> {
    const cached = this._ensuredUsers[suggestion.key];
    if (cached) return cached;
    try {
      const user = await this._sp.web.ensureUser(suggestion.key);
      const person: IPersonValue = {
        id: user.Id,
        title: user.Title || suggestion.displayName,
        email: user.Email || suggestion.email,
        loginName: user.LoginName
      };
      this._ensuredUsers[suggestion.key] = person;
      return person;
    } catch (error) {
      throw toAppError(error, { action: 'select this person' });
    }
  }
}
