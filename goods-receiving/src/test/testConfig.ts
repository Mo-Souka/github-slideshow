import { buildAppConfig } from '../config/appConfig';
import { IAppConfig, IFieldConfig } from '../models/IFieldConfig';
import { IUserContext } from '../models/IUserContext';
import { FieldValues, IReceivingRecord } from '../models/IReceivingRecord';

export function testConfig(): IAppConfig {
  return buildAppConfig({ siteUrl: 'https://contoso.sharepoint.com/sites/warehouse/' });
}

export function field(config: IAppConfig, key: string): IFieldConfig {
  const result = config.fields.filter((f) => f.key === key)[0];
  if (!result) throw new Error(`No field ${key}`);
  return result;
}

export function user(overrides: Partial<IUserContext> = {}): IUserContext {
  return {
    id: 10,
    loginName: 'i:0#.f|membership|jane@contoso.com',
    email: 'jane@contoso.com',
    displayName: 'Jane Receiver',
    isReceiver: true,
    isSupervisor: false,
    ...overrides
  };
}

export function record(values: FieldValues, id: number = 1): IReceivingRecord {
  return { id, values };
}
