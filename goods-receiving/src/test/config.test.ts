import { defaultFieldsConfig, defaultSolutionConfig, validateConfig, buildAppConfig } from '../config/appConfig';
import { IFieldsConfigFile } from '../models/IFieldConfig';
import { AppError } from '../models/AppError';

function cloneFields(): IFieldsConfigFile {
  return JSON.parse(JSON.stringify(defaultFieldsConfig));
}

describe('configuration', () => {
  it('ships a valid fields.json and solution.json', () => {
    expect(validateConfig(defaultFieldsConfig, defaultSolutionConfig)).toEqual([]);
  });

  it('sorts fields by order and applies web part overrides', () => {
    const config = buildAppConfig({ siteUrl: 'https://x.sharepoint.com/sites/a/', listTitle: 'Other list', pageSize: 50 });
    const orders = config.fields.map((f) => f.order);
    expect(orders).toEqual(orders.slice().sort((a, b) => a - b));
    expect(config.solution.list.title).toBe('Other list');
    expect(config.solution.listView.pageSize).toBe(50);
    expect(config.siteUrl).toBe('https://x.sharepoint.com/sites/a');
  });

  it('reports duplicate keys and internal names', () => {
    const fields = cloneFields();
    fields.fields.push({ ...fields.fields[1] });
    const problems = validateConfig(fields, defaultSolutionConfig);
    expect(problems.some((p) => p.indexOf('key is used more than once') >= 0)).toBe(true);
    expect(problems.some((p) => p.indexOf('internalName') >= 0)).toBe(true);
  });

  it('reports a missing field the workflow depends on', () => {
    const fields = cloneFields();
    fields.fields = fields.fields.filter((f) => f.key !== 'approvalStatus');
    const problems = validateConfig(fields, defaultSolutionConfig);
    expect(problems.some((p) => p.indexOf('"approvalStatus" is missing') >= 0)).toBe(true);
  });

  it('reports requiredWhen values that are not choices', () => {
    const fields = cloneFields();
    const comments = fields.fields.filter((f) => f.key === 'receivingComments')[0];
    comments.requiredWhen!.in = ['Damaged'];
    const problems = validateConfig(fields, defaultSolutionConfig);
    expect(problems.some((p) => p.indexOf('"Damaged"') >= 0)).toBe(true);
  });

  it('throws a ConfigurationMismatch AppError for an invalid configuration', () => {
    const fields = cloneFields();
    fields.fields[2].type = 'Banana' as never;
    expect(() => buildAppConfig({ siteUrl: 'https://x' }, fields)).toThrow(AppError);
  });

  it('accepts a new field added only in the configuration', () => {
    const fields = cloneFields();
    fields.fields.push({
      key: 'sealNumber',
      internalName: 'SealNumber',
      displayName: 'Seal Number',
      type: 'Text',
      required: false,
      order: 75,
      section: 'delivery',
      showInTable: false,
      includeInExport: true,
      searchable: true
    });
    const config = buildAppConfig({ siteUrl: 'https://x' }, fields);
    expect(config.fields.map((f) => f.key)).toContain('sealNumber');
  });
});
