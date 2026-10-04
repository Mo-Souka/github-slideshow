import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import { IRecordQuery } from '../../../../logic/caml';
import { buildExportFileName, buildExportSheet, downloadExcel, getExportFields } from '../../../../services/ExportService';
import { toAppError } from '../../../../services/errors';
import { useAppContext } from '../AppContext';
import { fireAndForget } from '../common/async';

export interface IExportButtonProps {
  query: IRecordQuery;
  disabled?: boolean;
}

/**
 * Exports ALL records matching the current search and filters (not only the
 * visible page) to an .xlsx file. Columns come from "includeInExport" in fields.json.
 */
export const ExportButton: React.FC<IExportButtonProps> = ({ query, disabled }) => {
  const { config, services, notify } = useAppContext();
  const [exporting, setExporting] = React.useState(false);
  const [loaded, setLoaded] = React.useState(0);
  const { maxRows, fileNamePrefix, sheetName } = config.solution.export;

  const runExport = async (): Promise<void> => {
    setExporting(true);
    setLoaded(0);
    try {
      const { records, truncated } = await services.records.queryAll(query, maxRows, (progress) => setLoaded(progress.loaded));
      if (records.length === 0) {
        notify({ type: 'info', text: 'There are no records to export for the current search and filters.' });
        return;
      }
      const sheet = buildExportSheet(records, getExportFields(config));
      await downloadExcel(sheet, sheetName, buildExportFileName(fileNamePrefix));
      notify(
        truncated
          ? {
              type: 'warning',
              text: `Only the first ${maxRows.toLocaleString('en-US')} records were exported. Narrow the date range or filters to export the rest.`
            }
          : { type: 'success', text: `Exported ${records.length.toLocaleString('en-US')} record(s) to Excel.` }
      );
    } catch (error) {
      notify({ type: 'warning', text: `Export failed. ${toAppError(error, { action: 'export records' }).userMessage}` });
    } finally {
      setExporting(false);
    }
  };

  return (
    <DefaultButton
      text={exporting ? (loaded > 0 ? `Exporting ${loaded.toLocaleString('en-US')}...` : 'Exporting...') : 'Export to Excel'}
      iconProps={{ iconName: 'ExcelDocument' }}
      disabled={disabled || exporting}
      onClick={() => fireAndForget(runExport())}
    />
  );
};
