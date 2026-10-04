import * as React from 'react';
import { PrimaryButton } from '@fluentui/react';
import { canCreateRecords } from '../../../../logic/workflow';
import { useAppContext } from '../AppContext';
import styles from '../App.module.scss';

export interface IRecordListProps {
  pendingOnly: boolean;
}

/** Placeholder until stage 4 (record list). */
export const RecordList: React.FC<IRecordListProps> = () => {
  const { user, navigate } = useAppContext();
  return (
    <div className={styles.header}>
      <h2 className={styles.title}>Goods Receiving</h2>
      {canCreateRecords(user) && <PrimaryButton text="New record" onClick={() => navigate({ name: 'new' })} />}
    </div>
  );
};
