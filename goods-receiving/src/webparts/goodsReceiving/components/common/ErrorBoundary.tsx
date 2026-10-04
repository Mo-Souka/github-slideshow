import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';

interface IErrorBoundaryState {
  error?: Error;
}

/** Last line of defence: shows a message instead of a blank web part if rendering fails. */
export class ErrorBoundary extends React.Component<React.PropsWithChildren<{}>, IErrorBoundaryState> {
  public constructor(props: React.PropsWithChildren<{}>) {
    super(props);
    this.state = {};
  }

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error): void {
    console.error('Goods Receiving web part error', error);
  }

  public render(): React.ReactNode {
    if (this.state.error) {
      return (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
          Something went wrong while showing this page. Please reload the page. If the problem continues, contact your
          administrator and mention: {this.state.error.message}
        </MessageBar>
      );
    }
    return this.props.children;
  }
}
