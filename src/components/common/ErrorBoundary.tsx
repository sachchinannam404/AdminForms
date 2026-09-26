import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Stack, Text } from '@fluentui/react';

export interface IErrorBoundaryProps {
  children: React.ReactNode;
  fallbackTitle?: string;
}

interface IState {
  hasError: boolean;
  error?: Error;
}

export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IState> {
  public state: IState = { hasError: false };

  public static getDerivedStateFromError(error: Error): IState {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('AdminForms ErrorBoundary:', error, info.componentStack);
  }

  private reset = (): void => {
    this.setState({ hasError: false, error: undefined });
  };

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 24 } }}>
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {this.props.fallbackTitle || 'Something went wrong in Admin Forms.'}
            {this.state.error?.message ? ` ${this.state.error.message}` : ''}
          </MessageBar>
          <Text variant="small">Check the browser console for details.</Text>
          <PrimaryButton text="Try again" onClick={this.reset} />
        </Stack>
      );
    }
    return this.props.children;
  }
}
