import React, { Component, ErrorInfo, ReactNode } from "react";

interface Props {
  children?: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("Uncaught error:", error, errorInfo);
  }

  public render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: "20px", color: "red", backgroundColor: "white", zIndex: 9999, position: "relative" }}>
          <h1>App Crashed</h1>
          <pre>{typeof this.state.error?.message === 'object' ? JSON.stringify(this.state.error?.message) : String(this.state.error?.message)}</pre>
          <pre>{typeof this.state.error?.stack === 'object' ? JSON.stringify(this.state.error?.stack) : String(this.state.error?.stack)}</pre>
        </div>
      );
    }

    return this.props.children;
  }
}
