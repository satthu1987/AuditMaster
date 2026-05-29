import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';

export interface ILoadingOverlayProps {
  message?: string;
}

const overlayStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  height: '100%',
  minHeight: 400,
  gap: 12
};

const LoadingOverlay: React.FC<ILoadingOverlayProps> = ({ message }) => (
  <div style={overlayStyle}>
    <Spinner size={SpinnerSize.large} />
    <span style={{ color: '#605e5c', fontSize: 14 }}>{message || 'Loading...'}</span>
  </div>
);

export default LoadingOverlay;
