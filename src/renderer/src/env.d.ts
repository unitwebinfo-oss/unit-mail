import type { OutlookApi } from '../../preload/index';

declare global {
  interface Window {
    outlook: OutlookApi;
  }
}

export {};
