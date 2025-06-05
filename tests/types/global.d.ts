/**
 * Global type definitions for test environment
 */

declare global {
  interface Window {
    // Add any window-specific test properties here
  }

  namespace NodeJS {
    interface Global {
      window: Window & typeof globalThis;
      document: Document;
      fetch: jest.MockedFunction<typeof fetch>;
      performance: Performance & {
        mark: jest.MockedFunction<(name: string) => void>;
        measure: jest.MockedFunction<(name: string, startMark?: string, endMark?: string) => void>;
        getEntriesByName: jest.MockedFunction<(name: string) => PerformanceEntry[]>;
        clearMarks: jest.MockedFunction<(name?: string) => void>;
        clearMeasures: jest.MockedFunction<(name?: string) => void>;
        now: jest.MockedFunction<() => number>;
      };
      testConfig: {
        sharePointSiteUrl: string;
        listName: string;
        extensionId: string;
        debugManifestUrl: string;
      };
    }
  }

  var global: NodeJS.Global;

  namespace jest {
    interface Matchers<R> {
      // Add custom matchers if needed
    }
  }

  // Extend globalThis as well for modern environments
  interface globalThis {
    fetch: jest.MockedFunction<typeof fetch>;
    performance: Performance & {
      mark: jest.MockedFunction<(name: string) => void>;
      measure: jest.MockedFunction<(name: string, startMark?: string, endMark?: string) => void>;
      getEntriesByName: jest.MockedFunction<(name: string) => PerformanceEntry[]>;
      clearMarks: jest.MockedFunction<(name?: string) => void>;
      clearMeasures: jest.MockedFunction<(name?: string) => void>;
      now: jest.MockedFunction<() => number>;
    };
  }
}

export {};
