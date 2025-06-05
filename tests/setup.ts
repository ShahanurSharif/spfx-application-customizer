/**
 * Test setup file for SPFx Extension E2E Tests
 */

/// <reference path="./types/global.d.ts" />

import { 
  DOMTestUtils, 
  BrowserAPIMocks 
} from './utils/TestUtils';

// Global test timeout
jest.setTimeout(30000);

// Mock console methods to reduce noise in tests
global.console = {
  ...console,
  // Uncomment to ignore specific console methods during tests
  // log: jest.fn(),
  // debug: jest.fn(),
  // info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
};

// Mock SharePoint context - extend global interfaces
declare global {
  namespace NodeJS {
    interface Global {
      window: Window & typeof globalThis;
      document: Document;
      testConfig: {
        sharePointSiteUrl: string;
        listName: string;
        extensionId: string;
        debugManifestUrl: string;
      };
    }
  }
}

// Set up global objects
(global as any).window = Object.create(window);
(global as any).document = Object.create(document);

// Set up global test data
(global as any).testConfig = {
  sharePointSiteUrl: process.env.SHAREPOINT_SITE_URL || 'https://yourtenant.sharepoint.com/sites/yoursite',
  listName: 'navbarcrud',
  extensionId: '5cf0b701-8c48-41ef-ad8c-cf66e6106423',
  debugManifestUrl: 'https://localhost:4321/temp/build/manifests.js'
};

// Initialize browser API mocks
BrowserAPIMocks.mockLocalStorage();
BrowserAPIMocks.mockPerformance();

// Helper function to create mock SharePoint context
export const createMockContext = () => ({
  serviceScope: {
    consume: jest.fn(),
    startNewChild: jest.fn(),
    finish: jest.fn()
  },
  pageContext: {
    web: {
      title: 'Test Site',
      absoluteUrl: (global as any).testConfig.sharePointSiteUrl,
      id: 'test-web-id'
    },
    list: {
      title: (global as any).testConfig.listName,
      id: 'test-list-id'
    },
    user: {
      displayName: 'Test User',
      email: 'test@example.com',
      loginName: 'test@example.com'
    }
  },
  spHttpClient: {
    get: jest.fn(),
    post: jest.fn(),
    fetch: jest.fn()
  }
});

// Helper function to create mock DOM elements
export const createMockDOMElements = () => {
  return DOMTestUtils.createSharePointDOM();
};

// Helper function to clean up DOM
export const cleanupDOM = () => {
  DOMTestUtils.cleanupDOM();
};
