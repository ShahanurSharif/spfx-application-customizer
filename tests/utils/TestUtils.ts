/**
 * Test Utilities for SPFx Extension E2E Tests
 */

import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

/**
 * SharePoint List Mock Factory
 */
export class SharePointListMockFactory {
  private static instance: SharePointListMockFactory;
  
  public static getInstance(): SharePointListMockFactory {
    if (!SharePointListMockFactory.instance) {
      SharePointListMockFactory.instance = new SharePointListMockFactory();
    }
    return SharePointListMockFactory.instance;
  }

  /**
   * Create a mock SharePoint context
   */
  public createMockSharePointContext(overrides: Partial<ApplicationCustomizerContext> = {}): ApplicationCustomizerContext {
    const defaultContext = {
      serviceScope: {
        consume: jest.fn(),
        startNewChild: jest.fn(),
        finish: jest.fn()
      },
      pageContext: {
        web: {
          title: 'Test Site',
          absoluteUrl: 'https://test.sharepoint.com',
          id: 'test-web-id'
        },
        list: {
          title: 'navbarcrud',
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
    };

    return { ...defaultContext, ...overrides } as ApplicationCustomizerContext;
  }

  /**
   * Create mock SharePoint list data
   */
  public createMockListData(itemCount: number = 5): any[] {
    const items = [];
    for (let i = 1; i <= itemCount; i++) {
      items.push({
        Id: i,
        Title: `Test Item ${i}`,
        value: i % 2 === 0 ? '#0078d4' : '16',
        Created: new Date(),
        Modified: new Date()
      });
    }
    return items;
  }
}

/**
 * DOM Test Utilities
 */
export class DOMTestUtils {
  /**
   * Create mock SharePoint DOM elements
   */
  public static createSharePointDOM(): HTMLElement[] {
    const elements: HTMLElement[] = [];

    // Site header
    const siteHeader = document.createElement('div');
    siteHeader.id = 'spSiteHeader';
    siteHeader.setAttribute('data-automationid', 'SiteHeader');
    document.body.appendChild(siteHeader);
    elements.push(siteHeader);

    // Navigation container
    const navContainer = document.createElement('div');
    navContainer.className = 'ms-siteHeader-container';
    siteHeader.appendChild(navContainer);
    elements.push(navContainer);

    // Horizontal nav item
    const horizontalNavItem = document.createElement('div');
    horizontalNavItem.className = 'ms-HorizontalNavItem';
    horizontalNavItem.setAttribute('data-automationid', 'HorizontalNav-link');
    navContainer.appendChild(horizontalNavItem);
    elements.push(horizontalNavItem);

    // Nav link text
    const linkText = document.createElement('span');
    linkText.className = 'ms-HorizontalNavItem-linkText';
    linkText.textContent = 'Test Nav Item';
    horizontalNavItem.appendChild(linkText);
    elements.push(linkText);

    return elements;
  }

  /**
   * Clean up DOM elements created during tests
   */
  public static cleanupDOM(): void {
    const elementsToRemove = [
      '#spSiteHeader',
      '#monarch360SettingsBtn',
      '.monarch360-settings-dialog',
      '.monarch360-dynamic-styles',
      'style[data-monarch360-styles]'
    ];
    
    elementsToRemove.forEach(selector => {
      const elements = document.querySelectorAll(selector);
      elements.forEach(element => element.remove());
    });
  }

  /**
   * Wait for element to appear in DOM
   */
  public static async waitForElement(selector: string, timeout: number = 5000): Promise<Element | null> {
    return new Promise((resolve) => {
      const element = document.querySelector(selector);
      if (element) {
        resolve(element);
        return;
      }

      const observer = new MutationObserver(() => {
        const element = document.querySelector(selector);
        if (element) {
          observer.disconnect();
          resolve(element);
        }
      });

      observer.observe(document.body, {
        childList: true,
        subtree: true
      });

      setTimeout(() => {
        observer.disconnect();
        resolve(null);
      }, timeout);
    });
  }
}

/**
 * Browser API Mocks
 */
export class BrowserAPIMocks {
  /**
   * Mock localStorage for tests
   */
  public static mockLocalStorage(): void {
    const localStorageMock = {
      getItem: jest.fn(),
      setItem: jest.fn(),
      removeItem: jest.fn(),
      clear: jest.fn(),
      length: 0,
      key: jest.fn()
    };

    Object.defineProperty(window, 'localStorage', {
      value: localStorageMock,
      writable: true
    });
  }

  /**
   * Mock performance API for tests
   */
  public static mockPerformance(): void {
    const performanceMock = {
      now: jest.fn(() => Date.now()),
      mark: jest.fn(),
      measure: jest.fn(),
      getEntriesByType: jest.fn(() => []),
      getEntriesByName: jest.fn(() => []),
      clearMarks: jest.fn(),
      clearMeasures: jest.fn()
    };

    Object.defineProperty(window, 'performance', {
      value: performanceMock,
      writable: true
    });

    Object.defineProperty(global, 'performance', {
      value: performanceMock,
      writable: true
    });
  }

  /**
   * Mock fetch API for tests
   */
  public static mockFetch(): void {
    (global as any).fetch = jest.fn(() =>
      Promise.resolve({
        ok: true,
        status: 200,
        json: () => Promise.resolve({}),
        text: () => Promise.resolve(''),
        headers: new Headers(),
        statusText: 'OK'
      })
    ) as jest.Mock;
  }
}

/**
 * Test Data Generators
 */
export class TestDataGenerators {
  /**
   * Generate random color
   */
  public static generateRandomColor(): string {
    const colors = ['#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#ffffff', '#000000'];
    return colors[Math.floor(Math.random() * colors.length)];
  }

  /**
   * Generate random font size
   */
  public static generateRandomFontSize(): number {
    return Math.floor(Math.random() * 20) + 10; // 10-30px
  }

  /**
   * Generate test user data
   */
  public static generateTestUser(): any {
    return {
      displayName: 'Test User',
      email: 'test@example.com',
      loginName: 'test@example.com',
      id: 'test-user-id'
    };
  }
}

/**
 * Performance Test Utilities
 */
export class PerformanceTestUtils {
  /**
   * Measure execution time of a function
   */
  public static async measureExecutionTime<T>(fn: () => Promise<T> | T): Promise<{ result: T; duration: number }> {
    const startTime = performance.now();
    const result = await fn();
    const endTime = performance.now();
    return {
      result,
      duration: endTime - startTime
    };
  }

  /**
   * Create memory usage snapshot
   */
  public static getMemorySnapshot(): any {
    if (typeof (window as any).performance?.memory !== 'undefined') {
      return {
        usedJSHeapSize: (window as any).performance.memory.usedJSHeapSize,
        totalJSHeapSize: (window as any).performance.memory.totalJSHeapSize,
        jsHeapSizeLimit: (window as any).performance.memory.jsHeapSizeLimit
      };
    }
    return null;
  }
}
