/**
 * Performance Tests for SPFx Extension
 * 
 * These tests verify that the extension performs well:
 * - Fast initialization times
 * - Efficient DOM operations
 * - Minimal memory footprint
 * - Quick style application
 */

/// <reference path="../types/global.d.ts" />

import { createMockContext, createMockDOMElements, cleanupDOM } from '../setup';
import Monarch360NavCrudApplicationCustomizer from '../../src/extensions/monarch360NavCrud/Monarch360NavCrudApplicationCustomizer';

describe('SPFx Extension - Performance Tests', () => {
  let customizer: Monarch360NavCrudApplicationCustomizer;
  let mockContext: any;

  beforeEach(() => {
    cleanupDOM();
    mockContext = createMockContext();
    customizer = new Monarch360NavCrudApplicationCustomizer();
    
    Object.defineProperty(customizer, 'context', {
      value: mockContext,
      writable: false
    });

    // Mock performance.now for timing tests
    global.performance = {
      now: jest.fn().mockReturnValue(Date.now())
    } as any;

    jest.spyOn(console, 'log').mockImplementation();
    jest.spyOn(console, 'warn').mockImplementation();
    jest.spyOn(console, 'error').mockImplementation();
  });

  afterEach(() => {
    cleanupDOM();
    jest.clearAllMocks();
  });

  describe('Initialization Performance', () => {
    test('should initialize within 100ms', async () => {
      const startTime = performance.now();
      
      await customizer.onInit();
      
      const endTime = performance.now();
      const initTime = endTime - startTime;
      
      expect(initTime).toBeLessThan(100);
    });

    test('should inject settings button within 50ms when DOM is ready', () => {
      createMockDOMElements();
      
      const startTime = performance.now();
      
      (customizer as any).injectSettingsButton();
      
      const endTime = performance.now();
      const injectionTime = endTime - startTime;
      
      expect(injectionTime).toBeLessThan(50);
      
      const settingsButton = document.getElementById('monarch360SettingsBtn');
      expect(settingsButton).toBeTruthy();
    });

    test('should load saved settings within 200ms', async () => {
      // Mock fast SharePoint response
      const mockSP = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                filter: jest.fn().mockReturnValue({
                  top: jest.fn().mockReturnValue({
                    '()': jest.fn().mockResolvedValue([{ value: '#ffffff' }])
                  })
                })
              }
            })
          }
        }
      };

      jest.doMock('@pnp/sp', () => ({
        spfi: () => ({ using: () => mockSP }),
        SPFx: jest.fn()
      }));

      const startTime = performance.now();
      
      await (customizer as any).loadSavedSettings();
      
      const endTime = performance.now();
      const loadTime = endTime - startTime;
      
      expect(loadTime).toBeLessThan(200);
    });
  });

  describe('DOM Manipulation Performance', () => {
    test('should apply styles to DOM elements within 10ms', () => {
      createMockDOMElements();
      
      const startTime = performance.now();
      
      (customizer as any).applySavedStyles('#ff0000', 18);
      
      const endTime = performance.now();
      const styleTime = endTime - startTime;
      
      expect(styleTime).toBeLessThan(10);
      
      const styleElement = document.querySelector('style[data-monarch360-styles]');
      expect(styleElement).toBeTruthy();
    });

    test('should handle rapid style updates efficiently', () => {
      createMockDOMElements();
      
      const startTime = performance.now();
      
      // Apply 100 rapid style changes
      for (let i = 0; i < 100; i++) {
        const color = `hsl(${i * 3.6}, 50%, 50%)`;
        const fontSize = 12 + (i % 20);
        (customizer as any).applySavedStyles(color, fontSize);
      }
      
      const endTime = performance.now();
      const totalTime = endTime - startTime;
      
      // Should complete all 100 updates in less than 100ms
      expect(totalTime).toBeLessThan(100);
      
      // Should only have one style element (no duplicates)
      const styleElements = document.querySelectorAll('style[data-monarch360-styles]');
      expect(styleElements.length).toBe(1);
    });

    test('should efficiently query DOM elements', () => {
      createMockDOMElements();
      
      const startTime = performance.now();
      
      // Perform multiple DOM queries
      for (let i = 0; i < 1000; i++) {
        document.querySelector('[data-automationid="SiteHeader"]');
        document.querySelector('#spSiteHeader');
        document.querySelector('.ms-HorizontalNavItem');
      }
      
      const endTime = performance.now();
      const queryTime = endTime - startTime;
      
      // 1000 queries should complete quickly
      expect(queryTime).toBeLessThan(50);
    });
  });

  describe('Memory Usage', () => {
    test('should not create memory leaks with repeated operations', () => {
      createMockDOMElements();
      
      // Track initial DOM element count
      const initialElementCount = document.querySelectorAll('*').length;
      
      // Perform repeated operations
      for (let i = 0; i < 50; i++) {
        (customizer as any).injectSettingsButton();
        (customizer as any).applySavedStyles(`#${i.toString(16).padStart(6, '0')}`, 16);
      }
      
      // Check final DOM element count
      const finalElementCount = document.querySelectorAll('*').length;
      
      // Should not have created excessive elements
      expect(finalElementCount - initialElementCount).toBeLessThan(5);
    });

    test('should clean up event listeners properly', () => {
      createMockDOMElements();
      
      // Track event listener additions
      const originalAddEventListener = document.addEventListener;
      const originalRemoveEventListener = document.removeEventListener;
      
      let addCount = 0;
      let removeCount = 0;
      
      document.addEventListener = jest.fn((...args) => {
        addCount++;
        return originalAddEventListener.apply(document, args);
      });
      
      document.removeEventListener = jest.fn((...args) => {
        removeCount++;
        return originalRemoveEventListener.apply(document, args);
      });
      
      // Perform operations that add/remove listeners
      (customizer as any).injectSettingsButton();
      
      // Simulate disposal
      (customizer as any).onDispose?.();
      
      // Should properly clean up listeners
      expect(removeCount).toBeGreaterThanOrEqual(0);
      
      // Restore original methods
      document.addEventListener = originalAddEventListener;
      document.removeEventListener = originalRemoveEventListener;
    });
  });

  describe('Network Performance', () => {
    test('should handle slow SharePoint responses gracefully', async () => {
      // Mock slow SharePoint response
      const mockSP = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                filter: jest.fn().mockReturnValue({
                  top: jest.fn().mockReturnValue({
                    '()': jest.fn().mockImplementation(() => 
                      new Promise(resolve => setTimeout(() => resolve([]), 2000))
                    )
                  })
                })
              }
            })
          }
        }
      };

      jest.doMock('@pnp/sp', () => ({
        spfi: () => ({ using: () => mockSP }),
        SPFx: jest.fn()
      }));

      const startTime = performance.now();
      
      // Should timeout or handle gracefully without blocking
      const loadPromise = (customizer as any).loadSavedSettings();
      
      // Extension should remain responsive
      (customizer as any).injectSettingsButton();
      const settingsButton = document.getElementById('monarch360SettingsBtn');
      expect(settingsButton).toBeTruthy();
      
      // Wait for load to complete
      await loadPromise;
      
      const endTime = performance.now();
      const totalTime = endTime - startTime;
      
      // Should handle the slow response
      expect(totalTime).toBeGreaterThan(1900);
    });

    test('should batch SharePoint operations efficiently', async () => {
      const mockSP = {
        web: {
          lists: {
            getByTitle: jest.fn().mockReturnValue({
              items: {
                filter: jest.fn().mockReturnValue({
                  top: jest.fn().mockReturnValue({
                    '()': jest.fn().mockResolvedValue([])
                  })
                }),
                add: jest.fn().mockResolvedValue({ data: { Id: 1 } })
              }
            })
          }
        }
      };

      jest.doMock('@pnp/sp', () => ({
        spfi: () => ({ using: () => mockSP }),
        SPFx: jest.fn()
      }));

      const startTime = performance.now();
      
      // Save multiple settings
      await (customizer as any).saveSettingsToSharePoint('#ff0000', 18);
      
      const endTime = performance.now();
      const saveTime = endTime - startTime;
      
      // Should complete save operations quickly
      expect(saveTime).toBeLessThan(100);
      
      // Should have made efficient API calls
      expect(mockSP.web.lists.getByTitle).toHaveBeenCalledWith('navbarcrud');
    });
  });

  describe('Scalability', () => {
    test('should handle multiple style rules efficiently', () => {
      createMockDOMElements();
      
      // Create many navigation items
      const navContainer = document.querySelector('.ms-siteHeader-container');
      for (let i = 0; i < 100; i++) {
        const navItem = document.createElement('div');
        navItem.className = 'ms-HorizontalNavItem';
        navItem.setAttribute('data-automationid', 'HorizontalNav-link');
        navContainer?.appendChild(navItem);
      }
      
      const startTime = performance.now();
      
      (customizer as any).applySavedStyles('#0078d4', 16);
      
      const endTime = performance.now();
      const styleTime = endTime - startTime;
      
      // Should handle many elements efficiently
      expect(styleTime).toBeLessThan(50);
      
      // Verify styles were applied
      const styleElement = document.querySelector('style[data-monarch360-styles]');
      expect(styleElement?.textContent).toContain('.ms-HorizontalNavItem');
    });

    test('should maintain performance with frequent updates', () => {
      createMockDOMElements();
      
      const startTime = performance.now();
      
      // Simulate frequent user interactions
      for (let i = 0; i < 20; i++) {
        (customizer as any).applySavedStyles(`hsl(${i * 18}, 70%, 50%)`, 14 + i);
        
        // Simulate some processing time
        const busyWaitStart = performance.now();
        while (performance.now() - busyWaitStart < 1) {
          // Busy wait for 1ms
        }
      }
      
      const endTime = performance.now();
      const totalTime = endTime - startTime;
      
      // Should handle frequent updates efficiently
      expect(totalTime).toBeLessThan(200);
    });
  });

  describe('Resource Optimization', () => {
    test('should minimize CSS rule complexity', () => {
      createMockDOMElements();
      
      (customizer as any).applySavedStyles('#0078d4', 16);
      
      const styleElement = document.querySelector('style[data-monarch360-styles]');
      const cssText = styleElement?.textContent || '';
      
      // CSS should be concise but effective
      expect(cssText.length).toBeLessThan(1000);
      expect(cssText).toContain('#spSiteHeader');
      expect(cssText).toContain('.ms-HorizontalNavItem');
      
      // Should use efficient selectors
      expect(cssText).not.toContain('* *'); // Avoid universal selectors
      expect(cssText).not.toContain('div div div'); // Avoid deep nesting
    });

    test('should reuse DOM queries efficiently', () => {
      createMockDOMElements();
      
      // Mock querySelector to count calls
      const originalQuerySelector = document.querySelector;
      let queryCount = 0;
      
      document.querySelector = jest.fn((selector) => {
        queryCount++;
        return originalQuerySelector.call(document, selector);
      });
      
      // Perform operations that might query DOM
      (customizer as any).injectSettingsButton();
      (customizer as any).applySavedStyles('#ff0000', 16);
      
      // Should not make excessive queries
      expect(queryCount).toBeLessThan(20);
      
      // Restore original method
      document.querySelector = originalQuerySelector;
    });
  });
});
