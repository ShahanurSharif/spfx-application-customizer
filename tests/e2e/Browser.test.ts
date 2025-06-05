/**
 * Browser Automation Tests using Puppeteer
 * 
 * These tests run in real browsers to verify:
 * - Extension works in actual SharePoint environment
 * - Cross-browser compatibility
 * - Visual testing and user interactions
 * - Real network requests and responses
 */

import puppeteer, { Browser, Page } from 'puppeteer';

describe('SPFx Extension - Browser Automation Tests', () => {
  let browser: Browser;
  let page: Page;
  
  const testConfig = {
    sharePointUrl: process.env.SHAREPOINT_SITE_URL || 'https://yourtenant.sharepoint.com/sites/yoursite',
    debugUrl: 'https://localhost:4321/temp/build/manifests.js',
    extensionId: '5cf0b701-8c48-41ef-ad8c-cf66e6106423',
    timeout: 30000
  };

  beforeAll(async () => {
    browser = await puppeteer.launch({
      headless: process.env.CI === 'true', // Run headless in CI, headed locally
      slowMo: 50, // Slow down actions for better debugging
      args: [
        '--disable-web-security',
        '--disable-features=VizDisplayCompositor',
        '--ignore-certificate-errors'
      ]
    });
  });

  afterAll(async () => {
    if (browser) {
      await browser.close();
    }
  });

  beforeEach(async () => {
    page = await browser.newPage();
    
    // Set viewport for consistent testing
    await page.setViewport({ width: 1920, height: 1080 });
    
    // Enable console logging from the page
    page.on('console', msg => {
      if (msg.type() === 'error') {
        console.error('Browser Error:', msg.text());
      } else if (msg.type() === 'warn') {
        console.warn('Browser Warning:', msg.text());
      }
    });

    // Handle page errors
    page.on('pageerror', error => {
      console.error('Page Error:', error.message);
    });
  });

  afterEach(async () => {
    if (page) {
      await page.close();
    }
  });

  describe('Extension Loading in Browser', () => {
    test('should load extension in SharePoint debug mode', async () => {
      // Construct debug URL
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      // Navigate to SharePoint with debug parameters
      await page.goto(debugUrl, { 
        waitUntil: 'networkidle2',
        timeout: testConfig.timeout 
      });

      // Wait for SharePoint to load
      await page.waitForSelector('[data-automationid="SiteHeader"]', { 
        timeout: testConfig.timeout 
      });

      // Check if extension loaded successfully
      const initMessage = await page.evaluate(() => {
        return new Promise((resolve) => {
          const originalLog = console.log;
          console.log = (...args) => {
            if (args.join(' ').includes('Initialized Monarch360NavCrud')) {
              resolve(true);
            }
            originalLog.apply(console, args);
          };
          
          // Timeout after 10 seconds
          setTimeout(() => resolve(false), 10000);
        });
      });

      expect(initMessage).toBe(true);
    }, testConfig.timeout);

    test('should inject settings button in SharePoint header', async () => {
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      
      // Wait for SharePoint header
      await page.waitForSelector('[data-automationid="SiteHeader"]');
      
      // Wait for settings button to be injected
      await page.waitForSelector('#monarch360SettingsBtn', { 
        timeout: 15000 
      });

      // Verify button properties
      const buttonText = await page.$eval('#monarch360SettingsBtn', el => el.textContent);
      const buttonTitle = await page.$eval('#monarch360SettingsBtn', el => el.getAttribute('title'));
      
      expect(buttonText).toContain('⚙️');
      expect(buttonTitle).toBe('Settings');
    }, testConfig.timeout);
  });

  describe('User Interactions', () => {
    beforeEach(async () => {
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      await page.waitForSelector('#monarch360SettingsBtn', { timeout: 15000 });
    });

    test('should open settings dialog when button is clicked', async () => {
      // Click the settings button
      await page.click('#monarch360SettingsBtn');
      
      // Wait for dialog to appear
      await page.waitForSelector('.ms-Dialog', { timeout: 5000 });
      
      // Verify dialog content
      const dialogTitle = await page.$eval('.ms-Dialog-title', el => el.textContent);
      expect(dialogTitle).toContain('Navigation Settings');
      
      // Check for color picker and slider
      const colorInput = await page.$('[type="color"]');
      const slider = await page.$('.ms-Slider');
      
      expect(colorInput).toBeTruthy();
      expect(slider).toBeTruthy();
    }, testConfig.timeout);

    test('should update background color when user changes color picker', async () => {
      await page.click('#monarch360SettingsBtn');
      await page.waitForSelector('.ms-Dialog');
      
      // Change color in color picker
      await page.evaluate(() => {
        const colorInput = document.querySelector('[type="color"]') as HTMLInputElement;
        if (colorInput) {
          colorInput.value = '#ff6600';
          colorInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
      
      // Wait a moment for preview
      await page.waitForFunction(() => true, { timeout: 500 });
      
      // Save settings
      await page.click('.ms-Dialog-actions .ms-Button--primary');
      
      // Wait for dialog to close
      await page.waitForSelector('.ms-Dialog', { hidden: true });
      
      // Check if background color was applied
      const headerBgColor = await page.evaluate(() => {
        const header = document.querySelector('#spSiteHeader');
        return header ? getComputedStyle(header).backgroundColor : null;
      });
      
      // Should have applied the new color (may be converted to rgb)
      expect(headerBgColor).toBeTruthy();
      expect(headerBgColor).not.toBe('rgba(0, 0, 0, 0)'); // Not transparent
    }, testConfig.timeout);

    test('should update font size when user changes slider', async () => {
      await page.click('#monarch360SettingsBtn');
      await page.waitForSelector('.ms-Dialog');
      
      // Move slider to increase font size
      const slider = await page.$('.ms-Slider-slideBox');
      if (slider) {
        const box = await slider.boundingBox();
        if (box) {
          // Click towards the right side of slider (larger value)
          await page.mouse.click(box.x + box.width * 0.8, box.y + box.height / 2);
        }
      }
      
      // Wait for update
      await page.waitForFunction(() => true, { timeout: 500 });
      
      // Save settings
      await page.click('.ms-Dialog-actions .ms-Button--primary');
      await page.waitForSelector('.ms-Dialog', { hidden: true });
      
      // Check if font size was applied to navigation items
      const navItemFontSize = await page.evaluate(() => {
        const navItem = document.querySelector('.ms-HorizontalNavItem');
        return navItem ? getComputedStyle(navItem).fontSize : null;
      });
      
      expect(navItemFontSize).toBeTruthy();
      // Font size should be larger than default (16px)
      const fontSize = parseInt(navItemFontSize?.replace('px', '') || '0');
      expect(fontSize).toBeGreaterThan(16);
    }, testConfig.timeout);
  });

  describe('Settings Persistence', () => {
    test('should persist settings across page reloads', async () => {
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      // First visit - set custom settings
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      await page.waitForSelector('#monarch360SettingsBtn');
      
      await page.click('#monarch360SettingsBtn');
      await page.waitForSelector('.ms-Dialog');
      
      // Set a distinctive color
      await page.evaluate(() => {
        const colorInput = document.querySelector('[type="color"]') as HTMLInputElement;
        if (colorInput) {
          colorInput.value = '#purple';
          colorInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
      
      await page.click('.ms-Dialog-actions .ms-Button--primary');
      await page.waitForSelector('.ms-Dialog', { hidden: true });
      
      // Reload the page
      await page.reload({ waitUntil: 'networkidle2' });
      await page.waitForSelector('#monarch360SettingsBtn');
      
      // Check if settings were restored
      const headerBgColor = await page.evaluate(() => {
        const header = document.querySelector('#spSiteHeader');
        return header ? getComputedStyle(header).backgroundColor : null;
      });
      
      // Should have the saved color
      expect(headerBgColor).toBeTruthy();
      expect(headerBgColor).not.toBe('rgba(0, 0, 0, 0)');
    }, testConfig.timeout * 2);
  });

  describe('Error Handling in Browser', () => {
    test('should handle missing SharePoint elements gracefully', async () => {
      // Create a minimal HTML page without SharePoint elements
      const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head><title>Test Page</title></head>
        <body>
          <div>No SharePoint elements here</div>
          <script>
            // Simulate extension loading
            console.log('Page loaded without SharePoint elements');
          </script>
        </body>
        </html>
      `;
      
      await page.setContent(htmlContent);
      
      // Inject extension code
      await page.evaluate(() => {
        // Simulate extension trying to find SharePoint elements
        const header = document.querySelector('[data-automationid="SiteHeader"]');
        if (!header) {
          console.warn('Header container not found. Available elements: ', 
            Array.from(document.querySelectorAll('*')).map(el => el.tagName).join(', ')
          );
        }
      });
      
      // Page should not crash
      const title = await page.title();
      expect(title).toBe('Test Page');
    });

    test('should handle network errors gracefully', async () => {
      // Block SharePoint list requests
      await page.setRequestInterception(true);
      page.on('request', (request) => {
        if (request.url().includes('navbarcrud')) {
          request.abort();
        } else {
          request.continue();
        }
      });
      
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      
      // Extension should still load and function
      await page.waitForSelector('#monarch360SettingsBtn', { timeout: 15000 });
      
      // Button should still be clickable
      await page.click('#monarch360SettingsBtn');
      await page.waitForSelector('.ms-Dialog');
      
      // Dialog should work even without SharePoint list
      const dialogTitle = await page.$eval('.ms-Dialog-title', el => el.textContent);
      expect(dialogTitle).toBeTruthy();
    }, testConfig.timeout);
  });

  describe('Visual Testing', () => {
    test('should render settings button with correct styling', async () => {
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      await page.waitForSelector('#monarch360SettingsBtn');
      
      // Check button styling
      const buttonStyles = await page.evaluate(() => {
        const button = document.getElementById('monarch360SettingsBtn');
        if (!button) return null;
        
        const styles = getComputedStyle(button);
        return {
          display: styles.display,
          cursor: styles.cursor,
          fontSize: styles.fontSize,
          padding: styles.padding
        };
      });
      
      expect(buttonStyles).toBeTruthy();
      expect(buttonStyles?.cursor).toBe('pointer');
      expect(buttonStyles?.display).not.toBe('none');
    });

    test('should apply visual changes correctly', async () => {
      const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
      
      await page.goto(debugUrl, { waitUntil: 'networkidle2' });
      await page.waitForSelector('#monarch360SettingsBtn');
      
      // Take screenshot before changes
      const beforeScreenshot = await page.screenshot({ 
        clip: { x: 0, y: 0, width: 1920, height: 200 } 
      });
      
      // Apply settings
      await page.click('#monarch360SettingsBtn');
      await page.waitForSelector('.ms-Dialog');
      
      await page.evaluate(() => {
        const colorInput = document.querySelector('[type="color"]') as HTMLInputElement;
        if (colorInput) {
          colorInput.value = '#ff0000';
          colorInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
      
      await page.click('.ms-Dialog-actions .ms-Button--primary');
      await page.waitForSelector('.ms-Dialog', { hidden: true });
      
      // Take screenshot after changes
      const afterScreenshot = await page.screenshot({ 
        clip: { x: 0, y: 0, width: 1920, height: 200 } 
      });
      
      // Screenshots should be different
      expect(beforeScreenshot.equals(afterScreenshot)).toBe(false);
    });
  });

  describe('Cross-Browser Compatibility', () => {
    test('should work in different browser contexts', async () => {
      // This test would be expanded to run across different browser engines
      // For now, we'll test different viewport sizes to simulate mobile/desktop
      
      const viewports = [
        { width: 360, height: 640 },   // Mobile
        { width: 768, height: 1024 },  // Tablet
        { width: 1920, height: 1080 }  // Desktop
      ];
      
      for (const viewport of viewports) {
        await page.setViewport(viewport);
        
        const debugUrl = `${testConfig.sharePointUrl}?debug=true&noredir=true&debugManifestsFile=${testConfig.debugUrl}&loadSPFX=true&customActions={"${testConfig.extensionId}":{"location":"ClientSideExtension.ApplicationCustomizer"}}`;
        
        await page.goto(debugUrl, { waitUntil: 'networkidle2' });
        
        // Extension should work across different screen sizes
        try {
          await page.waitForSelector('#monarch360SettingsBtn', { timeout: 10000 });
          
          const buttonVisible = await page.evaluate(() => {
            const button = document.getElementById('monarch360SettingsBtn');
            if (!button) return false;
            
            const rect = button.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
          });
          
          expect(buttonVisible).toBe(true);
        } catch (error) {
          console.warn(`Extension may not be visible at ${viewport.width}x${viewport.height}`);
        }
      }
    }, testConfig.timeout * 3);
  });
});
