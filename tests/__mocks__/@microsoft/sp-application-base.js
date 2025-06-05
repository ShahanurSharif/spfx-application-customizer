/**
 * Mock for @microsoft/sp-application-base
 */

module.exports = {
  BaseApplicationCustomizer: class MockBaseApplicationCustomizer {
    constructor() {
      this.context = null;
      this.properties = null;
    }

    onInit() {
      return Promise.resolve();
    }

    onDispose() {
      // Mock dispose
    }
  }
};
