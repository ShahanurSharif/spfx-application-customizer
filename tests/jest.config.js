/**
 * Jest Configuration for SPFx Extension E2E Tests
 */
module.exports = {
  preset: 'ts-jest/presets/default',
  testEnvironment: 'jsdom',
  rootDir: '../',
  
  // Test files pattern
  testMatch: [
    '<rootDir>/tests/**/*.test.ts',
    '<rootDir>/tests/**/*.test.js'
  ],

  // Module resolution
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json'],
  
  // Transform TypeScript files with updated ts-jest configuration
  transform: {
    '^.+\\.(ts|tsx)$': ['ts-jest', {
      tsconfig: 'tests/tsconfig.json'
    }]
  },

  // Transform ignore patterns for ES modules
  transformIgnorePatterns: [
    'node_modules/(?!(@pnp|@microsoft))'
  ],

  // Setup files
  setupFilesAfterEnv: ['<rootDir>/tests/setup.ts'],

  // Module name mapping for SPFx modules (corrected property name)
  moduleNameMapper: {
    '^@microsoft/sp-(.*)$': '<rootDir>/node_modules/@microsoft/sp-$1',
    '^@pnp/sp$': '<rootDir>/tests/__mocks__/@pnp/sp.ts',
    '^@pnp/sp/(.*)$': '<rootDir>/tests/__mocks__/@pnp/sp.ts',
    '^@fluentui/(.*)$': '<rootDir>/node_modules/@fluentui/$1',
    '^Monarch360NavCrudApplicationCustomizerStrings$': '<rootDir>/tests/__mocks__/Monarch360NavCrudApplicationCustomizerStrings.ts'
  },

  // Coverage configuration
  collectCoverageFrom: [
    'src/**/*.{ts,tsx}',
    '!src/**/*.d.ts',
    '!src/**/index.ts'
  ],

  // Coverage thresholds
  coverageThreshold: {
    global: {
      branches: 70,
      functions: 70,
      lines: 70,
      statements: 70
    }
  },

  // Coverage output
  coverageReporters: ['text', 'lcov', 'html'],
  coverageDirectory: 'tests/coverage',

  // Timeout for tests (increased for E2E)
  testTimeout: 30000,

  // Verbose output
  verbose: true
};
