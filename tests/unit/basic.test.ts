/// <reference path="../types/global.d.ts" />

describe('Basic Type Test', () => {
  test('should have global fetch mock', () => {
    // Arrange
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ test: true })
    } as Response);

    // Assert
    expect(global.fetch).toBeDefined();
    expect(typeof global.fetch).toBe('function');
  });

  test('should have global performance mock', () => {
    // Arrange
    global.performance = {
      mark: jest.fn(),
      measure: jest.fn(),
      getEntriesByName: jest.fn().mockReturnValue([]),
      clearMarks: jest.fn(),
      clearMeasures: jest.fn(),
      now: jest.fn().mockReturnValue(123.456)
    } as any;

    // Assert
    expect(global.performance).toBeDefined();
    expect(typeof global.performance.now).toBe('function');
  });
});
