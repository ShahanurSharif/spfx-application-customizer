/**
 * Mock for @pnp/sp library
 */

// Create a proper Jest mock function
const mockSpfi = jest.fn();

export const spfi = mockSpfi;

// Set default mock implementation
mockSpfi.mockReturnValue({
  using: jest.fn().mockReturnValue({
    web: {
      lists: {
        getByTitle: jest.fn().mockReturnValue({
          items: {
            filter: jest.fn().mockReturnValue({
              top: jest.fn().mockReturnValue(
                jest.fn().mockResolvedValue([]) // The final () call
              )
            }),
            add: jest.fn().mockResolvedValue({ data: { Id: 1 } }),
            getById: jest.fn().mockReturnValue({
              update: jest.fn().mockResolvedValue({})
            })
          }
        })
      }
    }
  })
});

export const SPFx = jest.fn();

// Mock the side-effect imports
jest.mock('@pnp/sp/webs', () => ({}));
jest.mock('@pnp/sp/lists', () => ({}));
jest.mock('@pnp/sp/items', () => ({}));
