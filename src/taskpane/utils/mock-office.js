// Mock implementation of the Word API for development
export const createMockWordAPI = () => {
    return {
      run: async (callback) => {
        const context = {
          document: {
            body: {
              insertParagraph: (text, location) => {
                console.log(`Inserting paragraph: ${text} at ${location}`)
                return {
                  font: {
                    set: (options) => {
                      console.log(`Setting font options:`, options)
                    },
                  },
                  leftIndent: 0,
                  styleBuiltIn: null,
                }
              },
              insertText: (text, location) => {
                console.log(`Inserting text: ${text} at ${location}`)
              },
            },
            properties: {
              customProperties: {
                items: [],
                load: () => {},
                add: (key, value) => {
                  console.log(`Adding custom property: ${key} = ${value}`)
                  // Store in localStorage for development
                  localStorage.setItem(key, value)
                },
                delete: (key) => {
                  console.log(`Deleting custom property: ${key}`)
                  localStorage.removeItem(key)
                },
                getItem: (key) => {
                  return localStorage.getItem(key)
                },
              },
            },
          },
          sync: async () => {},
        }
  
        await callback(context)
        return context
      },
      Style: {
        heading1: "Heading1",
        heading2: "Heading2",
      },
    }
  }
  
  // Export a mock Word object if the real one isn't available
  const mockWord = createMockWordAPI()
  
  // Check if we're in a browser environment
  const isBrowser = typeof window !== "undefined"
  
  // Use the real Word object if available, otherwise use the mock
  export const Word = isBrowser && window.Word ? window.Word : mockWord