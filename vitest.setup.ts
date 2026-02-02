// Vitest setup file
// Add global test utilities and mocks here

// Mock environment variables for tests
process.env.PUPIL_ENCRYPTION_KEY = '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef'
process.env.NEXTAUTH_SECRET = 'test-secret-key-for-testing-only'
process.env.NEXTAUTH_URL = 'http://localhost:3000'
// NODE_ENV is automatically set to 'test' by Vitest
