import { z } from 'zod'

/**
 * Environment variable validation schema
 * Validates all required configuration at startup
 */
const ConfigSchema = z.object({
  // Database
  DATABASE_URL: z.string().min(1, 'DATABASE_URL is required'),
  
  // Encryption
  PUPIL_ENCRYPTION_KEY: z
    .string()
    .length(64, 'PUPIL_ENCRYPTION_KEY must be exactly 64 characters')
    .regex(/^[0-9a-f]{64}$/, 'PUPIL_ENCRYPTION_KEY must be a valid hex string'),
  
  // NextAuth
  NEXTAUTH_SECRET: z
    .string()
    .min(32, 'NEXTAUTH_SECRET must be at least 32 characters'),
  NEXTAUTH_URL: z.string().url('NEXTAUTH_URL must be a valid URL'),
  
  // AI webhook
  AI_WEBHOOK_URL: z.string().url('AI_WEBHOOK_URL must be a valid URL'),

  // Site configuration
  SITE_NAME: z.string().optional().default('Comment Bank'),
  
  // Node environment
  NODE_ENV: z
    .enum(['development', 'production', 'test'])
    .optional()
    .default('development')
})

export type Config = z.infer<typeof ConfigSchema>

/**
 * Validate and parse environment variables
 * Throws an error if validation fails
 */
function validateConfig(): Config {
  try {
    return ConfigSchema.parse(process.env)
  } catch (error) {
    if (error instanceof z.ZodError) {
      const errorMessages = error.issues
        .map((err) => `  - ${err.path.join('.')}: ${err.message}`)
        .join('\n')
      
      throw new Error(
        `Environment variable validation failed:\n${errorMessages}\n\n` +
        'Please check your .env file and ensure all required variables are set correctly.'
      )
    }
    throw error
  }
}

/**
 * Validated configuration object
 * Access config values through this export
 */
export const config = validateConfig()
