/**
 * Custom error classes for standardized error handling
 */

export class AppError extends Error {
  constructor(
    message: string,
    public code: string,
    public statusCode: number = 500,
    public details?: any
  ) {
    super(message)
    this.name = 'AppError'
    Error.captureStackTrace(this, this.constructor)
  }
}

export class ValidationError extends AppError {
  constructor(message: string, details?: any) {
    super(message, 'VALIDATION_ERROR', 400, details)
    this.name = 'ValidationError'
  }
}

export class AuthError extends AppError {
  constructor(message: string) {
    super(message, 'AUTH_ERROR', 401)
    this.name = 'AuthError'
  }
}

export class ForbiddenError extends AppError {
  constructor(message: string) {
    super(message, 'FORBIDDEN', 403)
    this.name = 'ForbiddenError'
  }
}

export class NotFoundError extends AppError {
  constructor(message: string) {
    super(message, 'NOT_FOUND', 404)
    this.name = 'NotFoundError'
  }
}

/**
 * Handle errors from server actions and return standardized response
 */
export function handleServerActionError(error: unknown): {
  success: false
  error: string
  code: string
  details?: any
} {
  if (error instanceof AppError) {
    return {
      success: false,
      error: error.message,
      code: error.code,
      details: error.details
    }
  }

  if (error instanceof Error) {
    return {
      success: false,
      error: error.message,
      code: 'INTERNAL_ERROR'
    }
  }

  return {
    success: false,
    error: 'An unexpected error occurred',
    code: 'UNKNOWN_ERROR'
  }
}
