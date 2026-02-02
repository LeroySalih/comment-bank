/**
 * Standard server action response types
 */

export type ServerActionSuccess<T = void> = {
  success: true
} & (T extends void ? {} : { data: T })

export type ServerActionError = {
  success: false
  error: string
  code: string
  details?: any
}

export type ServerActionResult<T = void> = ServerActionSuccess<T> | ServerActionError

/**
 * Type guard to check if a server action result is an error
 */
export function isServerActionError(result: ServerActionResult<any>): result is ServerActionError {
  return !result.success
}

/**
 * Type guard to check if a server action result is successful
 */
export function isServerActionSuccess<T>(result: ServerActionResult<T>): result is ServerActionSuccess<T> {
  return result.success
}
