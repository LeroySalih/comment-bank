/**
 * Centralized logging utility
 * Provides structured logging with context
 */

type LogLevel = 'error' | 'warn' | 'info' | 'debug'

interface LogContext {
  [key: string]: any
}

class Logger {
  private log(level: LogLevel, message: string, context?: LogContext) {
    const timestamp = new Date().toISOString()
    const logData = {
      timestamp,
      level,
      message,
      ...context
    }

    // In development, use console with formatting
    if (process.env.NODE_ENV === 'development') {
      const prefix = `[${timestamp}] [${level.toUpperCase()}]`
      
      switch (level) {
        case 'error':
          console.error(prefix, message, context || '')
          break
        case 'warn':
          console.warn(prefix, message, context || '')
          break
        case 'info':
          console.info(prefix, message, context || '')
          break
        case 'debug':
          console.debug(prefix, message, context || '')
          break
      }
    } else {
      // In production, output JSON for log aggregation
      console.log(JSON.stringify(logData))
    }
  }

  error(message: string, context?: LogContext) {
    this.log('error', message, context)
  }

  warn(message: string, context?: LogContext) {
    this.log('warn', message, context)
  }

  info(message: string, context?: LogContext) {
    this.log('info', message, context)
  }

  debug(message: string, context?: LogContext) {
    this.log('debug', message, context)
  }
}

export const logger = new Logger()
