import crypto from 'node:crypto'

const ALGORITHM = 'aes-256-gcm'
const IV_LENGTH = 12
const AUTH_TAG_LENGTH = 16

/**
 * Encrypts a string using AES-256-GCM
 */
export function encrypt(text: string): string {
  const key = process.env.PUPIL_ENCRYPTION_KEY
  if (!key || key.length !== 64) {
    throw new Error('PUPIL_ENCRYPTION_KEY must be a 64-character (32-byte) hex string')
  }

  const iv = crypto.randomBytes(IV_LENGTH)
  const cipher = crypto.createCipheriv(ALGORITHM, Buffer.from(key, 'hex'), iv)
  
  let encrypted = cipher.update(text, 'utf8', 'hex')
  encrypted += cipher.final('hex')
  
  const authTag = cipher.getAuthTag().toString('hex')
  
  // Format: iv:authTag:encrypted
  return `${iv.toString('hex')}:${authTag}:${encrypted}`
}

/**
 * Check if a string is encrypted (has the expected format)
 */
export function isEncrypted(text: string | null | undefined): boolean {
  if (!text) return false
  const parts = text.split(':')
  return parts.length === 3 && parts.every(part => part.length > 0)
}

/**
 * Decrypts a string using AES-256-GCM
 * Throws an error if decryption fails
 */
export function decrypt(encryptedText: string | null | undefined): string {
  if (!encryptedText) return ''
  const key = process.env.PUPIL_ENCRYPTION_KEY
  if (!key || key.length !== 64) {
    throw new Error('PUPIL_ENCRYPTION_KEY must be a 64-character (32-byte) hex string')
  }

  // If the text doesn't look encrypted, it might be legacy unencrypted data
  if (!isEncrypted(encryptedText)) {
    // For now, return as-is to support migration
    // TODO: Remove this once all data is encrypted
    return encryptedText
  }

  try {
    const [ivHex, authTagHex, encryptedData] = encryptedText.split(':')

    const iv = Buffer.from(ivHex, 'hex')
    const authTag = Buffer.from(authTagHex, 'hex')
    const decipher = crypto.createDecipheriv(ALGORITHM, Buffer.from(key, 'hex'), iv)
    
    decipher.setAuthTag(authTag)
    
    let decrypted = decipher.update(encryptedData, 'hex', 'utf8')
    decrypted += decipher.final('utf8')
    
    return decrypted
  } catch (error) {
    // Log the error but don't expose sensitive details
    const errorMessage = error instanceof Error ? error.message : 'Unknown error'
    throw new Error(`Decryption failed: ${errorMessage}`)
  }
}
