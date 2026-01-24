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
 * Decrypts a string using AES-256-GCM
 */
export function decrypt(encryptedText: string): string {
  const key = process.env.PUPIL_ENCRYPTION_KEY
  if (!key || key.length !== 64) {
    throw new Error('PUPIL_ENCRYPTION_KEY must be a 64-character (32-byte) hex string')
  }

  try {
    const [ivHex, authTagHex, encryptedData] = encryptedText.split(':')
    if (!ivHex || !authTagHex || !encryptedData) {
      return encryptedText // Not encrypted or invalid format
    }

    const iv = Buffer.from(ivHex, 'hex')
    const authTag = Buffer.from(authTagHex, 'hex')
    const decipher = crypto.createDecipheriv(ALGORITHM, Buffer.from(key, 'hex'), iv)
    
    decipher.setAuthTag(authTag)
    
    let decrypted = decipher.update(encryptedData, 'hex', 'utf8')
    decrypted += decipher.final('utf8')
    
    return decrypted
  } catch (error) {
    console.error('Decryption failed:', error)
    return encryptedText // Fallback to raw text if decryption fails
  }
}
