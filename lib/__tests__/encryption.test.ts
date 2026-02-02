import { describe, it, expect, beforeEach } from 'vitest'
import { encrypt, decrypt, isEncrypted } from '../encryption'

describe('Encryption Functions', () => {
  const testData = 'John Doe'
  
  describe('encrypt', () => {
    it('should encrypt a string', () => {
      const encrypted = encrypt(testData)
      expect(encrypted).not.toBe(testData)
      expect(encrypted).toContain(':') // Should have format iv:authTag:encrypted
    })

    it('should produce different ciphertexts for the same input', () => {
      const encrypted1 = encrypt(testData)
      const encrypted2 = encrypt(testData)
      expect(encrypted1).not.toBe(encrypted2) // Different IVs
    })

    it('should throw error if encryption key is invalid', () => {
      const originalKey = process.env.PUPIL_ENCRYPTION_KEY
      process.env.PUPIL_ENCRYPTION_KEY = 'invalid'
      
      expect(() => encrypt(testData)).toThrow()
      
      process.env.PUPIL_ENCRYPTION_KEY = originalKey
    })
  })

  describe('decrypt', () => {
    it('should decrypt an encrypted string', () => {
      const encrypted = encrypt(testData)
      const decrypted = decrypt(encrypted)
      expect(decrypted).toBe(testData)
    })

    it('should return unencrypted text as-is (for migration support)', () => {
      const plaintext = 'Not Encrypted'
      const result = decrypt(plaintext)
      expect(result).toBe(plaintext)
    })

    it('should throw error for invalid encrypted data', () => {
      const invalidEncrypted = 'abc:def:ghi' // Valid format but invalid data
      expect(() => decrypt(invalidEncrypted)).toThrow()
    })

    it('should throw error if encryption key is invalid', () => {
      const encrypted = encrypt(testData)
      const originalKey = process.env.PUPIL_ENCRYPTION_KEY
      process.env.PUPIL_ENCRYPTION_KEY = 'invalid'
      
      expect(() => decrypt(encrypted)).toThrow()
      
      process.env.PUPIL_ENCRYPTION_KEY = originalKey
    })
  })

  describe('isEncrypted', () => {
    it('should return true for encrypted data', () => {
      const encrypted = encrypt(testData)
      expect(isEncrypted(encrypted)).toBe(true)
    })

    it('should return false for plaintext', () => {
      expect(isEncrypted('John Doe')).toBe(false)
      expect(isEncrypted('test')).toBe(false)
    })

    it('should return false for invalid format', () => {
      expect(isEncrypted('abc:def')).toBe(false) // Only 2 parts
      expect(isEncrypted('abc:def:ghi:jkl')).toBe(false) // Too many parts
      expect(isEncrypted('abc::def')).toBe(false) // Empty part
    })
  })

  describe('encrypt/decrypt roundtrip', () => {
    const testCases = [
      'Simple Name',
      'Name with Spaces',
      'Ñame with Spëcial Çharacters',
      '名前', // Japanese
      'اسم', // Arabic
      'A'.repeat(100), // Long string
    ]

    testCases.forEach(testCase => {
      it(`should correctly encrypt and decrypt: "${testCase.substring(0, 20)}..."`, () => {
        const encrypted = encrypt(testCase)
        const decrypted = decrypt(encrypted)
        expect(decrypted).toBe(testCase)
      })
    })
  })
})
