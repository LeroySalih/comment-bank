"use client"

export interface PupilPII {
  admissionNumber: string
  firstName: string
  lastName: string
  gender: string
  form?: string
}

const STORAGE_KEY = "comment-bank-pii-v1"

export function getStoredPII(): Record<string, PupilPII> {
  if (typeof window === "undefined") return {}
  const stored = localStorage.getItem(STORAGE_KEY)
  if (!stored) return {}
  try {
    return JSON.parse(stored)
  } catch (e) {
    console.error("Failed to parse stored PII", e)
    return {}
  }
}

export function savePII(piiList: PupilPII[]) {
  const piiMap = getStoredPII()
  piiList.forEach(p => {
    piiMap[p.admissionNumber] = p
  })
  localStorage.setItem(STORAGE_KEY, JSON.stringify(piiMap))
  // Dispatch custom event to notify listeners
  window.dispatchEvent(new Event("pii-store-updated"))
}

export function clearPII() {
  localStorage.removeItem(STORAGE_KEY)
  window.dispatchEvent(new Event("pii-store-updated"))
}
