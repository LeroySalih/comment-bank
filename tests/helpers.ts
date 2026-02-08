import { Page, expect } from '@playwright/test';

/**
 * Login helper — navigates to /login, fills credentials, and waits for redirect.
 */
export async function login(page: Page, username: string, password: string): Promise<void> {
  await page.goto('/login');
  await page.fill('input[name="username"]', username);
  await page.fill('input[name="password"]', password);
  await page.click('button[type="submit"]');

  // Wait for successful login (redirected away from login page)
  await expect(page).not.toHaveURL(/\/login/);
}

/**
 * Test user credentials matching prisma/seed.ts
 */
export const TEST_USERS = {
  admin: { username: 'admin', password: 'password' },
  hod: { username: 'leroysalih', password: 'password' },
  teacher: { username: 'teacher', password: 'password' },
  teacher2: { username: 'teacher2', password: 'password' },
  teacher3: { username: 'teacher3', password: 'password' },
} as const;

/**
 * Role-specific login helpers
 */
export async function loginAsAdmin(page: Page): Promise<void> {
  await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
}

export async function loginAsHoD(page: Page): Promise<void> {
  await login(page, TEST_USERS.hod.username, TEST_USERS.hod.password);
}

export async function loginAsTeacher(page: Page): Promise<void> {
  await login(page, TEST_USERS.teacher.username, TEST_USERS.teacher.password);
}

/**
 * Expects page to redirect to login
 */
export async function expectRedirectToLogin(page: Page): Promise<void> {
  await expect(page).toHaveURL(/\/login/);
}
