import { test, expect } from '@playwright/test';

test('admin can sign in and access admin dashboard', async ({ page }) => {
  await page.goto('/login');

  // Fill in credentials
  await page.fill('input[name="username"]', 'admin');
  await page.fill('input[name="password"]', 'password');

  // Click sign in
  await page.click('button[type="submit"]');

  // Expect to be redirected to admin dashboard (or home page then admin)
  // Assuming login redirects to home, we verify we can go to /admin or are redirected there if role based redirect exists.
  // The login page code showed `router.push('/')` on success.
  
  // Wait for navigation and verify logged in state
  // Instead of strict URL check, wait for a UI element that confirms we are logged in.
  // This is more robust against fast/slow redirects.
  await expect(page.getByRole('button', { name: 'Sign Out' })).toBeVisible();
  // OR just check we are not on login
  await expect(page).not.toHaveURL(/\/login/);
  
  // Navigate to admin
  await page.goto('/admin');
  
  // Verify Admin Dashboard title
  await expect(page.getByRole('heading', { name: 'Admin Dashboard' })).toBeVisible();

  // Sign out
  await page.click('button:has-text("Sign Out")');

  // Verify redirected to login
  await expect(page).toHaveURL(/\/login/);
});

test('HOD cannot access admin dashboard', async ({ page }) => {
  await page.goto('/login');

  // Fill in credentials for HOD
  await page.fill('input[name="username"]', 'leroysalih');
  await page.fill('input[name="password"]', 'password');

  // Click sign in
  await page.click('button[type="submit"]');

  // Expect to be redirected to home (or at least NOT login)
  await expect(page).not.toHaveURL(/\/login/);
  
  // Try to navigate to admin
  await page.goto('/admin');
  
  // Verify redirected to login (due to middleware protection)
  // Or check for "Sign in to Comment Bank" text or similar unique login page identifier
  await expect(page).toHaveURL(/\/login/);
});

test('HOD can access HOD dashboard', async ({ page }) => {
  await page.goto('/login');

  // Fill in credentials for HOD
  await page.fill('input[name="username"]', 'leroysalih');
  await page.fill('input[name="password"]', 'password');

  // Click sign in
  await page.click('button[type="submit"]');

  // Expect to be redirected to home (or at least NOT login)
  await expect(page).not.toHaveURL(/\/login/);
  
  // Navigate to HOD dashboard
  await page.goto('/hod');
  
  // Verify access granted by checking for the dashboard title
  await expect(page.getByRole('heading', { name: 'Head of Department Dashboard' })).toBeVisible();
});

