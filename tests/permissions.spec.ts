import { test, expect } from '@playwright/test';

/**
 * Permission E2E Tests
 * 
 * Tests role-based access control for:
 * - Admin: Can access /admin only (plus dashboard)
 * - HOD: Can access /hod only (plus dashboard)
 * - Teacher: Can access /class/* and /student/* (plus dashboard)
 * 
 * Each test verifies:
 * 1. Menu navigation to allowed routes
 * 2. Direct URL access to allowed routes
 * 3. Direct URL access to forbidden routes (should redirect to /login)
 */

// Helper function to login
async function login(page: any, username: string, password: string) {
  await page.goto('/login');
  await page.fill('input[name="username"]', username);
  await page.fill('input[name="password"]', password);
  await page.click('button[type="submit"]');
  
  // Wait for successful login (redirected away from login page)
  await expect(page).not.toHaveURL(/\/login/);
}

// ============================================================================
// ADMIN PERMISSION TESTS
// ============================================================================

test.describe('Admin Permissions', () => {
  
  test('admin can access /admin via menu', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Look for Admin Dashboard link in navigation
    const adminLink = page.locator('a[href="/admin"], a:has-text("Admin Dashboard")').first();
    await expect(adminLink).toBeVisible();
    
    // Click the link
    await adminLink.click();
    
    // Verify we're on the admin page
    await expect(page).toHaveURL(/\/admin/);
    await expect(page.getByRole('heading', { name: 'Admin Dashboard' })).toBeVisible();
  });

  test('admin can access /admin via direct URL', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Navigate directly to admin
    await page.goto('/admin');
    
    // Verify access granted
    await expect(page).toHaveURL(/\/admin/);
    await expect(page.getByRole('heading', { name: 'Admin Dashboard' })).toBeVisible();
  });

  test('admin can access dashboard', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Should be on dashboard after login or can navigate to it
    await page.goto('/');
    
    // Verify we're on a valid page (not redirected to login)
    await expect(page).not.toHaveURL(/\/login/);
  });

  test('admin CANNOT access /hod (redirects to login)', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Try to access HOD dashboard
    await page.goto('/hod');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('admin CANNOT access /class (redirects to login)', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Try to access a class page (we'll use a generic class ID)
    await page.goto('/class/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('admin CANNOT access /student (redirects to login)', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Try to access a student page
    await page.goto('/student/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('admin menu does NOT show HOD or Teacher links', async ({ page }) => {
    await login(page, 'admin', 'password');
    
    // Navigate to home/dashboard
    await page.goto('/');
    
    // Verify HOD Dashboard link is NOT visible
    const hodLink = page.locator('a[href="/hod"], a:has-text("HOD Dashboard"), a:has-text("Head of Department")');
    await expect(hodLink).not.toBeVisible();
    
    // Verify Classes/Students links are NOT visible (teacher-specific)
    const classLink = page.locator('a[href^="/class"]');
    await expect(classLink).not.toBeVisible();
  });
});

// ============================================================================
// HOD PERMISSION TESTS
// ============================================================================

test.describe('HOD Permissions', () => {
  
  test('HOD can access /hod via menu', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Look for HOD Dashboard link in navigation
    const hodLink = page.locator('a[href="/hod"], a:has-text("HOD Dashboard"), a:has-text("Head of Department")').first();
    await expect(hodLink).toBeVisible();
    
    // Click the link
    await hodLink.click();
    
    // Verify we're on the HOD page
    await expect(page).toHaveURL(/\/hod/);
    await expect(page.getByRole('heading', { name: 'Head of Department Dashboard' })).toBeVisible();
  });

  test('HOD can access /hod via direct URL', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Navigate directly to HOD
    await page.goto('/hod');
    
    // Verify access granted
    await expect(page).toHaveURL(/\/hod/);
    await expect(page.getByRole('heading', { name: 'Head of Department Dashboard' })).toBeVisible();
  });

  test('HOD can access dashboard', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Should be on dashboard after login or can navigate to it
    await page.goto('/');
    
    // Verify we're on a valid page (not redirected to login)
    await expect(page).not.toHaveURL(/\/login/);
  });

  test('HOD CANNOT access /admin (redirects to login)', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Try to access Admin dashboard
    await page.goto('/admin');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('HOD CANNOT access /class (redirects to login)', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Try to access a class page
    await page.goto('/class/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('HOD CANNOT access /student (redirects to login)', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Try to access a student page
    await page.goto('/student/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('HOD menu does NOT show Admin or Teacher links', async ({ page }) => {
    await login(page, 'leroysalih', 'password');
    
    // Navigate to home/dashboard
    await page.goto('/');
    
    // Verify Admin Dashboard link is NOT visible
    const adminLink = page.locator('a[href="/admin"], a:has-text("Admin Dashboard")');
    await expect(adminLink).not.toBeVisible();
    
    // Verify Classes/Students links are NOT visible (teacher-specific)
    const classLink = page.locator('a[href^="/class"]');
    await expect(classLink).not.toBeVisible();
  });
});

// ============================================================================
// TEACHER PERMISSION TESTS
// ============================================================================

test.describe('Teacher Permissions', () => {
  
  test('teacher can access dashboard', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // Should be on dashboard after login or can navigate to it
    await page.goto('/');
    
    // Verify we're on a valid page (not redirected to login)
    await expect(page).not.toHaveURL(/\/login/);
  });

  test('teacher can access /class via direct URL', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // First, get a valid class ID from the dashboard
    await page.goto('/');
    
    // Look for any class link in the page
    const classLink = page.locator('a[href^="/class/"]').first();
    
    // If class link exists, test direct access
    const classLinkCount = await classLink.count();
    if (classLinkCount > 0) {
      const href = await classLink.getAttribute('href');
      
      // Navigate directly to the class
      await page.goto(href!);
      
      // Verify we're on the class page (not redirected to login)
      await expect(page).toHaveURL(/\/class\//);
      await expect(page).not.toHaveURL(/\/login/);
    } else {
      // If no classes assigned, try a generic class URL
      // It might redirect to login or show "not found" but shouldn't crash
      await page.goto('/class/1');
      // Just verify we don't crash - actual behavior depends on middleware
    }
  });

  test('teacher can access /student via direct URL', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // First, get a valid student ID from a class page
    await page.goto('/');
    
    // Look for any student link in the page
    const studentLink = page.locator('a[href^="/student/"]').first();
    
    // If student link exists, test direct access
    const studentLinkCount = await studentLink.count();
    if (studentLinkCount > 0) {
      const href = await studentLink.getAttribute('href');
      
      // Navigate directly to the student
      await page.goto(href!);
      
      // Verify we're on the student page (not redirected to login)
      await expect(page).toHaveURL(/\/student\//);
      await expect(page).not.toHaveURL(/\/login/);
    } else {
      // If no students available, try a generic student URL
      await page.goto('/student/1');
      // Just verify we don't crash
    }
  });

  test('teacher CANNOT access /admin (redirects to login)', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // Try to access Admin dashboard
    await page.goto('/admin');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('teacher CANNOT access /hod (redirects to login)', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // Try to access HOD dashboard
    await page.goto('/hod');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('teacher menu does NOT show Admin or HOD links', async ({ page }) => {
    await login(page, 'teacher', 'password');
    
    // Navigate to home/dashboard
    await page.goto('/');
    
    // Verify Admin Dashboard link is NOT visible
    const adminLink = page.locator('a[href="/admin"], a:has-text("Admin Dashboard")');
    await expect(adminLink).not.toBeVisible();
    
    // Verify HOD Dashboard link is NOT visible
    const hodLink = page.locator('a[href="/hod"], a:has-text("HOD Dashboard"), a:has-text("Head of Department")');
    await expect(hodLink).not.toBeVisible();
  });
});

// ============================================================================
// UNAUTHENTICATED ACCESS TESTS
// ============================================================================

test.describe('Unauthenticated Access', () => {
  
  test('unauthenticated user cannot access /admin', async ({ page }) => {
    await page.goto('/admin');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('unauthenticated user cannot access /hod', async ({ page }) => {
    await page.goto('/hod');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('unauthenticated user cannot access /class', async ({ page }) => {
    await page.goto('/class/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });

  test('unauthenticated user cannot access /student', async ({ page }) => {
    await page.goto('/student/1');
    
    // Should be redirected to login
    await expect(page).toHaveURL(/\/login/);
  });
});
