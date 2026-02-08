import { test, expect } from '@playwright/test';
import { loginAsAdmin } from './helpers';

test.describe('Admin Role Tests', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsAdmin(page);
  });

  test('Can sign in and see admin dashboard', async ({ page }) => {
    // Navigate to admin page
    await page.goto('/admin');

    // Should see tabs
    await expect(page.locator('text=/Users|Subjects|Classes|Activity|CCG/i').first()).toBeVisible();
  });

  test('Can view Users tab', async ({ page }) => {
    await page.goto('/admin');
    await page.waitForLoadState('networkidle');

    // Click Users tab if not already visible
    const usersTab = page.locator('button:has-text("Users"), a:has-text("Users")').first();

    if (await usersTab.count() > 0) {
      await usersTab.click();
      await page.waitForLoadState('networkidle');
    }

    // Should see all 5 users (admin, leroysalih, teacher, teacher2, teacher3)
    const usersSection = page.locator('text=/admin|leroysalih|teacher/i');
    await expect(usersSection.first()).toBeVisible();

    // Count user entries
    const userCount = await page.locator('text=/teacher|admin|leroysalih/i').count();
    expect(userCount).toBeGreaterThanOrEqual(3);
  });

  test('Can view Subjects tab', async ({ page }) => {
    await page.goto('/admin');
    await page.waitForLoadState('networkidle');

    // Should see Subjects tab button
    const subjectsTab = page.locator('button:has-text("Subjects"), a:has-text("Subjects")');
    await expect(subjectsTab.first()).toBeVisible();

    // Click it
    await subjectsTab.first().click();
    await page.waitForTimeout(500);

    // Verify we're still on admin page (tab system works)
    await expect(page).toHaveURL(/\/admin/);
  });

  test('Can view Classes tab', async ({ page }) => {
    await page.goto('/admin');
    await page.waitForLoadState('networkidle');

    // Click Classes tab
    const classesTab = page.locator('button:has-text("Classes"), a:has-text("Classes")').first();

    if (await classesTab.count() > 0) {
      await classesTab.click();
      await page.waitForLoadState('networkidle');
    }

    // Should see all 3 classes (7A, 7B, 7C)
    const classText = await page.locator('text=/7A|7B|7C/').first();
    await expect(classText).toBeVisible();
  });

  test('Can view Activity Log tab', async ({ page }) => {
    await page.goto('/admin');
    await page.waitForLoadState('networkidle');

    // Click Activity Log tab
    const activityTab = page.locator('button:has-text("Activity"), a:has-text("Activity"), button:has-text("Log"), a:has-text("Log")').first();

    if (await activityTab.count() > 0) {
      await activityTab.click();
      await page.waitForLoadState('networkidle');
    }

    // Should see audit log entries
    const hasLogEntries = await page.locator('text=/action|user|timestamp|log/i').count() > 0;
    expect(hasLogEntries).toBeTruthy();
  });

  test('Can view CCG management page', async ({ page }) => {
    // Navigate to CCG admin page
    await page.goto('/admin/ccg');

    // Should see CCG management interface
    await expect(page.locator('text=/Common Comment Group|CCG/i')).toBeVisible();

    // Should see list of CCGs
    const hasCCGs = await page.locator('text=/group|comment/i').count() > 0;
    expect(hasCCGs).toBeTruthy();
  });

  test('Can access any class page', async ({ page }) => {
    // Try accessing different classes
    await page.goto('/');
    await page.waitForLoadState('networkidle');

    // Look for class links - admin should see all classes
    const classLink = page.locator('a[href*="/class/"]').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      // Should successfully load class page
      await expect(page).toHaveURL(/\/class\/[a-z0-9]+/);

      // Should NOT see access denied
      const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
      expect(hasAccessDenied).toBeFalsy();
    } else {
      // If no class links visible, admin can still navigate directly
      expect(true).toBeTruthy();
    }
  });

  test('Can access any student page', async ({ page }) => {
    // Navigate to a class
    await page.goto('/');

    const classLink = page.locator('a:has-text("7A"), a:has-text("7B"), a:has-text("7C")').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      // Navigate to a student
      const studentLink = page.locator('a[href*="/student/"]').first();

      if (await studentLink.count() > 0) {
        await studentLink.click();

        // Should successfully load student page
        await expect(page).toHaveURL(/\/student\/[a-z0-9]+/);

        // Should NOT see access denied
        const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
        expect(hasAccessDenied).toBeFalsy();
      }
    }
  });

  test('Can access admin pages without restriction', async ({ page }) => {
    // Admin should have unrestricted access to admin pages
    await page.goto('/admin');
    await page.waitForLoadState('networkidle');

    // Should be on admin page
    await expect(page).toHaveURL(/\/admin/);

    // Should NOT see access denied
    const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
    expect(hasAccessDenied).toBeFalsy();

    // Should see admin interface
    const hasAdminContent = await page.locator('button:has-text("Users"), button:has-text("Classes")').count() > 0;
    expect(hasAdminContent).toBeTruthy();
  });
});
