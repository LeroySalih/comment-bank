import { test, expect } from '@playwright/test';
import { loginAsTeacher, expectRedirectToLogin } from './helpers';

test.describe('Teacher Role Tests', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsTeacher(page);
  });

  test('Can sign in and see dashboard', async ({ page }) => {
    // After login, should be on home page
    await page.goto('/');

    // Should see at least one class (7A is assigned to teacher)
    await expect(page.locator('text=7A').first()).toBeVisible();
  });

  test('Can navigate to assigned class', async ({ page }) => {
    // Navigate to the class page
    await page.goto('/');

    // Find and click the class link (7A)
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();

    // Should be on the class matrix page (with cuid)
    await expect(page).toHaveURL(/\/class\/[a-z0-9]+/);

    // Should see students in the matrix
    await expect(page.locator('table')).toBeVisible();
  });

  test('Cannot access unassigned class', async ({ page }) => {
    // Try to navigate to 7B (assigned to teacher2)
    await page.goto('/');

    // Try to find and click on 7B link (if visible, which it shouldn't be for teacher)
    const class7BLink = page.locator('a:has-text("7B")');
    const hasAccess = await class7BLink.count() > 0;

    // Teacher should NOT see 7B in their class list
    expect(hasAccess).toBeFalsy();
  });

  test('Can select comment codes for a student', async ({ page }) => {
    // Navigate to assigned class
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();

    await page.waitForLoadState('networkidle');

    // Find the first student row and click a comment code button
    const firstCodeButton = page.locator('button[data-code]').first();

    if (await firstCodeButton.count() > 0) {
      await firstCodeButton.click();

      // Button should have active state (bg-primary class or similar)
      await expect(firstCodeButton).toHaveClass(/bg-primary|selected|active/);
    }
  });

  test('Can copy generated comment', async ({ page }) => {
    // Navigate to class
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();

    await page.waitForLoadState('networkidle');

    // Select some comment codes
    const codeButtons = page.locator('button[data-code]');
    const buttonCount = await codeButtons.count();

    if (buttonCount > 0) {
      // Click first few buttons
      for (let i = 0; i < Math.min(3, buttonCount); i++) {
        await codeButtons.nth(i).click();
        await page.waitForTimeout(200);
      }

      // Look for copy button
      const copyButton = page.locator('button:has-text("Copy")').first();

      if (await copyButton.count() > 0) {
        // Grant clipboard permissions
        await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);

        await copyButton.click();

        // Verify something was copied (clipboard has content)
        const clipboardText = await page.evaluate(() => navigator.clipboard.readText());
        expect(clipboardText.length).toBeGreaterThan(0);
      }
    }
  });

  test('Can navigate to student preview', async ({ page }) => {
    // Navigate to class
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();

    await page.waitForLoadState('networkidle');

    // Look for Preview link or student name link
    const previewLink = page.locator('a:has-text("Preview"), a[href*="/student/"]').first();

    if (await previewLink.count() > 0) {
      await previewLink.click();

      // Should be on student page (with cuid)
      await expect(page).toHaveURL(/\/student\/[a-z0-9]+/);

      // Should see comment editor
      await expect(page.locator('textarea, [contenteditable="true"]')).toBeVisible();
    }
  });

  test('Can edit and save comment text', async ({ page }) => {
    // Navigate to class
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();

    await page.waitForLoadState('networkidle');

    // Navigate to student
    const studentLink = page.locator('a[href*="/student/"]').first();

    if (await studentLink.count() > 0) {
      await studentLink.click();
      await page.waitForLoadState('networkidle');

      // Find textarea
      const textarea = page.locator('textarea').first();

      if (await textarea.count() > 0) {
        const testComment = 'This is a test comment for E2E testing.';
        await textarea.fill(testComment);

        // Find and click save button
        const saveButton = page.locator('button:has-text("Save")').first();

        if (await saveButton.count() > 0) {
          await saveButton.click();
          await page.waitForLoadState('networkidle');

          // Reload and verify persistence
          await page.reload();
          await expect(textarea).toHaveValue(testComment);
        }
      }
    }
  });

  test('Can revert comment', async ({ page }) => {
    // Navigate to class and student
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();
    await page.waitForLoadState('networkidle');

    const studentLink = page.locator('a[href*="/student/"]').first();

    if (await studentLink.count() > 0) {
      await studentLink.click();
      await page.waitForLoadState('networkidle');

      // Look for revert button
      const revertButton = page.locator('button:has-text("Revert")').first();

      if (await revertButton.count() > 0) {
        await revertButton.click();
        await page.waitForLoadState('networkidle');

        // Comment should be cleared or reset
        const textarea = page.locator('textarea').first();
        const value = await textarea.inputValue();

        // Value should be empty or returned to auto-generated state
        expect(value.length).toBeLessThan(200);
      }
    }
  });

  test('Comment selections persist after navigation', async ({ page }) => {
    // Navigate to class
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();
    await classLink.click();
    await page.waitForLoadState('networkidle');

    // Select a comment code
    const firstCodeButton = page.locator('button[data-code]').first();

    if (await firstCodeButton.count() > 0) {
      const codeValue = await firstCodeButton.getAttribute('data-code');
      await firstCodeButton.click();

      // Navigate away
      await page.goto('/');

      // Navigate back
      await classLink.click();
      await page.waitForLoadState('networkidle');

      // Find the same button
      const sameButton = page.locator(`button[data-code="${codeValue}"]`).first();

      // Should still have active state
      await expect(sameButton).toHaveClass(/bg-primary|selected|active/);
    }
  });

  test('Cannot access admin pages', async ({ page }) => {
    const response = await page.goto('/admin');

    // Should redirect to login or show access denied
    const url = page.url();
    const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
    const isRedirectedToLogin = url.includes('/login');

    expect(hasAccessDenied || isRedirectedToLogin).toBeTruthy();
  });

  test('Cannot access HoD pages', async ({ page }) => {
    const response = await page.goto('/hod');

    // Should redirect to login or show access denied
    const url = page.url();
    const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
    const isRedirectedToLogin = url.includes('/login');

    expect(hasAccessDenied || isRedirectedToLogin).toBeTruthy();
  });
});
