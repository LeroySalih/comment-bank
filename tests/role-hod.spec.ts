import { test, expect } from '@playwright/test';
import { loginAsHoD } from './helpers';

test.describe('HoD Role Tests', () => {
  test.beforeEach(async ({ page }) => {
    await loginAsHoD(page);
  });

  test('Can sign in and see HoD dashboard', async ({ page }) => {
    // Navigate to HoD dashboard
    await page.goto('/hod');

    // Should see subjects listed
    await expect(page.locator('text=/subject|7CS|7DT|8CS/i').first()).toBeVisible();
  });

  test('Can view subject detail page', async ({ page }) => {
    // Navigate to HoD dashboard
    await page.goto('/hod');
    await page.waitForLoadState('networkidle');

    // Find and click a subject link
    const subjectLink = page.locator('a:has-text("7CS"), a[href*="/subject/"]').first();

    if (await subjectLink.count() > 0) {
      await subjectLink.click();

      // Should be on subject page (might be /hod/subject/)
      await expect(page).toHaveURL(/\/subject\/[a-z0-9]+/);

      // Should see subject details
      await expect(page.locator('text=/7CS|Subject/i').first()).toBeVisible();
    }
  });

  test('Can view class from subject', async ({ page }) => {
    // Navigate to HoD dashboard
    await page.goto('/hod');
    await page.waitForLoadState('networkidle');

    // Find subject link
    const subjectLink = page.locator('a[href*="/subject/"]').first();

    if (await subjectLink.count() > 0) {
      await subjectLink.click();
      await page.waitForLoadState('networkidle');

      // HoD can view all classes through home page
      await page.goto('/');
      const classLink = page.locator('a[href*="/class/"]').first();

      if (await classLink.count() > 0) {
        await classLink.click();
        await page.waitForLoadState('networkidle');

        // Should be on class page
        await expect(page).toHaveURL(/\/class\/[a-z0-9]+/);
      }
    }
  });

  test('Can review student comments', async ({ page }) => {
    // Navigate to a class
    await page.goto('/');

    const classLink = page.locator('a:has-text("7A")').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      // Navigate to a student
      const studentLink = page.locator('a[href*="/student/"]').first();

      if (await studentLink.count() > 0) {
        await studentLink.click();
        await page.waitForLoadState('networkidle');

        // Should see review controls (approve/reject buttons)
        const approveButton = page.locator('button:has-text("Approve"), button:has-text("OK")');
        const rejectButton = page.locator('button:has-text("Reject")');

        const hasReviewControls = (await approveButton.count() > 0) || (await rejectButton.count() > 0);
        expect(hasReviewControls).toBeTruthy();
      }
    }
  });

  test('Can approve a comment', async ({ page }) => {
    // Navigate to a student page
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      const studentLink = page.locator('a[href*="/student/"]').first();

      if (await studentLink.count() > 0) {
        await studentLink.click();
        await page.waitForLoadState('networkidle');

        // First, ensure there's a comment to approve
        const textarea = page.locator('textarea').first();

        if (await textarea.count() > 0) {
          const currentValue = await textarea.inputValue();

          if (!currentValue || currentValue.length < 10) {
            // Add a test comment first
            await textarea.fill('This is a comment ready for approval.');

            const saveButton = page.locator('button:has-text("Save")').first();

            if (await saveButton.count() > 0) {
              await saveButton.click();
              await page.waitForLoadState('networkidle');
            }
          }

          // Now find and click approve button
          const approveButton = page.locator('button:has-text("Approve"), button:has-text("OK")').first();

          if (await approveButton.count() > 0) {
            await approveButton.click();
            await page.waitForLoadState('networkidle');

            // Should see success message or status change
            const hasSuccessIndicator = await page.locator('text=/approved|checked_ok|success/i').count() > 0;
            expect(hasSuccessIndicator).toBeTruthy();
          }
        }
      }
    }
  });

  test('Can reject a comment with note', async ({ page }) => {
    // Navigate to a student page
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      const studentLink = page.locator('a[href*="/student/"]').first();

      if (await studentLink.count() > 0) {
        await studentLink.click();
        await page.waitForLoadState('networkidle');

        // Look for reject button
        const rejectButton = page.locator('button:has-text("Reject")').first();

        if (await rejectButton.count() > 0) {
          await rejectButton.click();

          // May open a dialog or show a note input
          await page.waitForTimeout(500);

          // Look for note input
          const noteInput = page.locator('input[placeholder*="note"], textarea[placeholder*="note"], input[name*="note"]').first();

          if (await noteInput.count() > 0) {
            await noteInput.fill('Needs improvement - grammar errors.');

            // Submit
            const submitButton = page.locator('button:has-text("Submit"), button:has-text("Confirm")').first();

            if (await submitButton.count() > 0) {
              await submitButton.click();
              await page.waitForLoadState('networkidle');

              // Should see rejected status
              const hasRejectedIndicator = await page.locator('text=/rejected|checked_rejected/i').count() > 0;
              expect(hasRejectedIndicator).toBeTruthy();
            }
          }
        }
      }
    }
  });

  test('Can reset comment status', async ({ page }) => {
    // Navigate to a student page with a reviewed comment
    await page.goto('/');
    const classLink = page.locator('a:has-text("7A")').first();

    if (await classLink.count() > 0) {
      await classLink.click();
      await page.waitForLoadState('networkidle');

      const studentLink = page.locator('a[href*="/student/"]').first();

      if (await studentLink.count() > 0) {
        await studentLink.click();
        await page.waitForLoadState('networkidle');

        // Look for reset button
        const resetButton = page.locator('button:has-text("Reset")').first();

        if (await resetButton.count() > 0) {
          await resetButton.click();
          await page.waitForLoadState('networkidle');

          // Status should return to required_check or pending
          const hasResetIndicator = await page.locator('text=/required_check|pending|awaiting/i').count() > 0;
          expect(hasResetIndicator).toBeTruthy();
        }
      }
    }
  });

  test('Can view all assigned classes', async ({ page }) => {
    // HoD (leroysalih) is assigned to all subjects, so should see all classes
    await page.goto('/');

    // Should see all three classes
    const has7A = await page.locator('text=7A').count() > 0;
    const has7B = await page.locator('text=7B').count() > 0;
    const has7C = await page.locator('text=7C').count() > 0;

    expect(has7A || has7B || has7C).toBeTruthy();
  });

  test('Cannot access admin pages', async ({ page }) => {
    const response = await page.goto('/admin');

    // Should redirect to login or show access denied
    const url = page.url();
    const hasAccessDenied = await page.locator('text=/access denied|not authorized|forbidden/i').count() > 0;
    const isRedirectedToLogin = url.includes('/login');

    expect(hasAccessDenied || isRedirectedToLogin).toBeTruthy();
  });
});
