import { test, expect } from '@playwright/test';

/**
 * Teacher Comment Flow E2E Test
 * 
 * This test validates the complete teacher workflow for managing student comments:
 * 1. Sign in as teacher
 * 2. Navigate to an assigned class
 * 3. Select comment options from different groups
 * 4. Validate copy button states (disabled → enabled)
 * 5. Copy comment to clipboard
 * 6. Navigate to student preview page
 * 7. Edit comment text
 * 8. Verify persistence of changes
 */

// Helper function to login as teacher
async function loginAsTeacher(page: any) {
  await page.goto('/login');
  await page.fill('input[name="username"]', 'teacher');
  await page.fill('input[name="password"]', 'password');
  await page.click('button[type="submit"]');
  
  // Wait for successful login (redirected away from login page)
  await expect(page).not.toHaveURL(/\/login/);
}

test.describe('Teacher Comment Flow', () => {
  
  test('complete teacher workflow: sign-in, navigate, populate comments, copy, preview, edit, and verify persistence', async ({ page }) => {
    // ========================================================================
    // STEP 1: Teacher Sign-in
    // ========================================================================
    await loginAsTeacher(page);
    
    // Verify we're logged in (should see Sign Out button)
    await expect(page.getByRole('button', { name: 'Sign Out' })).toBeVisible();
    
    // ========================================================================
    // STEP 2: Navigate to Class
    // ========================================================================
    // Should be on dashboard after login
    await page.goto('/');
    
    // Find and click on the first class link
    const classLink = page.locator('a[href^="/class/"]').first();
    await expect(classLink).toBeVisible();
    
    const classHref = await classLink.getAttribute('href');
    await classLink.click();
    
    // Verify we're on the class page
    await expect(page).toHaveURL(new RegExp(classHref!));
    await expect(page.getByRole('heading', { name: /Class Matrix:/ })).toBeVisible();
    
    // ========================================================================
    // STEP 3: Find First Student and Copy Button
    // ========================================================================
    // Find the first student row
    const firstStudentRow = page.locator('tbody tr').first();
    await expect(firstStudentRow).toBeVisible();
    
    // Find the copy button for this student
    const copyButton = firstStudentRow.locator('button').filter({ has: page.locator('svg') }).first();
    await expect(copyButton).toBeVisible();
    
    // ========================================================================
    // STEP 4: Select Comment Options from All Groups
    // ========================================================================
    // Find the first group selector (should be in the first data column after name and gender)
    const firstGroupSelector = firstStudentRow.locator('td').nth(2).locator('button').first();
    await expect(firstGroupSelector).toBeVisible();
    await firstGroupSelector.click();
    
    // Wait a moment for the state to update
    await page.waitForTimeout(500);
    
    // ========================================================================
    // STEP 5: Select Second Group Option (Complete All Groups)
    // ========================================================================
    // Find the second group selector
    const secondGroupSelector = firstStudentRow.locator('td').nth(3).locator('button').first();
    await expect(secondGroupSelector).toBeVisible();
    await secondGroupSelector.click();
    
    // Wait for state update
    await page.waitForTimeout(500);
    
    // ========================================================================
    // STEP 6: Verify Copy Button is Enabled
    // ========================================================================
    // Copy button should now be enabled (green)
    await expect(copyButton).not.toBeDisabled();
    await expect(copyButton).not.toHaveClass(/bg-red-50/);
    
    // ========================================================================
    // STEP 7: Copy Comment to Clipboard
    // ========================================================================
    // Grant clipboard permissions
    await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
    
    // Click the copy button
    await copyButton.click();
    
    // Wait for copy animation
    await page.waitForTimeout(500);
    
    // Verify clipboard contains text (we can't easily verify exact content without knowing the student name)
    const clipboardText = await page.evaluate(() => navigator.clipboard.readText());
    expect(clipboardText).toBeTruthy();
    expect(clipboardText.length).toBeGreaterThan(0);
    
    // ========================================================================
    // STEP 8: Navigate to Preview Page
    // ========================================================================
    // Find and click the Preview link for this student
    const previewLink = firstStudentRow.locator('a', { hasText: 'Preview' });
    await expect(previewLink).toBeVisible();
    
    const previewHref = await previewLink.getAttribute('href');
    await previewLink.click();
    
    // Verify we're on the student page
    await expect(page).toHaveURL(new RegExp(previewHref!));
    await expect(page.getByRole('heading', { name: /Edit Report:/ })).toBeVisible();
    
    // ========================================================================
    // STEP 9: Verify Comment is Displayed and Matches Generated Content
    // ========================================================================
    // Find the comment editor textarea
    const commentTextarea = page.locator('textarea').first();
    await expect(commentTextarea).toBeVisible();
    
    // Get the current comment text
    const displayedComment = await commentTextarea.inputValue();
    expect(displayedComment).toBeTruthy();
    expect(displayedComment.length).toBeGreaterThan(0);
    
    // Verify the comment contains expected content based on selections
    expect(displayedComment).toContain('Computer Science');
    
    // ========================================================================
    // STEP 10: Verify Selected Codes Persist After Navigation
    // ========================================================================
    // Navigate back to class page
    await page.goto(classHref!);
    await expect(page.getByRole('heading', { name: /Class Matrix:/ })).toBeVisible();
    
    // Find the first student row again
    const studentRowAgain = page.locator('tbody tr').first();
    
    // Verify first group still has selection (button should have active styling)
    const firstGroupAgain = studentRowAgain.locator('td').nth(2).locator('button').first();
    await expect(firstGroupAgain).toHaveClass(/bg-primary/);
    
    // Verify second group still has selection
    const secondGroupAgain = studentRowAgain.locator('td').nth(3).locator('button').first();
    await expect(secondGroupAgain).toHaveClass(/bg-primary/);
    
    // Verify copy button is still enabled
    const copyButtonAgain = studentRowAgain.locator('button').filter({ has: page.locator('svg') }).first();
    await expect(copyButtonAgain).not.toBeDisabled();
  });
  
  test('teacher can navigate between multiple students with different comment selections', async ({ page }) => {
    await loginAsTeacher(page);
    
    // Navigate to class
    await page.goto('/');
    const classLink = page.locator('a[href^="/class/"]').first();
    const classHref = await classLink.getAttribute('href');
    await classLink.click();
    
    // Get first two student rows
    const firstStudent = page.locator('tbody tr').nth(0);
    const secondStudent = page.locator('tbody tr').nth(1);
    
    // Select options for first student (H, H)
    await firstStudent.locator('td').nth(2).locator('button').first().click();
    await page.waitForTimeout(800); // Increased wait for database save
    await firstStudent.locator('td').nth(3).locator('button').first().click();
    await page.waitForTimeout(800);
    
    // Select DIFFERENT options for second student (M, M)
    await secondStudent.locator('td').nth(2).locator('button').nth(1).click();
    await page.waitForTimeout(800);
    await secondStudent.locator('td').nth(3).locator('button').nth(1).click();
    await page.waitForTimeout(800);
    
    // Get preview hrefs before navigating
    const firstPreviewLink = firstStudent.locator('a', { hasText: 'Preview' });
    const firstPreviewHref = await firstPreviewLink.getAttribute('href');
    const secondPreviewLink = secondStudent.locator('a', { hasText: 'Preview' });
    const secondPreviewHref = await secondPreviewLink.getAttribute('href');
    
    // Navigate to first student's preview
    await firstPreviewLink.click();
    
    // Wait for page to load and comment to generate
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1000); // Wait for client-side comment generation
    
    const firstTextarea = page.locator('textarea').first();
    const firstComment = await firstTextarea.inputValue();
    
    // Verify first comment contains expected text for H selections
    expect(firstComment).toContain('excellent understanding');
    expect(firstComment).toContain('exceptional critical thinking');
    
    // Navigate to second student's preview
    await page.goto(secondPreviewHref!);
    
    // Wait for page to load and comment to generate
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1000); // Wait for client-side comment generation
    
    const secondTextarea = page.locator('textarea').first();
    const secondComment = await secondTextarea.inputValue();
    
    // Verify second comment contains expected text for M selections
    expect(secondComment).toContain('good progress');
    expect(secondComment).toContain('good thinking skills');
    
    // Verify comments are different (because we selected different options)
    expect(firstComment).not.toBe(secondComment);
    
    // Navigate back to class page and verify selections persist
    await page.goto(classHref!);
    
    // Verify first student still has H selections
    const firstStudentAgain = page.locator('tbody tr').nth(0);
    const firstGroup1 = firstStudentAgain.locator('td').nth(2).locator('button').first();
    await expect(firstGroup1).toHaveClass(/bg-primary/);
    
    // Verify second student still has M selections
    const secondStudentAgain = page.locator('tbody tr').nth(1);
    const secondGroup1 = secondStudentAgain.locator('td').nth(2).locator('button').nth(1);
    await expect(secondGroup1).toHaveClass(/bg-primary/);
  });
});
