import { test, expect } from '@playwright/test';
import { login, TEST_USERS } from './helpers';

test.describe('Admin CCG Management', () => {

  test('CCG tab shows seeded groups', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    // Click the CCG tab
    await page.getByRole('button', { name: 'CCG' }).click();

    // Verify seeded CCG groups are visible
    await expect(page.getByText('Academic').first()).toBeVisible();
    await expect(page.getByText('Effort').first()).toBeVisible();
    await expect(page.getByText('Behaviour').first()).toBeVisible();
    await expect(page.getByText('Homework').first()).toBeVisible();
    await expect(page.getByText('Overall').first()).toBeVisible();
  });

  test('create a new comment group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    // Click Add Group
    await page.getByRole('button', { name: 'Add Group' }).click();

    // Fill form — name (code) and title fields
    await page.locator('input').filter({ hasText: '' }).nth(0).fill('Test Group');
    await page.locator('input').filter({ hasText: '' }).nth(1).fill('Test Group Title');

    // Submit
    await page.getByRole('button', { name: 'Create Group' }).click();

    // Verify it appears on the page
    await expect(page.getByText('Test Group').first()).toBeVisible();
  });

  test('add an option to a group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    // Find the "Test Group" row and click its Add button
    const testGroupRow = page.locator('text=Test Group Title').locator('..');
    await testGroupRow.getByText('Add').click();

    // Fill the option form — code and text
    await page.locator('input[placeholder="e.g. H"]').fill('T1');
    await page.locator('textarea').last().fill('<Name> shows great test skills.');

    // Submit
    await page.getByRole('button', { name: 'Add Comment' }).click();

    // Verify option appears
    await expect(page.getByText('T1').first()).toBeVisible();
    await expect(page.getByText('<Name> shows great test skills.').first()).toBeVisible();
  });

  test('edit a group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    // Find the Test Group card and click Edit group button
    const testGroupRow = page.locator('text=Test Group Title').locator('..');
    await testGroupRow.getByTitle('Edit group').click();

    // Update the title
    const titleInput = page.locator('input').filter({ hasText: /Test Group Title/ });
    await titleInput.clear();
    await titleInput.fill('Updated Test Title');
    await page.getByRole('button', { name: 'Save Changes' }).click();

    // Verify the updated title appears
    await expect(page.getByText('Updated Test Title').first()).toBeVisible();
  });

  test('delete an option', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    // Expand the Test Group to see options
    const testGroupRow = page.locator('text=Updated Test Title').locator('..');
    // Click the chevron to expand
    await testGroupRow.locator('button').first().click();

    // Handle the browser confirm dialog
    page.on('dialog', dialog => dialog.accept());

    // Find the option and delete it
    const optionRow = page.locator('text=<Name> shows great test skills.').locator('..');
    await optionRow.getByTitle('Delete').click();

    // Verify option is removed
    await expect(page.getByText('<Name> shows great test skills.')).not.toBeVisible();
  });

  test('delete a group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    // Handle the browser confirm dialog
    page.on('dialog', dialog => dialog.accept());

    // Find Test Group and click delete
    const testGroupRow = page.locator('text=Updated Test Title').locator('..');
    await testGroupRow.getByTitle('Delete group').click();

    // Verify group is removed
    await expect(page.getByText('Updated Test Title')).not.toBeVisible();
  });

  test('Format tab shows template editor', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Format' }).click();

    // Verify format management UI is visible
    await expect(page.getByText('Comment Format Template').first()).toBeVisible();
    await expect(page.getByText('Available Tags').first()).toBeVisible();
    await expect(page.getByText('Expansion').first()).toBeVisible();

    // Verify the textarea has a template value
    const textarea = page.locator('textarea');
    await expect(textarea).toBeVisible();
  });

  test('update format template', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Format' }).click();

    // Find the format template textarea and update it
    const textarea = page.locator('textarea');
    await textarea.fill('<Academic>\n\n<Effort> <Behaviour>\n\n<Subject>\n\n<Overall>');

    // Save
    await page.getByRole('button', { name: 'Save Template' }).click();

    // Verify save confirmation
    await expect(page.getByText('Saved!')).toBeVisible();
  });
});
