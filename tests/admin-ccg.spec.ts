import { test, expect } from '@playwright/test';
import { login, TEST_USERS } from './helpers';

test.describe('Admin CCG Management', () => {

  test('navigate to CCG page from admin dashboard', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    // Click the Common Comment Groups link
    await page.getByText('Common Comment Groups').first().click();

    await expect(page).toHaveURL(/\/admin\/ccg/);
    await expect(page.getByRole('heading', { name: 'Common Comment Groups' })).toBeVisible();
  });

  test('CCG page shows seeded groups', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Verify seeded CCG groups are visible
    await expect(page.getByText('Academic Performance').first()).toBeVisible();
    await expect(page.getByText('Effort').first()).toBeVisible();
    await expect(page.getByText('Behaviour').first()).toBeVisible();
    await expect(page.getByText('Homework').first()).toBeVisible();
    await expect(page.getByText('Overall').first()).toBeVisible();
  });

  test('create a new comment group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Click Add Group
    await page.getByRole('button', { name: 'Add Group' }).click();

    // Fill form
    await page.fill('input[name="name"]', 'Test Group');
    await page.fill('input[name="title"]', 'Test Group Title');
    await page.selectOption('select[name="paragraphPosition"]', 'p1');

    // Submit
    await page.getByRole('button', { name: 'Create Group' }).click();

    // Verify it appears on the page
    await expect(page.getByText('Test Group').first()).toBeVisible();
  });

  test('add an option to a group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find the "Test Group" card and click its Add Option button
    // We target the last "Add Option" link since Test Group was just created
    const testGroupCard = page.locator('h3:has-text("Test Group")').locator('..').locator('..').locator('..');
    await testGroupCard.getByText('Add Option').click();

    // Fill the option form
    await testGroupCard.locator('input[name="code"]').fill('T1');
    await testGroupCard.locator('input[name="text"]').fill('<Name> shows great test skills.');

    // Submit
    await testGroupCard.getByRole('button', { name: 'Add' }).click();

    // Verify option appears
    await expect(page.getByText('T1').first()).toBeVisible();
    await expect(page.getByText('<Name> shows great test skills.').first()).toBeVisible();
  });

  test('edit a group title', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find the Test Group card and click edit
    const testGroupHeader = page.locator('h3:has-text("Test Group")').locator('..').locator('..');
    await testGroupHeader.getByTitle('Edit group').click();

    // The edit form should appear — update the title
    await page.locator('input[name="title"]').fill('Updated Test Title');
    await page.getByRole('button', { name: 'Update Group' }).click();

    // Verify the updated title appears
    await expect(page.getByText('Updated Test Title').first()).toBeVisible();
  });

  test('edit an option text', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find option T1 in Test Group and click its edit button
    const optionRow = page.locator('text=<Name> shows great test skills.').locator('..').locator('..');
    await optionRow.getByTitle('Edit').click();

    // Update the text
    await page.locator('input[name="text"]').last().fill('<Name> has updated test skills.');
    await page.getByRole('button', { name: 'Save' }).click();

    // Verify updated text
    await expect(page.getByText('<Name> has updated test skills.').first()).toBeVisible();
  });

  test('delete an option', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find option T1 and click delete
    const optionRow = page.locator('text=<Name> has updated test skills.').locator('..').locator('..');

    // Handle the browser confirm dialog
    page.on('dialog', dialog => dialog.accept());

    await optionRow.getByTitle('Delete').click();

    // Verify option is removed
    await expect(page.getByText('<Name> has updated test skills.')).not.toBeVisible();
  });

  test('delete a group', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find Test Group and click delete
    const testGroupHeader = page.locator('h3:has-text("Test Group")').locator('..').locator('..');
    await testGroupHeader.getByTitle('Delete group').click();

    // Confirm in the modal
    await page.getByRole('button', { name: 'Delete Group' }).click();

    // Verify group is removed
    await expect(page.locator('h3:has-text("Test Group")')).not.toBeVisible();
  });

  test('update P2 wrapper template', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin/ccg');

    // Find the wrapper template textarea
    const templateSection = page.locator('h3:has-text("P2 Wrapper Template")').locator('..').locator('..');
    const textarea = templateSection.locator('textarea');

    // Clear and type new template
    await textarea.fill('<Effort> <Behaviour> <Homework>');

    // Save
    await templateSection.getByRole('button', { name: 'Save Template' }).click();

    // Verify save confirmation
    await expect(templateSection.getByText('Saved!')).toBeVisible();
  });
});
