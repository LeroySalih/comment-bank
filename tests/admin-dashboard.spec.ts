import { test, expect } from '@playwright/test';
import { login, TEST_USERS } from './helpers';

test.describe('Admin Dashboard', () => {

  test('dashboard loads with heading and tabs', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await expect(page.getByRole('heading', { name: 'Admin Dashboard' })).toBeVisible();

    // Verify tab buttons are visible
    await expect(page.getByRole('button', { name: 'Users' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Subjects' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Classes' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Deadlines' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Activity Log' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'CCG' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Format' })).toBeVisible();
  });

  test('Users tab shows seeded users', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Users' }).click();

    // Verify seeded usernames appear
    await expect(page.getByText('admin')).toBeVisible();
    await expect(page.getByText('leroysalih')).toBeVisible();
    await expect(page.getByText('teacher')).toBeVisible();
  });

  test('Subjects tab shows seeded subjects', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Subjects' }).click();

    await expect(page.getByText('7CS')).toBeVisible();
    await expect(page.getByText('7DT')).toBeVisible();
    await expect(page.getByText('8CS')).toBeVisible();
  });

  test('Classes tab shows seeded classes', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Classes' }).click();

    await expect(page.getByText('7A')).toBeVisible();
    await expect(page.getByText('7B')).toBeVisible();
    await expect(page.getByText('7C')).toBeVisible();
  });

  test('Deadlines tab loads', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Deadlines' }).click();

    // Tab should load without error — just verify we're still on /admin
    await expect(page).toHaveURL(/\/admin/);
  });

  test('CCG tab loads with common groups', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'CCG' }).click();

    await expect(page.getByText('Common Comment Groups')).toBeVisible();
    await expect(page.getByRole('button', { name: 'Add Group' })).toBeVisible();
  });

  test('Format tab loads with template editor', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Format' }).click();

    await expect(page.getByText('Comment Format Template')).toBeVisible();
    await expect(page.getByRole('button', { name: 'Save Template' })).toBeVisible();
  });

  test('Activity Log tab loads', async ({ page }) => {
    await login(page, TEST_USERS.admin.username, TEST_USERS.admin.password);
    await page.goto('/admin');

    await page.getByRole('button', { name: 'Activity Log' }).click();

    await expect(page).toHaveURL(/\/admin/);
  });

  test('non-admin user is redirected away from /admin', async ({ page }) => {
    await login(page, TEST_USERS.teacher.username, TEST_USERS.teacher.password);
    await page.goto('/admin');

    // Should be redirected to /login by middleware
    await expect(page).toHaveURL(/\/login/);
  });
});
