import { test, expect } from '@playwright/test';

test.describe('HOD Management', () => {
  test.beforeEach(async ({ page }) => {
    // Login as HOD
    await page.goto('/login');
    await page.fill('input[name="username"]', 'leroysalih');
    await page.fill('input[name="password"]', 'password');
    await page.click('button[type="submit"]');
    // Wait for login to complete
    // Wait for login to complete
    await expect(page).not.toHaveURL(/\/login/);
  });

  test('can create and delete a subject', async ({ page }) => {
    await page.goto('/hod');

    // --- CREATE ---
    await page.getByRole('button', { name: '+ Add Subject' }).click();
    const subjectName = 'Test Subject ' + Date.now();
    await page.fill('input[name="name"]', subjectName);
    await page.fill('textarea[name="introduction"]', 'Test intro');
    await page.getByRole('button', { name: 'Add Subject', exact: true }).click();

    // Verify it appears
    await expect(page.getByText(subjectName)).toBeVisible();

    // --- DELETE ---
    // Hover to show buttons
    const subjectCard = page.locator('div.group').filter({ hasText: subjectName });
    await subjectCard.hover();
    
    const deleteBtn = subjectCard.locator('button[title="Delete Subject"]');
    
    // Handle confirm dialog
    page.on('dialog', dialog => dialog.accept());
    await deleteBtn.click();

    // Verify it's gone
    await expect(page.getByText(subjectName)).not.toBeVisible();
  });

  test('can create, edit, and delete a comment group', async ({ page }) => {
    await page.goto('/hod');

    // Select a subject (assuming at least one exists, or seed one)
    // We'll click the first subject card found.
    // If no subject, this will fail - we assume data seeded or available.
    // Let's create a subject first if we can, or rely on seeded "11CS" if standard seed used, 
    // but better to rely on what's visible.
    
    // In seed.ts "11CS" is created.
    // Let's check for "11CS" specifically or any subject.
    // We'll prioritize clicking a known one or creating if dashboard empty logic is handled, 
    // but simplicity: click first one.
    const subjectCard = page.locator('a[href^="/hod/subject/"]').first();
    await expect(subjectCard).toBeVisible();
    await subjectCard.click();

    // --- CREATE ---
    // Click Add Group
    await page.getByRole('button', { name: '+ Add Group' }).click();
    
    const groupName = 'Test Group ' + Date.now();
    await page.fill('input[name="name"]', groupName);
    await page.getByRole('button', { name: 'Add', exact: true }).click();

    // Verify it appears
    await expect(page.getByText(groupName)).toBeVisible();

    // --- EDIT ---
    // Hover over the item to show edit button? Or just find the button near the text.
    // Our implementation puts it in the flex row.
    // We need to target the edit button specifically for this item.
    // Locator: link containing text groupName -> find pencil button within it
    const groupItem = page.locator(`a:has-text("${groupName}")`);
    await groupItem.hover(); // Just in case, though we didn't use opacity-0 group-hover logic yet
    
    // Pencil button has title "Edit Group Name"
    const editBtn = groupItem.locator('button[title="Edit Group Name"]');
    await editBtn.click();
    
    const newGroupName = groupName + ' Edited';
    await page.fill('input[value="' + groupName + '"]', newGroupName);
    await page.getByRole('button', { name: 'Save' }).click();
    
    // Verify updated name
    await expect(page.getByText(newGroupName)).toBeVisible();
    await expect(page.getByText(groupName, { exact: true })).not.toBeVisible();

    // --- DELETE ---
    const updatedGroupItem = page.locator(`a:has-text("${newGroupName}")`);
    const deleteBtn = updatedGroupItem.locator('button[title="Delete Group"]');
    
    // Handle confirm dialog
    page.on('dialog', dialog => dialog.accept());
    
    await deleteBtn.click();
    
    // Verify removed
    await expect(page.getByText(newGroupName)).not.toBeVisible();
  });

  test('can create, edit, and delete a class', async ({ page }) => {
    await page.goto('/hod');
    const subjectCard = page.locator('a[href^="/hod/subject/"]').first();
    await subjectCard.click();

    // --- CREATE ---
    await page.getByRole('button', { name: '+ Add Class' }).click();
    
    const className = 'Class ' + Math.floor(Math.random() * 1000);
    await page.fill('input[name="name"]', className);
    await page.getByRole('button', { name: 'Add', exact: true }).click();

    await expect(page.getByText(className)).toBeVisible();

    // --- EDIT ---
    const classItem = page.locator(`div:has-text("${className}")`).filter({ hasText: 'Students' }); 
    // ^ refine locator as class items are divs with text
    
    const editBtn = classItem.locator('button[title="Edit Class Name"]');
    await classItem.hover();
    await editBtn.click();
    
    const newClassName = className + 'v2';
    await page.fill('input[value="' + className + '"]', newClassName);
    await page.getByRole('button', { name: 'Save' }).click();

    await expect(page.getByText(newClassName)).toBeVisible();
    await expect(page.getByText(className, { exact: true })).not.toBeVisible();

    // --- DELETE ---
    const updatedClassItem = page.locator(`div:has-text("${newClassName}")`).filter({ hasText: 'Students' });
    const deleteBtn = updatedClassItem.locator('button[title="Delete Class"]');
    
    page.on('dialog', dialog => dialog.accept());
    await deleteBtn.click();

    await expect(page.getByText(newClassName)).not.toBeVisible();
  });
});
