import { test, expect, type Page } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { execSync } from 'child_process';

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Generate a sample XER file using the Python sample_generator and return the path.
 */
function generateXer(filename: string): string {
  const outPath = path.join(os.tmpdir(), filename);
  if (!fs.existsSync(outPath)) {
    execSync(
      `cd ${path.join(__dirname, '../../schedule_comparison')} && ` +
      `.venv/bin/python -c "from sample_generator import generate_sample_xer; generate_sample_xer('${outPath}')"`,
      { stdio: 'pipe' }
    );
  }
  return outPath;
}

// ── Upload page ───────────────────────────────────────────────────────────────

test.describe('Upload page', () => {
  test('loads and shows two file inputs', async ({ page }) => {
    await page.goto('/');
    await expect(page).toHaveTitle(/Schedule/i);
    const inputs = page.locator('input[type="file"]');
    await expect(inputs).toHaveCount(2);
  });

  test('shows baseline and updated labels', async ({ page }) => {
    await page.goto('/');
    const body = await page.textContent('body');
    expect(body?.toLowerCase()).toContain('baseline');
    expect(body?.toLowerCase()).toContain('updated');
  });

  test('shows error when no files uploaded', async ({ page }) => {
    await page.goto('/');
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await expect(page.locator('body')).toContainText(/upload|file|required/i);
  });

  test('shows error for non-XER file', async ({ page }) => {
    const tmpTxt = path.join(os.tmpdir(), 'test.txt');
    fs.writeFileSync(tmpTxt, 'not an xer file');

    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(tmpTxt);
    await inputs.nth(1).setInputFiles(tmpTxt);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await expect(page.locator('body')).toContainText(/xer|invalid|format/i);
  });
});

// ── Comparison dashboard ──────────────────────────────────────────────────────

test.describe('Comparison dashboard', () => {
  let xerPath: string;

  test.beforeAll(() => {
    xerPath = generateXer('pw_baseline.xer');
  });

  test('uploads same file twice and shows dashboard', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();

    // Wait for the dashboard to load (up to 30s for comparison + AI)
    await page.waitForLoadState('networkidle', { timeout: 30000 });
    await expect(page.locator('body')).not.toContainText(/error|failed/i);
  });

  test('dashboard shows summary cards', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    // At least one of the summary stat words should appear
    const body = await page.textContent('body');
    const hasSummary = ['added', 'deleted', 'changed', 'unchanged', 'total']
      .some(w => body?.toLowerCase().includes(w));
    expect(hasSummary).toBe(true);
  });

  test('dashboard shows activity table', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    // Navigate to the Drilldown tab where the activity table is rendered
    await page.locator('[data-tab="drilldown"]').click();
    await page.waitForTimeout(500);

    // Table rows should appear in the drilldown matrix
    const rows = page.locator('table tr, .dd-row');
    const count = await rows.count();
    expect(count).toBeGreaterThan(0);
  });

  test('identical schedules show zero changed activities', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    const body = await page.textContent('body');
    // "changed: 0" or "0 changed" should appear
    const hasZeroChanged = /\b0\b.{0,20}changed|changed.{0,20}\b0\b/i.test(body ?? '');
    expect(hasZeroChanged).toBe(true);
  });
});

// ── Export endpoints ──────────────────────────────────────────────────────────

test.describe('Export endpoints', () => {
  let xerPath: string;

  test.beforeAll(() => {
    xerPath = generateXer('pw_export.xer');
  });

  async function getCacheKey(page: Page): Promise<string> {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    const csvLink = page.locator('a[href*="/export/csv"]').first();
    const href = await csvLink.getAttribute('href') ?? '';
    const match = href.match(/key=([^&]+)/);
    return match ? match[1] : '';
  }

  test('CSV export link is present on dashboard', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    const csvLink = page.locator('a[href*="/export/csv"]');
    await expect(csvLink).toHaveCount(1);
  });

  test('JSON export link is present on dashboard', async ({ page }) => {
    await page.goto('/');
    const inputs = page.locator('input[type="file"]');
    await inputs.nth(0).setInputFiles(xerPath);
    await inputs.nth(1).setInputFiles(xerPath);
    await page.locator('button[type="submit"], input[type="submit"]').click();
    await page.waitForLoadState('networkidle', { timeout: 30000 });

    const jsonLink = page.locator('a[href*="/export/json"]');
    await expect(jsonLink).toHaveCount(1);
  });

  test('CSV download returns valid CSV content', async ({ page }) => {
    const key = await getCacheKey(page);
    expect(key).not.toBe('');

    const resp = await page.request.get(`/export/csv?key=${key}`);
    expect(resp.status()).toBe(200);
    const body = await resp.text();
    expect(body).toContain('task_code');  // CSV header row
  });

  test('JSON export returns valid JSON with activity_variances', async ({ page }) => {
    const key = await getCacheKey(page);
    expect(key).not.toBe('');

    const resp = await page.request.get(`/export/json?key=${key}`);
    expect(resp.status()).toBe(200);
    const data = await resp.json();
    expect(data).toHaveProperty('activity_variances');
    expect(Array.isArray(data.activity_variances)).toBe(true);
  });
});
