/// <reference types="vitest/globals" />

import { SheetWriter } from '@xlsxflow/core';
import { createMockActionContext } from '@activepieces/pieces-framework';
import { excelToCsvAction } from '../src/lib/actions/convert-excel-to-csv';

async function makeXlsxBase64(sheets: Record<string, (string | number)[][]>): Promise<string> {
  const writer = new SheetWriter();
  for (const [name, rows] of Object.entries(sheets)) {
    writer.addSheet(name, rows);
  }
  return Buffer.from(await new Response(writer.write()).arrayBuffer()).toString('base64');
}

describe('convert_excel_to_csv output schema', () => {
  test('declares a schema', () => {
    expect(excelToCsvAction.outputSchema).toBeDefined();
  });

  test('every described path resolves, and nothing is left undescribed', async () => {
    const base64 = await makeXlsxBase64({
      Employees: [['name'], ['Alice']],
      Products: [['sku'], ['P001']],
    });
    const ctx = createMockActionContext({
      propsValue: { file: { base64, extension: 'xlsx', filename: 'test.xlsx' }, sheet_name: 'Products', delimiter_type: ',' },
    });
    const output = (await excelToCsvAction.run(ctx)) as Record<string, unknown>;
    const described = (excelToCsvAction.outputSchema?.fields ?? []).map((f) => f.value ?? f.key);

    for (const path of described) {
      expect(output[path], `"${path}" does not resolve`).toBeDefined();
    }
    for (const key of Object.keys(output)) {
      expect(described.includes(key), `"${key}" is not described`).toBe(true);
    }
  });

  test('available_sheets lists every sheet in the workbook', async () => {
    const base64 = await makeXlsxBase64({
      Employees: [['name'], ['Alice']],
      Products: [['sku'], ['P001']],
    });
    const ctx = createMockActionContext({
      propsValue: { file: { base64, extension: 'xlsx', filename: 'test.xlsx' }, sheet_name: '', delimiter_type: ',' },
    });
    const output = await excelToCsvAction.run(ctx);
    expect(output.available_sheets).toEqual(['Employees', 'Products']);
  });
});
