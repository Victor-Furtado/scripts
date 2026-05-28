#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function printHelp() {
  console.log(`Uso:
  node merge-excels.js [opções] <arquivo1.xlsx> <arquivo2.xlsx> ...

Opções:
  -o, --output <arquivo>  Caminho do arquivo de saída (padrão: merged.xlsx)
  --dedupe <coluna>       Remove linhas duplicadas usando o nome da coluna ou o índice da coluna começando em 1
  --sort <coluna>         Ordena usando o nome da coluna ou o índice da coluna começando em 1
  --desc                  Ordena em ordem decrescente
  --sheet <nome|índice>   Usa uma aba específica de cada pasta de trabalho (padrão: primeira aba)
  -h, --help              Mostra esta ajuda

Exemplos:
  node merge-excels.js a.xlsx b.xlsx -o merged.xlsx
  node merge-excels.js --dedupe Email --sort Name *.xlsx
  node merge-excels.js --sheet 2 --sort 3 file1.xlsx file2.xlsx
`);
}

function parseArgs(argv) {
  const options = {
    output: 'merged.xlsx',
    dedupe: null,
    sort: null,
    desc: false,
    sheet: null,
    files: [],
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];

    if (arg === '-h' || arg === '--help') {
      options.help = true;
      continue;
    }

    if (arg === '-o' || arg === '--output') {
      options.output = argv[++i];
      continue;
    }

    if (arg === '--dedupe') {
      options.dedupe = argv[++i];
      continue;
    }

    if (arg === '--sort') {
      options.sort = argv[++i];
      continue;
    }

    if (arg === '--desc') {
      options.desc = true;
      continue;
    }

    if (arg === '--sheet') {
      options.sheet = argv[++i];
      continue;
    }

    if (arg.startsWith('-')) {
      throw new Error(`Unknown option: ${arg}`);
    }

    options.files.push(arg);
  }

  return options;
}

function isNumeric(value) {
  return /^\d+$/.test(String(value));
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  return value;
}

function resolveColumnIndex(columns, identifier) {
  if (identifier === null || identifier === undefined) {
    return -1;
  }

  if (isNumeric(identifier)) {
    const index = Number(identifier) - 1;
    return index >= 0 ? index : -1;
  }

  return columns.findIndex((column) => String(column).trim() === String(identifier).trim());
}

function getSheetName(workbook, sheetSelector) {
  if (!sheetSelector) {
    return workbook.SheetNames[0];
  }

  if (isNumeric(sheetSelector)) {
    const index = Number(sheetSelector) - 1;
    return workbook.SheetNames[index];
  }

  return workbook.SheetNames.find((name) => name === sheetSelector);
}

function rowIsEmpty(row) {
  return row.every((value) => value === null || value === undefined || String(value).trim() === '');
}

function collectExcelFiles(inputPath, outputPath = null) {
  const resolvedInputPath = path.resolve(inputPath);

  if (!fs.existsSync(resolvedInputPath)) {
    throw new Error(`File or folder not found: ${inputPath}`);
  }

  const stats = fs.statSync(resolvedInputPath);

  if (stats.isFile()) {
    return resolvedInputPath.toLowerCase().endsWith('.xlsx') ? [resolvedInputPath] : [];
  }

  const discoveredFiles = [];

  function walk(directoryPath) {
    const entries = fs.readdirSync(directoryPath, { withFileTypes: true });

    for (const entry of entries) {
      const entryPath = path.join(directoryPath, entry.name);

      if (outputPath && path.resolve(entryPath) === outputPath) {
        continue;
      }

      if (entry.isDirectory()) {
        walk(entryPath);
        continue;
      }

      if (entry.isFile() && entry.name.toLowerCase().endsWith('.xlsx')) {
        discoveredFiles.push(entryPath);
      }
    }
  }

  walk(resolvedInputPath);
  return discoveredFiles.sort((left, right) => left.localeCompare(right));
}

function main() {
  const options = parseArgs(process.argv.slice(2));

  if (options.help) {
    printHelp();
    return;
  }

  if (options.files.length === 0) {
    printHelp();
    process.exitCode = 1;
    return;
  }

  const outputPath = path.resolve(options.output);
  const inputFiles = options.files.flatMap((inputPath) => collectExcelFiles(inputPath, outputPath));

  if (inputFiles.length === 0) {
    throw new Error('No .xlsx files found in the provided input.');
  }

  const mergedColumns = [];
  const columnSet = new Set();
  const rows = [];

  for (const filePath of inputFiles) {
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = getSheetName(workbook, options.sheet);

    if (!sheetName) {
      throw new Error(`Sheet not found in file: ${filePath}`);
    }

    const sheet = workbook.Sheets[sheetName];
    const sheetRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      defval: '',
    });

    if (sheetRows.length === 0) {
      continue;
    }

    const headers = sheetRows[0].map((value, index) => {
      const header = String(value || `Column ${index + 1}`).trim();
      return header || `Column ${index + 1}`;
    });

    for (const header of headers) {
      if (!columnSet.has(header)) {
        columnSet.add(header);
        mergedColumns.push(header);
      }
    }

    for (const rawRow of sheetRows.slice(1)) {
      if (rowIsEmpty(rawRow)) {
        continue;
      }

      const row = {};
      headers.forEach((header, index) => {
        row[header] = normalizeCellValue(rawRow[index]);
      });
      rows.push(row);
    }
  }

  if (rows.length === 0) {
    throw new Error('No data rows found to merge.');
  }

  let filteredRows = rows;

  if (options.dedupe !== null && options.dedupe !== undefined) {
    const dedupeIndex = resolveColumnIndex(mergedColumns, options.dedupe);

    if (dedupeIndex < 0 || dedupeIndex >= mergedColumns.length) {
      throw new Error(`Dedupe column not found: ${options.dedupe}`);
    }

    const dedupeColumn = mergedColumns[dedupeIndex];
    const seen = new Set();

    filteredRows = filteredRows.filter((row) => {
      const value = row[dedupeColumn];
      const key = `${dedupeColumn}::${String(value)}`;

      if (seen.has(key)) {
        return false;
      }

      seen.add(key);
      return true;
    });
  }

  if (options.sort !== null && options.sort !== undefined) {
    const sortIndex = resolveColumnIndex(mergedColumns, options.sort);

    if (sortIndex < 0 || sortIndex >= mergedColumns.length) {
      throw new Error(`Sort column not found: ${options.sort}`);
    }

    const sortColumn = mergedColumns[sortIndex];
    filteredRows = [...filteredRows].sort((left, right) => {
      const a = normalizeCellValue(left[sortColumn]);
      const b = normalizeCellValue(right[sortColumn]);

      if (a === b) {
        return 0;
      }

      if (a === '') {
        return 1;
      }

      if (b === '') {
        return -1;
      }

      if (typeof a === 'number' && typeof b === 'number') {
        return options.desc ? b - a : a - b;
      }

      const comparison = String(a).localeCompare(String(b), undefined, {
        numeric: true,
        sensitivity: 'base',
      });

      return options.desc ? -comparison : comparison;
    });
  }

  const outputRows = [mergedColumns];

  for (const row of filteredRows) {
    outputRows.push(mergedColumns.map((column) => row[column] ?? ''));
  }

  const outputWorkbook = XLSX.utils.book_new();
  const outputSheet = XLSX.utils.aoa_to_sheet(outputRows);
  XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'Merged');

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  XLSX.writeFile(outputWorkbook, outputPath);

  console.log(`Merged ${inputFiles.length} file(s) into ${outputPath}`);
  console.log(`Rows written: ${filteredRows.length}`);
}

try {
  main();
} catch (error) {
  console.error(error.message);
  process.exitCode = 1;
}
