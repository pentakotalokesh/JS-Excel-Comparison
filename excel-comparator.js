/**
 * Excel File Comparison Tool - JavaScript/Node.js Version
 * Compares multiple sheets from two Excel files with automatic key detection
 *
 * Required packages:
 * npm install xlsx exceljs lodash
 */

const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const _ = require("lodash");
const fs = require("fs");
const path = require("path");

class ExcelComparator {
  constructor(config) {
    this.file1Path = config.file1Path;
    this.file2Path = config.file2Path;
    this.sheets = config.sheets || null;
    this.headerRows = config.headerRows || {};
    this.keyColumns = config.keyColumns || {};
    this.file1Label = config.file1Label || "Old Version";
    this.file2Label = config.file2Label || "New Version";
    this.results = [];

    // Validate files exist
    if (!fs.existsSync(this.file1Path)) {
      throw new Error(`File not found: ${this.file1Path}`);
    }
    if (!fs.existsSync(this.file2Path)) {
      throw new Error(`File not found: ${this.file2Path}`);
    }
  }

  log(message) {
    const timestamp = new Date().toISOString();
    console.log(`${timestamp} - ${message}`);
  }

  normalizeColumnName(name) {
    return String(name).trim().toLowerCase().replace(/\s+/g, "_");
  }

  normalize(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toISOString().split("T")[0];
    let v = String(value).trim();
    // Convert numbers to consistent string format
    if (!isNaN(v) && v !== "") {
      v = Number(v).toString();
    }
    return v.toLowerCase();
  }

  normalizeData(data) {
    return data.map((row) => {
      const normalized = {};
      Object.keys(row).forEach((key) => {
        const normalizedKey = this.normalizeColumnName(key);
        normalized[normalizedKey] = this.normalize(row[key]);
      });
      return normalized;
    });
  }

  readSheet(filePath, sheetName) {
    try {
      const workbook = XLSX.readFile(filePath);

      if (!workbook.SheetNames.includes(sheetName)) {
        this.log(`Sheet '${sheetName}' not found in ${filePath}`);
        return [];
      }

      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, {
        defval: "",
        raw: false,
      });

      // Remove unnamed columns
      const cleanedData = data
        .map((row) => {
          const cleaned = {};
          Object.keys(row).forEach((key) => {
            if (!key.toLowerCase().includes("unnamed")) {
              cleaned[key] = row[key];
            }
          });
          return cleaned;
        })
        .filter((row) => Object.keys(row).length > 0);

      const normalized = this.normalizeData(cleanedData);

      this.log(
        `Read sheet '${sheetName}' from ${path.basename(filePath)}: ${
          normalized.length
        } rows, ${
          normalized.length > 0 ? Object.keys(normalized[0]).length : 0
        } columns`
      );
      return normalized;
    } catch (error) {
      this.log(`Error reading sheet '${sheetName}': ${error.message}`);
      return [];
    }
  }

  detectKeyColumn(data, sheetName) {
    if (data.length === 0) return null;

    // Check user-specified keys
    if (this.keyColumns[sheetName]) {
      this.log(
        `Using user-specified key columns for '${sheetName}': ${this.keyColumns[
          sheetName
        ].join(", ")}`
      );
      return this.keyColumns[sheetName];
    }

    const columns = Object.keys(data[0]);

    // Try single column with high uniqueness
    for (const col of columns) {
      const values = data.map((row) => row[col]).filter((v) => v !== "");
      const uniqueRatio = _.uniq(values).length / values.length;

      if (uniqueRatio > 0.95) {
        this.log(`Auto-detected key column for '${sheetName}': ${col}`);
        return col;
      }
    }

    // Try composite key (first 2-3 columns)
    for (
      let numCols = 2;
      numCols <= 3 && numCols <= columns.length;
      numCols++
    ) {
      const testCols = columns.slice(0, numCols);
      const compositeKeys = data.map((row) =>
        testCols.map((col) => row[col]).join("||")
      );
      const uniqueRatio = _.uniq(compositeKeys).length / compositeKeys.length;

      if (uniqueRatio > 0.95) {
        this.log(
          `Using composite key for '${sheetName}': ${testCols.join(", ")}`
        );
        return testCols;
      }
    }

    this.log(
      `No unique key found for '${sheetName}', using full-row comparison`
    );
    return null;
  }

  createRowHash(row, keyCol) {
    if (Array.isArray(keyCol)) {
      // Composite key
      return keyCol.map((col) => this.normalize(row[col] || "")).join("||");
    } else if (keyCol === null) {
      // Full row hash
      return Object.values(row)
        .map((v) => this.normalize(v))
        .join("||");
    } else {
      // Single key
      return this.normalize(row[keyCol] || "");
    }
  }

  findNewRecords(data1, data2, keyCol) {
    if (data2.length === 0) return [];
    if (data1.length === 0) return data2;

    const keys1 = new Set(data1.map((row) => this.createRowHash(row, keyCol)));
    return data2.filter((row) => !keys1.has(this.createRowHash(row, keyCol)));
  }

  findDeletedRecords(data1, data2, keyCol) {
    if (data1.length === 0) return [];
    if (data2.length === 0) return data1;

    const keys2 = new Set(data2.map((row) => this.createRowHash(row, keyCol)));
    return data1.filter((row) => !keys2.has(this.createRowHash(row, keyCol)));
  }

  findModifiedRecords(data1, data2, keyCol) {
    if (data1.length === 0 || data2.length === 0) {
      return { records: [], columnChanges: {} };
    }

    // Create maps by key
    const map1 = new Map();
    const map2 = new Map();

    data1.forEach((row, idx) => {
      const key = this.createRowHash(row, keyCol);
      map1.set(key, { row, originalIndex: idx });
    });

    data2.forEach((row, idx) => {
      const key = this.createRowHash(row, keyCol);
      map2.set(key, { row, originalIndex: idx });
    });

    const modifiedRecords = [];
    const columnChanges = {};

    // Find common keys
    const commonKeys = [...map1.keys()].filter((key) => map2.has(key));

    commonKeys.forEach((key) => {
      const entry1 = map1.get(key);
      const entry2 = map2.get(key);
      const row1 = entry1.row;
      const row2 = entry2.row;

      const changes = {};
      let hasChanges = false;

      // Get all columns from both rows
      const allCols = new Set([...Object.keys(row1), ...Object.keys(row2)]);

      allCols.forEach((col) => {
        const val1 = this.normalize(row1[col]);
        const val2 = this.normalize(row2[col]);

        if (val1 !== val2) {
          // Only add if values are not empty
          if (val1 !== "" || val2 !== "") {
            changes[`${col}_old`] = val1;
            changes[`${col}_new`] = val2;
            columnChanges[col] = (columnChanges[col] || 0) + 1;
            hasChanges = true;
          }
        }
      });

      if (hasChanges) {
        // Add key information
        if (Array.isArray(keyCol)) {
          keyCol.forEach((col) => {
            changes[col] = row1[col];
          });
        } else if (keyCol) {
          changes[keyCol] = row1[keyCol];
        }

        modifiedRecords.push(changes);
      }
    });

    return { records: modifiedRecords, columnChanges };
  }

  findDuplicates(data, keyCol) {
    if (data.length === 0) return [];

    const keyMap = new Map();

    data.forEach((row) => {
      const key = this.createRowHash(row, keyCol);
      if (!keyMap.has(key)) {
        keyMap.set(key, []);
      }
      keyMap.get(key).push(row);
    });

    const duplicates = [];
    keyMap.forEach((rows, key) => {
      if (rows.length > 1) {
        duplicates.push(...rows);
      }
    });

    return duplicates;
  }

  compareSheets() {
    this.log(
      `Starting comparison: ${path.basename(this.file1Path)} vs ${path.basename(
        this.file2Path
      )}`
    );

    // Get sheet names
    if (!this.sheets) {
      const workbook = XLSX.readFile(this.file1Path);
      this.sheets = workbook.SheetNames;
      this.log(`Comparing all sheets: ${this.sheets.join(", ")}`);
    }

    this.sheets.forEach((sheetName) => {
      this.log(`\n${"=".repeat(60)}`);
      this.log(`Processing sheet: ${sheetName}`);
      this.log("=".repeat(60));

      try {
        const data1 = this.readSheet(this.file1Path, sheetName);
        const data2 = this.readSheet(this.file2Path, sheetName);

        if (data1.length === 0 && data2.length === 0) {
          this.log(`Sheet '${sheetName}' is empty in both files. Skipping.`);
          return;
        }

        const keyCol = this.detectKeyColumn(
          data1.length > 0 ? data1 : data2,
          sheetName
        );

        // Debug logging
        if (keyCol === null) {
          this.log("Using full row hash comparison");
          if (data1.length > 0 && data2.length > 0) {
            const sample1 = data1
              .slice(0, 2)
              .map((row) => this.createRowHash(row, keyCol));
            const sample2 = data2
              .slice(0, 2)
              .map((row) => this.createRowHash(row, keyCol));
            this.log(`Sample File1 hashes: ${sample1.join(", ")}`);
            this.log(`Sample File2 hashes: ${sample2.join(", ")}`);
          }
        }

        const newRecords = this.findNewRecords(data1, data2, keyCol);
        const deletedRecords = this.findDeletedRecords(data1, data2, keyCol);
        const { records: modifiedRecords, columnChanges } =
          this.findModifiedRecords(data1, data2, keyCol);
        const duplicates1 = this.findDuplicates(data1, keyCol);
        const duplicates2 = this.findDuplicates(data2, keyCol);

        const keyDisplay = Array.isArray(keyCol)
          ? keyCol.join(", ")
          : keyCol || "Full Row Hash";

        this.results.push({
          sheetName,
          keyColumn: keyDisplay,
          rowCountFile1: data1.length,
          rowCountFile2: data2.length,
          colCountFile1: data1.length > 0 ? Object.keys(data1[0]).length : 0,
          colCountFile2: data2.length > 0 ? Object.keys(data2[0]).length : 0,
          newRecords,
          deletedRecords,
          modifiedRecords,
          duplicatesFile1: duplicates1,
          duplicatesFile2: duplicates2,
          columnChanges,
        });

        this.log(`\nSheet '${sheetName}' Summary:`);
        this.log(`  Key Column: ${keyDisplay}`);
        this.log(`  Row Count: ${data1.length} vs ${data2.length}`);
        this.log(`  New Records: ${newRecords.length}`);
        this.log(`  Deleted Records: ${deletedRecords.length}`);
        this.log(`  Modified Records: ${modifiedRecords.length}`);
        this.log(`  Duplicates in File1: ${duplicates1.length}`);
        this.log(`  Duplicates in File2: ${duplicates2.length}`);
      } catch (error) {
        this.log(`Error processing sheet '${sheetName}': ${error.message}`);
      }
    });

    return this.results;
  }

  formatDetailRecords(records, sheetName) {
    return records
      .map((row, index) => {
        const detailsParts = [];

        Object.entries(row).forEach(([key, value]) => {
          const normalized = this.normalize(value);
          if (normalized !== "" && normalized !== "nan") {
            detailsParts.push(`${key}: ${value}`);
          }
        });

        return {
          "Sheet Name": sheetName,
          "Row Number": index + 2,
          "Full Details": detailsParts.join(" | "),
        };
      })
      .filter((record) => record["Full Details"] !== "");
  }

  async generateReport(outputPath) {
    if (this.results.length === 0) {
      this.log("No comparison results to report.");
      return null;
    }

    if (!outputPath) {
      const timestamp = new Date()
        .toISOString()
        .replace(/[:.]/g, "-")
        .slice(0, -5);
      outputPath = `comparison_report_${timestamp}.xlsx`;
    }

    this.log(`\nGenerating report: ${outputPath}`);

    const workbook = new ExcelJS.Workbook();

    // Create Validation Summary sheet
    await this.createValidationSummary(workbook);

    // Consolidate records
    const allNew = [];
    const allModified = [];
    const allDeleted = [];
    const allDuplicates = [];

    this.results.forEach((result) => {
      if (result.newRecords.length > 0) {
        allNew.push(
          ...this.formatDetailRecords(result.newRecords, result.sheetName)
        );
      }
      if (result.modifiedRecords.length > 0) {
        allModified.push(
          ...this.formatDetailRecords(result.modifiedRecords, result.sheetName)
        );
      }
      if (result.deletedRecords.length > 0) {
        allDeleted.push(
          ...this.formatDetailRecords(result.deletedRecords, result.sheetName)
        );
      }
      // Collect duplicates separately
      if (result.duplicatesFile1.length > 0) {
        const dupsWithLabel = this.formatDetailRecords(
          result.duplicatesFile1,
          result.sheetName
        ).map((rec) => ({
          ...rec,
          File: this.file1Label,
        }));
        allDuplicates.push(...dupsWithLabel);
      }
      if (result.duplicatesFile2.length > 0) {
        const dupsWithLabel = this.formatDetailRecords(
          result.duplicatesFile2,
          result.sheetName
        ).map((rec) => ({
          ...rec,
          File: this.file2Label,
        }));
        allDuplicates.push(...dupsWithLabel);
      }
    });

    // Create consolidated sheets
    if (allNew.length > 0) {
      this.createDetailSheet(workbook, "New", allNew);
    }
    if (allModified.length > 0) {
      this.createDetailSheet(workbook, "Modified", allModified);
    }
    if (allDeleted.length > 0) {
      this.createDetailSheet(workbook, "Deleted", allDeleted);
    }
    if (allDuplicates.length > 0) {
      this.createDuplicateSheet(workbook, "Duplicates", allDuplicates);
    }

    await workbook.xlsx.writeFile(outputPath);
    this.log(`Report generated successfully: ${outputPath}`);
    return outputPath;
  }

  async createValidationSummary(workbook) {
    const sheet = workbook.addWorksheet("Validation Summary");

    // Set column widths
    sheet.columns = [
      { width: 25 },
      { width: 20 },
      { width: 20 },
      { width: 60 },
    ];

    const headerStyle = {
      font: { bold: true, color: { argb: "FFFFFFFF" } },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF4472C4" },
      },
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "center", vertical: "middle" },
    };

    let currentRow = 1;

    // Title row
    const titleRow = sheet.getRow(currentRow);
    ["Tab Validation Summary", "", "", "Comments"].forEach((val, idx) => {
      titleRow.getCell(idx + 1).value = val;
      titleRow.getCell(idx + 1).style = headerStyle;
    });
    currentRow++;

    // Headers
    const headerRow = sheet.getRow(currentRow);
    ["Validations", this.file1Label, this.file2Label, ""].forEach(
      (val, idx) => {
        headerRow.getCell(idx + 1).value = val;
        headerRow.getCell(idx + 1).style = headerStyle;
      }
    );
    currentRow++;

    // Total Tabs Count
    const totalSheets = this.results.length;
    this.addValidationRow(
      sheet,
      currentRow++,
      "Total Tabs Count",
      totalSheets,
      totalSheets,
      "Count Match",
      true
    );

    // Tabs Added/Removed
    this.addYesNoRow(sheet, currentRow++, "Tabs Added", false, "No new tabs");
    this.addYesNoRow(
      sheet,
      currentRow++,
      "Tabs Removed",
      false,
      "No tabs removed"
    );

    // Process each sheet result
    this.results.forEach((result) => {
      currentRow++; // Blank row

      // Sheet title
      const sheetTitleRow = sheet.getRow(currentRow);
      [`Tab Name: ${result.sheetName}`, "", "", "Comments"].forEach(
        (val, idx) => {
          sheetTitleRow.getCell(idx + 1).value = val;
          sheetTitleRow.getCell(idx + 1).style = headerStyle;
        }
      );
      currentRow++;

      // Sub-headers
      const subHeaderRow = sheet.getRow(currentRow);
      ["Validations", this.file1Label, this.file2Label, ""].forEach(
        (val, idx) => {
          subHeaderRow.getCell(idx + 1).value = val;
          subHeaderRow.getCell(idx + 1).style = headerStyle;
        }
      );
      currentRow++;

      // Row Count
      const rowMatch = result.rowCountFile1 === result.rowCountFile2;
      this.addValidationRow(
        sheet,
        currentRow++,
        "Row Count",
        result.rowCountFile1,
        result.rowCountFile2,
        rowMatch ? "Row Count is match" : "Row Count is mismatch",
        rowMatch
      );

      // Column Count
      const colMatch = result.colCountFile1 === result.colCountFile2;
      this.addValidationRow(
        sheet,
        currentRow++,
        "Column Count",
        result.colCountFile1,
        result.colCountFile2,
        colMatch ? "Column Count is match" : "Column Count is mismatch",
        colMatch
      );

      // New Records
      const hasNew = result.newRecords.length > 0;
      this.addYesNoRow(
        sheet,
        currentRow++,
        "New Records",
        hasNew,
        hasNew
          ? `${result.newRecords.length}-New Records available in New Records tab`
          : ""
      );

      // Modified Records
      const hasModified = result.modifiedRecords.length > 0;
      this.addYesNoRow(
        sheet,
        currentRow++,
        "Modified Records",
        hasModified,
        hasModified
          ? `${result.modifiedRecords.length}-Modified Records available in Modified Record tab`
          : ""
      );

      // Deleted Records
      const hasDeleted = result.deletedRecords.length > 0;
      this.addYesNoRow(
        sheet,
        currentRow++,
        "Deleted Records",
        hasDeleted,
        hasDeleted
          ? `${result.deletedRecords.length}-Deleted Record details available in Deleted Records Data tab`
          : ""
      );

      // Duplicates
      const hasDup =
        result.duplicatesFile1.length > 0 || result.duplicatesFile2.length > 0;
      this.addYesNoRow(sheet, currentRow++, "Duplicate Records", hasDup, "");
    });
  }

  addValidationRow(sheet, rowNum, label, val1, val2, comment, isMatch) {
    const row = sheet.getRow(rowNum);
    row.getCell(1).value = label;
    row.getCell(2).value = val1;
    row.getCell(3).value = val2;
    row.getCell(4).value = comment;

    const matchStyle = isMatch
      ? {}
      : {
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF6B6B" },
          },
          font: { color: { argb: "FFFFFFFF" } },
        };

    [2, 3].forEach((cellNum) => {
      row.getCell(cellNum).style = {
        ...matchStyle,
        border: {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        },
        alignment: { horizontal: "center" },
      };
    });

    row.getCell(1).style = {
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "center" },
    };

    row.getCell(4).style = {
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "left" },
    };
  }

  addYesNoRow(sheet, rowNum, label, hasIssue, comment) {
    const row = sheet.getRow(rowNum);
    row.getCell(1).value = label;
    row.getCell(2).value = "No";
    row.getCell(3).value = hasIssue ? "Yes" : "No";
    row.getCell(4).value = comment;

    const noStyle = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFD3D3D3" },
      },
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "center" },
    };

    const yesStyle = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF6B6B" },
      },
      font: { bold: true, color: { argb: "FFFFFFFF" } },
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "center" },
    };

    row.getCell(1).style = {
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "center" },
    };

    row.getCell(2).style = noStyle;
    row.getCell(3).style = hasIssue ? yesStyle : noStyle;

    row.getCell(4).style = {
      border: {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      },
      alignment: { horizontal: "left" },
    };
  }

  createDetailSheet(workbook, sheetName, data) {
    const sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      { header: "Sheet Name", key: "Sheet Name", width: 25 },
      { header: "Row Number", key: "Row Number", width: 12 },
      { header: "Full Details", key: "Full Details", width: 100 },
    ];

    // Style header
    sheet.getRow(1).eachCell((cell) => {
      cell.style = {
        font: { bold: true, color: { argb: "FFFFFFFF" } },
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF4472C4" },
        },
        border: {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        },
        alignment: { horizontal: "center", vertical: "middle" },
      };
    });

    // Add data
    data.forEach((record) => {
      sheet.addRow(record);
    });
  }

  createDuplicateSheet(workbook, sheetName, data) {
    const sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      { header: "File", key: "File", width: 20 },
      { header: "Sheet Name", key: "Sheet Name", width: 25 },
      { header: "Row Number", key: "Row Number", width: 12 },
      { header: "Full Details", key: "Full Details", width: 100 },
    ];

    // Style header
    sheet.getRow(1).eachCell((cell) => {
      cell.style = {
        font: { bold: true, color: { argb: "FFFFFFFF" } },
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF4472C4" },
        },
        border: {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        },
        alignment: { horizontal: "center", vertical: "middle" },
      };
    });

    // Add data
    data.forEach((record) => {
      sheet.addRow(record);
    });
  }
}

// Example usage
async function main() {
  const comparator = new ExcelComparator({
    file1Path: "SampleData.xlsx",
    file2Path: "SampleData1.xlsx",
    sheets: ["Sample Orders"],
    headerRows: {
      "Sample Orders": 0,
    },
    file1Label: "Old Version",
    file2Label: "New Version",
    keyColumns: {
      // Specify composite key if needed, otherwise auto-detection will be used
      // 'Sample Orders': ['orderdate', 'region', 'rep', 'item']
    },
  });

  comparator.compareSheets();
  await comparator.generateReport();

  console.log("\n" + "=".repeat(60));
  console.log("Comparison complete!");
  console.log("=".repeat(60));
}

// Run if executed directly
if (require.main === module) {
  main().catch(console.error);
}

module.exports = ExcelComparator;
