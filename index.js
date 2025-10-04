const ExcelComparator = require("./excel-comparator");

const comparator = new ExcelComparator({
  file1Path: "SampleData.xlsx", // OLD file
  file2Path: "SampleData1.xlsx", // NEW file
  sheets: ["Sample Orders"],
  file1Label: "Old Version",
  file2Label: "New Version",
  keyColumns: {
    // Leave empty for auto-detection, or specify:
    "Sample Orders": ["orderdate", "region", "rep", "item"],
  },
});

comparator.compareSheets();
comparator.generateReport();
