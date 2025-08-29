const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const inputFile = process.argv[2];

if (!inputFile) {
  console.error("Please provide a file to convert.");
  process.exit(1);
}

function convertJsonToExcel(jsonPath, outputPath) {
  const rawData = fs.readFileSync(jsonPath, "utf8");
  const jsonData = JSON.parse(rawData);

  const rows = [];

  jsonData.sections.forEach((section, index) => {
    const sectionTitle = section.title?.en || "";
    const sectionNumber = index + 1;

    const sectionRows = Array.isArray(section.rows)
      ? section.rows
      : Object.keys(section.rows || {})
          .sort((a, b) => Number(a) - Number(b))
          .map((k) => section.rows[k]);

    sectionRows.forEach((row) => {
      Object.entries(row).forEach(([fieldKey, fieldData]) => {
        const rowData = {
          section: sectionNumber,
          section_title: sectionTitle,
          field: fieldKey,
          label: fieldData.label?.en || "",
          type: fieldData.type || "",
          options: "",
        };

        if (fieldData.type === "select" && Array.isArray(fieldData.values)) {
          const optionTitles = fieldData.values
            .map((option) => option?.title?.en || "")
            .filter((title) => title !== "");
          rowData.options = optionTitles.join(", ");
        }
        rows.push(rowData);
      });
    });
  });

  const worksheet = xlsx.utils.json_to_sheet(rows, {
    header: ["section", "section_title", "field", "label", "type", "options"],
  });
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Form Data");
  xlsx.writeFile(workbook, outputPath);
  console.log(`Excel file '${outputPath}' has been created.`);
}

function convertExcelToJson(excelPath, outputPath) {
  const workbook = xlsx.readFile(excelPath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  const sectionsMap = new Map();

  data.forEach((row) => {
    const sectionNumber = Number(row.section) || 0;

    if (!sectionsMap.has(sectionNumber)) {
      sectionsMap.set(sectionNumber, {
        title: row.section_title ? { en: row.section_title } : undefined,
        rows: [],
      });
    }

    if (!row.field) return;

    const fieldObj = {
      [row.field]: {
        label: row.label ? { en: String(row.label) } : undefined,
        type: row.type || undefined,
      },
    };
    if (String(row.type).toLowerCase() === "checkbox") {
      fieldObj[row.field].mandatory = true;
    }

    if (String(row.type).toLowerCase() === "select" && row.options) {
      const options = row.options
        .split(",")
        .map((option) => option.trim())
        .filter((option) => option !== "");
      fieldObj[row.field].values = options.map((option, index) => ({
        title: { en: option },
        value: index === 0 ? "" : option,
      }));
    }
    sectionsMap.get(sectionNumber).rows.push(fieldObj);
  });

  const sections = Array.from(sectionsMap.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([, section]) => {
      const rowsObj = {};
      section.rows.forEach((row, index) => {
        rowsObj[index + 1] = row;
      });
      return {
        ...section,
        rows: rowsObj,
      };
    });
  const jsonOutput = { sections };
  fs.writeFileSync(outputPath, JSON.stringify(jsonOutput, null, 2));
  console.log(`JSON file '${outputPath}' has been created.`);
}

function main() {
  const ext = path.extname(inputFile).toLowerCase();

  if (!fs.existsSync(inputFile)) {
    console.error(`File ${inputFile} does not exist.`);
    process.exit(1);
  }

  try {
    if (ext === ".json") {
      const outputFile = inputFile.replace(/\.json$/i, "_output.xlsx");
      convertJsonToExcel(inputFile, outputFile);
    } else if (ext === ".xlsx") {
      const outputFile = inputFile.replace(/\.xlsx$/i, "_output.json");
      convertExcelToJson(inputFile, outputFile);
    } else {
      console.error(
        "Unsupported file format. Please provide a .json or .xlsx file."
      );
    }
  } catch (error) {
    console.log("Error processing file:", error.message);
    process.exit(1);
  }
}

main();

module.exports = { convertJsonToExcel, convertExcelToJson };
