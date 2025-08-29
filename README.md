# Json <-> Excel form converter

A script to convert between JSON and Excel files

## Instalation

1. Clone this repository

```bash
git clone https://github.com/carolf32/json-excel-converter
```

2. Install dependencies

```bash
npm install
```

## Dependencies

- [xlsx](https://www.npmjs.com/package/xlsx) for Excel file handling

## Usage

Run the script with either a JSON or Excel file:

```bash
node convert.js form_example.json
# Excel file 'form_example_output.xlsx' has been created.
node convert.js form_example.xlsx
# JSON file 'form_example_output.xlsx' has been created.
```

The script will automatically detects the input format and creates an output with "\_output" sufix and the appropriate extension

## Output files

1. Input: form_example.json -> Output: form_example_output.xlsx
2. Input: form_example.xlsx -> Output: form_example_output.json

## File formats

### Json file

The JSON file should follow this structure:

```json
{
  "sections": [
    {
      "title": { "en": "Section Title" },
      "rows": {
        "1": {
          "field_name": {
            "label": { "en": "Field Label" },
            "type": "field_type"
          }
        }
      }
    }
  ]
}
```

### Excel file

The Excel file should follow this format:

| section        | section_title | field            | label            | type                                    | options                                            |
| -------------- | ------------- | ---------------- | ---------------- | --------------------------------------- | -------------------------------------------------- |
| section number | section title | field identifier | field label text | field type(text, select, checkbox, etc) | options: comma-separated options for select fields |
