# Json <-> Excel form converter

A script to convert between JSON and Excel files

## Instalation

1. Clone this repository

```bash
git clone https://github.com/carolf32/json-excel-converter
```

2. Enter in the project folder

```bash
cd json-excel-converter
```

3. Install dependencies

```bash
npm install
```

## Important

- Place the JSON/Excel file you want to convert <strong> in the same folder </strong> as `convert.js`
- Run the commands <strong> from inside this folder </strong>
- If you don't know how to navigate to the place this folder is using command line:

1. Click "Show in explorer" on the Github page/Github Desktop
2. Copy the folder path (for example, `C:\Users\carolina\OneDrive\Área de Trabalho\json-excel-converter`)
3. In the terminal, type `cd "PASTE_YOUR_PATH_HERE"` (<strong> with quotes if the path contains spaces </strong>) </br>
   Example:

```bash
cd "C:\Users\carolina\OneDrive\Área de Trabalho\json-excel-converter"
```

## Dependencies

- `xlsx` for Excel file handling

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
