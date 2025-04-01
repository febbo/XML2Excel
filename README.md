# XML2Excel

[![Italiano](https://img.shields.io/badge/lang-it-green.svg)](README.it.md)
[![English](https://img.shields.io/badge/lang-en-red.svg)](README.md)
[![Python](https://img.shields.io/badge/python-3.6%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

[Versione italiana](README.it.md)

This Python script converts XML files to Excel format, creating a separate worksheet for each second-level element (direct children of the root element).

## Features

- Automatically converts XML files to Excel spreadsheets
- Creates a separate Excel worksheet for each second-level element in the XML
- Automatically generates column headers based on the name attributes of column tags
- Maintains correct values for each record
- Supports hierarchical XML structure with multiple records per section

## Requirements

- Python 3.6+
- Libraries: pandas, openpyxl, xml.etree.ElementTree
- To install the necessary dependencies:

```
pip install pandas openpyxl
```

- Alternatively, you can install all dependencies from the requirements.txt file:

```
pip install -r requirements.txt
```

## Usage

- Place the XML file and Python script in the same directory
- Run the script:

```
python script.py
```

- Enter the XML file name when prompted
- The Excel file will be generated in the same directory with the same name as the XML file but with .xlsx extension

- Supported XML Structure

The script is designed to work with an XML structure like this:

```xml
<root>
    <section1>
        <record>
            <column type='data_type' name='column_name1'>value1</column>
            <column type='data_type' name='column_name2'>value2</column>
            <!-- Other fields -->
        </record>
        <record>
            <column type='data_type' name='column_name1'>value3</column>
            <column type='data_type' name='column_name2'>value4</column>
            <!-- Other fields -->
        </record>
    </section1>
    <section2>
        <record>
            <column type='data_type' name='column_name3'>value5</column>
            <column type='data_type' name='column_name4'>value6</column>
            <!-- Other fields -->
        </record>
        <!-- Other records -->
    </section2>
    <!-- Other sections -->
</root>
```

## Output

- Each worksheet will contain all records corresponding to the XML sections
- In each worksheet, the columns will be those defined by the name attribute in the column tags

## Notes

- Excel worksheet names are limited to 31 characters (Excel limitation)
- In case of duplicate worksheet names, a progressive number will be added
- If a field is empty in the XML file, the corresponding cell in the Excel sheet will be empty
- The type attribute in column tags is ignored during conversion

## Troubleshooting

If the script generates an error:

- Make sure the XML file is in the same directory as the script
- Verify that the XML file is well-formed and follows the expected structure
- Check that you have installed all the required libraries
- Verify that you have permission to write to the current directory