# Excel-Formulas-Guide

**Overview**

This repository contains a collection of newly introduced Excel formulas that enhance data manipulation and organization. These formulas allow users to stack tables, split text, extract specific text portions, insert images, and transform data structures efficiently.

**List of New Excel Formulas**

1. =VSTACK

Stacks multiple tables into a single table in a vertical format.

2. =HSTACK

Stacks multiple tables into a single table in a horizontal format.

3. =TEXTSPLIT(VALUE, " ")

Splits text into multiple parts (e.g., first name and last name) in a column.

For rows: =TEXTSPLIT(VALUE, , " ")

4. =TEXTBEFORE(VALUE, "DELIMITER")

Extracts the text appearing before a specified delimiter.

Similar function: =TEXTAFTER(VALUE, "DELIMITER") (extracts text after the delimiter).

5. =IMAGE("LINK")

Displays an image from a specified URL directly within an Excel cell.

Image to Text Conversion

Navigate to Data → Picture → Picture from File, then select the image to convert into text.

6. =TOCOL(ARRAY)

Converts data from multiple columns into a single column (similar to VSTACK).

Similar function: =TOROW(ARRAY), which converts data into a single row.

7. =TAKE(ARRAY, ROW_NUM, COL_NUM)

Extracts a specific number of rows and columns from an array.

Similar function: =DROP(ARRAY, ROW_NUM, COL_NUM), which removes specified rows and columns.

8. =EXPAND(ARRAY, ROW_NUM, COL_NUM, "N/A")

Expands an array to a defined number of rows and columns, filling extra spaces with a specified value.

**How to Use**

1. Clone this repository:

git clone https://github.com/your-username/Excel-Formulas-Guide.git

2. Open the 9 NEW EXCEL FORMULA.txt file to explore the formulas and their usage.

**Contribution**

Feel free to contribute by adding more examples, use cases, or explanations. Fork the repository, make changes, and submit a pull request!

**License**

This project is licensed under the MIT License.

