#  Library Management System (Excel + VBA)

##  Description
A simple Library Management System built with **Microsoft Excel** and **VBA (Visual Basic for Applications)**. It allows users to add, search, update, and delete book records, with search results displayed on a dedicated worksheet.



##  Requirements
- Microsoft Excel (2016 or later recommended)
- Macros must be enabled
- Basic knowledge of Excel and VBA is helpful



##  Project Structure

- **UserForm**: Main form interface with input fields for:
  - `Book ID`
  - `Title`
  - `Author`
  - `Category` *(Dropdown: Science, History, Literature, Technology, Philosophy)*
  - `Year`
  - `Copies Left`
  - `Price`

- **Worksheets:**
  - `Books`: Stores all book records
  - `Search`: Displays filtered search results



##  Features

- **Add Book**: Adds a new book entry if the ID is unique
- **Search Book**:
  - By `ID` (displays data in the form)
  - By `multiple filters` (displays results in the "Search" sheet)
- **Update Book**: Modifies existing book details based on the Book ID
- **Delete Book**: Deletes a record by Book ID
- **Clear Fields**: Clears the form fields for a new operation
- **Welcome Message**: Displays date/time greeting when the workbook is opened
- **Backup Reminder**: Alerts the user on Fridays to back up files



##  How to Use

1. Open the workbook and **enable macros**.
2. Use the form to:
   - Enter new book data and click `Add`
   - Search by ID or use filters to locate books
   - Edit details and click `Update`
   - Click `Delete` to remove a book
3. Search results will be shown in the `Search` sheet.
4. Use the `Clear` button to reset the form fields.



## Protection

- The `Search` sheet is automatically unprotected during search and protected again afterward using the password `readonly`.



##  Notes

- Make sure the `Books` and `Search` sheets exist in the workbook.
- IDs must be numeric and unique.
- Year, Copies, and Price fields must be numeric.
- Password for sheet protection: `readonly`
