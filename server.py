from mcp.server.fastmcp import FastMCP 
import xlwings as xw
import sys ,json

mcp = FastMCP("ExcelServer")

print("Excel MCP Server starting...", file=sys.stderr)

@mcp.tool()
def read_excel() -> dict:
    """
    Read data from the Excel file.
    Returns:
        dict: Data from the Excel file in dictionary format.
    """
    try:
        app = xw.App(visible=True, add_book=False)
        book = app.books.open(r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx")
        sheet = book.sheets[0]
        
        # Get all data including headers
        data = sheet.used_range.value
        book.save()
        
        if not data or len(data) < 2:
            return {"error": "No data found in Excel file"}
        
        # Convert to dict format with headers as keys
        headers = data[0]
        rows = data[1:]
        result = {header: [row[i] for row in rows] for i, header in enumerate(headers)}
        
        book.close()
        app.quit()
        return result
        
    except Exception as e:
        return {"error": f"Error reading Excel: {str(e)}"}

@mcp.tool()
def write_excel(data: dict) -> str:
    """
    Write data to an Excel file.
    Args:
        data (dict): Data to write to Excel. Can be:
            - Simple dict: {"key": "value"}
            - Tabular dict: {"Column1": [val1, val2], "Column2": [val1, val2]}
    Returns:
        str: Confirmation message of successful write operation.
    """
    try:
        if xw.apps:
            app = xw.apps.active
            book = app.books.open(r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx")
            sheet = book.sheets[0]
        else:
            # Open Excel (it will be visible by default)
            app = xw.App(visible=True, add_book=False)
            book = app.books.open(r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx")
            sheet = book.sheets[0]

        sheet.clear()
        
        # Check if it's tabular data (dict with lists as values)
        if data and isinstance(next(iter(data.values())), list):
            # Convert to list of lists format for xlwings
            headers = list(data.keys())
            values = list(data.values())
            # Transpose: zip creates rows from columns
            rows = [headers] + list(zip(*values))
            sheet.range('A1').value = rows
        else:
            # Simple dict
            sheet.range('A1').value = list(data.items())
        
        return json.dumps({
            "status": "success",
            "message": "Data written to Excel. Review the file and confirm if you want to save.",
            "next_action": "Call 'save_excel' tool if user confirms save or Call 'discard_changes' tool to discard changes if user confirms not to save or Continue with other instructions as given by the user."
        })

    except Exception as e:
        return json.dumps({
            "status": "error",
            "message": str(e)
        })
    
@mcp.tool()    
def save_excel() -> str:
    """
    Saves and closes the currently open Excel workbook after user confirmation.
    Returns:
        str: Confirmation message after saving and closing the workbook.
    """
    try:
        app_excel = xw.apps.active  # Connect to running Excel instance
        book = None
        for wb in app_excel.books:
            if wb.fullname == (r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx"):
                book = wb
                break

        if not book:
            return json.dumps({
                "status": "error",
                "message": "No open workbook found. Please write data first before saving."
            })

        # Save and close the workbook
        book.save()
        book.close()
        app_excel.quit()

        return json.dumps({
            "status": "success",
            "message": "Excel file saved and closed successfully."
        })
    except Exception as e:
        return json.dumps({
            "status": "error",
            "message": f"Failed to save/close Excel: {str(e)}"
        })


@mcp.tool()
def discard_changes() -> str:
    """Close the Excel workbook without saving changes."""
    try:
        app_excel = xw.apps.active
        book = None
        for wb in app_excel.books:
            if wb.fullname == (r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx"):
                book = wb
                break

        if not book:
            return json.dumps({
                "status": "error",
                "message": "No open workbook found."
            })

        book.close(save=False)
        app_excel.quit()

        return json.dumps({
            "status": "success",
            "message": "Workbook closed without saving changes."
        })
    except Exception as e:
        return json.dumps({
            "status": "error",
            "message": f"Failed to discard changes: {str(e)}"
        })
    
@mcp.tool()
def apply_formula(formula: str, cell: str) -> str:
    """
    Apply a formula to a specific cell in the Excel file.
    Args:
        formula (str): The formula to apply (e.g., '=SUM(A1:A10)').
        cell (str): The cell address where the formula should be applied (e.g., "A11").
    Returns:
        str: Confirmation message after applying the formula.
    """
    try:
        if xw.apps:
            app = xw.apps.active
            book = app.books.open(r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx")
            sheet = book.sheets[0]
        else:
            # Open Excel (it will be visible by default)
            app = xw.App(visible=True, add_book=False)
            book = app.books.open(r"C:\Users\PRITHVI RAJ\Downloads\Data-for-Practice.xlsx")
            sheet = book.sheets[0]
        
        # Apply the formula to the specified cell
        sheet.range(cell).formula = formula

    
        return json.dumps({
            "status": "success",
            "message": "Formula {formula} applied to cell {cell} successfully. Review the file and confirm if you want to save.",
            "next_action": "Call 'save_excel' tool if user confirms save or Call 'discard_changes' tool to discard changes if user confirms not to save or Continue with other instructions as given by the user."
        })    
        
        
    except Exception as e:
        return json.dumps({
            "status": "error",
            "message": f"Error applying formula: {str(e)}"
        })

if __name__ == "__main__":
    print("Running MCP server...", file=sys.stderr)
    mcp.run(transport='stdio')