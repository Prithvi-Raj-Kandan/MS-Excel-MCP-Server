from mcp.server.fastmcp import FastMCP 
import xlwings as xw
import sys ,json

mcp = FastMCP("ExcelServer")

print("Excel MCP Server starting...", file=sys.stderr)

@mcp.tool()
def read_excel() -> dict:
    """
    Reads sheet as a full 2D matrix preserving blank cells storing it as a nested list eg: [["Date","Sales Rep","Shift","Cost Price","Selling Price","Profit","Profit%","Commission","Weekday"], ["","ABC","Day",9845,9855,10,"","",""], ["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""]]
    Returns:
        dict: A dictionary containing the status, the nested list matrix of cell values, and the used range address.
    """

    file_path = r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx"

    try:
        if xw.apps:
            app = xw.apps.active
            book = app.books.open(file_path)
            sheet = book.sheets[0]
        else:
            # Open Excel (it will be visible by default)
            app = xw.App(visible=True, add_book=False)
            book = app.books.open(file_path)
            sheet = book.sheets[0]

        used = sheet.used_range
        used_address = used.address
        values = sheet.range(used.address).value  # Nested list with blanks preserved
        
        book.close()
        app.quit()

        return {
            "status": "success",
            "matrix": values,
            "range": used_address
        }

    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})

@mcp.tool()
def write_excel(matrix: list, used_address: str) -> str:
    """
    Overwrites entire used-range with updated matrix (a nested list of the values of cells in the used range).
    Args:
        matrix (list): a nested list of the values of cells in the used range. eg: [["Date","Sales Rep","Shift","Cost Price","Selling Price","Profit","Profit%","Commission","Weekday"], ["","ABC","Day",9845,9855,10,"","",""], ["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""],["","ABC","Day",9845,9855,10,"","",""]]
        used(str): Used range of the sheet.
    Returns:    
        str: Confirmation message of successful write operation.
    """

    file_path = r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx"
    try:
       if xw.apps:
            app = xw.apps.active
            book = app.books.open(file_path)
            sheet = book.sheets[0]
       else:
            # Open Excel (it will be visible by default)
            app = xw.App(visible=True, add_book=False)
            book = app.books.open(file_path)
            sheet = book.sheets[0]

        # Clear then rewrite
       sheet.range(used_address).value = matrix

       return json.dumps({
            "status": "success",
            "message": "Full range rewritten. Review & save."
        })

    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})
   
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
            if wb.fullname == (r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx"):
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
            if wb.fullname == (r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx"):
                book = wb
                break

        if not book:
            return json.dumps({
                "status": "error",
                "message": "No open workbook found."
            })

        book.close()
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
def apply_formula(formula: str, cell: str, data_range: str) -> str:
    """
    Apply a formula to a specific cell in the Excel file.Determine the address of the cell based on user input.
    Args:
        formula (str): The formula to apply (e.g., '=SUM(A1:A10)').
        cell (str): The cell address where the formula should be applied (e.g., "A11").The cell address should be within the provided data range.
        data_range (str): The range in Excel where data should be written (e.g., "A1" or "B4:K90").
    Returns:
        str: Confirmation message after applying the formula.
    """
    try:
        if xw.apps:
            app = xw.apps.active
            book = app.books.open(r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx")
            sheet = book.sheets[0]
        else:
            # Open Excel (it will be visible by default)
            app = xw.App(visible=True, add_book=False)
            book = app.books.open(r"C:\Users\PRITHVI RAJ\MS Excel MCP\Data-for-Practice.xlsx")
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
