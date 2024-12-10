from excel_handler import ExcelHandler

def main():
    # Create an instance of ExcelHandler
    handler = ExcelHandler()
    
    # Import the Excel file
    file_path = '20241125.xlsx'
    success, message = handler.import_cluster_expense(file_path)
    
    if success:
        print("Successfully imported data:", message)
    else:
        print("Error importing data:", message)

if __name__ == '__main__':
    main()
