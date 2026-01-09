from openpyxl import load_workbook, Workbook
from copy import copy
from src.app.excel import DataProcessing, ExcelDataLoader
import pandas as pd
import os


class ExcelProcessorCLI:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.loader = ExcelDataLoader()
        # self.raw_st = load_workbook(folder_path)
        self.raw_df = None
        self.sheet_list = []
        # self.loader = ExcelDataLoader(folder_path)

    
    def list_xlsx_files(self):
        """Return list of .xlsx file paths in data_dir"""
        if not os.path.isdir(self.folder_path):
            raise FileNotFoundError(f"Folder not found: {self.folder_path}")

        files = [
            f for f in os.listdir(self.folder_path)
            if f.lower().endswith(".xlsx")
        ]

        if not files:
            print("❌ No Excel (.xlsx) files found.")
            return []

        print("\n📂 Available Excel files:")

        for idx, filename in enumerate(files, start=1):
            print(f"\t[{idx}] - {filename}")

        return files

    def _clear_terminal(self):
        os.system("cls" if os.name == "nt" else "clear")

    def select_file(self):
        files = self.list_xlsx_files()

        if files == []:
            return None, None
        
        while True:
            choice = input("\n👉 Choose input file by number (or 'q' to cancel): ").strip()

            if choice.lower() == "q":
                return None, None

            if not choice.isdigit():
                print("⚠️ Please enter a valid number.")
                continue
            
            choice = int(choice)

            if 1 <= choice <= len(files):
                file_path = os.path.join(self.folder_path, files[choice - 1])
                print(f"✅ Selected file: {file_path}\n")
                return file_path, files[choice - 1] #Path to file, file name

            print("⚠️ Number out of range.")

        

    def single_file_processing(self, i=1):
        input_file, file_name = self.select_file()

        if not input_file:
            print("Không tìm thấy excel file")
            return
        
        self.loader.load_file(input_file)
        self.loader.show_summary()

        # Select Sheet
        while True:
            choice = input("👉 Choose excel sheet by number (or 'q' to cancel): ").strip()

            if choice.lower() == "q":
                return False
            
            if not choice.isdigit():
                print("Vui lòng nhập lại")
                continue
            
            sheet_index = int(choice)
            if 1 > sheet_index > len(self.loader.sheets):
                print("Vui lòng nhập lại")
                continue
            break   

        sheet = self.loader.load_sheet(sheet_index)
        ws_template = self.loader.load_template(sheet_index)
        self.loader.show_info(sheet)

        data = DataProcessing(file_name, [sheet], ws_template, "data/output")

        col_indexs = input("\nChọn các cột mục tiêu (cách nhau bởi khoảng trắng): ").strip()
        nums = []
        for x in col_indexs.split():
            try:
                nums.append(int(x))
            except ValueError:
                pass

        if i == 1: 
            data.delete_by_index(nums)
        elif i == 2:
            data.reorder_by_index(nums)
        elif i == 3:
            data.split_to_sheets_by_col(nums)
        elif i == 0:
            pass
        else:
            return

    def merge_files(self):
        input_files = self.list_xlsx_files()

        if input_files == []:
            print("Không tìm thấy excel file")
            return
        else:
            choise = input("\nBạn có chắc muốn gộp tất cả file trên không? (y/n): ").strip().lower()
            if choise != "y":
                return
            
        loader = ExcelDataLoader()
        loader.load_file(os.path.join(self.folder_path, input_files[0]))
        file_names = [input_files[0]]
        input_sheets = [loader.load_sheet(1)]
        ws_template = loader.load_template(1)

        for f in input_files[1:]:
            loader.load_file(os.path.join(self.folder_path, f))
            sheet = loader.load_sheet(1)
            input_sheets.append(sheet)
            file_names.append(f[:-5])  # Remove .xlsx extension
        
        print(file_names)
        
        print(len(input_sheets))
        data = DataProcessing(
            base_name="test1.xlsx", 
            raw_df=input_sheets, 
            raw_ws=ws_template, 
            raw_df_name=file_names
            )
        data.merge_table() 

    def file_processing_menu(self):
        self._clear_terminal()
        while True:
            print("""
            ------ File Processing ------
            1. Xóa cột
            2. Sắp xếp cột
            3. Tách bảng theo cột -> sheets
            4. Gộp files -> một sheet
            *. Back
            """)
            choice = input("Select option: ").strip()

            if choice == "1":
                self.single_file_processing(1)
                break
            elif choice == "2":
                self.single_file_processing(2)
                break
            elif choice == "3":
                self.single_file_processing(3)
                break
            elif choice == "4":
                self.merge_files()
                break
            elif choice == "5":
                print("Merging files to multiple sheets...")
                break
            else:
                print("❌ Invalid option")
                break
    # -----------------------------
    # CLI loop
    # -----------------------------
    def run(self):
        while True:
            print("""
            ========= Main Menu =========
            1. Xử lý file
            2. Kiểm tra file
            4. Exit
            """)
            choice = input("Select option: ").strip()

            if choice == "1":
                self.file_processing_menu()
            elif choice == "2":
                print("Save not implemented yet")
            elif choice == "3":
                print("Save not implemented yet")
            elif choice == "4":
                print("Goodbye!")
                break
            else:
                print("❌ Invalid option")



if __name__ == "__main__":
    app = ExcelProcessorCLI("data\input")
    app.run()