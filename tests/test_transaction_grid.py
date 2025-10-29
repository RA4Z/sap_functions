import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.sap_functions import SAP
import pytest
from dotenv import load_dotenv
import os

load_dotenv()

sap = SAP()

def test_transaction():
   sap.select_transaction(os.getenv("transaction_1"))

def test_insert_data_transaction():
   sap.write_text_field(os.getenv("transaction_1_field_1_name"), os.getenv("transaction_1_field_1_value"))

def test_run_transaction():
   sap.run_actual_transaction()


grid = None
def test_get_grid():
   global grid
   grid = sap.get_grid()

def test_grid_layout():
   with pytest.raises(Exception):
      exec(os.getenv("transaction_1_before_grid_layout"))
      grid.select_layout(os.getenv("transaction_1_grid_not_existant_layout"))
   exec(os.getenv("transaction_1_after_grid_layout"))
   exec(os.getenv("transaction_1_before_grid_layout"))
   grid.select_layout(os.getenv("transaction_1_grid_layout"))

def test_grid_get_content():
   content = grid.get_grid_content()
   assert type(content.get("header")).__name__ == "list"
   assert type(content.get("content")).__name__ == "list"

def test_grid_count_rows():
   rows = grid.count_rows()
   assert type(rows).__name__ == "int"

def test_grid_get_cell_value():
   cell_value = grid.get_cell_value(0, os.getenv("transaction_1_grid_column_id"))
   assert type(cell_value).__name__ == "str"

def test_grid_press_button():
   grid.press_button(os.getenv("transaction_1_grid_button"))
   exec(os.getenv("transaction_1_after_press_button"))
   grid.press_nested_button(*os.getenv("transaction_1_grid_nested_button").split(","))
   exec(os.getenv("transaction_1_after_press_button"))

def test_grid_select_actions():
   grid.select_all_content()
   grid.select_column(os.getenv("transaction_1_grid_column_id"))   
   grid.click_cell(0, os.getenv("transaction_1_grid_column_id"))
   