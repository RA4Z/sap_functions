import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parents[1]))

from src.sap_functions import SAP
import pytest
from dotenv import load_dotenv
import os

load_dotenv()

database_path = r'Q:\GROUPS\BR_SC_JGS_WM_LOGISTICA\SISTEMAS_PCP\Dashboards\_WEM\Indicador_CRP_WEM\Database'

sap = SAP()
# texto = sap.get.text_at_side('Lista', 1)
# print(texto)
# node = sap.get.node()
# node.expand_selected_node()
