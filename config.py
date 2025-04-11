import os
import sys

def get_base_dir():
    """Determine the base directory of the application, even when frozen as an executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
DATAFILES_DIR = os.path.join(BASE_DIR, "DataFiles")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
LOG_DIR = os.path.join(BASE_DIR, "logs")

TEMPLATE_EXPECTED_HEADERS = [
    "Material_ID", "Description", "Buyer_Group", "Technical Text", "Standard_Material_Set", "COPIC_Number",
    "NIIN", "Manufacturing_Part_No", "Part_No", "Maturity", "ITAR", "TypeDescription", "MaterialType",
    "ExternalMaterialStatus", "StandardMaterialCategory", "TechnicalResponsibleUser", "MatSpec", "GFX",
    "NatoStockID", "ProcPack", "Weight", "Height", "Width", "Depth", "StandardMaterialClass", "CageCode",
    "ContPartNo", "ContPartName", "ContNo", "ProcClass", "EAR600", "EAR", "NonStdRtl", "Unit", "Certificate",
    "StockShelfLife", "HazardousMaterial", "Equivalentmaterial", "machined", "machiningStrategy", "BuildPhaseID",
    "LockoutLoadout", "hotwork", "UnitBlockBreak", "Flushed", "PipeInstallationTestMedium", "PipeInstallationTestPressure",
    "PipeShopTestMedium", "PipeShopTestPressure", "PipeFlushingMedium", "PipeFlushingAcceptanceCriteria",
    "PipeAdditionalTestPressCrit", "PipeAdditionalTestMedium", "AuthoringApplication", "Identifier", "Interface",
    "Inspection_Codes", "CommodityCode", "Min_Order_Qty", "Max_Order_Qty", "Supplier_ID"
]

UNWANTED_COLUMNS = [
    "ITAR", "TypeDescription", "StandardMaterialClass", "Certificate",
    "StockShelfLife", "AuthoringApplication", "Identifier", "Interface"
]