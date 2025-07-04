import os
import sys
import yaml
import pandas as pd
import pytest

# Setup path to import from src/osdag
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../src")))

from osdag.design_type.connection.fin_plate_connection import FinPlateConnection

# Load expected values from Excel
def load_expected_values(path):
    df = pd.read_excel(path, sheet_name="Fin Plate Connection", header=None)
    expected = {}
    for _, row in df.iterrows():
        try:
            test_id = str(row[1]).strip()
            if not test_id.startswith("FinPlateTest"):
                continue
            filename = f"{test_id}.osi"
            expected[filename] = {
                "weld_size": float(row[3]),
                "shear_capacity": float(row[5]),
                "bearing_capacity": float(row[7]),
                "bolt_shear": float(row[9]),
                "weld_stress": float(row[19]),
                "moment_capacity": float(row[21])
            }
        except Exception as e:
            print(f"Skipping row due to error: {e}")
            continue
    print("✅ Loaded expected data for:", list(expected.keys()))
    return expected

# Run the full design
def run_design(filepath):
    with open(filepath, "r") as f:
        data = yaml.safe_load(f)
    obj = FinPlateConnection()
    obj.set_input_values(data)
    # If your class requires a design method, call it here
    # obj.design_finplate_connection()
    return obj

# Parametrize based on all FinPlateTest*.osi files in test_cases/
@pytest.mark.parametrize("filename", [
    f for f in os.listdir(os.path.join(os.path.dirname(__file__), "test_cases"))
    if f.startswith("FinPlateTest") and f.endswith(".osi")
])
def test_fin_plate_outputs(filename):
    expected_path = os.path.join(os.path.dirname(__file__), "ExpectedOutputs.xlsx")
    if not os.path.exists(expected_path):
        pytest.skip("ExpectedOutputs.xlsx not found, skipping test.", allow_module_level=True)

    expected_data = load_expected_values(expected_path)

    assert filename in expected_data, f"{filename} not found in Excel sheet"
    expected = expected_data[filename]

    filepath = os.path.join(os.path.dirname(__file__), "test_cases", filename)
    assert os.path.exists(filepath), f"Missing OSI file: {filepath}"

    obj = run_design(filepath)

    assert round(obj.weld.weld_size, 2) == expected["weld_size"]
    assert round(obj.shear_capacity, 2) == expected["shear_capacity"]
    assert round(obj.bearing_capacity, 2) == expected["bearing_capacity"]
    assert round(obj.bolt_shear_capacity, 2) == expected["bolt_shear"]
    assert round(obj.weld_stress, 2) == expected["weld_stress"]
    assert round(obj.moment_capacity, 2) == expected["moment_capacity"]