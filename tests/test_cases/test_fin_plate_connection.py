import pytest
import yaml
from osdag.design_type.connection.fin_plate_connection import FinPlateConnection
import pandas as pd
import glob
import os


if not (os.path.exists('tests/test_cases/FinPlateTest1.osi') and os.path.exists('tests/Osdag Data for Unit Tests.xlsx')):
    pytest.skip('Required OSI or Excel file missing, skipping test.', allow_module_level=True)

def get_expected_from_excel(excel_path, osi_filename):
    df = pd.read_excel(excel_path)
    # Find the row matching the OSI file name
    row = df[df['OSI File Name'] == osi_filename].iloc[0]
    # Map Excel columns to output fields (update as needed)
    expected = {
        'Plate.Thickness': row['Gusset Plate Thickness Value'],
        'Plate.Height': row['Gusset Plate Min Height Value'],
        'Plate.Length': row['Gusset Plate Length Value'],
        'Weld.Size': row['Size of Weld Value'],
        'Weld.Strength': row['Weld Strength Value'],
        'Weld.Stress': row['Weld Strength Value'],
        
    }
    return expected

def extract_results_from_output(output):
    # Convert output list of tuples to a dict for easy comparison
    return {k: v for (k, _, _, v, *_) in output if k}

@pytest.mark.parametrize("osi_path", sorted(glob.glob("tests/test_cases/FinPlateTest*.osi")))
def test_fin_plate_connection_vs_excel(osi_path):
    osi_filename = os.path.basename(osi_path)
    excel_path = "tests/Osdag Data for Unit Tests.xlsx"
    osi_data = yaml.safe_load(open(osi_path))
    conn = FinPlateConnection()
    conn.set_input_values(osi_data)
    output = conn.output_values(flag=True)
    results = extract_results_from_output(output)
    expected = get_expected_from_excel(excel_path, osi_filename)
    for key, exp_val in expected.items():
        assert key in results, f"Missing result for {key}"
        # Use tolerance for float comparison
        if isinstance(exp_val, float):
            assert abs(results[key] - exp_val) < 1e-2, f"Mismatch for {key}: {results[key]} != {exp_val}"
        else:
            assert results[key] == exp_val, f"Mismatch for {key}: {results[key]} != {exp_val}"

def test_fin_plate_connection():
    # 1. Load OSI file
    with open('tests/test_cases/FinPlateTest1.osi', 'r') as f:
        osi_data = yaml.safe_load(f)
        print("OSI Data:", osi_data)  # Debug print

    # 2. Create FinPlateConnection instance
    conn = FinPlateConnection()

    # 3. Set input values (mapping is handled in set_input_values)
    conn.set_input_values(osi_data)

    # 4. Assert bolt slip factor from OSI file
    assert osi_data['Bolt.Slip_Factor'] == '0.3'

    # 6. Check output fields are populated
    output = conn.output_values(flag=True)
    assert any(o[0] == 'Bolt.Diameter' and o[3] for o in output)
    assert any(o[0] == 'Bolt.Grade_Provided' and o[3] for o in output)
    assert any(o[0] == 'Bolt.Shear' and o[3] for o in output)
    assert any(o[0] == 'Bolt.Bearing' and o[3] for o in output)
    assert any(o[0] == 'Bolt.Capacity' and o[3] for o in output)
    assert any(o[0] == 'Plate.Thickness' and o[3] for o in output)
    assert any(o[0] == 'Plate.Height' and o[3] for o in output)
    assert any(o[0] == 'Plate.Length' and o[3] for o in output)
    assert any(o[0] == 'Weld.Size' and o[3] for o in output)
    assert any(o[0] == 'Weld.Strength' and o[3] for o in output)
    assert any(o[0] == 'Weld.Stress' and o[3] for o in output)

    # 7. Check CAD generation (mocked or real)
    try:
        if hasattr(conn, 'get_3d_components'):
            conn.get_3d_components()
        cad_success = True
    except Exception as e:
        cad_success = False
    assert cad_success, 'CAD generation failed (OCC dependency or other error)'

    # 8. Check report generation (mocked or real)
    try:
        if hasattr(conn, 'save_design'):
            conn.save_design(popup_summary=False)
        report_success = True
    except Exception as e:
        report_success = False
    assert report_success, 'Report generation failed (LaTeX or other error)'

    if not os.path.exists('osi_files/TensionBoltedTest1.osi'):
        pytest.skip("Missing OSI file, skipping test.")

def load_expected_values(path):
    df = pd.read_excel(path)
    expected = {
        # Plate Section
        'Plate.Thickness': df['Gusset Plate Thickness Value'],
        'Plate.Height': df['Gusset Plate Min Height Value'],
        'Plate.Length': df['Gusset Plate Length Value'],
        
        # Weld Section
        'Weld.Size': df['Size of Weld Value'],
        'Weld.Strength': df['Weld Strength Value'],
        'Weld.Stress': df['Weld Strength Value'],
        
        # Bolt Section
        'Bolt.Diameter': df['Bolt Diameter Value'],
        'Bolt.Grade': df['Bolt Grade Value'],
        'Bolt.Shear': df['Bolt Shear Capacity Value'],
        'Bolt.Bearing': df['Bolt Bearing Capacity Value'],
        'Bolt.Capacity': df['Bolt Total Capacity Value'],
        
        # Member Section
        'Member.Supported_Section.Designation': df['Supported Section Value'],
        'Member.Supporting_Section.Designation': df['Supporting Section Value'],
        
        # Load Section
        'Load.Shear': df['Shear Load Value'],
        'Load.Axial': df['Axial Load Value'],
        
        # Material Section
        'Material': df['Material Grade Value'],
        'Connector.Material': df['Connector Material Value'],
        
        # Design Section
        'Design.Design_Method': df['Design Method Value'],
        
        # Detailing Section
        'Detailing.Gap': df['Gap Value'],
        'Detailing.Edge_type': df['Edge Type Value'],
        'Detailing.Corrosive_Influences': df['Corrosive Influences Value']
    }
    return expected 