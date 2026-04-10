import pandas as pd
import sys
import os

# Add project root to path
sys.path.append(os.path.abspath("."))

from utils.data_engine import DataEngine

def test_diagnostics():
    # Create dummy data with 10 rows
    data = {
        "시설구분": ["대상", "비대상", "대상", "대상", "비대상", "대상", "대상", "대상", "대상", "기타"],
        "이름": ["A", "A", "B", "C", "D", "E", "F", "G", "H", "I"]
    }
    df = pd.DataFrame(data)
    
    # 1. Test Auto Target only
    config = {"auto_target": True, "custom_filters": []}
    res, diag = DataEngine.apply_filters(df, config)
    
    print("--- Test 1: Auto Target ---")
    print(f"Initial: {diag['initial']}")
    print(f"Auto Target Removed: {diag['auto_target_removed']}")
    print(f"Final: {diag['final']}")
    
    # Valdiate expected: Initial 10, Auto target (not '대상') are 3 rows (비대상, 비대상, 기타)
    assert diag['initial'] == 10
    assert diag['auto_target_removed'] == 3
    assert diag['final'] == 7
    
    # 2. Test Custom Filter on top
    config = {
        "auto_target": True, 
        "custom_filters": [{"column": "이름", "values": ["A", "B"], "mode": "exclude"}]
    }
    res, diag = DataEngine.apply_filters(df, config)
    
    print("\n--- Test 2: Auto Target + Custom Filter ---")
    print(f"Initial: {diag['initial']}")
    print(f"Auto Target Removed: {diag['auto_target_removed']}")
    print(f"Custom Filter Removed: {diag['custom_filter_removed']}")
    print(f"Final: {diag['final']}")
    
    # Initial 10
    # Auto Target removes 3 rows -> 7 left
    # Custom filter excludes 'A', 'B' from those 7.
    # Rows left after Auto Target: (시설구분='대상') 이람=["A", "B", "C", "E", "F", "G", "H"] (Total 7)
    # Exclude A, B -> ["C", "E", "F", "G", "H"] remain (Total 5)
    # Custom Filter Removed should be 2
    assert diag['initial'] == 10
    assert diag['auto_target_removed'] == 3
    assert diag['custom_filter_removed'] == 2
    assert diag['final'] == 5

    print("\nDiagnostics logic verified successfully!")

if __name__ == "__main__":
    test_diagnostics()
