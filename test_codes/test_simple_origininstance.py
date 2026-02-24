"""
Simple test for OriginInstance refactoring validation.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_basic_import():
    """Test that OriginInstance can be imported from origin_instance.py"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        print("SUCCESS: OriginInstance imported from origin_instance.py")
        return True
    except Exception as e:
        print(f"FAILED: Import error - {e}")
        return False

def test_class_structure():
    """Test that OriginInstance class has expected methods"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Check for key methods
        required_methods = [
            'get_root_dir',
            'new_workbook', 
            'save',
            'close'
        ]
        
        missing = []
        for method in required_methods:
            if not hasattr(OriginInstance, method):
                missing.append(method)
        
        if missing:
            print(f"FAILED: Missing methods - {missing}")
            return False
        
        print("SUCCESS: All required methods present")
        return True
        
    except Exception as e:
        print(f"FAILED: Class structure check failed - {e}")
        return False

def main():
    """Run simple tests"""
    print("=== Simple OriginInstance Refactoring Test ===")
    
    results = []
    results.append(test_basic_import())
    results.append(test_class_structure())
    
    passed = sum(results)
    total = len(results)
    
    print(f"\nResults: {passed}/{total} tests passed")
    
    if passed == total:
        print("SUCCESS: OriginInstance refactoring appears functional")
    else:
        print("FAILED: Some issues detected")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
