"""
Test new_workbook method with direct execution in OriginInstance.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_direct_execution():
    """Test new_workbook method with direct execution"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance with unique name
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_direct_execution_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test new_workbook method with direct execution
        print('\n=== Testing new_workbook with direct execution ===')
        wb = origin.new_workbook('direct_test')
        print(f'new_workbook returned: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created!')
            try:
                worksheets = wb.worksheets()
                print(f'Worksheets: {worksheets}')
                
                # Test workbook methods
                print(f'Workbook name: {wb.get_name()}')
                print(f'Workbook long name: {wb.get_long_name()}')
                
            except Exception as e:
                print(f'Error testing workbook methods: {e}')
        else:
            print('FAILED: Workbook is None')
        
        origin.close()
        return wb is not None
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_with_template():
    """Test new_workbook method with template"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance with unique name
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_with_template_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test new_workbook method with template
        print('\n=== Testing new_workbook with template ===')
        wb = origin.new_workbook('template_test', 'Origin')
        print(f'new_workbook returned: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created!')
        else:
            print('FAILED: Workbook is None')
        
        origin.close()
        return wb is not None
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def main():
    """Run all tests"""
    print("=== Test Direct Execution ===")
    print("=" * 40)
    
    results = []
    
    # Test 1: Direct execution
    results.append(test_direct_execution())
    
    print("\n" + "=" * 40)
    
    # Test 2: With template
    results.append(test_with_template())
    
    # Summary
    print("\n" + "=" * 40)
    print("TEST SUMMARY")
    print("=" * 40)
    
    test_names = [
        "Direct execution",
        "With template"
    ]
    
    for i, (name, result) in enumerate(zip(test_names, results)):
        status = "PASSED" if result else "FAILED"
        print(f"{name}: {status}")
    
    passed = sum(results)
    total = len(results)
    print(f"\nOverall: {passed}/{total} tests passed")
    
    if passed == total:
        print("All tests PASSED!")
    else:
        print("Some tests FAILED")
    
    return passed == total

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
