"""
Test error handling in new_workbook method with invalid characters.
"""
import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

def test_successful_creation():
    """Test successful workbook creation"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_success_v2_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test successful workbook creation
        print('\n=== Testing successful workbook creation ===')
        wb = origin.new_workbook('success_test')
        print(f'Workbook created: {wb}')
        
        if wb is not None:
            print('SUCCESS: Workbook created successfully!')
        else:
            print('UNEXPECTED: Workbook is None')
        
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

def test_error_handling():
    """Test error handling with invalid characters"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.base import OriginPageGenerationError
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_error_v2_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test error handling by trying to create workbook with invalid name
        print('\n=== Testing error handling with invalid characters ===')
        try:
            # Try to create workbook with invalid characters (should fail)
            wb = origin.new_workbook('test/workbook')  # Contains slash
            print(f'UNEXPECTED: Invalid name workbook created: {wb}')
            origin.close()
            return False
        except OriginPageGenerationError as e:
            print(f'SUCCESS: OriginPageGenerationError raised: {e}')
            origin.close()
            return True
        except Exception as e:
            print(f'Other error (might be expected): {e}')
            # Check if it's a different type of error that indicates failure
            if 'failed' in str(e).lower() or 'error' in str(e).lower():
                print('SUCCESS: Some form of error was raised')
                origin.close()
                return True
            else:
                print('UNEXPECTED: No clear error indication')
                origin.close()
                return False
        
    except Exception as e:
        print(f'ERROR: {e}')
        if 'origin' in locals():
            try:
                origin.close()
            except:
                pass
        return False

def test_duplicate_name():
    """Test error handling with duplicate name"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.base import OriginNameConflictError
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_duplicate_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test duplicate name handling
        print('\n=== Testing duplicate name handling ===')
        try:
            # Create first workbook
            wb1 = origin.new_workbook('duplicate_test')
            print(f'First workbook: {wb1}')
            
            # Try to create another with same name
            wb2 = origin.new_workbook('duplicate_test')
            print(f'UNEXPECTED: Duplicate name workbook created: {wb2}')
            origin.close()
            return False
        except OriginNameConflictError as e:
            print(f'SUCCESS: OriginNameConflictError raised: {e}')
            origin.close()
            return True
        except Exception as e:
            print(f'Other error: {e}')
            origin.close()
            return False
        
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
    print("=== Error Handling Test V2 ===")
    print("=" * 45)
    
    results = []
    
    # Test 1: Successful creation
    results.append(test_successful_creation())
    
    print("\n" + "=" * 45)
    
    # Test 2: Error handling with invalid characters
    results.append(test_error_handling())
    
    print("\n" + "=" * 45)
    
    # Test 3: Duplicate name handling
    results.append(test_duplicate_name())
    
    # Summary
    print("\n" + "=" * 45)
    print("TEST SUMMARY")
    print("=" * 45)
    
    test_names = [
        "Successful creation",
        "Invalid characters handling",
        "Duplicate name handling"
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
