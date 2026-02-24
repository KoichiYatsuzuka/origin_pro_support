"""
Test error handling in new_workbook method.
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
        test_path = current_dir / f'test_success_{timestamp}.opju'
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
    """Test error handling with invalid name"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.base import OriginPageGenerationError
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_error_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Test error handling by trying to create workbook with problematic name
        print('\n=== Testing error handling ===')
        try:
            # Try to create workbook with empty name (should fail)
            wb = origin.new_workbook('')
            print('UNEXPECTED: Empty name workbook created')
            origin.close()
            return False
        except OriginPageGenerationError as e:
            print(f'SUCCESS: OriginPageGenerationError raised: {e}')
            origin.close()
            return True
        except Exception as e:
            print(f'UNEXPECTED error: {e}')
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

def test_subfolder_error():
    """Test error handling in subfolder"""
    try:
        from origin_pro_support.origin_instance import OriginInstance
        from origin_pro_support.base import OriginPageGenerationError
        
        # Create instance
        current_dir = Path.cwd()
        import time
        timestamp = int(time.time())
        test_path = current_dir / f'test_subfolder_error_{timestamp}.opju'
        origin = OriginInstance(str(test_path))
        print('SUCCESS: OriginInstance created')
        
        # Create subfolder
        root_folder = origin.get_root_dir()
        subfolder = root_folder.create_folder('ErrorTest')
        
        # Test error handling in subfolder
        print('\n=== Testing subfolder error handling ===')
        try:
            # Try to create workbook with problematic name in subfolder
            wb = subfolder.new_workbook('')
            print('UNEXPECTED: Empty name workbook created in subfolder')
            origin.close()
            return False
        except OriginPageGenerationError as e:
            print(f'SUCCESS: OriginPageGenerationError raised in subfolder: {e}')
            origin.close()
            return True
        except Exception as e:
            print(f'UNEXPECTED error: {e}')
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
    print("=== Error Handling Test ===")
    print("=" * 40)
    
    results = []
    
    # Test 1: Successful creation
    results.append(test_successful_creation())
    
    print("\n" + "=" * 40)
    
    # Test 2: Error handling
    results.append(test_error_handling())
    
    print("\n" + "=" * 40)
    
    # Test 3: Subfolder error handling
    results.append(test_subfolder_error())
    
    # Summary
    print("\n" + "=" * 40)
    print("TEST SUMMARY")
    print("=" * 40)
    
    test_names = [
        "Successful creation",
        "Error handling",
        "Subfolder error handling"
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
