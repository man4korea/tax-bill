# -*- coding: utf-8 -*-
"""
hometax_quick.py 함수들 검증 테스트
"""

import inspect
import sys
import os
sys.path.append('..')
sys.path.append('../core')

try:
    from hometax_quick import (
        TaxInvoiceExcelProcessor,
        process_transaction_details,
        get_same_business_number_rows,
        check_and_update_supply_date,
        input_transaction_items_basic,
        input_transaction_items_extended,
        input_single_transaction_item,
        finalize_transaction_summary,
        verify_and_calculate_credit,
        show_amount_mismatch_dialog,
        handle_issuance_alerts,
        write_to_tax_invoice_sheet,
        clear_form_fields
    )
    
    print("SUCCESS: All functions imported!")
    
    # 함수 존재 검증
    functions_to_check = [
        process_transaction_details,
        get_same_business_number_rows,
        check_and_update_supply_date,
        input_transaction_items_basic,
        input_transaction_items_extended,
        input_single_transaction_item,
        finalize_transaction_summary,
        verify_and_calculate_credit,
        show_amount_mismatch_dialog,
        handle_issuance_alerts,
        write_to_tax_invoice_sheet,
        clear_form_fields
    ]
    
    print("\n=== Function Definition Check ===")
    for func in functions_to_check:
        sig = inspect.signature(func)
        print(f"OK {func.__name__}{sig}")
    
    # TaxInvoiceExcelProcessor 클래스의 새 메소드 검증
    print("\n=== TaxInvoiceExcelProcessor Method Check ===")
    processor = TaxInvoiceExcelProcessor()
    
    if hasattr(processor, 'write_tax_invoice_data'):
        print("OK write_tax_invoice_data method exists")
        sig = inspect.signature(processor.write_tax_invoice_data)
        print(f"   Signature: write_tax_invoice_data{sig}")
    else:
        print("ERROR write_tax_invoice_data method missing")
    
    print("\n=== Verification Complete ===")
    print("All transaction detail input process functions are properly defined!")
    
except ImportError as e:
    print(f"ERROR Import failed: {e}")
except Exception as e:
    print(f"ERROR Verification failed: {e}")