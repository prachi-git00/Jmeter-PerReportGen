import sys
import traceback

try:
    print("Starting imports test...")
    
    import streamlit as st
    print("[OK] streamlit import OK")
    
    from Sample_template import create_sample_template
    print("[OK] Sample_template import OK")
    
    from Cal_TPH_RespTime_jtl import (
        calculate_average_throughput,
        generate_jtl_summary_sheets,
        load_jtl_data,
        load_template,
    )
    print("[OK] Cal_TPH_RespTime_jtl imports OK")
    
    from DB_Server import DB_Server_Utilization
    print("[OK] DB_Server import OK")
    
    from Errors import Errors
    print("[OK] Errors import OK")
    
    from ExeSummary import ExecutiveSummary
    print("[OK] ExeSummary import OK")
    
    from JTLPreview import show_jtl_preview
    print("[OK] JTLPreview import OK")
    
    from PodService import podsServiceUtilization
    print("[OK] PodService import OK")
    
    from PreTestChanges import PreTestChanges
    print("[OK] PreTestChanges import OK")
    
    from db_manager import (
        add_pretest_change,
        create_project,
        delete_project,
        delete_pretest_change,
        get_all_projects,
        get_pretest_changes_filtered,
        init_database,
        update_pretest_change,
    )
    print("[OK] db_manager imports OK")
    
    print("\n[SUCCESS] All imports successful with exception handling in place!")
    
except Exception as e:
    print(f"\n[ERROR] {e}")
    traceback.print_exc()
    sys.exit(1)
