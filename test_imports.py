import sys
import traceback

try:
    print("Starting imports test...")
    
    import streamlit as st
    print("✓ streamlit import OK")
    
    from Sample_template import create_sample_template
    print("✓ Sample_template import OK")
    
    template_data = create_sample_template()
    print(f"✓ create_sample_template() execution OK (type: {type(template_data).__name__})")
    
    from Cal_TPH_RespTime_jtl import (
        calculate_average_throughput,
        generate_jtl_summary_sheets,
        load_jtl_data,
        load_template,
    )
    print("✓ Cal_TPH_RespTime_jtl imports OK")
    
    from DB_Server import DB_Server_Utilization
    print("✓ DB_Server import OK")
    
    from Errors import Errors
    print("✓ Errors import OK")
    
    from ExeSummary import ExecutiveSummary
    print("✓ ExeSummary import OK")
    
    from JTLPreview import show_jtl_preview
    print("✓ JTLPreview import OK")
    
    from PodService import podsServiceUtilization
    print("✓ PodService import OK")
    
    from PreTestChanges import PreTestChanges
    print("✓ PreTestChanges import OK")
    
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
    print("✓ db_manager imports OK")
    
    print("\n✅ All imports successful!")
    
except Exception as e:
    print(f"\n❌ ERROR: {e}")
    traceback.print_exc()
    sys.exit(1)
