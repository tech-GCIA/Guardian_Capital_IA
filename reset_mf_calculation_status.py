#!/usr/bin/env python
"""
Reset MF Metrics calculation status to allow new updates
"""

import os
import sys
import django
from pathlib import Path

# Add the project directory to Python path
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

# Now we can access the views module and reset the global variable
try:
    from gcia_app.views import calculation_progress
    
    print("=== Current MF Metrics Calculation Status ===")
    print(f"Status: {calculation_progress['status']}")
    print(f"Total funds: {calculation_progress['total_funds']}")
    print(f"Processed funds: {calculation_progress['processed_funds']}")
    print(f"Successful funds: {calculation_progress['successful_funds']}")
    print(f"Current fund: {calculation_progress.get('current_fund', 'None')}")
    print(f"Error message: {calculation_progress.get('error_message', 'None')}")
    
    if calculation_progress['status'] != 'idle':
        print(f"\nISSUE FOUND: Status is '{calculation_progress['status']}' instead of 'idle'")
        print("Resetting calculation status to 'idle'...")
        
        # Reset the global variable to idle state
        calculation_progress.update({
            'status': 'idle',
            'total_funds': 0,
            'processed_funds': 0,
            'successful_funds': 0,
            'partial_funds': 0,
            'failed_funds': 0,
            'current_fund': '',
            'error_message': '',
            'log_id': None
        })
        
        print("✓ Successfully reset calculation status to 'idle'")
        print("\nNow you should be able to see the 'Update All MF Metrics' button!")
        
    else:
        print("\n✓ Status is already 'idle' - no reset needed")
        
except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc()