"""
RISA 3D Load Case ISO View Plot Automation
-----------------------------------------
This script automates the generation of ISO view plots showing applied loads for all 
Basic Load Cases in a RISA 3D model.

Requirements:
- RISA 3D installed
- pywin32 package (pip install pywin32)
"""

import os
import time
import win32com.client
from datetime import datetime


class RISAAutomation:
    def __init__(self):
        """Initialize the RISA 3D automation interface."""
        try:
            # Connect to RISA 3D application
            self.risa = win32com.client.Dispatch("RISA3D.Application")
            print("Successfully connected to RISA 3D")
        except Exception as e:
            print(f"Failed to connect to RISA 3D: {e}")
            raise

    def get_active_model(self):
        """Get the currently open model in RISA 3D."""
        try:
            model = self.risa.ActiveModel
            if model:
                print(f"Active model: {model.FileName}")
                return model
            else:
                print("No active model found. Please open a model in RISA 3D.")
                return None
        except Exception as e:
            print(f"Error accessing active model: {e}")
            return None

    def get_basic_load_cases(self, model):
        """Get all Basic Load Cases from the model."""
        try:
            # Access the load cases in the model
            load_cases = model.GetLoadCases()
            
            # Filter to get only Basic Load Cases
            basic_load_cases = [lc for lc in load_cases if lc.Type == 0]  # Type 0 is for Basic Load Cases
            
            print(f"Found {len(basic_load_cases)} Basic Load Cases")
            return basic_load_cases
        except Exception as e:
            print(f"Error retrieving load cases: {e}")
            return []

    def set_iso_view(self, model):
        """Set the view to isometric."""
        try:
            # Set to ISO view
            model.SetIsometricView()
            # Allow time for view to update
            time.sleep(0.5)
            print("Set to ISO view")
            return True
        except Exception as e:
            print(f"Error setting ISO view: {e}")
            return False

    def show_applied_loads(self, model, show=True):
        """Toggle display of applied loads."""
        try:
            # Show applied loads (node forces, distributed loads, etc.)
            model.ShowAppliedLoads = show
            # Allow time for view to update
            time.sleep(0.5)
            print(f"Applied loads display: {'On' if show else 'Off'}")
            return True
        except Exception as e:
            print(f"Error toggling applied loads display: {e}")
            return False

    def generate_load_case_plots(self, model, output_dir):
        """Generate ISO view plots showing applied loads for all Basic Load Cases."""
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        # Get load cases
        load_cases = self.get_basic_load_cases(model)
        if not load_cases:
            return False
        
        # Set up view settings
        self.set_iso_view(model)
        self.show_applied_loads(model, True)
        
        # Generate plots for each load case
        for lc in load_cases:
            lc_name = lc.Label.replace(" ", "_")
            print(f"Processing load case: {lc_name}")
            
            try:
                # Set current load case
                model.SetCurrentLoadCase(lc.Label)
                
                # Allow time for view to update
                time.sleep(1)
                
                # Generate filename
                filename = os.path.join(output_dir, f"{lc_name}_ISO_Applied_Loads.png")
                
                # Export view as image
                model.ExportView(filename)
                print(f"  - Exported: {filename}")
            except Exception as e:
                print(f"  - Failed to export view for {lc_name}: {e}")
        
        return True

    def close_connection(self):
        """Close the connection to RISA 3D."""
        try:
            # Release COM object
            self.risa = None
            print("Connection to RISA 3D closed")
        except Exception as e:
            print(f"Error closing connection: {e}")


def main():
    # Create output directory based on current date/time
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"RISA_LoadCase_Plots_{timestamp}")
    
    print("=" * 50)
    print("RISA 3D Load Case ISO View Plot Automation")
    print("=" * 50)
    print(f"Plot outputs will be saved to: {output_dir}")
    
    # Initialize RISA automation
    risa_auto = RISAAutomation()
    
    try:
        # Get active model
        model = risa_auto.get_active_model()
        if model:
            # Generate plots
            success = risa_auto.generate_load_case_plots(model, output_dir)
            if success:
                print(f"\nPlot generation complete! All plots saved to: {output_dir}")
            else:
                print("\nPlot generation failed.")
    finally:
        # Clean up connection
        risa_auto.close_connection()


if __name__ == "__main__":
    main()