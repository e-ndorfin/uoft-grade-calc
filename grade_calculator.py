from config import ClassConfig, CategoryConfig
from excel_generator import generate_grade_file, create_mat137_calculator
from gui import GradeCalculatorApp

def main():
    """
    Main entry point for the Grade Calculator application.
    Launches the GUI interface.
    """
    app = GradeCalculatorApp()
    app.mainloop()

if __name__ == "__main__":
    main()
