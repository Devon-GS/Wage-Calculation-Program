# 💼 Automated Payroll & Payslip Generator (v2.0)

A Python-based automation tool designed to streamline payroll management. This programme automatically calculates employee hours, processes tax deductions, updates your payroll Excel sheets, and generates individual payslips.

## 🚀 What's New in Version 2.0
The application has undergone a major architectural and visual overhaul:
- **Modernized GUI:** Completely redesigned using **CustomTkinter**, replacing the older layout with a sleek, responsive, and dark-mode-ready interface.
- **Class-Based Architecture:** The backend has been fully refactored into robust, Object-Oriented classes (e.g., `DatabaseManager`, `WageProcessor`, `PayrollManager`, `PayslipGenerator`). This modular approach minimizes code duplication, improves maintainability, and strictly separates the UI from the business logic.

## ✨ Features

- **Automated Time Tracking:** Calculates total hours worked based on employee rosters and clock in/out times.
- **Excel Integration:** Automatically exports and formats time data directly into your existing payroll Excel sheets.
- **Tax Calculation:** Accurately calculates tax amounts due based on total earnings.
- **Payslip Generation:** Automatically creates formatted payslips for employees (Includes support for company logos).

## 🚀 Getting Started

### Prerequisites
Make sure you have [Python](https://www.python.org/downloads/) installed on your machine.

### Installation

**1. Clone the repository and navigate to the project directory:**
*(If applicable, otherwise just navigate to the folder)*
```bash
cd your-project-folder
```

**2. Create and activate a virtual environment:**
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python -m venv venv
source venv/bin/activate
```
**3. Install Poetry (Dependency Manager):**
```bash
pip install poetry
```

**4. Install Project Dependencies:**
```bash
poetry install
```

## 💻 Usage

Once your environment is set up and activated, you can run the main script. Make sure you are in the correct directory:

```bash
python Project/main.py
```
*(Alternatively, you can `cd Project` and run `python main.py`)*

## 🗺️ Roadmap / In Progress

- [x] Minimize code by creating classes (Completed in v2.0)
- [x] Update gui to more modernized and sleek layout (Completed in v2.0)
- [ ] Implement automated backups for database data
- [ ] Export payroll summaries to PDF

## ✅ Completed Phases (v2.0 Overhaul)

### Phase 1: Analysis & Architectural Design
* **Code Audit:** Identified redundant functions, global variables, and "spaghetti code" that could be grouped into logical objects.
* **Class Mapping:** Defined a new Class Hierarchy (`DatabaseManager`, `WageProcessor`, etc.) to handle specific responsibilities, adhering to the Single Responsibility Principle.
* **UI Benchmarking:** Researched modern design languages to establish a "Sleek & Professional" style guide.

### Phase 2: Backend Refactoring (Structural Cleanup)
* **Object-Oriented Migration:** Converted procedural scripts into reusable classes.
* **Encapsulation:** Protected data by making class variables private, ensuring that the GUI only interacts with the backend through controlled logic.
* **API/Service Layer:** Separated the core logic from the UI so that the backend runs independently of the layout updates.

### Phase 3: GUI Modernization (Visual Overhaul)
* **Layout Restructuring:** Replaced fixed-size windows with responsive grids that adapt to different screen resolutions using CustomTkinter.
* **Visual Refresh:**
    * Switched to clean, sans-serif fonts.
    * Implemented a balanced palette with a focus on "Dark Mode" compatibility.
    * Added modern elements like rounded corners and generous white space to reduce cognitive load.

### Phase 4: Optimization & Deployment
* **Component Binding:** Connected the new Class Methods to the modernized UI components.
* **Performance Testing:** Verified that the "minimized code" results in faster load times and lower memory usage.

## 📝 Changelog & Fixed Issues

*Note: Patches are released periodically to address issues and add features.*

### 2026
- **25 March 2026 [v2.0 Update]:** Major UI overhaul using CustomTkinter and backend refactor into classes.
- **13 March 2026 [Patch014]:** Fixed issues 33, 34
- **09 January 2026[Patch013]:** Fixed issues 30, 31
- **05 January 2026 [Patch012]:** Fixed issues 26, 27

### 2025
- **10 December 2025 [Patch011]:** Fixed issues 25
- **27 October 2025 [Patch010]:** Fixed issue 24 *(Added company logo on payslips)*
- **24 October 2025 [Patch009]:** Fixed issues 23
- **24 October 2025 [Patch008]:** Fixed issues 18, 22
- **17 October 2025 [Patch007]:** Fixed issue 21
- **16 October 2025 [Patch006]:** Fixed issues 15, 19, 20

### 2024
- **11 November 2024:** Fixed issues 12, 14
- **08 August 2024:** Fixed issues 8, 11
- **06 August 2024:** Fixed issues 6, 10
- **30 July 2024:** Fixed issue 7
- **01 March 2024:** Fixed issues 1, 2, 3, 4
- **20 February 2024:** Fixed issues 5, 13, 16