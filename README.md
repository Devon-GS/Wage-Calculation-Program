# 💼 Automated Payroll & Payslip Generator (v2.0.170)

A Python-based automation tool designed to streamline payroll management. This programme automatically calculates employee hours, processes tax deductions, updates your payroll Excel sheets, and generates individual payslips.

## ⚠️ Important Notice: Branch Changes

Please note that the branch structure for this repository has recently been updated:

* **New Default Branch (`main`)**: The branch previously known as `refactor` has been renamed to `main` and is now set as the default branch. This contains the latest refactored code.

* **Original Code (`original-program`)**: The old `main` branch, which contains the original version of the program, has been renamed to `original-program`. 

If you are looking for the original version of the program, you will need to switch to the `original-program` branch. 

You can do this locally by running:
```bash
git checkout original-program
```

## 🚀 What's New in Version 2.0
 **📢 Current Status:** We have officially moved to the testing phase for v2. We are actively evaluating the system and fixing bugs.
 
The application has undergone a major architectural and visual overhaul:
- **Modernized GUI:** Completely redesigned using **CustomTkinter**, replacing the older layout with a sleek, responsive, and dark-mode-ready interface.
- **Architecture:** The backend has been fully refactored into a robust **structure**, eliminating redundancy through **centralized logic** and **reusable code patterns**.

## ✨ Features

- **Automated Time Tracking:** Calculates total hours worked based on employee rosters and clock in/out times.
- **Excel Integration:** Automatically exports and formats time data directly into your existing payroll Excel sheet.
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

---
## 🗺️ Roadmap / In Progress

- **Please see** [TODO.md](TODO.md)

---
## 🛠️ Phases (v2.0 Overhaul)

### Phase 1: Analysis & Architectural Design
* **Code Audit:** Identified redundant functions, global variables, and "spaghetti code" that could be grouped into logical objects.
* **Modular Architecture:** Organized core logic into dedicated modules to ensure a clear Separation of Concerns. This structure follows the **Single Responsibility Principle**, making the codebase easier to test and maintain.
* **UI Benchmarking:** Researched modern design languages to establish a "Sleek & Professional" style guide.

### Phase 2: Backend Refactoring (Structural Cleanup)
* **Encapsulation:** Protected data by making class variables private, ensuring that the GUI only interacts with the backend through controlled logic.
* **API/Service Layer:** Separated the core logic from the UI so that the backend runs independently of the layout updates.

### Phase 3: GUI Modernization (Visual Overhaul)
* **Layout Restructuring:** Replaced fixed-size windows with responsive grids that adapt to different screen resolutions using CustomTkinter.
* **Visual Refresh:**
    * Switched to clean, sans-serif fonts.
    * Implemented a balanced palette with a focus on "Dark Mode" compatibility.
    * Added modern elements like rounded corners and generous white space to reduce cognitive load.

### Phase 4: Optimization & Deployment
* **Component Binding:** Connected the new functions to the modernized UI components.
* **Performance Testing:** Verified that the "minimized code" results in faster load times and lower memory usage.

---
## 🤖 Phases (v3.0 Testing and AI)

### Phase 1: v2.0 Testing & Error Resolution
* **System Evaluation:** Conducted comprehensive testing on the v2.0 release to identify functional anomalies, edge cases, and unexpected behaviors.
* **Bug Fixing:** Debugged and deployed targeted patches for all identified errors to ensure a stable baseline before further optimization.

### Phase 2: AI-Driven Optimization (v3.x)
* **Algorithmic Correction:** Processed the stabilized v2.0 codebase through AI tools to detect deeper inefficiencies and automatically correct structural flaws.
* **Runtime Minimization:** Minimized overall code footprint and streamlined functions to significantly reduce execution times and resource consumption.

### Phase 3: v3.x Validation & Accuracy Testing
* **Regression Testing:** Rigorously tested the newly optimized v2.1 build to ensure no new bugs were introduced during the AI refactoring process.
* **Accuracy Verification:** Validated all system outputs to confirm that the minimized code maintains strict functional accuracy and reliability.
---