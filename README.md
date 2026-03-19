# 💼 Automated Payroll & Payslip Generator

A Python-based automation tool designed to streamline payroll management. This programme automatically calculates employee hours, processes tax deductions, updates your payroll Excel sheets, and generates individual payslips.

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

- Minimize code by creating classes
- Update gui to more modernized and sleek layout

## Phase 1: Analysis & Architectural Design
*Before touching the code, we define the "blueprint" for the new structure.*

* **Code Audit:** Identify redundant functions, global variables, and "spaghetti code" that can be grouped into logical objects.
* **Class Mapping:** Define a new Class Hierarchy (e.g., `User`, `DataManager`, `UIController`) to handle specific responsibilities, adhering to the Single Responsibility Principle.
* **UI Benchmarking:** Research modern design languages (e.g., Google’s Material 3, Apple’s Glassmorphism, or Microsoft’s Fluent Design) to establish a "Sleek & Professional" style guide.

## Phase 2: Backend Refactoring (Structural Cleanup)
*The goal here is to reduce the "weight" of the codebase and improve maintainability.*

* **Object-Oriented Migration:** Convert procedural scripts into reusable classes. Use **Inheritance** to share common traits and **Methods** to eliminate duplicated logic.
* **Encapsulation:** Protect data by making class variables private, ensuring that the GUI only interacts with the backend through controlled "getters" and "setters."
* **API/Service Layer:** Separate the core logic from the UI so that the backend can run independently of the layout updates.

## Phase 3: GUI Modernization (Visual Overhaul)
*Moving away from dated, clunky layouts toward a fluid, minimalist experience.*

* **Layout Restructuring:** Replace fixed-size windows with **Responsive Grid Layouts** that adapt to different screen resolutions.
* **Visual Refresh:**
    * **Typography:** Switch to clean, sans-serif fonts (e.g., Inter, Roboto).
    * **Color Palette:** Implement a balanced palette with a focus on "Dark Mode" compatibility and high-contrast accessibility.
    * **Sleek Elements:** Add rounded corners, subtle shadows (soft depth), and generous white space to reduce cognitive load.
* **Micro-interactions:** Add smooth transitions for button clicks, hover states, and loading sequences to make the app feel "alive."

## Phase 4: Optimization & Deployment
*Fine-tuning the integration between the new classes and the new layout.*

* **Component Binding:** Connect the new Class Methods to the modernized UI components using an **Observer pattern** (ensuring the UI updates automatically when data changes).
* **Performance Testing:** Verify that the "minimized code" results in faster load times and lower memory usage.
* **User Acceptance (UAT):** Gather feedback on the new sleek layout to ensure it improves the actual workflow rather than just looking good.

## 📝 Changelog & Fixed Issues

*Note: Patches are released periodically to address issues and add features.*

### 2026
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