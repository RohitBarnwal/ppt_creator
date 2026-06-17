# Paytm & PPSL Organization Chart Creator Utility

A modern, standalone Desktop Application (using `customtkinter`) that dynamically reverse-engineers and generates beautifully styled, hierarchical organization chart PowerPoint decks from standard Excel databases.

## 🚀 Features
- **Visual Color Gradient Coding:** Automatically applies hierarchical executive color pallets (VSS Navy $\rightarrow$ HOD Medium Blue $\rightarrow$ Manager Light Blue $\rightarrow$ IC Very Light Blue) matching professional design standards.
- **Dual-Level Visualization Limit:** Restricts each sub-org sheet to 2 levels (HOD + direct reports) to maintain a highly readable layout.
- **Rollup Summaries:** Automatically detects and displays total rollup counts (recursive child counts) inside manager cards (e.g., `DR- 13`) rather than cluttering slides.
- **programmatic Elbow Routing:** Utilizes native PowerPoint Elbow Connectors that auto-snap to standard boundary ports.
- **Multi-Sheet Support:** Automatically concatenates divided worksheets (e.g., `Leads` and `Team`) into a single, unified tree database.

## 🛠️ Installation & Windows/macOS Setup
1. Download and install Python (ensure you check **"Add Python to PATH"** during install).
2. Open terminal/cmd, navigate here, and install requirements:
   ```cmd
   pip install -r requirements.txt
   ```
3. Run the desktop application:
   ```cmd
   python org_chart_app.py
   ```
