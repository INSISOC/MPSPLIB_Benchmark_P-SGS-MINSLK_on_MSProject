# VBA Script – P-SGS/MINSLK Heuristic for Microsoft Project

## Info

This script provides the full implementation of the P-SGS/MINSLK heuristic in Microsoft Visual Basic for Applications (VBA). The script is intended to be embedded in Microsoft Project and can be executed directly by the user. It performs resource leveling for multi-project environments using a general-purpose, priority-rule-based scheduling strategy.

The heuristic implemented here was originally proposed in:

- Villafáñez, F. A., Poza, D., López-Paredes, A., Pajares, J., & del Olmo, R. (2019). A generic heuristic for multi-project scheduling problems with global and local resource constraints (RCMPSP). ***Soft Computing*, 23(9), 3465–3479. [https://doi.org/10.1007/s00500-017-3003-y](https://doi.org/10.1007/s00500-017-3003-y)**

If you use this script in your research, please cite both this original work and the accompanying article (see the main [README](https://github.com/INSISOC/MPSPLIB_Benchmark_P-SGS-MINSLK_on_MSProject/blob/main/README.md)).

## How to Use the Script

1. Open your Microsoft Project file.
2. Press **Alt + F11** to open the VBA editor.
3. In the menu, go to **File → Import File…** and select the file **Rescheduling_Module.bas**, available in this folder.
4. Return to Microsoft Project and run the script from the Developer tab or via the Macros dialog (**Alt + F8**).

The script will:

- Read project tasks, dependencies, and resource assignments.
- Apply the MINSLK priority rule to resolve resource overallocations.
- Update task start dates accordingly.

Additionally, you can add a user-friendly button to Microsoft Project's ribbon to launch the script directly:

1. Go to **File → Options → Customize Ribbon**.
2. Click on **Import/Export** and select **Import customization file**.
3. Navigate to the file **MSProject_customizations.exportedUI**, available in this folder, and confirm the import.

This will add a new Schedule Optimization tab to the Microsoft Project ribbon, containing a Resource Leveling Engine button that launches the script directly.

During execution, the status bar at the bottom of the Microsoft Project window shows the current progress (number of iterations and percentage of processed tasks). Once completed, it displays the total computation time in seconds.

## Note on Commercial Use

All contents of this repository are released under the **[Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)](https://creativecommons.org/licenses/by-nc/4.0/)** licence.

This means you are free to:

- **Share** — copy and redistribute the material in any medium or format.
- **Adapt** — remix, transform, and build upon the material.

**Under the following terms:**

- **Attribution** — You must give appropriate credit, provide a link to the licence, and indicate if changes were made.
- **NonCommercial** — You may not use the material for commercial purposes.

