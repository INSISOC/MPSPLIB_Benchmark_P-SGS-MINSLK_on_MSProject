##How to use the script
This script provides the full implementation of the P-SGS/MINSLK heuristic in Microsoft Visual Basic for Applications (VBA). The script is intended to be embedded in Microsoft Project and can be executed directly by the user. It performs resource leveling for multi-project environments using a general-purpose, priority-rule-based scheduling strategy.
How to use the script:
1.	Open your Microsoft Project file.
2.	Press Alt + F11 to open the VBA editor.
3.	In the menu, go to File → Import File… and select the provided file Rescheduling_Module.bas [link to Github here]
4.	Return to Microsoft Project and run the script from the Developer tab or via the Macros dialog (Alt + F8). 
The script will:
•	Prompt the user to select a leveling horizon (full project or a limited number of days).
•	Read project tasks, dependencies, and resource assignments.
•	Apply the MINSLK priority rule to resolve resource overallocations.
•	Update task start dates accordingly.
Additionally, you can add a user-friendly button to Microsoft Project’s ribbon to launch the script directly:
1.	Go to File → Options → Customize Ribbon.
2.	Click on Import/Export and select Import customization file.
3.	Navigate to the provided file Project_customizations.exportedUI and confirm the import.
This VBA module enables practitioners to apply an academically validated heuristic directly within Microsoft Project, without requiring additional software or advanced programming knowledge.
The code is available as a standalone file and can be reused or adapted for different project scheduling needs.
