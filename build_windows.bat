@echo off
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --onefile --windowed --name "HLI_Project_GHG_Calcs_Builder" hli_project_ghg_calcs_builder.py
pause
