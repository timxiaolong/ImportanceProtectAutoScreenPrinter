@echo off
conda init
conda create -n PrintScreen python=3.8
conda activate ScreenPrinter
pip install -r requirement.txt
python main.py
pause
