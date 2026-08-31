@echo off
echo ==============================================
echo Instalando las librerias necesarias para el robot...
echo ==============================================
python -m pip install pywin32 pyautogui pynput pyperclip
echo.
echo ==============================================
echo Ejecutando el Robot...
echo ==============================================
start pythonw robot_rma.py
exit
