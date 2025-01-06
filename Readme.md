Dependencies: PsWriteHTML 
Create a Task in TaskScheduler:
1) Win + R: taskschd.msc
2) Create a task : 
    * Triggers: At long on + Delay 10 seconds (For taskbar to load)
    * Actions: Start a program: conhost.exe ,Arguments:  --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Path to Main.ps1"
    Note: Creating a headless conhost is important as it will not spawn a console window. If you create a console window then if you exit it, it will not calculate the application run time.

For time location and screen position, set it in config.json 
