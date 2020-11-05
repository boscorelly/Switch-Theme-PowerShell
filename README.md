# Switch between Light and Dark theme with PowerShell

It's pretty simple :
- Put the script in the folder of your choice
- Edit the settings $lat and $long (coordinates on https://sunrise-sunset.org) if you plan to use the sun position
- Edit the $ProcessList if you want to use it (usefull for Office Apps)
- Create two scheduled tasks in a folder called ChangeTheme
    - One task SetLightTheme
    - One task SetDarkTheme
- Define your trigger :
    - Any date and time you want if you plan to use auto update
    - Your desired date and time
- Set the arguments you want :
    - -Image "c:\path\to\image-without-extension"
    - -Type PNG|JPG|JPEG
    - -Style Dark|Light
    - -Update True|False (for auto update of triggers)
    - -KillProcess True|False
    - -RestartProcess True|False
- Try it
