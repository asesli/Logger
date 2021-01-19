# Logger
Automated time logging (Temp, these files are the core files of 2 exe files)

This is the source code for automated time logging. All the filepaths will need to be re-routed. All external linkages will need to be fixed. Parameters needs to be promoted to make it .bat friendly. Currently written to be compiled in an exe. 


The logger code will try to run "process_log.exe", which is the same file as the process_log.py, compiled into an exe. 


the code essentially does the following:

- every M seconds it detects the active window and saves its name to memory.
- every R seconds it takes all the windows recorded so far, and takes the item that appears the most
- every S seconds it takes all the windows that has been recorded the most and saves this as a file. this is essentially a save file.
- every F seconds, process_log is called against all saved data and uploads it to Ftrack
