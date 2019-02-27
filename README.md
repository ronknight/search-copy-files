# Search directory for a specific filename or a file type and copy files to new location using VBScript and batch file

###### Date Created: January 10, 2018
###### Author: Louise13w
###### Edited by: Ronknight
###### Edit Date: 01/02/2019
###### License: MIT
###### Tags: Search, Copy, Files, Images, VBScript, vbs, Batch, bat, CLI
###### Script: search-copy-files.vbs
###### Description: Copies all *-large.jpg and *.xlsx to the new location

---

## Instructions

First, You’ll start by editing the **search-copy-files.vbs** file.

1. Open the vbs file using your favorite file editor.
2. Change the **image_base_path** to your *soure-path*. Eg. Z:\Images\Newest Site Product Images\whitebackground\
3. Change the **image_target_new_location** to your *new-path*. Eg. E:\whitebackground\
4. Change **keywords** and **file-type** to your *new-keyword* and *new-file-type*.
5. Save and close *search-copy-files.vbs*.
---

Secondly, edit the **Run-Seach-Copy-Program.bat** file.

1. Open the BATCH file using your favorite file editor.
2. Update the **location** of your vbs file.
3. Save and close *Run-Seach-Copy-Program.bat*.
---

Lastly, Run the Program

1. Double cLick **Run-Seach-Copy-Program.bat** to execute the vbs program.
2. Wait for vbs to finish the task.
3. Watch as windows copies your files to your new location.