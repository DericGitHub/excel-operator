# Excel-operator User Manual
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/Window.png)
## File Operation
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/File%20Operation.png)
* **Open**  
If the file has not been modified, pop up a file selection dialog for open function.  
If the file has been modified, pop up an attention dialog first. You could choose save, save as or don't save here. Then pop up a file selection dialog for open function.  
* **Save**  
Save the file with it's own path.
* **Saveas**  
Pop up a file selection dialog, save file with the given file name.
* **Undo**  
Revert the last action applied on file. 10 steps at most.
* **File name**  
The opened file name displayed in the file name text box.
* **Sheet selection**  
Once a file is opened, the sheet selection combo box will be filled up by the sheet names.  
User is able to change the current sheet by choosing item in the sheet selection combo box.
## Comparison
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/Comparison.png)
* **Comparison start**  
The comparison will be automatically executed when the current sheet switches to another sheet or the content of current sheet has been modified.  
* **Section A**  
The **Section A** is used to show the outdated xml names in PS file which won't be used in CAS file anymore.
* **Delete**  
Delete the selected items of **Section A** in PS file. If user press the button without selected any items in **Section A**, there should be a prompt to notify user.
* **Section B**  
The Section B is used to show the extra xml names in CAS file which are not recorded by PS file.
* **Append**  
Append the selected items of **Section B** to CAS file. If user press the button without selected any items in **Section B**, there should be a prompt to notify user.
* **Select all**  
The **Select all** check box below **Section A**/**Section B** is used to check/unchecked all the items in **Section A**/**Section B**.
## Sync
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/Sync.png)
* **Section A**  
The **Section A** is used to show all the headers in current PS sheet.
* **PS->CAS**  
Copy the cells of specified columns from current PS sheet and paste them to the current CAS sheet according to the xml names in current CAS sheet. If user press the button without selected any items in **Section A**, there should be a prompt to notify user.
* **Section B**  
The Section B is used to show all the headers in current CAS sheet.
* **CAS->PS**  
Copy the cells of specified columns from current CAS sheet and paste them to the current PS sheet according to the xml names in current CAS sheet. If user press the button without selected any items in **Section B**, there should be a prompt to notify user.
* **Select all**  
The **Select all** check box below **Section A**/**Section B** is used to check/unchecked all the items in **Section A**/**Section B**.
## Preview
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/Preview.png)
* **Section A**  
Show the content of four columns of current PS sheet.
* **Section B**  
Input interface for user to enter the keyword of which they are looking for.
* **Search**  
Press button to find next position which matches the searching pattern. If there is no matching cell, there should be a prompt to notify user.
* **Add**  
Append a blank row under the selected cell of **Section A**. If user press the button without selected cell **Section A**, there should be a prompt to notify user.
* **Delete**  
Delete the entire row that contains the selected cell of **Section A**. If user press the button without selected cell **Section A**, there should be a prompt to notify user.
* **Lock POR**  
Step 1:Unlock the current PS sheet.  
Step 2:Set the protection mode of all the cell in current PS sheet to unlocked.  
Step 3:Iterate all the status. If the status is "POR", set the protection mode of entire row to locked.  
Step 4:Lock the current PS sheet.
* **Preview**  
Open a new window to display the full content of current PS sheet.
## Information
![](https://raw.githubusercontent.com/DericGitHub/excel-operator/master/Information.png)
Show the progress bar and the necessary informations for user.
- - -
## Setup.ini
A configuration file located in the same directory of **Excel Operator.exe**.
### Content
```
[setup]
MAX_ROW:1000
COLOR=Orange
```
### Usage
1. Change **MAX_ROW** to limit the max row read by program while loading the worksheet.  
2. Change **COLOR** to set the color which will be used to fill up the cell.  
  Optional color:
**White**, **Black**, **Gray-25%**, **Blue-Gray**, **Blue**, **Orange**, **Gray-50%**, **Gold**, **Blue2**, **Green**.
