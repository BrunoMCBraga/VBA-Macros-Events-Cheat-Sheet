# VBA-Macros-Events-Cheat-Sheet
Cheat-Sheet with events to look out for when analysing malicious Office documents. It is focused on Excel and Word since these are the most common ways to distribute malware.

This is a spin-off of the material found on: https://bettersolutions.com (not mine). For further description of each directive, refer to the aforementioned website or MSDN.

## Word
### Application Level
* Application_DocumentBeforeClose
* Application_DocumentChange
* Application_NewDocument
* Application_Quit
* Application_WindowActivate/WindowDeactivate
* Application_WindowBeforeDoubleClick/WindowBeforeRightClick


### Document Level
* Document_New
* Document_Open
* Document_Close

### Chain of Execution
#### New Document
Application_WindowActivate->Document_New->Application_NewDocument
#### Opening a Document 
Application_WindowActivate->Document_Open->Application_DocumentOpen

#### Closing a Document
Application_DocumentBeforeClose->Document_Close->Application_WindowDeactivate

## Excel
### Application Level
* Application_NewWorkbook
* Application_SheetActivate/SheetDeactivate
* Application_SheetBeforeDoubleClick/SheetBeforeRightClick
* Application_WindowActivate/WindowDeactivate
* Application_WorkbookActivate/WorkbookDeactivate
* Application_WorkbookOpen

### Workbook Level
Workbook_Activate/Deactivate
Workbook_BeforeClose
Workbook_Open
Workbook_SheetActivate/SheetDeactivate
Workbook_WindowActivate/WindowDeactivate


### Worksheet Level
Worksheet_Activate/Deactivate
Worksheet_BeforeDoubleClick/BeforeRightClick
Worksheet_SelectionChange

## Word Auto Macros https://msdn.microsoft.com/en-us/vba/word-vba/articles/auto-macros
* AutoExec
* AutoNew
* AutoOpen
* AutoClose
* AutoExit

## Macro directives before events were introduced. Some of these are only available for Word or Excel.
* Auto_Open
* Auto_Close
* Auto_Activate
* Auto_Deactivate
* AutoNew
* AutoExec
* AutoExit

