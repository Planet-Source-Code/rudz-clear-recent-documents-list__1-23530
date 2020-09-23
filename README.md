<div align="center">

## Clear Recent Documents List


</div>

### Description

Clears the recent documents list with a single command. Easy to implement in a text editor, or in some sort of trace-deleter program..
 
### More Info
 
place in a module, set 'Sub Main' as startup, press F5.

0 if no error occoured.

if there's _many_ entries in recent folder, have patience :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[rudz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rudz.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rudz-clear-recent-documents-list__1-23530/archive/master.zip)

### API Declarations

```
Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
```


### Source Code

```
' Name : Clear all recent documents
' By  : Rudy Alex Kohn
'   [rudyalexkohn@hotmail.com]
Public Function ClearRecent()
 ' Clear the 'Recent Document' list
 ' Returns 0 if successfull
 ClearRecent = SHAddToRecentDocs(0, 0)
End Function
Sub Main()
 If MsgBox("This will clean the 'Recent Documents', proceed?", 68, "Clear Recent Documents List") = 7 Then End
 If ClearRecent <> 0 Then MsgBox "Error.."
End Sub
```

