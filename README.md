<div align="center">

## ADVANCED Form Loaded Times


</div>

### Description

Gets how many times your form was loaded
 
### More Info
 
Returns as Long


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demian-net.md)
**Level**          |Advanced
**User Rating**    |3.0 (6 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demian-net-advanced-form-loaded-times__1-3624/archive/master.zip)





### Source Code

```
Public Sub SetLoaded()
  'put this in your main forms' Load procedure
  'this will set the count
  Dim lTemp As Long, sPath As String
  lTemp& = GetLoaded&
  If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
  Open sPath$ For Output As #1
  Print #1, lTemp& + 1
  Close #1
 End Sub
 Public Function GetLoaded() As Long
  'call this to get how many times program has been loaded
  On Error Resume Next
  Dim sPath As String, sTemp As String
  If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & App.EXEName & ".tmp" Else sPath$ = App.Path & App.EXEName & ".tmp"
  Open sPath$ For Input As #1
  sTemp$ = Input(LOF(1), #1)
  Close #1
  If sTemp$ = "" Then GetLoaded& = 0 Else GetLoaded& = CLng(sTemp$)
 End Function
```

