<div align="center">

## MS Project Function Lib


</div>

### Description

These functions can be used to read MS Project tasks into an array and vice versa.

I use them within a LotusScript agent to realize an export-to-MSProject feature.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Petra Tenzlinger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/petra-tenzlinger.md)
**Level**          |Advanced
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/petra-tenzlinger-ms-project-function-lib__1-34844/archive/master.zip)





### Source Code

```
Attribute VB_Name = "mdlMSProject"
Option Explicit
' These functions are used to read/write tasks to/from an
' MS Project file.
' If you use this code from Word or Excel, don't forget to link
' Microsoft Project 8.0 Object Library to use early binding.
' I used this code within a Lotus Notes agent (LotusScript), so
' I had to use late binding. That's why you sometimes find some
' function or variable declarations in comments.
' The project start date will not be copied
'
' Author:    Petra Tenzlinger
' Date:     05/10/2002
' Copyright Petra Tenzlinger, 2002
' I used this constants within LotusScript because Notes doesn't have them.
'Const pjDoNotSave = 0
'Const pjSave = 1
'Const pjPrompSave = 2
' Describes one task which is one row in a MS Project document.
' An array of this type can describe the whole project.
Type MspTask
  Name As String     ' task name
  Start As String     ' task start date
  Finish As String    ' task end date
  ResourceNames As String ' resource name
  Level As Integer    ' task level
  isMS As Variant     ' Is task a milestone?
  isSummary As Variant  ' Is task a summary?
End Type
Function openMspDocument(ByVal strFilename As String, ByVal isVisible As Boolean, objMspApp As MSProject.Application, objMspDoc As MSProject.Project) As Boolean
'Function openMspDocument(Byval strFilename As String, Byval isVisible As Variant, objMspApp As Variant, objMspDoc As Variant, intPIdx As Integer) As Variant
' Opens given MS Project file via COM.
' Returns application and document (project).
'
' Arguments:
' Name     Type  In/Out Description
' strFilename  String In   Name and path of MS Project file (*.mpp) to open.
'                Empty string opens new file.
' isVisible   Variant In   true: MS Project Application opens visibly
'                false: MS Project opens in background
' objMspApp   Variant Out   MS Project application object
' objMspDoc   Variant Out   MS Project document object (Project)
' return value Variant Return true: no errors occured, false: an error occured
  On Error GoTo CreateObject
  Set objMspApp = GetObject(, "MSProject.Application")
  objMspApp.Visible = True
Continue:
  ' here we have an application object
  On Error GoTo Func_Err
  ' turn out the annoying message boxes
  Call objMspApp.Alerts(False)
  ' open project
  If strFilename <> "" Then
    Call objMspApp.FileOpen(strFilename)
  Else  ' new project
    Call objMspApp.FileNew
  End If
  Set objMspDoc = objMspApp.ActiveProject
  openMspDocument = True
  Exit Function
CreateObject:
  Set objMspApp = CreateObject("MSProject.Application")
  objMspApp.Visible = True
  Resume Continue
Func_Exit:
  Exit Function
Func_Err:
  MsgBox Error & " (" & Err & ") in line " & Erl
  openMspDocument = False
  Resume Func_Exit
End Function
Function saveMspDocument(objMspApp As MSProject.Application, ByVal strFilename As String, ByVal withBaseline As Boolean) As Boolean
'Function saveMspDocument(objMspApp As Variant, Byval strFilename As String, Byval withBaseline As Variant) As Variant
' Saves active project of the given MS Project application.
' If filename empty, save as current file.
'
' Arguments:
' Name     Type  In/Out Description
' objMspApp   Variant Out   Application which active project will be saved.
' strFilename  String In   File name (incl. path) to save as. Empty string if current name
'                and location should be used.
' withBaseline Variant In   True: Saves with base line, false: Saves without baseline
' return value Variant Return true: no errors occured, false: an error occured
  On Error GoTo Func_Err
  If withBaseline Then
    Call objMspApp.BaselineSave(True, 0, 0)
  End If
  If strFilename = "" Then
    ' save if changes made
    If Not objMspApp.ActiveProject.Saved Then
      If objMspApp.ActiveProject.Path <> "" Then
        Call objMspApp.FileSave
      Else
        ' cannot save
        MsgBox "No path found. Changes cannot be saved."
        saveMspDocument = False
        GoTo Func_Exit
      End If
    End If
  Else
    Call objMspApp.FileSaveAs(strFilename)
  End If
  saveMspDocument = True
Func_Exit:
  Exit Function
Func_Err:
  MsgBox Error & " (" & Err & ") in line " & Erl
  saveMspDocument = False
  Resume Func_Exit
End Function
Function closeMspDocument(objMspDoc As MSProject.Project, ByVal withSave As Boolean, ByVal quitApp As Boolean) As Boolean
'Function closeMspDocument(objMspDoc As Variant, ByVal withSave As Variant, ByVal quitApp As Variant) As Variant
' Closes given MS Project file (project).
'
' Arguments:
' Name     Type  In/Out Description
' objMspDoc   Variant Out   File (project), that will be closed.
' withSave   Variant In   true: changes will be saved, false: changes won't be saved.
' quitApp    Variant In   true: application will be quited, false: applications remains open.
' return value Variant Return true: no errors occured, false: an error occured
  Dim objMspApp As MSProject.Application
  'Dim objMspApp As Variant
  On Error GoTo Func_Err
  If objMspDoc Is Nothing Then GoTo Func_Exit
  Set objMspApp = objMspDoc.Application
  ' make project current project
  Call objMspDoc.Activate
  If withSave Then
    ' save if changes made
    If Not objMspDoc.Saved Then
      If objMspDoc.Path <> "" Then
        Call objMspApp.FileClose(pjSave)
      Else
        ' cannot save
        MsgBox "No path found. Changes cannot be saved."
        closeMspDocument = False
        GoTo Func_Exit
      End If
    End If
  Else
    ' close without saving
    Call objMspApp.FileClose(pjDoNotSave)
  End If
  Set objMspDoc = Nothing
  If quitApp Then
    ' close all other files without saving
    Call objMspApp.FileCloseAll(pjDoNotSave)
    Call objMspApp.Quit
    Set objMspApp = Nothing
  End If
  closeMspDocument = True
Func_Exit:
  Exit Function
Func_Err:
  MsgBox Error & " (" & Err & ") in line " & Erl
  closeMspDocument = False
  Resume Func_Exit
End Function
Sub setTaskLevel(objMspTask As MSProject.Task, ByVal intLevel As Integer)
'Sub setTaskLevel(objMspTask As Variant, ByVal intLevel As Integer)
' Sets level of given task to given value.
'
' Arguments:
' Name     Type  In/Out Description
' objMspTask  Variant Out   Task object to set level.
' intLevel   Integer In   Level to set.
  Dim intDiff As Integer
  Dim i As Integer
  intDiff = objMspTask.OutlineLevel - intLevel
  If intDiff > 0 Then   ' task too far right
    'Call objMspTask.OutlineOutdent(intDiff)  ' doesn't work :-(
    For i = 1 To intDiff
      Call objMspTask.OutlineOutdent
    Next
  ElseIf intDiff < 0 Then   ' task too far left
    'Call objMspTask.OutlineIndent(Abs(intDiff))
    For i = 1 To Abs(intDiff)
      Call objMspTask.OutlineIndent
    Next
  End If
End Sub
Function deleteEmptyTasks(objMspDoc As MSProject.Project) As Integer
'Function deleteEmptyTasks(objMspDoc As Variant) As Integer
' Deletes all empty tasks (without task names). They make problems!
'
' Arguments:
' Name     Type  In/Out Description
' objMspDoc   Variant Out   Project to delete empty tasks.
' return value Integer Return Number of deleted tasks.
  Const STR_VIEW = "Balkendiagramm (Gantt)"  'sorry I only know in german
  'Dim objMspApp As Variant
  Dim objMspApp As MSProject.Application
  Dim i As Integer
  Dim intNoDeleted As Integer
  Set objMspApp = objMspDoc.Application
  ' change view
  If objMspDoc.CurrentView <> STR_VIEW Then
    Call objMspApp.ViewApply(STR_VIEW)
  End If
  i = 1
  Do While i <= objMspDoc.Tasks.Count
    ' ... give empty task a name and delete it
    Call objMspApp.SelectTaskField(i, "Name", False)
    If objMspApp.ActiveCell.Text = "" Then
      Call objMspApp.SetActiveCell("@EMPTY@")
      Call objMspApp.ActiveCell.Task.Delete
      i = i - 1    ' after deletion subsequent tasks move up
      intNoDeleted = intNoDeleted + 1
    End If
    i = i + 1
  Loop
  deleteEmptyTasks = intNoDeleted
End Function
Function ArrayToMsp(objMspDoc As MSProject.Project, aTasks() As MspTask) As Integer
'Function ArrayToMsp(objMspDoc As Variant, aTasks() As MspTask) As Integer
' Imports array data into MS Project document objMspDoc.
'
' Arguments:
' Name     Type    In/Out Description
' objMspDoc   Variant   Out   MS Project document to fill.
' aTasks    Array of  In   Array of tasks to import.
'        MspTask
' return value Integer   Return Number of imported tasks.
'
' Notice:
' – the name field cannot be empty
' - the finish date will be ignored in summary tasks (it is computed by MS Project)
' – Level must be > 0
' – the start date must be before the end date
' – a milestone (isMS) cannot be a summary task (isSummary)
  Dim intI As Integer
  'Dim objMspTask As Variant
  Dim objMspTask As MSProject.Task
  On Error GoTo Func_Err
  For intI = LBound(aTasks) To UBound(aTasks)
'    ' Projektbeginn
'    If istErstesFeldProjekt And intI = LBound(aTasks) Then
'      objMspDoc.ProjectStart = aTasks(intI).Start
'    End If
    ' new task
    Set objMspTask = objMspDoc.Tasks.Add(aTasks(intI).Name)
    ' task start and end date
'    If Not istErstesFeldProjekt Or intI > LBound(aTasks) Then
      objMspTask.Start = aTasks(intI).Start
      ' finish not in summary tasks
      If Not aTasks(intI).isSummary Then
        objMspTask.Finish = aTasks(intI).Finish
      End If
'    End If
    ' milestones
    If aTasks(intI).isMS Then
      objMspTask.Milestone = True
      objMspTask.Duration = 0
    End If
        ' resource names
    objMspTask.ResourceNames = aTasks(intI).ResourceNames
    ' task level
    'If intI > LBound(aTasks) Then
      Call setTaskLevel(objMspTask, aTasks(intI).Level)
    'End If
  Next intI
  ArrayToMsp = intI + 1
Func_Exit:
  Exit Function
Func_Err:
  MsgBox Error & " (" & Err & ") in line " & Erl
  ArrayToMsp = 0
  Resume Func_Exit
End Function
Function MspToArray(objMspDoc As MSProject.Project, aTasks() As MspTask) As Integer
'Function MspToArray(objMspDoc As Variant, aTasks() As MspTask) As Integer
' Exports MS Project data into array aTasks.
'
' Arguments:
' Name     Type    In/Out Description
' objMspDoc   Variant   Out   MS Project document to read.
' aTasks    Array of  In   Array of tasks to be filled.
'        MspTask
' return value Integer   Return Number of exported tasks.
  Dim intI As Integer
  'Dim objMspTask As Variant
  Dim objMspTask As MSProject.Task
  On Error GoTo Func_Err
  ' empty tasks make problems
  Call deleteEmptyTasks(objMspDoc)
  ReDim aTasks(0 To objMspDoc.Tasks.Count - 1)
  For intI = 0 To objMspDoc.Tasks.Count - 1
    aTasks(intI).Name = objMspDoc.Tasks(intI + 1).Name
    aTasks(intI).Start = objMspDoc.Tasks(intI + 1).Start
    aTasks(intI).Finish = objMspDoc.Tasks(intI + 1).Finish
    aTasks(intI).ResourceNames = objMspDoc.Tasks(intI + 1).ResourceNames
    aTasks(intI).Level = objMspDoc.Tasks(intI + 1).OutlineLevel
    aTasks(intI).isMS = objMspDoc.Tasks(intI + 1).Milestone
    aTasks(intI).isSummary = objMspDoc.Tasks(intI + 1).Summary
  Next intI
  MspToArray = intI + 1
Func_Exit:
  Exit Function
Func_Err:
  MsgBox Error & " (" & Err & ") in line " & Erl
  MspToArray = 0
  Resume Func_Exit
End Function
Sub exampleCopyMPP()
' Copies a MS Project document into another one.
' Needs a MS Project file test.mpp in C:\temp,
' creates a MS Project file copy.mpp in C:\temp.
  Dim objMspApp As MSProject.Application
  Dim objMyProject As MSProject.Project
  Dim objNewProject As MSProject.Project
  Dim aTasks() As MspTask
  ' open project to read (import)
  Call openMspDocument("C:\temp\test.mpp", True, objMspApp, objMyProject)
  ' import data into array
  Call MspToArray(objMyProject, aTasks())
  ' open new mpp file
  Call openMspDocument("", True, objMspApp, objNewProject)
  ' export array data to new file
  Call ArrayToMsp(objNewProject, aTasks)
  ' save new project
  Call saveMspDocument(objMspApp, "C:\temp\copy.mpp", False)
  ' close new file
  Call closeMspDocument(objNewProject, False, False)
  ' close other project and application
  Call closeMspDocument(objMyProject, True, True)
End Sub
```

