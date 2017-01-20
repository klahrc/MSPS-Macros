Attribute VB_Name = "DeleteEmptyTasks"

' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment
' --------------------------------------------------------------
' Initial version
'


' **************************************************************
' Module Constant Declarations Follow
' **************************************************************

Option Explicit
'Author: Cesar Klahr
'Change selected tasks to ASAP
'Pending: Allow to change selected or ALL tasks


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
'
'
' Authors:      Cesar Klahr

'
' Comment
' --------------------------------------------------------------
' Initial version
'
Sub DeleteEmptyTasks()
    Dim pj As Project
    Dim t As Task
    Dim i As Integer

    Set pj = ActiveProject

    i = 0
      
    For Each t In ActiveSelection.Tasks
        If Not t Is Nothing Then
            If Left(t.Name, 3) = "(E)" Then
                i = i + 1
                t.Delete
            End If
        End If
    Next
    
    MsgBox "Deleted: " + Str(i) + " Tasks"
    
    Set pj = Nothing


End Sub



