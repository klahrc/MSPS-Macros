Attribute VB_Name = "MI_Add_Task_Prefix"

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
Sub MI_Add_Task_Prefix()
    Dim pj As Project
    Dim t As Task
    Dim i As Integer
    Dim sPrefix

    Set pj = ActiveProject

    i = 0

    sPrefix = InputBox("Please enter Task Prefix: ", "Add Task Prefix")

    If Len(sPrefix) > 0 Then

        For Each t In ActiveSelection.Tasks
            If Not t Is Nothing Then
                i = i + 1
                t.Name = sPrefix + t.Name
            End If
        Next

        MsgBox "Added Prefix to: " + Str(i) + " Tasks"

    End If
    Set pj = Nothing


End Sub



