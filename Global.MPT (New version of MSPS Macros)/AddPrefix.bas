Attribute VB_Name = "AddPrefix"

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
'Add prefix to selected tasks
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
Sub AddPrefix()
    Dim pj As Project
    Dim t As Task
    Dim i As Integer
    Dim sPrefix As String

    Set pj = ActiveProject

    i = 0
             
    sPrefix = InputBox("Please enter Prefix to use: ")
    If sPrefix = "" Then
        End
    End If

        
    For Each t In ActiveSelection.Tasks
        If Not t Is Nothing Then
            t.Name = sPrefix + t.Name
            i = i + 1
        End If
    Next
    
    MsgBox "Added prefix " + sPrefix + " to " + Str(i) + " Tasks"
    
    Set pj = Nothing

End Sub




