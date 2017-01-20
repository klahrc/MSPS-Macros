Attribute VB_Name = "MI_ChgSelection2ASAP"

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
Sub MI_ChgSelection2ASAP()
    Dim pj As Project
    Dim t As Task
    Dim i As Integer

    Set pj = ActiveProject

    i = 0
      
    
      
    For Each t In ActiveSelection.Tasks
        If Not t Is Nothing Then
            If t.ConstraintType <> pjASAP Then
                i = i + 1
                t.ConstraintType = pjASAP
            End If
        End If
    Next
    
    MsgBox "Changed: " + Str(i) + " Tasks to ASAP"
    
    Set pj = Nothing


End Sub



