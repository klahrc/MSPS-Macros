VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashForm 
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   OleObjectBlob   =   "SplashForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Private Sub cmdCancel_Click()
    Unload Me
    
    FileClose pjDoNotSave
End Sub


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
Private Sub cmdContinue_Click()
    Unload Me
End Sub


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
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Call MsgBox("Please click Cancel or Continue", vbOKOnly + vbExclamation, "Warning")
        Cancel = True
    End If
    
End Sub


