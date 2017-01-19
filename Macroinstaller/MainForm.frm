VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Copy MI Macros to Global.MPT"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
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

Option Explicit
Option Base 0

Const MAX_NUM_MACROS = 100

Dim m_asModules(MAX_NUM_MACROS)
Dim m_asDescriptions(MAX_NUM_MACROS)
Dim m_nNumCopies

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
Sub CreateCustomTab()
    Dim strXML As String
    strXML = "<mso:customUI xmlns:mso=" _
             & """http://schemas.microsoft.com/office/2009/07/customui"">" _
             & "<mso:ribbon>" _
             & "<mso:tabs>" _
             & "<mso:tab id=""myTab"" label=""MI MSP Macros""  " _
             & "insertBeforeQ=""mso:TabView"">" _
             & "<mso:group id=""group1"" label=""MI PM Career Centre"">" _
             & "<mso:button id=""TaskDriver"" label=" & """Show what is driving the selected Task"" " & "imageMso=""EquationLinearFormat"" " & "onAction=""TaskDriversMain"" />" _
             & "<mso:button id=""TaskDriverDelete"" label=" & """Delete this tab"" " & "imageMso=""EquationLinearFormat"" " & "onAction=""DeleteCustomTab"" />" _
             & "</mso:group>" _
             & "</mso:tab>" _
             & "</mso:tabs>" _
             & "</mso:ribbon>" _
             & "</mso:customUI>"

    ActiveProject.SetCustomUI (strXML)
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
Public Sub DeleteCustomTab()
Dim strXML As String
    strXML = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com" _
                & "/office/2009/07/customui"">"
    strXML = strXML & "<mso:ribbon></mso:ribbon></mso:customUI>"
    ActiveProject.SetCustomUI (strXML)
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
Sub mpsVersion()
    'Should return an error if not connected to any MSP server
    
    URL = ActiveProject.ServerURL
    If Application.GetProjectServerVersion(URL) = pjServerVersionInfo_P10 Then
        ActiveProject.MakeServerURLTrusted
        xmlStream = Application.GetProjectServerSettings( _
                    RequestXML:="<ProjectServerSettingsRequest>" _
                                & "<AdminDefaultTrackingMethod /><AdminTrackingLocked />" _
                                & "<ProjectIDInProjectServer />" _
                                & "<ProjectManagerHasTransactions />" _
                                & "<ProjectManagerHasTransactionsForCurrentProject />" _
                                & "<TimePeriodGranularity /><GroupsForCurrentProjectManager />" _
                                & "</ProjectServerSettingsRequest>")
        MsgBox xmlStream
    Else
        MsgBox "This macro returns information from Project " _
               & "Server. Please choose 'Collaborate using Project " _
               & "Server' and specify a valid Project Server URL " _
               & "for this project in Collaboration Options (Collaborate menu)."
        Exit Sub
    End If
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
Private Sub cmdAll_Click()
    Dim i As Integer

    For i = 0 To lstMacros.ListCount - 1
        lstMacros.Selected(i) = True
    Next
    cmdOK.Enabled = True

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
Private Sub cmdNone_Click()
    Dim i As Integer

    For i = 0 To lstMacros.ListCount - 1
        lstMacros.Selected(i) = False
    Next
    cmdOK.Enabled = False

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
Private Sub cmdOK_Click()
    Dim i As Integer

    ' Inform what modules will be copied
    ' Inform existing modules will be asked to be overwritten
    ' Include Misc module ALWAYS
    ' Create Custom Tab to show macros
    ' Ensure the References needed are included

    ' Application.DisplayAlerts = False

    m_nNumCopies = 0
    With lstMacros
        For i = 0 To .ListCount - 1
            If (.Selected(i)) Then
                If InStr(UCase(.List(i)), "WIP") = 0 Then                                     ' WIP is not part of the name
                    OrganizerMoveItem Type:=pjModules, FileName:=ThisProject.FullName, ToFileName:="Global.MPT", Name:=m_asModules(i)
                    
                    ' !!!!!!!Verify if the copy was succesful!!!!!
                    m_nNumCopies = m_nNumCopies + 1
                End If
            End If
        Next
    End With

    ''''' FINALIZE THIS!!!!!
    ''''''''''''''Call CreateCustomTab

    MsgBox "Macro install is completed" + vbCrLf + vbCrLf + "   Click OK to finish", vbOKOnly
    Unload Me
    FileClose pjDoNotSave

    ' Application.DisplayAlerts = True

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
Private Sub SplashScreen()
 SplashForm.Show
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
Private Sub lstMacros_Change()
    Dim i As Integer

    lblDescription.Caption = m_asDescriptions(lstMacros.ListIndex)

    With lstMacros
        For i = 0 To lstMacros.ListCount - 1
            If .Selected(i) = True Then
                cmdOK.Enabled = True
                Exit Sub
            End If
        Next
    End With

    'If I get here, is because coulnd't find any selection at all
    cmdOK.Enabled = False

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
Private Sub lstMacros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lblDescription.Caption = m_asDescriptions(lstMacros.ListIndex)
End Sub

Private Sub lstMacros_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

        lblDescription.Caption = m_asDescriptions(lstMacros.ListIndex)

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
Private Sub UserForm_Initialize()
    Dim sWarning1 As String
    Dim sWarning2 As String
    Dim sWarning3 As String
    Dim nTaskNum As Integer
    Dim nDescIndex As Integer
    Dim vArr As Variant

    sWarning1 = "Macro Installer will install macros in your Global Template"
    sWarning2 = "Please ensure you opened Microsoft Project in " + """Computer Mode"""
    sWarning3 = "If not sure, please click on Cancel. Please note that you will have to close Microsoft Project and open it again in " + """Computer Mode"""

    ' Call mpsVersion                                                                      ' This is to validate if MSP is connected to any server

    ' Call AddReferences

    SplashForm.lblSplashWarning1.Caption = sWarning1
    SplashForm.lblSplashWarning2.Caption = sWarning2
    SplashForm.lblSplashWarning3.Caption = sWarning3

    SplashForm.Show

    With lblHeaderDesc
        .Font.Name = "Calibri"
        .Caption = "Macro Description"
        .BackColor = RGB(150, 0, 0)
        .ForeColor = vbWhite
        .Font.Size = 12
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
    End With

    With lblExplanation
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(150, 0, 0)
        .Caption = "Please select the macros you wish to copy to your Global.mpt"
    End With

    cmdOK.Enabled = False

    With lstMacros
        .Font.Name = "Calibri"
        .Font.Size = 14
        .BackColor = RGB(215, 230, 245)

        nDescIndex = -1
        For nTaskNum = 1 To ThisProject.Tasks.Count
            '''vArr = Split(ThisProject.Tasks(nTaskNum).Name, ",")
            '''.AddItem Trim(vArr(0)) + " (" + Trim(vArr(1)) + ")"
            '''m_asModules(nTaskNum - 1) = Trim(vArr(1))

            If (InStr(Trim(ThisProject.Tasks(nTaskNum).Name), "MI_")) = 1 Then
                .AddItem (Trim(ThisProject.Tasks(nTaskNum).Name))
                m_asModules(nTaskNum - 1) = (Trim(ThisProject.Tasks(nTaskNum).Name))

                If InStr(UCase(m_asModules(nTaskNum - 1)), "MI_MISCELLANEOUS") > 0 Then
                    lstMacros.Selected(nTaskNum - 1) = True
                    cmdOK.Enabled = True
                End If
                nDescIndex = nDescIndex + 1
                m_asDescriptions(nDescIndex) = Trim(ThisProject.Tasks(nTaskNum).Notes)
            Else
                MsgBox "Task: " + Trim(ThisProject.Tasks(nTaskNum).Name) + " not following Naming convention in Installer" + _
             vbCrLf + vbCrLf + "Click OK to continue", vbOKOnly
            End If
        Next
    End With

    lblDescription.Font.Name = "Calibri"
    lblDescription.Font.Size = 11

    If lstMacros.ListCount > 0 Then
        lstMacros.ListIndex = 0
        lstMacros.SetFocus
        lblDescription.Caption = m_asDescriptions(0)
    Else
        MsgBox "No macros are present in MacroInstaller" + vbCrLf + vbCrLf + " Ask for assistance", vbOKOnly
        Unload Me
        FileClose pjDoNotSave
    End If
End Sub
