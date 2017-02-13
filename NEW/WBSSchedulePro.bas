Attribute VB_Name = "WBSSchedulePro"
#If VBA7 Then
    Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpctstr As String, ByVal ulOptions As Long, ByVal samDesired As Long, phKey As Long) As Long
    Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long
    Declare PtrSafe Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, LPType As Long, LPData As Any, lpcbData As Long) As Long
#Else
    Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpctstr As String, phKey As Long) As Long
    Declare Function RegCloseKey Lib "advapi32" (ByVal HKey As Long) As Long
    Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, LPType As Long, LPData As Any, lpcbData As Long) As Long
#End If
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 260
Function GetProgramPath() As String
    Dim lngRegError As Long, HKey As Long, lngDataType As Long, lngLenPgmName As Long
    Dim strFileAssn As String, strPgmName As String
    
    strFileAssn = "WBS.Pro\shell\open\command"
    #If VBA7 Then
        lngRegError = RegOpenKeyEx(HKEY_CLASSES_ROOT, strFileAssn, 0, &H20219, HKey)
    #Else
        lngRegError = RegOpenKey(HKEY_CLASSES_ROOT, strFileAssn, HKey)
    #End If
    If lngRegError = ERROR_SUCCESS Then
        lngLenPgmName = MAX_PATH + 1
        strPgmName = String(lngLenPgmName, 0)
        lngRegError = RegQueryValueEx(HKey, "", 0, lngDataType, ByVal strPgmName, lngLenPgmName)
        GetProgramPath = Left(strPgmName, lngLenPgmName)
        RegCloseKey (HKey)
    End If
End Function
Sub GotoWBSChart()
Alerts False
On Error Resume Next

    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 12.0"
    
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 12.0"
    
    OptionsSecurityEx LegacyFileFormats:=2
    
    If Projects.Count > 0 Then
        If ActiveProject.LastSaveDate = "" Then
                response = MsgBox("This project has not been saved. Saving this project allows WBS Schedule Pro to dynamically link to it. Do you want to save the project so WBS Schedule Pro can link to it?", vbYesNo)
                If response = vbYes Then
                    FileSave
                    If ActiveProject.LastSaveDate = "" Then
                        Exit Sub 'not saved though we tried - probably cancelled by user
                    End If
                End If
        End If
    End If
    Err = 0
    
    DDEInitiate "WBSPro", "System"
    If Err > 0 Then
        ' MsgBox "WBS Schedule Pro is not running"
        szPath$ = GetProgramPath()
        ' MsgBox szPath$
        Err = 0
        Shell szPath$, 1
        If Err > 0 Then
            MsgBox "Cannot run WBS Schedule Pro. Run WBS Schedule Pro standalone to initialize Project, then repeat this operation."
            Exit Sub
        End If
    
        ' now retry the DDEInitiate
        Err = 0
        DDEInitiate "WBSPro", "System"
        If Err > 0 Then
            Err = 0
            DDEInitiate "WBSPro", "System"
            If Err > 0 Then
                MsgBox "WBS Schedule Pro not responding."
                Exit Sub
            End If
        End If
    End If
    
    szCommand$ = "[open(" + Chr$(34) + "::MSProject::" + Version() + ",0" + Chr$(34) + ")]"
    'MsgBox szCommand$
    Err = 0
    DDEExecute szCommand$, 1
    If Err > 0 Then
       MsgBox "WBS Schedule Pro is busy (perhaps showing a dialog or in Print Preview). Switch to WBS Schedule Pro, correct the problem, then repeat this operation."
    End If
    DDETerminate
End Sub

Sub GotoNetworkChart()
Alerts False
On Error Resume Next

    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 12.0"
    
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 12.0"
    
    OptionsSecurityEx LegacyFileFormats:=2
    
    If Projects.Count > 0 Then
        If ActiveProject.LastSaveDate = "" Then
                response = MsgBox("This project has not been saved. Saving this project allows WBS Schedule Pro to dynamically link to it. Do you want to save the project so WBS Schedule Pro can link to it?", vbYesNo)
                If response = vbYes Then
                    FileSave
                    If ActiveProject.LastSaveDate = "" Then
                        Exit Sub 'not saved though we tried - probably cancelled by user
                    End If
                End If
        End If
    End If
    Err = 0
    
    DDEInitiate "WBSPro", "System"
    If Err > 0 Then
        ' MsgBox "WBS Schedule Pro is not running"
        szPath$ = GetProgramPath()
        ' MsgBox szPath$
        Err = 0
        Shell szPath$, 1
        If Err > 0 Then
            MsgBox "Cannot run WBS Schedule Pro. Run WBS Schedule Pro standalone to initialize Project, then repeat this operation."
            Exit Sub
        End If
    
        ' now retry the DDEInitiate
        Err = 0
        DDEInitiate "WBSPro", "System"
        If Err > 0 Then
            Err = 0
            DDEInitiate "WBSPro", "System"
            If Err > 0 Then
                MsgBox "WBS Schedule Pro not responding."
                Exit Sub
            End If
        End If
    End If
    
    szCommand$ = "[open(" + Chr$(34) + "::MSProject::" + Version() + ",1" + Chr$(34) + ")]"
    'MsgBox szCommand$
    Err = 0
    DDEExecute szCommand$, 1
    If Err > 0 Then
       MsgBox "WBS Schedule Pro is busy (perhaps showing a dialog or in Print Preview). Switch to WBS Schedule Pro, correct the problem, then repeat this operation."
    End If
    DDETerminate
End Sub

Sub GotoTaskSheet()
Alerts False
On Error Resume Next

    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Export Table 12.0"
    
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 4.1"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 8.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 9.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 10.0"
    OrganizerDeleteItem Type:=1, FileName:=ActiveProject.FullName, Name:="Critical Tools Import Table2 12.0"
    
    OptionsSecurityEx LegacyFileFormats:=2
    
    If Projects.Count > 0 Then
        If ActiveProject.LastSaveDate = "" Then
                response = MsgBox("This project has not been saved. Saving this project allows WBS Schedule Pro to dynamically link to it. Do you want to save the project so WBS Schedule Pro can link to it?", vbYesNo)
                If response = vbYes Then
                    FileSave
                    If ActiveProject.LastSaveDate = "" Then
                        Exit Sub 'not saved though we tried - probably cancelled by user
                    End If
                End If
        End If
    End If
    Err = 0
    
    DDEInitiate "WBSPro", "System"
    If Err > 0 Then
        ' MsgBox "WBS Schedule Pro is not running"
        szPath$ = GetProgramPath()
        ' MsgBox szPath$
        Err = 0
        Shell szPath$, 1
        If Err > 0 Then
            MsgBox "Cannot run WBS Schedule Pro. Run WBS Schedule Pro standalone to initialize Project, then repeat this operation."
            Exit Sub
        End If
    
        ' now retry the DDEInitiate
        Err = 0
        DDEInitiate "WBSPro", "System"
        If Err > 0 Then
            Err = 0
            DDEInitiate "WBSPro", "System"
            If Err > 0 Then
                MsgBox "WBS Schedule Pro not responding."
                Exit Sub
            End If
        End If
    End If
    
    szCommand$ = "[open(" + Chr$(34) + "::MSProject::" + Version() + ",2" + Chr$(34) + ")]"
    'MsgBox szCommand$
    Err = 0
    DDEExecute szCommand$, 1
    If Err > 0 Then
       MsgBox "WBS Schedule Pro is busy (perhaps showing a dialog or in Print Preview). Switch to WBS Schedule Pro, correct the problem, then repeat this operation."
    End If
    DDETerminate
End Sub
