VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Project_Open(ByVal pj As Project)
    '''''''AddToolbar
  
    MainForm.Show
    
End Sub

Private Sub Project_BeforeClose(ByVal pj As Project)
    ''''''''DeleteBar
    
    
End Sub




