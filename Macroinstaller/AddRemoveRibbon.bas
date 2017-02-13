Attribute VB_Name = "AddRemoveRibbon"
Sub ShowMenu()

    frmToolbox.Show

End Sub

Sub AddToolboxRibbon()
    Dim ribbonXml As String

        ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
        ribbonXml = ribbonXml + "  <mso:ribbon>"
        ribbonXml = ribbonXml + "    <mso:qat/>"
        ribbonXml = ribbonXml + "    <mso:tabs>"
        ribbonXml = ribbonXml + "      <mso:tab id=""Toolbox"" label=""MSP Toolbox"" insertBeforeQ=""mso:TabFormat"">"
        ribbonXml = ribbonXml + "        <mso:group id=""TBGroup"" label=""Toolbox"" autoScale=""true"">"
        ribbonXml = ribbonXml + "          <mso:button id=""highlightManualTasks"" label=""Click to open menu"" size=""large"" "
        ribbonXml = ribbonXml + "imageMso=""ShowFrom"" onAction=""ShowMenu""/>"
        ribbonXml = ribbonXml + "        </mso:group>"
        ribbonXml = ribbonXml + "      </mso:tab>"
        ribbonXml = ribbonXml + "    </mso:tabs>"
        ribbonXml = ribbonXml + "  </mso:ribbon>"
        ribbonXml = ribbonXml + "</mso:customUI>"

    ActiveProject.SetCustomUI (ribbonXml)
End Sub

Sub RemoveToolboxRibbon()
    Dim ribbonXml As String

    ''SetCustomUI where the parameter contains an empty mso:ribbon element
    ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    ribbonXml = ribbonXml + "<mso:ribbon></mso:ribbon></mso:customUI>"

    ActiveProject.SetCustomUI (ribbonXml)
End Sub



