Attribute VB_Name = "MI_Miscellaneous"

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

' Description: v0.9 Supports Late Binding to avoid issues with missing References
Private Declare Function SHGetFolderPath Lib "shfolder.dll" _
   Alias "SHGetFolderPathA" ( _
       ByVal hwndOwner As Long, _
       ByVal nFolder As Long, _
       ByVal hToken As Long, _
       ByVal dwReserved As Long, _
       ByVal lpszPath As String) As Long

Private Const CSIDL_PERSONAL As Long = &H5


'//////////////// LATE BINDING CONSTANTS////////////////
Public Const xlValuesLB = -4163
Public Const xlWholeLB = 1
Public Const xlByRowsLB = 1
Public Const xlNextLB = 1
Public Const xlA1LB = 1
Public Const xlR1C1LB = -4150
Public Const xlUpLB = -4162
Public Const xlMaximizedLB = -4137
Public Const xlContinuousLB = 1
Public Const xlThinLB = 2
Public Const xlThickLB = 4
Public Const xlEdgeTopLB = 8
Public Const xlEdgeLeftLB = 7
Public Const xlEdgeRightLB = 10
Public Const xlEdgeBottomLB = 9
Public Const xlPasteValuesAndNumberFormatsLB = 12

Public Const xlInsideVerticalLB = 11
Public Const xlInsideHorizontalLB = 12

Public Const xlExpressionLB = 2

Public Const xlCalculationManualLB = -4135
Public Const xlCalculationAutomaticLB = -4105
Public Const xlCalculationSemiautomaticLB = 2

Public Const xlConditionValueLowestValueLB = 1
Public Const xlConditionValuePercentileLB = 5
Public Const xlConditionValueHighestValueLB = 2
'//////////////////////////////////////////////////////////


Option Explicit

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
Public Function IsNumber(ByVal Value As String) As Boolean
    Dim DP As String
    Dim TS As String
    '   Get local setting for decimal point
    DP = Format$(0, ".")
    '   Get local setting for thousand's separator and eliminate them. Remove the next two lines
    '   if you don't want your users being able to type in the thousands separator at all.
    TS = Mid$(Format$(1000, "#,###"), 2, 1)
    Value = Replace$(Value, TS, "")
    '   Leave the next statement out if you don't want to provide for plus/minus signs
    If Value Like "[+-]*" Then Value = Mid$(Value, 2)
    IsNumber = Not Value Like "*[!0-9" & DP & "]*" And Not Value Like "*" & DP & "*" & DP & "*" And Len(Value) > 0 And Value <> DP
End Function


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
Public Function SetCurrencyFormat(pj As Project)
    Dim currencyformat As String
    Dim i As Integer


    ' Set currency number format
    currencyformat = ""

    Select Case pj.CurrencySymbolPosition
    Case pjBefore
        currencyformat = """" & pj.CurrencySymbol & """"
    Case pjBeforeWithSpace
        currencyformat = """" & pj.CurrencySymbol & """" & " "
    End Select

    currencyformat = currencyformat & "#,##0"

    If ActiveProject.CurrencyDigits > 0 Then
        currencyformat = currencyformat & "."
        For i = 1 To pj.CurrencyDigits
            currencyformat = currencyformat & "0"
        Next i
    End If

    Select Case pj.CurrencySymbolPosition
    Case pjAfter
        currencyformat = currencyformat & """" & pj.CurrencySymbol & """"
    Case pjAfterWithSpace
        currencyformat = currencyformat & " " & """" & pj.CurrencySymbol & """"
    End Select

    SetCurrencyFormat = currencyformat

End Function


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
Public Function MyDocuments() As String
    Dim pos As Long
    Dim sBuffer As String
    sBuffer = Space$(260)
    If SHGetFolderPath(0&, CSIDL_PERSONAL, -1, 0&, sBuffer) = 0 Then
        pos = InStr(1, sBuffer, Chr(0))
        MyDocuments = Left$(sBuffer, pos - 1)
    End If
End Function

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
Public Function Ceiling(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
' X is the value you want to round
' Factor is the multiple to which you want to round
    
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor

End Function

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
'//////////////// EARLY BINDING DECLARATIONS////////////////
' Public Sub TurnCalculationOnAndOff(App As Excel.Application)
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Public Sub TurnCalculationOnAndOff(App As Object)
'////////////////////////////////////////////////////////////

    Select Case App.Calculation

    Case xlCalculationManualLB
        App.Calculation = xlCalculationAutomaticLB
    Case xlCalculationAutomaticLB
        App.Calculation = xlCalculationManualLB
    Case xlCalculationSemiautomaticLB
        MsgBox ("Calculation is set to semiautomatic. No change will take place")
    End Select

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

'//////////////// EARLY BINDING DECLARATIONS////////////////
'Public Sub AddConditionalFormatting(ws1 As Worksheet, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional bReverse As Boolean = False)
    ' Dim cs As ColorScale
'////////////////////////////////////////////////////////////

'//////////////// LATE BINDING DECLARATIONS////////////////
Public Sub AddConditionalFormatting(ws1 As Object, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional bReverse As Boolean = False)
    Dim cs As Object
'////////////////////////////////////////////////////////////
    With ws1
        .Range(.cells(x1, y1), .cells(x2, y2)).FormatConditions.Delete
        Set cs = .Range(.cells(x1, y1), .cells(x2, y2)).FormatConditions.AddColorScale(ColorScaleType:=3)

        ' Set the color of the lowest value, with a range up to the next scale criteria. The color should be red.
        With cs.ColorScaleCriteria(1)
            .Type = xlConditionValueLowestValueLB
            With .FormatColor
                If bReverse Then
                    .Color = &H6B69F8
                Else
                    .Color = &H7BBE63
                End If
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be red/green.
        ' Note that you can't set the Value property for all values of Type.
        With cs.ColorScaleCriteria(2)
            .Type = xlConditionValuePercentileLB
            .Value = 50
            With .FormatColor
                .Color = &H84EBFF
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be green.
        With cs.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValueLB
            With .FormatColor
                If bReverse Then
                    .Color = &H7BBE63
                Else
                    .Color = &H6B69F8
                End If

                .TintAndShade = 0
            End With
        End With
    End With

    Set cs = Nothing
End Sub








