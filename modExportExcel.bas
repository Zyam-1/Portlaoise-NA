Attribute VB_Name = "modExportExcel"
Option Explicit

Public Sub ExportFlexGrid(ByVal objGrid As MSFlexGrid, _
                          ByVal CallingForm As Form, _
                          Optional ByVal HeadingMatrix As String = "")


          Dim objXL As Object
          Dim objWB As Object
          Dim objWS As Object
          Dim R As Long
          Dim c As Long
          Dim T As Single

          'Assume the calling form has a MSFlexGrid (grdToExport),
          'CommandButton (cmdXL) and Label (lblExcelInfo) (Visible set to False)
          'In the calling form:
          'Private Sub cmdXL_Click()
          'ExportFlexGrid grdToExport, Me
          'End Sub

10        On Error GoTo ExportFlexGrid_Error

20        With CallingForm.lblExcelInfo
30            .Caption = "Exporting..."
40            .Visible = True
50            .Refresh
60        End With

70        Set objXL = CreateObject("Excel.Application")
80        Set objWB = objXL.Workbooks.Add
90        Set objWS = objWB.Worksheets(1)

          Dim intLineCount As Integer
          '****Change: Babar Shahzad 2007-11-19
          'Heading for export to excel can be passed as string which would be
          'a string having TABS as column breaks and CR as row break.

100       intLineCount = 0
110       If HeadingMatrix <> "" Then
120           With objWS
                  Dim strTokens() As String
130               strTokens = Split(HeadingMatrix, vbCr)
140               intLineCount = UBound(strTokens)

150               For R = LBound(strTokens) To UBound(strTokens) - 1
                      'For C = 0 To objGrid.Cols - 1
                      'The "'" is required to format the cells as text in Excel
                      'otherwise entries like "4/2" are interpreted as a date
160                   .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).MergeCells = True
170                   .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).HorizontalAlignment = 3
180                   .Range(.Cells(R + 1, 1), .Cells(R + 1, objGrid.Cols)).Font.Bold = True
190                   objWS.Cells(R + 1, 1) = "'" & strTokens(R)

200               Next
210           End With

220       End If

230       With objWS
240           For R = 0 To objGrid.Rows - 1
250               For c = 0 To objGrid.Cols - 1
                      'The "'" is required to format the cells as text in Excel
                      'otherwise entries like "4/2" are interpreted as a date
260                   If R = 0 Then
270                       .Range(.Cells(R + 1 + intLineCount, 1), .Cells(R + 1 + intLineCount, objGrid.Cols)).Font.Bold = True

280                   End If
290                   .Cells(R + 1 + intLineCount, c + 1) = "'" & objGrid.TextMatrix(R, c)
300               Next
310           Next

320           .Cells.Columns.AutoFit
330       End With

340       objXL.Visible = True

350       Set objWS = Nothing
360       Set objWB = Nothing
370       Set objXL = Nothing

380       CallingForm.lblExcelInfo.Visible = False

390       Exit Sub

ExportFlexGrid_Error:

          Dim strES As String
          Dim lngErr As Long

400       iMsg strES

410       With CallingForm.lblExcelInfo
420           .Caption = "Error " & Format(lngErr)
430           .Refresh
440           T = Timer
450           Do While Timer - T < 1: Loop
460           .Visible = False
470       End With

End Sub


