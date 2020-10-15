Imports ExcelComWrapper

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim objApp As Excel.Application = Nothing

        Me.Enabled = False
        Try
            objApp = ExcelWrapper.CreateInstance(True)
            Dim objBook As Excel.Workbook = objApp.Workbooks.Add()
            Dim objSheet As Excel.Worksheet = objBook.Sheets.Add
            objSheet.Cells(1, 1).Value = "This is a test."

            objSheet.Cells(1, 1).Copy()
            objSheet.Cells(2, 1).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteAll)

            Dim objRange As Excel.Range = objSheet.Range(objSheet.Cells(1, 1), objSheet.Cells(3, 3))
            objRange.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlDouble
            objRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlDouble
            objRange.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlDouble
            objRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble
            objRange.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            objRange.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
            objRange.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlDash
            objRange.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlDash
            objRange.Interior.Color = Color.Aqua

            objRange.Cut()
            Dim objDest As Excel.Worksheet = objBook.Sheets(1)
            'objDest.Paste(objDest.Cells(10, 5))
            Dim objTest As Excel.Range = objDest.Cells(11, 5)
            objDest.Cells(10, 5).Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
            objApp.CutCopyMode = Excel.XlCutCopyMode.False

            MsgBox("Address of test range=" & objTest.Address)
            Dim objDlg As New SaveFileDialog
            objDlg.Filter = "xlsx|xlsx"
            If objDlg.ShowDialog = DialogResult.OK Then
                objBook.SaveAs(Filename:=objDlg.FileName)
            End If

            objBook.Close(SaveChanges:=False)
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & ex.StackTrace)
        Finally
            Me.Enabled = True
            If Not objApp Is Nothing Then
                objApp.Quit()
                objApp = Nothing
            End If
        End Try
    End Sub
End Class
