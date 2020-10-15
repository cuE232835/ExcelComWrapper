Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Worksheet
        Private m_objSheets As Sheets

        Friend Sub New(ByRef rObjSheets As Sheets, ByVal vStrId As String)
            m_objSheets = rObjSheets
            m_strId = vStrId
        End Sub

#Region "properties"
        ''' ****************************************************************
        ''' <summary>COMオブジェクトID</summary>
        ''' ****************************************************************
        Public ReadOnly Property ID As String
            Get
                Return m_strId
            End Get
        End Property
        Private m_strId As String

        ''' ================================================================
        ''' <summary>COMオブジェクト</summary>
        ''' ================================================================
        Friend ReadOnly Property COM As Interop.Excel.Worksheet
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Worksheet)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objSheets.Application
            End Get
        End Property

        Public ReadOnly Property Parent As Workbook
            Get
                Return m_objSheets.Parent
            End Get
        End Property

        Public ReadOnly Property CodeName As String
            Get
                Return Me.COM.CodeName
            End Get
        End Property

        Public ReadOnly Property Cells As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Cells))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Columns As Range
            Get
                Return Me.Columns(Nothing)
            End Get
        End Property

        Public ReadOnly Property Columns(ByVal ColumnIndex As Object) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    If ColumnIndex Is Nothing Then
                        objRes = New Range(Me, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Columns))
                    Else
                        objRes = New Range(Me, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Columns(ColumnIndex)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Rows As Range
            Get
                Return Me.Rows(Nothing)
            End Get
        End Property

        Public ReadOnly Property Rows(ByVal RowIndex As Object) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    If RowIndex Is Nothing Then
                        objRes = New Range(Me, RangeTypeEnum.Rows, Me.Application.RegisterCom(Me.COM.Rows))
                    Else
                        objRes = New Range(Me, RangeTypeEnum.Rows, Me.Application.RegisterCom(Me.COM.Rows(RowIndex)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property Name As String
            Get
                Dim strRes As String = String.Empty

                Try
                    strRes = Me.COM.Name
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
                End Try

                Return strRes
            End Get
            Set(value As String)
                Try
                    Me.COM.Name = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
                End Try
            End Set
        End Property

        Public ReadOnly Property Range(ByVal Cell1 As Object, Optional ByVal Cell2 As Object = Nothing) As Range
            Get
                Dim objRes As Range = Nothing

                Dim objCell1 As Object = Nothing
                Dim objCell2 As Object = Nothing
                Try
                    If Not Cell1 Is Nothing Then
                        If TypeOf Cell1 Is Range Then
                            objCell1 = DirectCast(Cell1, Range).COM
                        Else
                            objCell1 = Cell1
                        End If
                    End If
                    If Not Cell2 Is Nothing Then
                        If TypeOf Cell2 Is Range Then
                            objCell2 = DirectCast(Cell2, Range).COM
                        Else
                            objCell2 = Cell2
                        End If
                    End If

                    If objCell2 Is Nothing Then
                        objRes = New Range(Me, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Range(objCell1)))
                    Else
                        objRes = New Range(Me, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Range(objCell1, objCell2)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(colRelease.ToArray)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property UsedRange As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.UsedRange))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property Visible As XlSheetVisibility
            Get
                Dim intRes As XlSheetVisibility = XlSheetVisibility.xlSheetVisible

                Try
                    Dim strName As String = [Enum].GetName(GetType(Interop.Excel.XlSheetVisibility), Me.COM.Visible)
                    [Enum].TryParse(Of XlSheetVisibility)(strName, intRes)
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
                End Try

                Return intRes
            End Get
            Set(value As XlSheetVisibility)
                Try
                    Dim strName As String = [Enum].GetName(GetType(XlSheetVisibility), value)
                    [Enum].TryParse(Of Interop.Excel.XlSheetVisibility)(strName, Me.COM.Visible)
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
                End Try
            End Set
        End Property
#End Region

#Region "methods"
        Public Sub Copy(Optional ByVal Before As Object = Nothing, Optional ByVal After As Object = Nothing)
            Dim objBefore As Object = Nothing
            Dim objAfter As Object = Nothing

            Try
                If Not Before Is Nothing Then
                    If TypeOf Before Is Worksheet Then
                        objBefore = DirectCast(Before, Worksheet).COM
                    Else
                        objBefore = Before
                    End If
                End If
                If Not After Is Nothing Then
                    If TypeOf After Is Worksheet Then
                        objAfter = DirectCast(After, Worksheet).COM
                    Else
                        objAfter = After
                    End If
                End If

                If objBefore Is Nothing Then
                    If objAfter Is Nothing Then
                        Me.COM.Copy()
                    Else
                        Me.COM.Copy(After:=objAfter)
                    End If
                Else
                    If objAfter Is Nothing Then
                        Me.COM.Copy(Before:=objBefore)
                    Else
                        Me.COM.Copy(Before:=objBefore, After:=objAfter)
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(colRelease.ToArray)
            End Try
        End Sub

        Public Sub Move(Optional ByVal Before As Object = Nothing, Optional ByVal After As Object = Nothing)
            Dim objBefore As Object = Nothing
            Dim objAfter As Object = Nothing

            Try
                If Not Before Is Nothing Then
                    If TypeOf Before Is Worksheet Then
                        objBefore = DirectCast(Before, Worksheet).COM
                    Else
                        objBefore = Before
                    End If
                End If
                If Not After Is Nothing Then
                    If TypeOf After Is Worksheet Then
                        objAfter = DirectCast(After, Worksheet).COM
                    Else
                        objAfter = After
                    End If
                End If

                If objBefore Is Nothing Then
                    If objAfter Is Nothing Then
                        Me.COM.Move()
                    Else
                        Me.COM.Move(After:=objAfter)
                    End If
                Else
                    If objAfter Is Nothing Then
                        Me.COM.Move(Before:=objBefore)
                    Else
                        Me.COM.Move(Before:=objBefore, After:=objAfter)
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(colRelease.ToArray)
            End Try
        End Sub

        Public Sub Delete()
            Try
                Me.COM.Delete()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
            End Try
        End Sub

        Public Sub Paste(Optional ByVal Destination As Object = Nothing, Optional ByVal Link As Object = Nothing)
            Dim objDest As Object = Nothing
            Dim colRelease As New List(Of Object)
            Try
                If Not Destination Is Nothing Then
                    If TypeOf Destination Is Range Then
                        If Link Is Nothing Then
                            Me.COM.Paste(DirectCast(Destination, Range).COM)
                        Else
                            Me.COM.Paste(DirectCast(Destination, Range).COM, Link)
                        End If
                    Else
                        If Link Is Nothing Then
                            Me.COM.Paste(Destination)
                        Else
                            Me.COM.Paste(Destination, Link)
                        End If
                    End If
                Else
                    If Link Is Nothing Then
                        Me.COM.Paste()
                    Else
                        Me.COM.Paste(Link)
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
            End Try
        End Sub

        Public Sub [Select](Optional ByVal Replace As Object = Nothing)
            Try
                Me.COM.Select(Replace)
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
            End Try
        End Sub

        Public Sub Protect(Optional ByVal Password As Object = Nothing, Optional ByVal DrawingObjects As Object = Nothing,
                       Optional ByVal Contents As Object = Nothing, Optional ByVal Scenarios As Object = Nothing,
                       Optional ByVal UserInterfaceOnly As Object = Nothing, Optional ByVal AllowFormattingCells As Object = Nothing,
                       Optional ByVal AllowFormattingColumns As Object = Nothing, Optional ByVal AllowFormattingRows As Object = Nothing,
                       Optional ByVal AllowInsertingColumns As Object = Nothing, Optional ByVal AllowInsertingRows As Object = Nothing,
                       Optional ByVal AllowInsertingHyperlinks As Object = Nothing, Optional ByVal AllowDeletingColumns As Object = Nothing,
                       Optional ByVal AllowDeletingRows As Object = Nothing, Optional ByVal AllowSorting As Object = Nothing,
                       Optional ByVal AllowFiltering As Object = Nothing, Optional ByVal AllowUsingPivotTables As Object = Nothing)
            Try

                Me.COM.Protect(Password:=Password, DrawingObjects:=DrawingObjects,
                                 Contents:=Contents, Scenarios:=Scenarios,
                                 UserInterfaceOnly:=UserInterfaceOnly, AllowFormattingCells:=AllowFormattingCells,
                                 AllowFormattingColumns:=AllowFormattingColumns, AllowFormattingRows:=AllowFormattingRows,
                                 AllowInsertingColumns:=AllowInsertingColumns, AllowInsertingRows:=AllowInsertingRows,
                                 AllowInsertingHyperlinks:=AllowInsertingHyperlinks, AllowDeletingColumns:=AllowDeletingColumns,
                                 AllowDeletingRows:=AllowDeletingRows, AllowSorting:=AllowSorting,
                                 AllowFiltering:=AllowFiltering, AllowUsingPivotTables:=AllowUsingPivotTables)
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
            End Try
        End Sub

        Public Sub Unprotect(Optional ByVal Password As Object = Nothing)
            Try
                Me.COM.Unprotect(Password)
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
            End Try
        End Sub

        Friend Function ConvertColToNumber(ByVal vStrColName As String) As Integer
            Dim intRes As Integer = 0

            Dim objRange As Interop.Excel.Range = Nothing
            Try
                objRange = Me.COM.Range(vStrColName & "1")

                intRes = objRange.Column
            Catch ex As Exception
                Throw
            Finally
                ExcelWrapper.ReleaseComObject(objRange)
            End Try

            Return intRes
        End Function
#End Region

    End Class
End Namespace

