Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Range
        Private m_objSheet As Worksheet

        Friend Sub New(ByRef rObjSheet As Worksheet, ByVal vIntRangeType As RangeTypeEnum, ByVal vStrId As String)
            m_objSheet = rObjSheet
            m_intRangeType = vIntRangeType
            m_strId = vStrId

            m_objInterior = New Interior(Me, Me.Application.RegisterCom(Me.COM.Interior))
            m_objAreas = New Areas(Me, Me.Application.RegisterCom(Me.COM.Areas))
            m_objBorders = New Borders(Me, Me.Application.RegisterCom(Me.COM.Borders))
            m_objFont = New Font(Me, Me.Application.RegisterCom(Me.COM.Font))
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
        Friend ReadOnly Property COM As Interop.Excel.Range
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Range)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return Me.Parent.Application
            End Get
        End Property

        Friend ReadOnly Property RangeType As RangeTypeEnum
            Get
                Return m_intRangeType
            End Get
        End Property
        Private m_intRangeType As RangeTypeEnum = RangeTypeEnum.Range

        Public ReadOnly Property Address As String
            Get
                Return Me.COM.Address
            End Get
        End Property

        Public ReadOnly Property Parent As Worksheet
            Get
                Return m_objSheet
            End Get
        End Property

        Public ReadOnly Property Areas As Areas
            Get
                Return m_objAreas
            End Get
        End Property
        Private m_objAreas As Areas

        Public ReadOnly Property Borders As Borders
            Get
                Return m_objBorders
            End Get
        End Property
        Private m_objBorders As Borders

        Public ReadOnly Property Cells As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Cells))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objCells, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Column As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    intRes = Me.COM.Column
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return intRes
            End Get
        End Property

        Public ReadOnly Property Row As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    intRes = Me.COM.Row
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return intRes
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
                        objRes = New Range(Me.Parent, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Columns))
                    Else
                        objRes = New Range(Me.Parent, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Columns(ColumnIndex)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objColumns, objRange)
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
                        objRes = New Range(Me.Parent, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Rows))
                    Else
                        objRes = New Range(Me.Parent, RangeTypeEnum.Columns, Me.Application.RegisterCom(Me.COM.Rows(RowIndex)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRows, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    Dim objAddress As New ExcelWrapper.Address(Me.Address)

                    If Me.RangeType = RangeTypeEnum.Range Then
                        If Not Me.Application.GetCom(Of Interop.Excel.Range)(Me.ID) Is Nothing Then
                            intRes = Me.Application.GetCom(Of Interop.Excel.Range)(Me.ID).Count
                        End If
                    ElseIf Me.RangeType = RangeTypeEnum.Columns Then
                        Dim intStart As Integer = Me.Parent.ConvertColToNumber(objAddress.StartColumn)
                        Dim intEnd As Integer = Me.Parent.ConvertColToNumber(objAddress.EndColumn)
                        intRes = intEnd - intStart + 1
                    ElseIf Me.RangeType = RangeTypeEnum.Rows Then
                        intRes = objAddress.EndRow - objAddress.StartRow + 1
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return intRes
            End Get
        End Property

        Public Property ColumnWidth As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.ColumnWidth
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ColumnWidth = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public ReadOnly Property [End](ByVal Direction As XlDirection) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.End(Direction)))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objEnd, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property EntireColumn As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.EntireColumn))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objEntire, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property EntireRow As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.EntireRow))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objEntire, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Font As Font
            Get
                Return m_objFont
            End Get
        End Property
        Private m_objFont As Font

        Public Property Formula As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Formula
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Formula = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property FormulaArray As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FormulaArray
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FormulaArray = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property FormulaHidden As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FormulaHidden
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FormulaHidden = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property FormulaLocal As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FormulaLocal
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FormulaLocal = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property FormulaR1C1 As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FormulaR1C1
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FormulaR1C1 = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property FormulaR1C1Local As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FormulaR1C1Local
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FormulaR1C1Local = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property HorizontalAlignment As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.HorizontalAlignment
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.HorizontalAlignment = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public ReadOnly Property Interior As Interior
            Get
                Return m_objInterior
            End Get
        End Property
        Private m_objInterior As Interior

        Public ReadOnly Property MergeArea As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = Me.COM.MergeArea
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property MergeCells As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.MergeCells
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.MergeCells = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property Name As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Name
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Name = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property NumberFormat As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.NumberFormat
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.NumberFormat = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property NumberFormatLocal As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.NumberFormatLocal
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.NumberFormatLocal = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public ReadOnly Property Offset(Optional ByVal RowOffset As Object = Nothing, Optional ByVal ColumnOffset As Object = Nothing) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = Me.COM.Offset(RowOffset, ColumnOffset)
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property Orientation As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Orientation
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Orientation = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public ReadOnly Property Text As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Text
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property Value(Optional ByVal RangeValueDataType As Object = Nothing) As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Value(RangeValueDataType)
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Value(RangeValueDataType) = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try
            End Set
        End Property

        Public Property Value2 As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Value2
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Value2 = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
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
                        objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Range(objCell1)))
                    Else
                        objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Range(objCell1, objCell2)))
                    End If

                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(colRelease.ToArray)
                End Try

                Return objRes
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal RowIndex As Object, Optional ByVal ColumnIndex As Object = Nothing) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    If ColumnIndex Is Nothing Then
                        objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Item(RowIndex)))
                    Else
                        objRes = New Range(Me.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Item(RowIndex, ColumnIndex)))
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objItem, objRange)
                End Try

                Return objRes
            End Get
        End Property
#End Region

#Region "methods"
        Public Function AutoFit() As Boolean
            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.AutoFit()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function Clear() As Boolean
            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.Clear()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Sub ClearComments()
            Try
                Me.COM.Clear()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try
        End Sub

        Public Function ClearContents() As Boolean
            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.ClearContents()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function ClearFormats() As Boolean
            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.ClearFormats()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Sub ClearHyperlinks()
            Try
                Me.COM.ClearHyperlinks()
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try
        End Sub

        Public Function Copy(Optional ByVal Destination As Object = Nothing) As Boolean
            Dim blnRes As Boolean = False

            Try
                If Destination Is Nothing Then
                    blnRes = Me.COM.Copy()
                Else
                    blnRes = Me.COM.Copy(Destination)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function Cut(Optional ByVal Destination As Object = Nothing) As Boolean
            Dim blnRes As Boolean = False

            Try
                If Destination Is Nothing Then
                    blnRes = Me.COM.Cut()
                Else
                    blnRes = Me.COM.Cut(Destination)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function Delete(Optional ByVal Shift As Object = Nothing) As Boolean
            Dim blnRes As Boolean = False

            Try
                If Shift Is Nothing Then
                    blnRes = Me.COM.Delete()
                Else
                    blnRes = Me.COM.Delete(Shift)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function Find(ByVal What As Object, Optional ByVal After As Range = Nothing,
                         Optional ByVal LookIn As XlFindLookIn = XlFindLookIn.xlValues, Optional ByVal LookAt As XlLookAt = XlLookAt.xlWhole,
                         Optional ByVal SearchOrder As XlSearchOrder = XlSearchOrder.xlByColumns,
                         Optional ByVal SearchDirection As XlSearchDirection = XlSearchDirection.xlNext,
                         Optional ByVal MatchCase As Boolean = False, Optional ByVal MatchByte As Boolean = True,
                         Optional ByVal SearchFormat As Object = Nothing) As Range
            Dim objRes As Range = Nothing

            Dim strId As String = String.Empty
            Try
                If After Is Nothing Then
                    strId = Me.Application.RegisterCom(Me.COM.Find(What,
                                                      LookIn:=LookIn, LookAt:=LookAt,
                                                      SearchOrder:=SearchOrder,
                                                      SearchDirection:=SearchDirection,
                                                      MatchCase:=MatchCase, MatchByte:=MatchByte))
                Else
                    strId = Me.Application.RegisterCom(Me.COM.Find(What, After:=After,
                                                      LookIn:=LookIn, LookAt:=LookAt,
                                                      SearchOrder:=SearchOrder,
                                                      SearchDirection:=SearchDirection,
                                                      MatchCase:=MatchCase, MatchByte:=MatchByte))
                End If
                If Not String.IsNullOrEmpty(strId) Then
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, strId)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objFound, objRange)
            End Try

            Return objRes
        End Function

        Public Function FindNext(Optional ByVal After As Range = Nothing) As Range
            Dim objRes As Range = Nothing

            Dim strId As String = String.Empty
            Try
                If After Is Nothing Then
                    strId = Me.Application.RegisterCom(Me.COM.FindNext())
                Else
                    strId = Me.Application.RegisterCom(Me.COM.FindNext(After))
                End If

                If Not String.IsNullOrEmpty(strId) Then
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, strId)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objFound, objRange)
            End Try

            Return objRes
        End Function

        Public Function FindPrevious(Optional ByVal After As Object = Nothing) As Range
            Dim objRes As Range = Nothing

            Dim strId As String = String.Empty
            Try
                If After Is Nothing Then
                    strId = Me.Application.RegisterCom(Me.COM.FindPrevious())
                Else
                    strId = Me.Application.RegisterCom(Me.COM.FindPrevious(After))
                End If

                If Not String.IsNullOrEmpty(strId) Then
                    objRes = New Range(Me.Parent, RangeTypeEnum.Range, strId)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objFound, objRange)
            End Try

            Return objRes
        End Function

        Public Function Insert(Optional ByVal Shift As Object = Nothing,
                           Optional ByVal CopyOrigin As XlInsertFormatOrigin = XlInsertFormatOrigin.xlFormatFromLeftOrAbove) As Boolean
            Dim blnRes As Boolean = False

            Try
                If Shift Is Nothing Then
                    blnRes = Me.COM.Insert(CopyOrigin:=CopyOrigin)
                Else
                    blnRes = Me.COM.Insert(Shift:=Shift, CopyOrigin:=CopyOrigin)
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function PasteSpecial(Optional ByVal Paste As XlPasteType = XlPasteType.xlPasteAll,
                                 Optional ByVal Operation As XlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                 Optional ByVal SkipBlanks As Boolean = False, Optional ByVal Transpose As Boolean = False) As Boolean

            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.PasteSpecial(Paste:=Paste, Operation:=Operation, SkipBlanks:=SkipBlanks, Transpose:=Transpose)
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function

        Public Function [Select]() As Boolean
            Dim blnRes As Boolean = False

            Try
                blnRes = Me.COM.Select
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objRange)
            End Try

            Return blnRes
        End Function
#End Region

    End Class
End Namespace

