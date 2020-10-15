Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Sheets
        Private m_objBook As Workbook

        Friend Sub New(ByRef rObjBook As Workbook, ByVal vStrId As String)
            m_objBook = rObjBook
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
        Friend ReadOnly Property COM As Interop.Excel.Sheets
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Sheets)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objBook.Application
            End Get
        End Property

        Public ReadOnly Property Parent As Workbook
            Get
                Return m_objBook
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal Index As Object) As Worksheet
            Get
                Dim objRes As Worksheet = Nothing

                Try
                    objRes = New Worksheet(Me, Me.Application.RegisterCom(Me.COM(Index)))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet, objSheets, objWorkbook, objWorkbooks)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    intRes = Me.COM.Count
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorkbooks, objWorkbook, objSheets)
                End Try

                Return intRes
            End Get
        End Property
#End Region

#Region "methods"
        Public Function Add(Optional Before As Object = Nothing, Optional After As Object = Nothing,
                            Optional Count As Integer = 1, Optional [Type] As XlSheetType = XlSheetType.xlWorksheet) As Object
            Dim objRes As Object = Nothing

            Try
                Dim objBefore As Object = Nothing
                Dim objAfter As Object = Nothing
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

                If Not objBefore Is Nothing Then
                    If Not objAfter Is Nothing Then
                        objRes = New Worksheet(Me.Parent.Sheets, Me.Application.RegisterCom(Me.COM.Add(Before:=objBefore, After:=objAfter, Count:=Count, Type:=Type)))
                    Else
                        objRes = New Worksheet(Me.Parent.Sheets, Me.Application.RegisterCom(Me.COM.Add(Before:=objBefore, Count:=Count, Type:=Type)))
                    End If
                Else
                    If Not objAfter Is Nothing Then
                        objRes = New Worksheet(Me.Parent.Sheets, Me.Application.RegisterCom(Me.COM.Add(After:=objAfter, Count:=Count, Type:=Type)))
                    Else
                        objRes = New Worksheet(Me.Parent.Sheets, Me.Application.RegisterCom(Me.COM.Add(Count:=Count, Type:=Type)))
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                '
            End Try

            Return objRes
        End Function
#End Region

    End Class
End Namespace

