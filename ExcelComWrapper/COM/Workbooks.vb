Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Workbooks

        Friend Sub New(ByRef rObjApp As Application, ByVal vStrId As String)
            m_objApp = rObjApp
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
        Friend ReadOnly Property COM As Interop.Excel.Workbooks
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Workbooks)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objApp
            End Get
        End Property
        Private m_objApp As Application

        Default Public ReadOnly Property Item(ByVal Index As Object) As Workbook
            Get
                Dim objRes As Workbook = Nothing

                Try
                    objRes = New Workbook(Me.Application, Me.Application.RegisterCom(Me.COM(Index)))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorkbook, objWorkbooks)
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
                    'Excel.ReleaseComObject(objWorkbooks)
                End Try

                Return intRes
            End Get
        End Property
#End Region

#Region "methods"
        Public Function Add() As Workbook
            Return Me.Add(Nothing)
        End Function

        Public Function Add(ByRef Template As Object) As Workbook
            Dim objRes As Workbook = Nothing

            Try
                If Template Is Nothing Then
                    objRes = New Workbook(Me.Application, Me.Application.RegisterCom(Me.COM.Add()))
                Else
                    objRes = New Workbook(Me.Application, Me.Application.RegisterCom(Me.COM.Add(Template)))
                End If
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorkbook, objWorkbooks)
            End Try

            Return objRes
        End Function

        'Public Function Open(Filename As String, Optional ByVal UpdateLinks As Object = Nothing,
        '                     Optional ByVal [ReadOnly] As Object = Nothing, Optional ByVal Format As Object = Nothing,
        '                     Optional ByVal Password As Object = Nothing, Optional ByVal WriteResPassword As Object = Nothing,
        '                     Optional ByVal IgnoreReadOnlyRecommended As Object = Nothing, Optional ByVal Origin As Object = Nothing,
        '                     Optional ByVal Delimiter As Object = Nothing, Optional ByVal Editable As Object = Nothing,
        '                     Optional ByVal Notify As Object = Nothing, Optional ByVal Converter As Object = Nothing,
        '                     Optional ByVal AddToMru As Object = Nothing, Optional ByVal Local As Object = Nothing,
        '                     Optional ByVal CorruptLoad As Object = Nothing) As Workbook
        '    Dim objRes As Workbook = Nothing

        '    Dim objWorkbooks As Interop.Excel.Workbooks = Nothing
        '    Dim objWorkbook As Interop.Excel.Workbook = Nothing
        '    Try
        '        objWorkbooks = Me.Application.COM.Workbooks
        '        objWorkbook = objWorkbooks.Open(Filename, UpdateLinks:=UpdateLinks,
        '                                        [ReadOnly]:=[ReadOnly], Format:=Format,
        '                                        Password:=Password, WriteResPassword:=WriteResPassword,
        '                                        IgnoreReadOnlyRecommended:=IgnoreReadOnlyRecommended, Origin:=Origin,
        '                                        Delimiter:=Delimiter, Editable:=Editable,
        '                                        Notify:=Notify, Converter:=Converter,
        '                                        AddToMru:=AddToMru, Local:=Local,
        '                                        CorruptLoad:=CorruptLoad)
        '        objRes = New Workbook(Me.Application, objWorkbook.CodeName)
        '    Catch ex As Exception
        '        Throw
        '    Finally
        '        Excel.ReleaseComObject(objWorkbook, objWorkbooks)
        '    End Try
        '    Return objRes
        'End Function
        Public Function Open(Filename As String, Optional ByVal [ReadOnly] As Boolean = False) As Workbook
            Dim objRes As Workbook = Nothing

            Try
                objRes = New Workbook(Me.Application, Me.Application.RegisterCom(Me.COM.Open(Filename, [ReadOnly]:=[ReadOnly])))
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorkbook, objWorkbooks)
            End Try
            Return objRes
        End Function

        Public Sub Close()
            Dim colList As New List(Of Workbook)
            Try
                For Each objWorkbook As Interop.Excel.Workbook In Me.COM
                    colList.Add(New Workbook(Me.Application, Me.Application.RegisterCom(objWorkbook)))
                Next
                For Each objBook As Workbook In colList
                    objBook.Close()
                Next
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorkbooks, objWorkbook)
            End Try
        End Sub
#End Region
    End Class
End Namespace

