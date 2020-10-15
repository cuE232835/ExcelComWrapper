Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Workbook
        Friend Sub New(ByRef rObjApp As Application, ByVal vStrId As String)
            m_objApp = rObjApp
            m_strId = vStrId
            m_objSheets = New Sheets(Me, Me.Application.RegisterCom(Me.COM.Sheets))
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
        Friend ReadOnly Property COM As Interop.Excel.Workbook
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Workbook)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objApp
            End Get
        End Property
        Private m_objApp As Application

        Public ReadOnly Property ActiveSheet As Worksheet
            Get
                Dim objRes As Worksheet = Nothing

                Try
                    objRes = New Worksheet(Me.Sheets, Me.Application.RegisterCom(Me.COM.ActiveSheet))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBook, objWorkbooks)
                End Try
                Return objRes
            End Get
        End Property

        Public ReadOnly Property CodeName As String
            Get
                Return Me.COM.CodeName
            End Get
        End Property

        Public ReadOnly Property FullName As String
            Get
                Dim strRes As String = String.Empty

                Try
                    strRes = Me.COM.FullName
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBook, objWorkbooks)
                End Try
                Return strRes
            End Get
        End Property

        Public ReadOnly Property Name As String
            Get
                Dim strRes As String = String.Empty

                Try
                    strRes = Me.COM.Name
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBook, objWorkbooks)
                End Try
                Return strRes
            End Get
        End Property

        Public ReadOnly Property Path As String
            Get
                Dim strRes As String = String.Empty

                Try
                    strRes = Me.COM.Path
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorkbook, objWorkbooks)
                End Try
                Return strRes
            End Get
        End Property

        Public ReadOnly Property Sheets As Sheets
            Get
                Return m_objSheets
            End Get
        End Property
        Private m_objSheets As Sheets
#End Region

#Region "methods"
        Public Sub Save()
            Try
                Me.COM.Save()
            Catch ex As Exception
                Throw
            Finally
                '
            End Try
        End Sub

        Public Sub SaveAs(ByVal Filename As String, Optional ByVal FileFormat As Excel.XlFileFormat = XlFileFormat.xlOpenXMLWorkbook)
            Try
                Me.COM.SaveAs(Filename:=Filename, FileFormat:=FileFormat)
            Catch ex As Exception
                Throw
            Finally
                '
            End Try
        End Sub

        Public Sub Close(Optional ByVal SaveChanges As Object = Nothing, Optional ByVal Filename As Object = Nothing, Optional ByVal RouteWorkbook As Object = Nothing)
            Try
                Me.COM.Close(SaveChanges:=SaveChanges, Filename:=Filename, RouteWorkbook:=RouteWorkbook)
            Catch ex As Exception
                Throw
            Finally
                'Excel.ReleaseComObject(objWorkbook, objWorkbooks)
            End Try
        End Sub
#End Region

    End Class
End Namespace

