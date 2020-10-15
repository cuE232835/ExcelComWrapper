Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Borders
        Private m_objRange As Range

        Friend Sub New(ByRef rObjRange As Range, ByVal vStrId As String)
            m_objRange = rObjRange
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
        Friend ReadOnly Property COM As Interop.Excel.Borders
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Borders)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objRange.Application
            End Get
        End Property

        Public ReadOnly Property Parent As Range
            Get
                Return m_objRange
            End Get
        End Property

        Public Property Color As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Color
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Color = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property ColorIndex As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.ColorIndex
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ColorIndex = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    intRes = Me.COM.Count
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return intRes
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal Index As XlBordersIndex) As Border
            Get
                Dim objRes As Border = Nothing

                Try
                    objRes = New Border(Me, Index, Me.Application.RegisterCom(Me.COM(Index)))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public Property LineStyle As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.LineStyle
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.LineStyle = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property ThemeColor As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.ThemeColor
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ThemeColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property TintAndShade As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.TintAndShade
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.TintAndShade = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property Value As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Value = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property Weight As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Weight
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Weight = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorders, objRange)
                End Try
            End Set
        End Property
#End Region

    End Class
End Namespace

