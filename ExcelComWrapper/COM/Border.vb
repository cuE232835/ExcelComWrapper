Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Border
        Private m_intIndex As XlBordersIndex
        Private m_objBorders As Borders

        Friend Sub New(ByRef rObjBorders As Borders, ByVal vIntIndex As XlBordersIndex, ByVal vStrId As String)
            m_intIndex = vIntIndex
            m_objBorders = rObjBorders
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
        Friend ReadOnly Property COM As Interop.Excel.Border
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Border)(Me.ID)
            End Get
        End Property

        Public ReadOnly Property Application As Application
            Get
                Return m_objBorders.Application
            End Get
        End Property

        Public ReadOnly Property Parent As Range
            Get
                Return m_objBorders.Parent
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
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Color = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
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
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ColorIndex = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try
            End Set
        End Property

        Public Property LineStyle As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.LineStyle
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.LineStyle = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
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
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ThemeColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
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
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.TintAndShade = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
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
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Weight = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objBorder, objBorders, objRange)
                End Try
            End Set
        End Property
#End Region

    End Class
End Namespace


