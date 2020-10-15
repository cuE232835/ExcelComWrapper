Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Interior
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
        Friend ReadOnly Property COM As Interop.Excel.Interior
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Interior)(Me.ID)
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
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Color = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
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
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ColorIndex = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property

        Public Property Pattern As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Pattern
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Pattern = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property

        Public Property PatternColor As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.PatternColor
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.PatternColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property

        Public Property PatternColorIndex As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.PatternColorIndex
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.PatternColorIndex = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property

        Public Property PatternThemeColor As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.PatternThemeColor
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.PatternThemeColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property

        Public Property PatternTintAndShade As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.PatternTintAndShade
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.PatternTintAndShade = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
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
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ThemeColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
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
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.TintAndShade = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objInterior, objRange)
                End Try
            End Set
        End Property
#End Region

    End Class
End Namespace

