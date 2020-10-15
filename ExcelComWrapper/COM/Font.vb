Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Font
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
        Friend ReadOnly Property COM As Interop.Excel.Font
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Font)(Me.ID)
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

        Public Property Background As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Background
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Background = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Bold As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Bold
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Bold = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Color As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Color
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Color = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
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
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ColorIndex = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property FontStyle As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.FontStyle
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.FontStyle = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Italic As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Italic
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Italic = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
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
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Name = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Size As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Size
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Size = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Strikethrough As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Strikethrough
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Strikethrough = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Subscript As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Subscript
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Subscript = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Superscript As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Superscript
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Superscript = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
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
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.ThemeColor = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property ThemeFont As Excel.XlThemeFont
            Get
                Dim objRes As Excel.XlThemeFont = Excel.XlThemeFont.xlThemeFontNone

                Try
                    objRes = Me.COM.ThemeFont
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Excel.XlThemeFont)
                Try
                    Me.COM.ThemeFont = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
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
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.TintAndShade = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property

        Public Property Underline As Object
            Get
                Dim objRes As Object = Nothing

                Try
                    objRes = Me.COM.Underline
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try

                Return objRes
            End Get
            Set(value As Object)
                Try
                    Me.COM.Underline = value
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objFont, objRange)
                End Try
            End Set
        End Property
#End Region

    End Class
End Namespace

