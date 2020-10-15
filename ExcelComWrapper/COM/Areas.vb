Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Areas
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
        Friend ReadOnly Property COM As Interop.Excel.Areas
            Get
                Return Me.Application.GetCom(Of Interop.Excel.Areas)(Me.ID)
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

        Public ReadOnly Property Count As Integer
            Get
                Dim intRes As Integer = 0

                Try
                    intRes = Me.COM.Count
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objAreas, objRange)
                End Try

                Return intRes
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal Index As Integer) As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.Parent.Parent, RangeTypeEnum.Range, Me.Application.RegisterCom(Me.COM.Item(Index)))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objARange, objAreas, objRange)
                End Try

                Return objRes
            End Get
        End Property
    End Class
#End Region

End Namespace

