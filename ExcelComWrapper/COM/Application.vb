Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal

Namespace Excel
    Public Class Application

#Region "COM Management"
        ''' <summary>COMオブジェクトディクショナリ(発番済みID検索用)</summary>
        Private m_colCom As New Dictionary(Of Object, String)
        ''' <summary>COMオブジェクトディクショナリ(ID発番済みオブジェクト検索用)</summary>
        Private m_colComR As New Dictionary(Of String, Object)

        ''' ================================================================
        ''' <summary>
        ''' COMオブジェクトを登録します。
        ''' 登録済みオブジェクトの場合は発番済みのIDを返却します。
        ''' </summary>
        ''' <param name="ComObject">登録対象のCOMオブジェクト</param>
        ''' <returns></returns>
        ''' ================================================================
        Friend Function RegisterCom(ByRef ComObject As Object) As String
            Dim strRes As String = String.Empty

            If Not ComObject Is Nothing Then
                If m_colCom.ContainsKey(ComObject) Then
                    strRes = m_colCom(ComObject)
                Else
                    strRes = ExcelWrapper.GetNewId()
                    m_colCom(ComObject) = strRes
                    m_colComR(strRes) = ComObject
                End If
            End If

            Return strRes
        End Function

        ''' ================================================================
        ''' <summary>
        ''' 登録済みCOMオブジェクトをIDで検索します。
        ''' </summary>
        ''' <typeparam name="T">COMオブジェクトのタイプ</typeparam>
        ''' <param name="ObjectID">RegisterComにより発番したCOMオブジェクトID</param>
        ''' <returns></returns>
        ''' ================================================================
        Friend Function GetCom(Of T)(ByVal ObjectID As String) As T
            If m_colComR.ContainsKey(ObjectID) Then
                Return DirectCast(m_colComR(ObjectID), T)
            Else
                Return Nothing
            End If
        End Function

        ''' ================================================================
        ''' <summary>
        ''' 登録済みCOMオブジェクトを一括して開放します。
        ''' </summary>
        ''' ================================================================
        Friend Sub ClearComPool()
            m_colCom.Clear()
            For Each objCom As Object In m_colComR.Values
                If Not objCom Is Nothing Then
                    FinalReleaseComObject(objCom)
                    objCom = Nothing
                End If
            Next
            m_colComR.Clear()
        End Sub
#End Region

        Friend Sub New()
            m_strId = Me.RegisterCom(New Interop.Excel.Application)
            m_objWorkbooks = New Workbooks(Me, Me.RegisterCom(Me.COM.Workbooks))
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
        Friend ReadOnly Property COM As Interop.Excel.Application
            Get
                Return Me.GetCom(Of Interop.Excel.Application)(m_strId)
            End Get
        End Property

        Public ReadOnly Property Workbooks As Workbooks
            Get
                Return m_objWorkbooks
            End Get
        End Property
        Private m_objWorkbooks As Workbooks

        Public ReadOnly Property ActiveCell As Range
            Get
                Dim objRes As Range = Nothing

                Try
                    objRes = New Range(Me.ActiveSheet, Excel.RangeTypeEnum.Range, Me.RegisterCom(Me.COM.ActiveCell))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objRange)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property ActiveSheet As Worksheet
            Get
                Dim objRes As Worksheet = Nothing

                Try
                    objRes = New Worksheet(Me.ActiveWorkbook.Sheets, Me.RegisterCom(Me.COM.ActiveSheet))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorksheet)
                End Try

                Return objRes
            End Get
        End Property

        Public ReadOnly Property ActiveWorkbook As Workbook
            Get
                Dim objRes As Workbook = Nothing

                Try
                    objRes = New Workbook(Me, Me.RegisterCom(Me.COM.ActiveWorkbook))
                Catch ex As Exception
                    Throw
                Finally
                    'Excel.ReleaseComObject(objWorkbook)
                End Try

                Return objRes
            End Get
        End Property

        Public Property CutCopyMode As Excel.XlCutCopyMode
            Get
                Dim intRes As Excel.XlCutCopyMode = Excel.XlCutCopyMode.xlCopy
                Dim strName As String = [Enum].GetName(GetType(Interop.Excel.XlCutCopyMode), Me.COM.CutCopyMode)
                [Enum].TryParse(Of Interop.Excel.XlCutCopyMode)(strName, intRes)
                Return intRes
            End Get
            Set(value As Excel.XlCutCopyMode)
                Me.COM.CutCopyMode = value
            End Set
        End Property

        Public Property DisplayAlerts As Boolean
            Get
                Return Me.COM.DisplayAlerts
            End Get
            Set(value As Boolean)
                Me.COM.DisplayAlerts = value
            End Set
        End Property

        Public Property EnableEvents As Boolean
            Get
                Return Me.COM.EnableEvents
            End Get
            Set(value As Boolean)
                Me.COM.EnableEvents = value
            End Set
        End Property

        Public ReadOnly Property PathSeparator As String
            Get
                Return Me.COM.PathSeparator
            End Get
        End Property

        Public Property ScreenUpdating As Boolean
            Get
                Return Me.COM.ScreenUpdating
            End Get
            Set(value As Boolean)
                Me.COM.ScreenUpdating = value
            End Set
        End Property

        Public ReadOnly Property Selection As Object
            Get
                Dim objRes As Range = Nothing

                Dim objSelection As Object = Nothing

                Dim colRelease As New List(Of Object)
                Try
                    objSelection = Me.COM.Selection
                    colRelease.Add(objSelection)

                    If Not objSelection Is Nothing AndAlso TypeOf objSelection Is Interop.Excel.Range Then
                        Dim strRId As String = Me.RegisterCom(DirectCast(objSelection, Interop.Excel.Range))
                        Dim strWsId As String = Me.RegisterCom(Me.GetCom(Of Interop.Excel.Range)(strRId).Parent)
                        Dim strWbId As String = Me.RegisterCom(Me.GetCom(Of Interop.Excel.Worksheet)(strWsId).Parent)
                        Dim strSId As String = Me.RegisterCom(Me.GetCom(Of Interop.Excel.Workbook)(strWbId).Sheets)

                        objRes = New Range(New Worksheet(New Sheets(New Workbook(Me.COM, strWbId), strWsId), strSId), Excel.RangeTypeEnum.Range, strRId)
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    ExcelWrapper.ReleaseComObject(colRelease.ToArray)
                End Try

                Return objRes
            End Get
        End Property

        Public Property Visible As Boolean
            Get
                Return Me.COM.Visible
            End Get
            Set(value As Boolean)
                Me.COM.Visible = value
            End Set
        End Property
#End Region

#Region "methods"
        Public Sub Quit(Optional ByVal vBlnRemainAppAlive As Boolean = False)
            ExcelWrapper.Quit(Me)
            If Not vBlnRemainAppAlive Then
                Me.COM.Quit()
            End If
            Me.ClearComPool()
        End Sub
#End Region
    End Class
End Namespace

