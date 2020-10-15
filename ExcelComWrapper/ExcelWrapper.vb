Imports Microsoft.Office
Imports System.Runtime.InteropServices.Marshal


Namespace Excel
#Region "constants"
    Friend Enum RangeTypeEnum
        Range
        Rows
        Columns
    End Enum

    Public Enum XlSheetVisibility
        xlSheetHidden = Interop.Excel.XlSheetVisibility.xlSheetHidden
        xlSheetVeryHidden = Interop.Excel.XlSheetVisibility.xlSheetVeryHidden
        xlSheetVisible = Interop.Excel.XlSheetVisibility.xlSheetVisible
    End Enum

    Public Enum XlBordersIndex
        xlDiagonalDown = Interop.Excel.XlBordersIndex.xlDiagonalDown
        xlDiagonalUp = Interop.Excel.XlBordersIndex.xlDiagonalUp
        xlEdgeLeft = Interop.Excel.XlBordersIndex.xlEdgeLeft
        xlEdgeTop = Interop.Excel.XlBordersIndex.xlEdgeTop
        xlEdgeBottom = Interop.Excel.XlBordersIndex.xlEdgeBottom
        xlEdgeRight = Interop.Excel.XlBordersIndex.xlEdgeRight
        xlInsideHorizontal = Interop.Excel.XlBordersIndex.xlInsideHorizontal
        xlInsideVertical = Interop.Excel.XlBordersIndex.xlInsideVertical
    End Enum

    Public Enum XlLineStyle
        xlContinuous = Interop.Excel.XlLineStyle.xlContinuous
        xlDash = Interop.Excel.XlLineStyle.xlDash
        xlDashDot = Interop.Excel.XlLineStyle.xlDashDot
        xlDashDotDot = Interop.Excel.XlLineStyle.xlDashDotDot
        xlDot = Interop.Excel.XlLineStyle.xlDot
        xlDouble = Interop.Excel.XlLineStyle.xlDouble
        xlLineStyleNone = Interop.Excel.XlLineStyle.xlLineStyleNone
        xlSlantDashDot = Interop.Excel.XlLineStyle.xlSlantDashDot
    End Enum

    Public Enum XlBorderWeight
        xlHairline = Interop.Excel.XlBorderWeight.xlHairline
        xlMedium = Interop.Excel.XlBorderWeight.xlMedium
        xlThick = Interop.Excel.XlBorderWeight.xlThick
        xlThin = Interop.Excel.XlBorderWeight.xlThin
    End Enum

    Public Enum XlCutCopyMode
        [False] = 0
        xlCopy = Interop.Excel.XlCutCopyMode.xlCopy
        xlCut = Interop.Excel.XlCutCopyMode.xlCut
    End Enum

    Public Enum XlDirection
        xlDown = Interop.Excel.XlDirection.xlDown
        xlToLeft = Interop.Excel.XlDirection.xlToLeft
        xlToRight = Interop.Excel.XlDirection.xlToRight
        xlUp = Interop.Excel.XlDirection.xlUp
    End Enum

    Public Enum XlSearchDirection
        xlNext = Interop.Excel.XlSearchDirection.xlNext
        xlPrevious = Interop.Excel.XlSearchDirection.xlPrevious
    End Enum

    Public Enum XlThemeFont
        xlThemeFontNone = Interop.Excel.XlThemeFont.xlThemeFontNone
        xlThemeFontMajor = Interop.Excel.XlThemeFont.xlThemeFontMajor
        xlThemeFontMinor = Interop.Excel.XlThemeFont.xlThemeFontMinor
    End Enum

    Public Enum XlPasteType
        xlPasteAll = Interop.Excel.XlPasteType.xlPasteAll
        xlPasteAllExceptBorders = Interop.Excel.XlPasteType.xlPasteAllExceptBorders
        xlPasteAllMergingConditionalFormats = Interop.Excel.XlPasteType.xlPasteAllMergingConditionalFormats
        xlPasteAllUsingSourceTheme = Interop.Excel.XlPasteType.xlPasteAllUsingSourceTheme
        xlPasteColumnWidths = Interop.Excel.XlPasteType.xlPasteColumnWidths
        xlPasteComments = Interop.Excel.XlPasteType.xlPasteComments
        xlPasteFormats = Interop.Excel.XlPasteType.xlPasteFormats
        xlPasteFormulas = Interop.Excel.XlPasteType.xlPasteFormulas
        xlPasteFormulasAndNumberFormats = Interop.Excel.XlPasteType.xlPasteFormulasAndNumberFormats
        xlPasteValidation = Interop.Excel.XlPasteType.xlPasteValidation
        xlPasteValues = Interop.Excel.XlPasteType.xlPasteValues
        xlPasteValuesAndNumberFormats = Interop.Excel.XlPasteType.xlPasteValuesAndNumberFormats
    End Enum

    Public Enum XlPasteSpecialOperation
        xlPasteSpecialOperationAdd = Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd
        xlPasteSpecialOperationDivide = Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationDivide
        xlPasteSpecialOperationMultiply = Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationMultiply
        xlPasteSpecialOperationNone = Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone
        xlPasteSpecialOperationSubtract = Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationSubtract
    End Enum

    Public Enum XlFindLookIn
        xlComments = Interop.Excel.XlFindLookIn.xlComments
        xlFormulas = Interop.Excel.XlFindLookIn.xlFormulas
        xlValues = Interop.Excel.XlFindLookIn.xlValues
    End Enum

    Public Enum XlLookAt
        xlPart = Interop.Excel.XlLookAt.xlPart
        xlWhole = Interop.Excel.XlLookAt.xlWhole
    End Enum

    Public Enum XlSearchOrder
        xlByColumns = Interop.Excel.XlSearchOrder.xlByColumns
        xlByRows = Interop.Excel.XlSearchOrder.xlByRows
    End Enum

    Public Enum XlInsertShiftDirection
        xlShiftDown = Interop.Excel.XlInsertShiftDirection.xlShiftDown
        xlShiftToRight = Interop.Excel.XlInsertShiftDirection.xlShiftToRight
    End Enum

    Public Enum XlInsertFormatOrigin
        xlFormatFromLeftOrAbove = Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove
        xlFormatFromRightOrBelow = Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow
    End Enum

    Public Enum XlSheetType
        xlChart = Interop.Excel.XlSheetType.xlChart
        xlDialogSheet = Interop.Excel.XlSheetType.xlDialogSheet
        xlExcel4IntlMacroSheet = Interop.Excel.XlSheetType.xlExcel4IntlMacroSheet
        xlExcel4MacroSheet = Interop.Excel.XlSheetType.xlExcel4MacroSheet
        xlWorksheet = Interop.Excel.XlSheetType.xlWorksheet
    End Enum

    Public Enum XlFileFormat
        xlCurrentPlatformText = Interop.Excel.XlFileFormat.xlCurrentPlatformText
        xlDBF2 = Interop.Excel.XlFileFormat.xlDBF2
        xlDBF3 = Interop.Excel.XlFileFormat.xlDBF3
        xlDBF4 = Interop.Excel.XlFileFormat.xlDBF4
        xlDIF = Interop.Excel.XlFileFormat.xlDIF
        xlExcel12 = Interop.Excel.XlFileFormat.xlExcel12
        xlExcel2 = Interop.Excel.XlFileFormat.xlExcel2
        xlExcel2FarEast = Interop.Excel.XlFileFormat.xlExcel2FarEast
        xlExcel3 = Interop.Excel.XlFileFormat.xlExcel3
        xlExcel4 = Interop.Excel.XlFileFormat.xlExcel4
        xlExcel4Workbook = Interop.Excel.XlFileFormat.xlExcel4Workbook
        xlExcel5 = Interop.Excel.XlFileFormat.xlExcel5
        xlExcel7 = Interop.Excel.XlFileFormat.xlExcel7
        xlExcel8 = Interop.Excel.XlFileFormat.xlExcel8
        xlExcel9795 = Interop.Excel.XlFileFormat.xlExcel9795
        xlHtml = Interop.Excel.XlFileFormat.xlHtml
        xlIntlAddIn = Interop.Excel.XlFileFormat.xlIntlAddIn
        xlIntlMacro = Interop.Excel.XlFileFormat.xlIntlMacro
        xlOpenDocumentSpreadsheet = Interop.Excel.XlFileFormat.xlOpenDocumentSpreadsheet
        xlOpenXMLAddIn = Interop.Excel.XlFileFormat.xlOpenXMLAddIn
        xlOpenXMLStrictWorkbook = Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook
        xlOpenXMLTemplate = Interop.Excel.XlFileFormat.xlOpenXMLTemplate
        xlOpenXMLTemplateMacroEnabled = Interop.Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled
        xlOpenXMLWorkbook = Interop.Excel.XlFileFormat.xlOpenXMLWorkbook
        xlOpenXMLWorkbookMacroEnabled = Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
        xlSYLK = Interop.Excel.XlFileFormat.xlSYLK
        xlTemplate = Interop.Excel.XlFileFormat.xlTemplate
        xlTemplate8 = Interop.Excel.XlFileFormat.xlTemplate8
        xlTextMac = Interop.Excel.XlFileFormat.xlTextMac
        xlTextMSDOS = Interop.Excel.XlFileFormat.xlTextMSDOS
        xlTextPrinter = Interop.Excel.XlFileFormat.xlTextPrinter
        xlTextWindows = Interop.Excel.XlFileFormat.xlTextWindows
        xlUnicodeText = Interop.Excel.XlFileFormat.xlUnicodeText
        xlWebArchive = Interop.Excel.XlFileFormat.xlWebArchive
        xlWJ2WD1 = Interop.Excel.XlFileFormat.xlWJ2WD1
        xlWJ3 = Interop.Excel.XlFileFormat.xlWJ3
        xlWJ3FJ3 = Interop.Excel.XlFileFormat.xlWJ3FJ3
        xlWK1 = Interop.Excel.XlFileFormat.xlWK1
        xlWK1ALL = Interop.Excel.XlFileFormat.xlWK1ALL
        xlWK1FMT = Interop.Excel.XlFileFormat.xlWK1FMT
        xlWK3 = Interop.Excel.XlFileFormat.xlWK3
        xlWK3FM3 = Interop.Excel.XlFileFormat.xlWK3FM3
        xlWK4 = Interop.Excel.XlFileFormat.xlWK4
        xlWKS = Interop.Excel.XlFileFormat.xlWKS
        xlWorkbookDefault = Interop.Excel.XlFileFormat.xlWorkbookDefault
        xlWorkbookNormal = Interop.Excel.XlFileFormat.xlWorkbookNormal
        xlWorks2FarEast = Interop.Excel.XlFileFormat.xlWorks2FarEast
        xlWQ1 = Interop.Excel.XlFileFormat.xlWQ1
        xlXMLSpreadsheet = Interop.Excel.XlFileFormat.xlXMLSpreadsheet
    End Enum

    Public Enum Constants
        xl3DBar = Interop.Excel.Constants.xl3DBar
        xl3DEffects1 = Interop.Excel.Constants.xl3DEffects1
        xl3DEffects2 = Interop.Excel.Constants.xl3DEffects2
        xl3DSurface = Interop.Excel.Constants.xl3DSurface
        xlAbove = Interop.Excel.Constants.xlAbove
        xlAccounting1 = Interop.Excel.Constants.xlAccounting1
        xlAccounting2 = Interop.Excel.Constants.xlAccounting2
        xlAccounting3 = Interop.Excel.Constants.xlAccounting3
        xlAccounting4 = Interop.Excel.Constants.xlAccounting4
        xlAdd = Interop.Excel.Constants.xlAdd
        xlAll = Interop.Excel.Constants.xlAll
        xlAllExceptBorders = Interop.Excel.Constants.xlAllExceptBorders
        xlAutomatic = Interop.Excel.Constants.xlAutomatic
        xlBar = Interop.Excel.Constants.xlBar
        xlBelow = Interop.Excel.Constants.xlBelow
        xlBidi = Interop.Excel.Constants.xlBidi
        xlBidiCalendar = Interop.Excel.Constants.xlBidiCalendar
        xlBoth = Interop.Excel.Constants.xlBoth
        xlBottom = Interop.Excel.Constants.xlBottom
        xlCascade = Interop.Excel.Constants.xlCascade
        xlCenter = Interop.Excel.Constants.xlCenter
        xlCenterAcrossSelection = Interop.Excel.Constants.xlCenterAcrossSelection
        xlChart4 = Interop.Excel.Constants.xlChart4
        xlChartSeries = Interop.Excel.Constants.xlChartSeries
        xlChartShort = Interop.Excel.Constants.xlChartShort
        xlChartTitles = Interop.Excel.Constants.xlChartTitles
        xlChecker = Interop.Excel.Constants.xlChecker
        xlCircle = Interop.Excel.Constants.xlCircle
        xlClassic1 = Interop.Excel.Constants.xlClassic1
        xlClassic2 = Interop.Excel.Constants.xlClassic2
        xlClassic3 = Interop.Excel.Constants.xlClassic3
        xlClosed = Interop.Excel.Constants.xlClosed
        xlColor1 = Interop.Excel.Constants.xlColor1
        xlColor2 = Interop.Excel.Constants.xlColor2
        xlColor3 = Interop.Excel.Constants.xlColor3
        xlColumn = Interop.Excel.Constants.xlColumn
        xlCombination = Interop.Excel.Constants.xlCombination
        xlComplete = Interop.Excel.Constants.xlComplete
        xlConstants = Interop.Excel.Constants.xlConstants
        xlContents = Interop.Excel.Constants.xlContents
        xlContext = Interop.Excel.Constants.xlContext
        xlCorner = Interop.Excel.Constants.xlCorner
        xlCrissCross = Interop.Excel.Constants.xlCrissCross
        xlCross = Interop.Excel.Constants.xlCross
        xlCustom = Interop.Excel.Constants.xlCustom
        xlDebugCodePane = Interop.Excel.Constants.xlDebugCodePane
        xlDefaultAutoFormat = Interop.Excel.Constants.xlDefaultAutoFormat
        xlDesktop = Interop.Excel.Constants.xlDesktop
        xlDiamond = Interop.Excel.Constants.xlDiamond
        xlDirect = Interop.Excel.Constants.xlDirect
        xlDistributed = Interop.Excel.Constants.xlDistributed
        xlDivide = Interop.Excel.Constants.xlDivide
        xlDoubleAccounting = Interop.Excel.Constants.xlDoubleAccounting
        xlDoubleClosed = Interop.Excel.Constants.xlDoubleClosed
        xlDoubleOpen = Interop.Excel.Constants.xlDoubleOpen
        xlDoubleQuote = Interop.Excel.Constants.xlDoubleQuote
        xlDrawingObject = Interop.Excel.Constants.xlDrawingObject
        xlEntireChart = Interop.Excel.Constants.xlEntireChart
        xlExcelMenus = Interop.Excel.Constants.xlExcelMenus
        xlExtended = Interop.Excel.Constants.xlExtended
        xlFill = Interop.Excel.Constants.xlFill
        xlFirst = Interop.Excel.Constants.xlFirst
        xlFixedValue = Interop.Excel.Constants.xlFixedValue
        xlFloating = Interop.Excel.Constants.xlFloating
        xlFormats = Interop.Excel.Constants.xlFormats
        xlFormula = Interop.Excel.Constants.xlFormula
        xlFullScript = Interop.Excel.Constants.xlFullScript
        xlGeneral = Interop.Excel.Constants.xlGeneral
        xlGray16 = Interop.Excel.Constants.xlGray16
        xlGray25 = Interop.Excel.Constants.xlGray25
        xlGray50 = Interop.Excel.Constants.xlGray50
        xlGray75 = Interop.Excel.Constants.xlGray75
        xlGray8 = Interop.Excel.Constants.xlGray8
        xlGregorian = Interop.Excel.Constants.xlGregorian
        xlGrid = Interop.Excel.Constants.xlGrid
        xlGridline = Interop.Excel.Constants.xlGridline
        xlHigh = Interop.Excel.Constants.xlHigh
        xlHindiNumerals = Interop.Excel.Constants.xlHindiNumerals
        xlIcons = Interop.Excel.Constants.xlIcons
        xlImmediatePane = Interop.Excel.Constants.xlImmediatePane
        xlInside = Interop.Excel.Constants.xlInside
        xlInteger = Interop.Excel.Constants.xlInteger
        xlJustify = Interop.Excel.Constants.xlJustify
        xlLast = Interop.Excel.Constants.xlLast
        xlLastCell = Interop.Excel.Constants.xlLastCell
        xlLatin = Interop.Excel.Constants.xlLatin
        xlLeft = Interop.Excel.Constants.xlLeft
        xlLeftToRight = Interop.Excel.Constants.xlLeftToRight
        xlLightDown = Interop.Excel.Constants.xlLightDown
        xlLightHorizontal = Interop.Excel.Constants.xlLightHorizontal
        xlLightUp = Interop.Excel.Constants.xlLightUp
        xlLightVertical = Interop.Excel.Constants.xlLightVertical
        xlList1 = Interop.Excel.Constants.xlList1
        xlList2 = Interop.Excel.Constants.xlList2
        xlList3 = Interop.Excel.Constants.xlList3
        xlLocalFormat1 = Interop.Excel.Constants.xlLocalFormat1
        xlLocalFormat2 = Interop.Excel.Constants.xlLocalFormat2
        xlLogicalCursor = Interop.Excel.Constants.xlLogicalCursor
        xlLong = Interop.Excel.Constants.xlLong
        xlLotusHelp = Interop.Excel.Constants.xlLotusHelp
        xlLow = Interop.Excel.Constants.xlLow
        xlLTR = Interop.Excel.Constants.xlLTR
        xlMacrosheetCell = Interop.Excel.Constants.xlMacrosheetCell
        xlManual = Interop.Excel.Constants.xlManual
        xlMaximum = Interop.Excel.Constants.xlMaximum
        xlMinimum = Interop.Excel.Constants.xlMinimum
        xlMinusValues = Interop.Excel.Constants.xlMinusValues
        xlMixed = Interop.Excel.Constants.xlMixed
        xlMixedAuthorizedScript = Interop.Excel.Constants.xlMixedAuthorizedScript
        xlMixedScript = Interop.Excel.Constants.xlMixedScript
        xlModule = Interop.Excel.Constants.xlModule
        xlMultiply = Interop.Excel.Constants.xlMultiply
        xlNarrow = Interop.Excel.Constants.xlNarrow
        xlNextToAxis = Interop.Excel.Constants.xlNextToAxis
        xlNoDocuments = Interop.Excel.Constants.xlNoDocuments
        xlNone = Interop.Excel.Constants.xlNone
        xlNotes = Interop.Excel.Constants.xlNotes
        xlOff = Interop.Excel.Constants.xlOff
        xlOn = Interop.Excel.Constants.xlOn
        xlOpaque = Interop.Excel.Constants.xlOpaque
        xlOpen = Interop.Excel.Constants.xlOpen
        xlOutside = Interop.Excel.Constants.xlOutside
        xlPartial = Interop.Excel.Constants.xlPartial
        xlPartialScript = Interop.Excel.Constants.xlPartialScript
        xlPercent = Interop.Excel.Constants.xlPercent
        xlPlus = Interop.Excel.Constants.xlPlus
        xlPlusValues = Interop.Excel.Constants.xlPlusValues
        xlReference = Interop.Excel.Constants.xlReference
        xlRight = Interop.Excel.Constants.xlRight
        xlRTL = Interop.Excel.Constants.xlRTL
        xlScale = Interop.Excel.Constants.xlScale
        xlSemiautomatic = Interop.Excel.Constants.xlSemiautomatic
        xlSemiGray75 = Interop.Excel.Constants.xlSemiGray75
        xlShort = Interop.Excel.Constants.xlShort
        xlShowLabel = Interop.Excel.Constants.xlShowLabel
        xlShowLabelAndPercent = Interop.Excel.Constants.xlShowLabelAndPercent
        xlShowPercent = Interop.Excel.Constants.xlShowPercent
        xlShowValue = Interop.Excel.Constants.xlShowValue
        xlSimple = Interop.Excel.Constants.xlSimple
        xlSingle = Interop.Excel.Constants.xlSingle
        xlSingleAccounting = Interop.Excel.Constants.xlSingleAccounting
        xlSingleQuote = Interop.Excel.Constants.xlSingleQuote
        xlSolid = Interop.Excel.Constants.xlSolid
        xlSquare = Interop.Excel.Constants.xlSquare
        xlStar = Interop.Excel.Constants.xlStar
        xlStError = Interop.Excel.Constants.xlStError
        xlStrict = Interop.Excel.Constants.xlStrict
        xlSubtract = Interop.Excel.Constants.xlSubtract
        xlSystem = Interop.Excel.Constants.xlSystem
        xlTextBox = Interop.Excel.Constants.xlTextBox
        xlTiled = Interop.Excel.Constants.xlTiled
        xlTitleBar = Interop.Excel.Constants.xlTitleBar
        xlToolbar = Interop.Excel.Constants.xlToolbar
        xlToolbarButton = Interop.Excel.Constants.xlToolbarButton
        xlTop = Interop.Excel.Constants.xlTop
        xlTopToBottom = Interop.Excel.Constants.xlTopToBottom
        xlTransparent = Interop.Excel.Constants.xlTransparent
        xlTriangle = Interop.Excel.Constants.xlTriangle
        xlVeryHidden = Interop.Excel.Constants.xlVeryHidden
        xlVisible = Interop.Excel.Constants.xlVisible
        xlVisualCursor = Interop.Excel.Constants.xlVisualCursor
        xlWatchPane = Interop.Excel.Constants.xlWatchPane
        xlWide = Interop.Excel.Constants.xlWide
        xlWorkbookTab = Interop.Excel.Constants.xlWorkbookTab
        xlWorksheet4 = Interop.Excel.Constants.xlWorksheet4
        xlWorksheetCell = Interop.Excel.Constants.xlWorksheetCell
        xlWorksheetShort = Interop.Excel.Constants.xlWorksheetShort
    End Enum
#End Region
End Namespace

''' ****************************************************************
''' <summary>
''' ExcelCOM操作ユーティリティクラス
''' </summary>
''' ****************************************************************
Public Class ExcelWrapper
    ''' <summary>COMオブジェクトID発番用同期オブジェクト</summary>
    Public Shared m_objSeqLock As New Object
    ''' <summary>COMオブジェクトID発番用ディクショナリ</summary>
    Private Shared m_colComId As New Dictionary(Of String, Integer)

    ''' <summary>生成済みアプリケーションのリスト</summary>
    Private Shared m_colApp As New List(Of Excel.Application)

    ''' ****************************************************************
    ''' <summary>
    ''' Excelの新しいインスタンスを起動します。
    ''' </summary>
    ''' <param name="Visible">アプリケーションのウィンドウを表示するにはTrueを指定します</param>
    ''' <returns>Exce.Applicationオブジェクト</returns>
    ''' ****************************************************************
    Public Shared Function CreateInstance(ByVal Visible As Boolean) As Excel.Application
        Dim objApp As New Excel.Application()
        objApp.Visible = Visible

        m_colApp.Add(objApp)
        Return objApp
    End Function

    ''' ****************************************************************
    ''' <summary>
    ''' オープンするファイルを指定してExcelの新しいインスタンスを起動します。
    ''' </summary>
    ''' <param name="File">オープンするファイルのパスを指定します</param>
    ''' <param name="[ReadOnly]">読み取り専用としてファイルを開くにはTrueを指定します</param>
    ''' <param name="Visible">アプリケーションのウィンドウを表示するにはTrueを指定します</param>
    ''' <returns>Exce.Applicationオブジェクト</returns>
    ''' ****************************************************************
    Public Shared Function CreateInstance(ByVal File As String, ByVal [ReadOnly] As Boolean, ByVal Visible As Boolean) As Excel.Application
        Dim objApp As Excel.Application = ExcelWrapper.CreateInstance(Visible)
        objApp.Workbooks.Open(File, [ReadOnly])

        Return objApp
    End Function

    ''' ================================================================
    ''' <summary>
    ''' COMオブジェクト用に新しいIDを発番します。
    ''' IDは毎秒最大999,999,999まで発番可能ですがそれだけのオブジェクトを
    ''' 保持するためのメモリ管理は行わないのでご注意ください。
    ''' </summary>
    ''' <returns>COMオブジェクトID</returns>
    ''' ================================================================
    Friend Shared Function GetNewId() As String
        Dim strRes As String = String.Empty
        Const FORMAT As String = "000000000"

        SyncLock m_objSeqLock
            Dim strTime As String = Date.Now.ToString("yyyyMMddHHmmss")
            Dim intId As Integer = 0
            If m_colComId.ContainsKey(strTime) Then
                intId = m_colComId(strTime)
            Else
                m_colComId.Clear()
            End If
            intId += 1
            m_colComId(strTime) = intId
            Dim strNo As String = intId.ToString(FORMAT)
            If strNo.Length > FORMAT.Length Then
                Throw New OverflowException("Failed to issue new COM Object ID.")
            End If

            strRes = strTime & strNo
        End SyncLock

        Return strRes
    End Function

    ''' ================================================================
    ''' <summary>
    ''' COMオブジェクトを開放します。
    ''' ※通常はオブジェクトを登録して使いまわすので最後にまとめて開放しますが、
    ''' 処理の中でやむを得ずCOMオブジェクトを所有してしまった場合は個別の開放
    ''' してください。
    ''' </summary>
    ''' <param name="ComObject"></param>
    ''' ================================================================
    Friend Shared Sub ReleaseComObject(ParamArray ByVal ComObject() As Object)
        For Each objCom As Object In ComObject
            If Not objCom Is Nothing Then
                FinalReleaseComObject(objCom)
                objCom = Nothing
            End If
        Next
    End Sub

    ''' ================================================================
    ''' <summary>
    ''' Excel終了時にアプリケーションオブジェクトをリストから削除するために呼び出します。
    ''' </summary>
    ''' <param name="Application"></param>
    ''' ================================================================
    Friend Shared Sub Quit(ByRef Application As Excel.Application)
        If m_colApp.Contains(Application) Then
            m_colApp.Remove(Application)
        End If
    End Sub

#Region "Friend Class Address"
    ''' ================================================================
    ''' <summary>
    ''' Rangeオブジェクトのアドレスを行/列・開始/終了の要素に分解します。
    ''' Rangeに複数の範囲が含まれる場合は正しく分解できないのでご注意ください。
    ''' </summary>
    ''' ================================================================
    Friend Class Address
        Friend Sub New(ByVal AddressA1 As String)
            Dim strSe() As String = AddressA1.Split(":"c)
            For intIdx As Integer = 0 To strSe.Length - 1
                Me.SplitColAndRow(strSe(intIdx), intIdx)
            Next
        End Sub

        ''' ----------------------------------------------------------------
        ''' <summary>
        ''' 要素に分解します。
        ''' </summary>
        ''' <param name="AddressElement"></param>
        ''' <param name="Index"></param>
        ''' ----------------------------------------------------------------
        Private Sub SplitColAndRow(ByVal AddressElement As String, ByVal Index As Integer)
            Dim strCr() As String = AddressElement.Split("$"c)
            If Index = 0 Then
                m_strStartCol = strCr(1)
                m_intStartRow = Integer.Parse(strCr(2))
            Else
                m_strEndCol = strCr(1)
                m_intEndRow = Integer.Parse(strCr(2))
            End If
        End Sub

        ''' <summary>開始列名</summary>
        Friend ReadOnly Property StartColumn As String
            Get
                Return m_strStartCol
            End Get
        End Property
        Private m_strStartCol As String

        ''' <summary>開始列インデックス</summary>
        Friend ReadOnly Property StartRow As Integer
            Get
                Return m_intStartRow
            End Get
        End Property
        Private m_intStartRow As Integer

        ''' <summary>終了列名</summary>
        Friend ReadOnly Property EndColumn As String
            Get
                Return m_strEndCol
            End Get
        End Property
        Private m_strEndCol As String

        ''' <summary>終了列インデックス</summary>
        Friend ReadOnly Property EndRow As Integer
            Get
                Return m_intEndRow
            End Get
        End Property
        Private m_intEndRow As Integer
    End Class
#End Region

End Class
