Option Explicit

Private Const CLASS_NAME As String = "modKeyValuePdfExport"
Private Const PDF_HEADER As String = "%PDF-1.4"
Private Const PDF_PAGE_WIDTH As Double = 1368#
Private Const PDF_PAGE_HEIGHT As Double = 684#
Private Const LOGO_X As Double = 40.52
Private Const LOGO_Y As Double = 654#
Private Const LOGO_WIDTH As Double = 85.52
Private Const LOGO_HEIGHT As Double = 17.96

Private Const TABLE_LEFT_X As Double = 40.52
Private Const TABLE_DIVIDER_X As Double = 365.04
Private Const TABLE_RIGHT_X As Double = 1323.04
Private Const TABLE_TOP_Y As Double = 642#
Private Const TABLE_BOTTOM_Y As Double = 92#
Private Const ROW_HEIGHT As Double = 22#
Private Const ROW_COUNT As Long = 25
Private Const TEXT_LEFT_PADDING As Double = 8#
Private Const TEXT_BASELINE_FROM_BOTTOM As Double = 7#
Private Const KEY_FONT_SIZE As Double = 10#
Private Const VALUE_FONT_SIZE As Double = 11#
Private Const VALUE_MIN_FONT_SIZE As Double = 8#
Private Const LINE_WIDTH As Double = 0.6
Private Const FIRST_DATA_ROW As Long = 7
Private Const VALUE_WIDTH_FACTOR As Double = 0.55
Private Const VALUE_RIGHT_PADDING As Double = 8#
Private Const PAGE_NUMBER_FONT_SIZE As Double = 9#
Private Const PAGE_NUMBER_Y As Double = 40#
Private Const PAGE_NUMBER_RIGHT_X As Double = 1323.04
Private Const PAGE_NUMBER_WIDTH_FACTOR As Double = 0.5
Private Const FOOTER_X As Double = 40.52
Private Const FOOTER_Y As Double = 40#
Private Const FOOTER_FONT_SIZE As Double = 9#
Private Const FOOTER_TEXT As String = "Diese Abrechnung wurde maschinell erstellt und ist ohne eigenhaendige Unterschrift gueltig"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-08-13
' Parameters:    targetPath As String, wksSource As Excel.Worksheet, logoPath As String
' Returns:       Boolean
' Description:   Generates a parser-friendly PDF 1.4 with exactly one transaction
'                per page and 25 fixed KEY/VALUE rows. The existing report
'                worksheet remains the single source of business-calculated data.
'-------------------------------------------------------------------------------
Public Function GenerateKeyValuePdfFromWorksheet(ByVal targetPath As String, ByVal wksSource As Excel.Worksheet, ByVal logoPath As String) As Boolean

    Const METHOD_NAME As String = "GenerateKeyValuePdfFromWorksheet"
    Dim arrBottomRows() As Long
    Dim arrTopRows() As Long
    Dim blnFileOpened As Boolean
    Dim colPageContents As Collection
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFileNumber As Long
    Dim lngLogoComponents As Long
    Dim lngLogoHeight As Long
    Dim lngLogoWidth As Long
    Dim lngTradeCount As Long
    Dim strDocument As String
    Dim strLogoHex As String

    On Error GoTo ErrorHandler

    GenerateKeyValuePdfFromWorksheet = False

    If wksSource Is Nothing Then
        Err.Raise 1001, METHOD_NAME, "The completed report worksheet is required."
    End If

    If Not ValidateTargetPath(targetPath) Then
        Err.Raise 1020, METHOD_NAME, "The target PDF path is invalid or inaccessible."
    End If

    wksSource.Calculate

    If Not CollectTransactionRows(wksSource, arrTopRows, arrBottomRows, lngTradeCount) Then
        Err.Raise 1001, METHOD_NAME, "No transaction row pairs were found in the report worksheet."
    End If

    If Not PrepareLogoImageFromPath(logoPath, strLogoHex, lngLogoWidth, lngLogoHeight, lngLogoComponents) Then
        Err.Raise 1013, METHOD_NAME, "The report logo could not be prepared as a JPEG image."
    End If

    If Not BuildAllPageContents( wksSource, arrTopRows, arrBottomRows, lngTradeCount, colPageContents) Then

        Err.Raise 1024, METHOD_NAME, "The key-value PDF page contents could not be built."
    End If

    If Not BuildPdfDocument(colPageContents, strLogoHex, lngLogoWidth, lngLogoHeight, lngLogoComponents, strDocument) Then
        Err.Raise 1024, METHOD_NAME, "The PDF object structure or XREF table could not be built."
    End If

    If Len(Dir$(targetPath)) > 0 Then
        Kill targetPath
    End If

    lngFileNumber = FreeFile
    Open targetPath For Output Access Write Lock Write As #lngFileNumber
    blnFileOpened = True
    Print #lngFileNumber, strDocument;
    Close #lngFileNumber
    blnFileOpened = False

    If Len(Dir$(targetPath)) = 0 Then
        Err.Raise 1011, METHOD_NAME, "The key-value PDF file was not created."
    End If

    If FileLen(targetPath) <= 0 Then
        Err.Raise 1011, METHOD_NAME, "The key-value PDF file is empty."
    End If

    If Not HasPdf14Header(targetPath) Then
        Err.Raise 1024, METHOD_NAME, "The generated file does not start with the PDF 1.4 header."
    End If

    GenerateKeyValuePdfFromWorksheet = True

ExitFunction:
    If blnFileOpened Then
        CloseFileHandle lngFileNumber
    End If

    If Not GenerateKeyValuePdfFromWorksheet Then
        DeleteFileIfPresent targetPath
    End If

    Set colPageContents = Nothing
    Set wksSource = Nothing
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    GenerateKeyValuePdfFromWorksheet = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "targetPath;logoPath;tradeCount", targetPath, logoPath, lngTradeCount
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Builds one content stream per transaction. No page contains more than one trade.
'-------------------------------------------------------------------------------
Private Function BuildAllPageContents( ByVal wksSource As Excel.Worksheet, ByRef arrTopRows() As Long, ByRef arrBottomRows() As Long, ByVal lngTradeCount As Long, ByRef colPageContents As Collection) As Boolean

    Const METHOD_NAME As String = "BuildAllPageContents"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngTradeIndex As Long
    Dim strPageContent As String

    On Error GoTo ErrorHandler

    BuildAllPageContents = False
    Set colPageContents = New Collection

    If lngTradeCount <= 0 Then
        Err.Raise 1001, METHOD_NAME, "At least one transaction is required."
    End If

    For lngTradeIndex = 1 To lngTradeCount
        strPageContent = BuildSingleTransactionPageContent(wksSource, arrTopRows(lngTradeIndex), arrBottomRows(lngTradeIndex), lngTradeIndex, lngTradeCount)

        If Len(strPageContent) = 0 Then
            Err.Raise 1024, METHOD_NAME, "A key-value page content stream was empty for transaction " & CStr(lngTradeIndex) & "."
        End If

        colPageContents.Add strPageContent
    Next lngTradeIndex

    BuildAllPageContents = (colPageContents.Count = lngTradeCount)

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    BuildAllPageContents = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "tradeCount;tradeIndex", lngTradeCount, lngTradeIndex
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Builds a single fixed 25-row KEY/VALUE page. Lines are emitted once using m/l/S.
'-------------------------------------------------------------------------------
Private Function BuildSingleTransactionPageContent(ByVal wksSource As Excel.Worksheet, ByVal lngTopRow As Long, ByVal lngBottomRow As Long, ByVal lngPageNumber As Long, ByVal lngPageCount As Long) As String

    Const METHOD_NAME As String = "BuildSingleTransactionPageContent"
    Dim arrKeys As Variant
    Dim arrColumns As Variant
    Dim arrRowOffsets As Variant
    Dim dblBaselineY As Double
    Dim dblCurrentY As Double
    Dim dblFontSize As Double
    Dim dblValueAvailableWidth As Double
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim lngSourceRow As Long
    Dim dblPageNumberX As Double
    Dim strContent As String
    Dim strPageNumber As String
    Dim strKey As String
    Dim strValue As String

    On Error GoTo ErrorHandler

    BuildSingleTransactionPageContent = vbNullString

    arrKeys = Array( "Depotnr.", "Fondsbezeichnung", "Geschaeftsart", "Schlusstag", "Valutatag", "WKN", "ISIN", "Wertpapierbezeichnung", "Nominale/Stuecke", "Whg (Nominale/Stuecke)", "Kurs", "Stueckzinsen", "Transaktionsnr.", "Maklergebuehren", "Whg (Maklergebuehren)", "Steuern", "Whg (Steuern)", "Abwicklungsprovision", "Whg (Abwicklungsprovision)", "Spesen", "Whg (Spesen)", "ausm. Betrag in Whg", "ausm. Betrag in Abrech. Whg", "Whg", "Devisenkurs")

    arrColumns = Array( 1, 2, 5, 7, 9, 10, 11, 13, 16, 18, 19, 20, 1, 2, 4, 5, 8, 9, 12, 13, 15, 16, 19, 20, 21)

    arrRowOffsets = Array( 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)

    If UBound(arrKeys) - LBound(arrKeys) + 1 <> ROW_COUNT Then
        Err.Raise 1024, METHOD_NAME, "The fixed KEY list does not contain exactly 25 rows."
    End If

    strContent = "q" & vbCrLf & PdfNumber(LOGO_WIDTH) & " 0 0 " & PdfNumber(LOGO_HEIGHT) & " " & PdfNumber(LOGO_X) & " " & PdfNumber(LOGO_Y) & " cm" & vbCrLf & "/Im0 Do" & vbCrLf & "Q" & vbCrLf
    strContent = strContent & "0 g" & vbCrLf & "0 G" & vbCrLf & PdfNumber(LINE_WIDTH) & " w" & vbCrLf

    ' Horizontal boundaries: 26 lines for 25 rows.
    For lngIndex = 0 To ROW_COUNT
        dblCurrentY = TABLE_TOP_Y - (CDbl(lngIndex) * ROW_HEIGHT)
        strContent = strContent & PdfNumber(TABLE_LEFT_X) & " " & PdfNumber(dblCurrentY) & " m " & PdfNumber(TABLE_RIGHT_X) & " " & PdfNumber(dblCurrentY) & " l S" & vbCrLf
    Next lngIndex

    ' Three vertical rules. Each is drawn once, so there are no overlapping cell borders.
    strContent = strContent & PdfNumber(TABLE_LEFT_X) & " " & PdfNumber(TABLE_BOTTOM_Y) & " m " & PdfNumber(TABLE_LEFT_X) & " " & PdfNumber(TABLE_TOP_Y) & " l S" & vbCrLf & PdfNumber(TABLE_DIVIDER_X) & " " & PdfNumber(TABLE_BOTTOM_Y) & " m " & PdfNumber(TABLE_DIVIDER_X) & " " & PdfNumber(TABLE_TOP_Y) & " l S" & vbCrLf & PdfNumber(TABLE_RIGHT_X) & " " & PdfNumber(TABLE_BOTTOM_Y) & " m " & PdfNumber(TABLE_RIGHT_X) & " " & PdfNumber(TABLE_TOP_Y) & " l S" & vbCrLf

    dblValueAvailableWidth = TABLE_RIGHT_X - TABLE_DIVIDER_X - TEXT_LEFT_PADDING - VALUE_RIGHT_PADDING

    For lngIndex = 0 To ROW_COUNT - 1
        strKey = CStr(arrKeys(lngIndex))

        If CLng(arrRowOffsets(lngIndex)) = 0 Then
            lngSourceRow = lngTopRow
        Else
            lngSourceRow = lngBottomRow
        End If

        strValue = GetWorksheetDisplayText( wksSource.Cells(lngSourceRow, CLng(arrColumns(lngIndex))))

        dblBaselineY = TABLE_TOP_Y - (CDbl(lngIndex + 1) * ROW_HEIGHT) + TEXT_BASELINE_FROM_BOTTOM

        strContent = strContent & BuildTextCommand( "/F0", KEY_FONT_SIZE, TABLE_LEFT_X + TEXT_LEFT_PADDING, dblBaselineY, strKey)

        dblFontSize = ResolveValueFontSize(strValue, dblValueAvailableWidth)

        strContent = strContent & BuildTextCommand( "/F1", dblFontSize, TABLE_DIVIDER_X + TEXT_LEFT_PADDING, dblBaselineY, strValue)
    Next lngIndex

    If lngPageNumber = lngPageCount Then
        strContent = strContent & BuildTextCommand("/F0", FOOTER_FONT_SIZE, FOOTER_X, FOOTER_Y, FOOTER_TEXT)
    End If

    strPageNumber = CStr(lngPageNumber) & "/" & CStr(lngPageCount)
    dblPageNumberX = PAGE_NUMBER_RIGHT_X - (CDbl(Len(strPageNumber)) * PAGE_NUMBER_FONT_SIZE * PAGE_NUMBER_WIDTH_FACTOR)
    strContent = strContent & BuildTextCommand("/F1", PAGE_NUMBER_FONT_SIZE, dblPageNumberX, PAGE_NUMBER_Y, strPageNumber)

    BuildSingleTransactionPageContent = strContent

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    BuildSingleTransactionPageContent = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "topRow;bottomRow;pageNumber;pageCount;keyIndex", lngTopRow, lngBottomRow, lngPageNumber, lngPageCount, lngIndex
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Returns the display text already calculated/formatted in wsOut.
'-------------------------------------------------------------------------------
Private Function GetWorksheetDisplayText(ByVal rngCell As Excel.Range) As String
    Const METHOD_NAME As String = "GetWorksheetDisplayText"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim strText As String

    On Error GoTo ErrorHandler

    GetWorksheetDisplayText = vbNullString

    If rngCell Is Nothing Then GoTo ExitFunction
    If IsError(rngCell.Value) Then GoTo ExitFunction

    strText = CStr(rngCell.Text)

    ' Excel can show #### only because of a worksheet display width limitation.
    ' The parser PDF should still receive the calculated value.
    If Len(strText) > 0 And Replace$(strText, "#", vbNullString) = vbNullString Then
        strText = CStr(rngCell.Value2)
    End If

    GetWorksheetDisplayText = Trim$(strText)

ExitFunction:
    Set rngCell = Nothing
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    GetWorksheetDisplayText = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Keeps VALUE on one physical text line. With the very wide value column this is
' normally 11 pt; long descriptions are reduced only as much as necessary.
'-------------------------------------------------------------------------------
Private Function ResolveValueFontSize( ByVal strText As String, ByVal dblAvailableWidth As Double) As Double

    Const METHOD_NAME As String = "ResolveValueFontSize"
    Dim dblEstimatedWidth As Double
    Dim dblResolvedSize As Double
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    dblResolvedSize = VALUE_FONT_SIZE

    If Len(strText) > 0 Then
        dblEstimatedWidth = CDbl(Len(strText)) * VALUE_FONT_SIZE * VALUE_WIDTH_FACTOR

        If dblEstimatedWidth > dblAvailableWidth And dblEstimatedWidth > 0# Then
            dblResolvedSize = VALUE_FONT_SIZE * dblAvailableWidth / dblEstimatedWidth
        End If
    End If

    If dblResolvedSize < VALUE_MIN_FONT_SIZE Then dblResolvedSize = VALUE_MIN_FONT_SIZE
    If dblResolvedSize > VALUE_FONT_SIZE Then dblResolvedSize = VALUE_FONT_SIZE

    ResolveValueFontSize = dblResolvedSize

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    ResolveValueFontSize = VALUE_FONT_SIZE

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "textLength;availableWidth", Len(strText), dblAvailableWidth
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' One text show per KEY and one per VALUE. Hex text keeps the full PDF ASCII-only.
'-------------------------------------------------------------------------------
Private Function BuildTextCommand( ByVal strFont As String, ByVal dblFontSize As Double, ByVal dblX As Double, ByVal dblY As Double, ByVal strText As String) As String

    Const METHOD_NAME As String = "BuildTextCommand"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    BuildTextCommand = "BT " & strFont & " " & PdfNumber(dblFontSize) & " Tf " & "1 0 0 1 " & PdfNumber(dblX) & " " & PdfNumber(dblY) & " Tm " & "<" & EncodePdfHex(strText) & "> Tj ET" & vbCrLf

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    BuildTextCommand = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "font;fontSize;x;y", strFont, dblFontSize, dblX, dblY
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Encodes VBA Unicode text as Windows ANSI bytes and then as PDF hex text.
'-------------------------------------------------------------------------------
Private Function EncodePdfHex(ByVal strText As String) As String
    Const METHOD_NAME As String = "EncodePdfHex"
    Dim arrBytes() As Byte
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngByte As Long
    Dim lngIndex As Long
    Dim strHex As String

    On Error GoTo ErrorHandler

    EncodePdfHex = vbNullString

    If Len(strText) = 0 Then GoTo ExitFunction

    ' Same ANSI byte conversion strategy as the existing raw Oracle module.
    arrBytes = StrConv(strText, vbFromUnicode)

    For lngIndex = LBound(arrBytes) To UBound(arrBytes)
        lngByte = CLng(arrBytes(lngIndex))
        strHex = strHex & Right$("0" & Hex$(lngByte), 2)
    Next lngIndex

    EncodePdfHex = strHex

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    EncodePdfHex = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "textLength", Len(strText)
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Collects the same two-row transaction blocks already built by modTradeversand.
' Separator rows and the footer are ignored.
'-------------------------------------------------------------------------------
Private Function CollectTransactionRows( ByVal wksSource As Excel.Worksheet, ByRef arrTopRows() As Long, ByRef arrBottomRows() As Long, ByRef lngTradeCount As Long) As Boolean

    Const METHOD_NAME As String = "CollectTransactionRows"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCapacity As Long
    Dim lngLastRow As Long
    Dim lngRow As Long

    On Error GoTo ErrorHandler

    CollectTransactionRows = False
    lngTradeCount = 0
    lngLastRow = GetLastUsedRow(wksSource)

    If lngLastRow < FIRST_DATA_ROW + 1 Then GoTo ExitFunction

    lngCapacity = 16
    ReDim arrTopRows(1 To lngCapacity)
    ReDim arrBottomRows(1 To lngCapacity)

    lngRow = FIRST_DATA_ROW

    Do While lngRow <= lngLastRow - 1
        If IsTransactionRowPair(wksSource, lngRow, lngLastRow) Then
            lngTradeCount = lngTradeCount + 1

            If lngTradeCount > lngCapacity Then
                lngCapacity = lngCapacity * 2
                ReDim Preserve arrTopRows(1 To lngCapacity)
                ReDim Preserve arrBottomRows(1 To lngCapacity)
            End If

            arrTopRows(lngTradeCount) = lngRow
            arrBottomRows(lngTradeCount) = lngRow + 1
            lngRow = lngRow + 2
        Else
            lngRow = lngRow + 1
        End If
    Loop

    If lngTradeCount = 0 Then GoTo ExitFunction

    ReDim Preserve arrTopRows(1 To lngTradeCount)
    ReDim Preserve arrBottomRows(1 To lngTradeCount)
    CollectTransactionRows = True

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    CollectTransactionRows = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "lastRow;tradeCount", lngLastRow, lngTradeCount
    errorManager.save
    Resume ExitFunction
End Function

Private Function IsTransactionRowPair( ByVal wksSource As Excel.Worksheet, ByVal lngTopRow As Long, ByVal lngLastRow As Long) As Boolean

    Const METHOD_NAME As String = "IsTransactionRowPair"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim strBottomKey As String
    Dim strTopKey As String

    On Error GoTo ErrorHandler

    IsTransactionRowPair = False

    If lngTopRow < FIRST_DATA_ROW Or lngTopRow + 1 > lngLastRow Then GoTo ExitFunction

    strTopKey = Trim$(CStr(wksSource.Cells(lngTopRow, 1).Text))
    strBottomKey = Trim$(CStr(wksSource.Cells(lngTopRow + 1, 1).Text))

    IsTransactionRowPair = (Len(strTopKey) > 0 And Len(strBottomKey) > 0)

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    IsTransactionRowPair = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "topRow;lastRow", lngTopRow, lngLastRow
    errorManager.save
    Resume ExitFunction
End Function

Private Function GetLastUsedRow(ByVal wksSource As Excel.Worksheet) As Long
    Const METHOD_NAME As String = "GetLastUsedRow"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim rngLast As Excel.Range

    On Error GoTo ErrorHandler

    GetLastUsedRow = 0

    Set rngLast = wksSource.Cells.Find( What:="*", After:=wksSource.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    If Not rngLast Is Nothing Then GetLastUsedRow = rngLast.Row

ExitFunction:
    Set rngLast = Nothing
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    GetLastUsedRow = 0

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Reads the configured JPG/JPEG logo directly from logoPath and prepares it as a
' PDF image XObject. No worksheet copy/paste or temporary Chart export is used.
' The same image object is reused on every transaction page.
'-------------------------------------------------------------------------------
Private Function PrepareLogoImageFromPath(ByVal logoPath As String, ByRef strLogoHex As String, ByRef lngLogoWidth As Long, ByRef lngLogoHeight As Long, ByRef lngLogoComponents As Long) As Boolean
    Const METHOD_NAME As String = "PrepareLogoImageFromPath"
    Dim arrBytes() As Byte
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDotPosition As Long
    Dim strExtension As String

    On Error GoTo ErrorHandler

    PrepareLogoImageFromPath = False
    strLogoHex = vbNullString
    lngLogoWidth = 0
    lngLogoHeight = 0
    lngLogoComponents = 0

    If Len(Trim$(logoPath)) = 0 Then
        Err.Raise 1001, METHOD_NAME, "The logo path is empty. Check bnp_logo and bnp_logo_name on the Tradeversand worksheet."
    End If

    If Len(Dir$(logoPath)) = 0 Then
        Err.Raise 1010, METHOD_NAME, "The configured logo file does not exist: " & logoPath
    End If

    lngDotPosition = InStrRev(logoPath, ".")
    If lngDotPosition > 0 Then
        strExtension = LCase$(Mid$(logoPath, lngDotPosition + 1))
    Else
        strExtension = vbNullString
    End If

    If strExtension <> "jpg" And strExtension <> "jpeg" Then
        Err.Raise 1013, METHOD_NAME, "The key-value RAW PDF logo must be a JPG or JPEG file. Configured file: " & logoPath
    End If

    arrBytes = ReadFileBytes(logoPath)

    If Not GetJpegInformation(arrBytes, lngLogoWidth, lngLogoHeight, lngLogoComponents) Then
        Err.Raise 1013, METHOD_NAME, "The configured logo is not a supported JPEG image: " & logoPath
    End If

    strLogoHex = BytesToAsciiHex(arrBytes)

    If Len(strLogoHex) <= 1 Then
        Err.Raise 1013, METHOD_NAME, "The configured logo contains no image data: " & logoPath
    End If

    PrepareLogoImageFromPath = True

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    PrepareLogoImageFromPath = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "logoPath", logoPath
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Reads a complete binary file into a byte array.
'-------------------------------------------------------------------------------
Private Function ReadFileBytes(ByVal filePath As String) As Byte()
    Const METHOD_NAME As String = "ReadFileBytes"
    Dim arrBytes() As Byte
    Dim blnFileOpened As Boolean
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFileNumber As Long
    Dim lngSize As Long

    On Error GoTo ErrorHandler

    If Len(Dir$(filePath)) = 0 Then
        Err.Raise 1010, METHOD_NAME, "The requested file does not exist."
    End If

    lngFileNumber = FreeFile
    Open filePath For Binary Access Read Lock Read As #lngFileNumber
    blnFileOpened = True
    lngSize = LOF(lngFileNumber)

    If lngSize <= 0 Then
        Err.Raise 1011, METHOD_NAME, "The requested file is empty."
    End If

    ReDim arrBytes(0 To lngSize - 1)
    Get #lngFileNumber, 1, arrBytes
    Close #lngFileNumber
    blnFileOpened = False
    ReadFileBytes = arrBytes

ExitFunction:
    If blnFileOpened Then
        CloseFileHandle lngFileNumber
    End If
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    ReDim arrBytes(0 To 0)
    ReadFileBytes = arrBytes

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "filePath", filePath
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Extracts JPEG dimensions and component count from a SOF marker.
'-------------------------------------------------------------------------------
Private Function GetJpegInformation(ByRef arrBytes() As Byte, ByRef lngWidth As Long, ByRef lngHeight As Long, ByRef lngComponents As Long) As Boolean
    Const METHOD_NAME As String = "GetJpegInformation"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMarker As Long
    Dim lngPosition As Long
    Dim lngSegmentLength As Long

    On Error GoTo ErrorHandler

    GetJpegInformation = False
    lngWidth = 0
    lngHeight = 0
    lngComponents = 0

    If UBound(arrBytes) < 10 Then GoTo ExitFunction
    If arrBytes(0) <> &HFF Or arrBytes(1) <> &HD8 Then GoTo ExitFunction

    lngPosition = 2

    Do While lngPosition <= UBound(arrBytes) - 8
        Do While lngPosition <= UBound(arrBytes) And arrBytes(lngPosition) <> &HFF
            lngPosition = lngPosition + 1
        Loop

        Do While lngPosition <= UBound(arrBytes) And arrBytes(lngPosition) = &HFF
            lngPosition = lngPosition + 1
        Loop

        If lngPosition > UBound(arrBytes) Then Exit Do

        lngMarker = CLng(arrBytes(lngPosition))
        lngPosition = lngPosition + 1

        If lngMarker = &HD9 Or lngMarker = &HDA Then Exit Do
        If lngMarker = &H1 Or (lngMarker >= &HD0 And lngMarker <= &HD7) Then GoTo ContinueLoop
        If lngPosition + 1 > UBound(arrBytes) Then Exit Do

        lngSegmentLength = CLng(arrBytes(lngPosition)) * 256 + CLng(arrBytes(lngPosition + 1))

        Select Case lngMarker
            Case &HC0, &HC1, &HC2, &HC3, &HC5, &HC6, &HC7, &HC9, &HCA, &HCB, &HCD, &HCE, &HCF
                If lngPosition + 7 <= UBound(arrBytes) Then
                    lngHeight = CLng(arrBytes(lngPosition + 3)) * 256 + CLng(arrBytes(lngPosition + 4))
                    lngWidth = CLng(arrBytes(lngPosition + 5)) * 256 + CLng(arrBytes(lngPosition + 6))
                    lngComponents = CLng(arrBytes(lngPosition + 7))
                    GetJpegInformation = (lngWidth > 0 And lngHeight > 0 And lngComponents > 0)
                    GoTo ExitFunction
                End If
        End Select

        If lngSegmentLength < 2 Then Exit Do
        lngPosition = lngPosition + lngSegmentLength
ContinueLoop:
    Loop

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    GetJpegInformation = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Converts JPEG bytes to an ASCIIHex stream terminated by greater-than.
'-------------------------------------------------------------------------------
Private Function BytesToAsciiHex(ByRef arrBytes() As Byte) As String
    Const METHOD_NAME As String = "BytesToAsciiHex"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim lngLength As Long
    Dim strHex As String

    On Error GoTo ErrorHandler

    BytesToAsciiHex = vbNullString
    lngLength = UBound(arrBytes) - LBound(arrBytes) + 1
    If lngLength <= 0 Then GoTo ExitFunction

    strHex = String$(lngLength * 2, "0")

    For lngIndex = 0 To lngLength - 1
        Mid$(strHex, (lngIndex * 2) + 1, 2) = Right$("0" & Hex$(arrBytes(LBound(arrBytes) + lngIndex)), 2)
    Next lngIndex

    BytesToAsciiHex = strHex & ">"

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    BytesToAsciiHex = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Creates a classic PDF 1.4 object/xref structure. Every stream is ASCII only,
' so VBA Len() equals the byte offset used in the xref table.
'-------------------------------------------------------------------------------
Private Function BuildPdfDocument(ByVal colPageContents As Collection, ByVal strLogoHex As String, ByVal lngLogoWidth As Long, ByVal lngLogoHeight As Long, ByVal lngLogoComponents As Long, ByRef strDocument As String) As Boolean

    Const METHOD_NAME As String = "BuildPdfDocument"
    Dim arrOffsets() As Long
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngContentObject As Long
    Dim lngObject As Long
    Dim lngObjectCount As Long
    Dim lngPage As Long
    Dim lngPageCount As Long
    Dim lngPageObject As Long
    Dim lngStartXref As Long
    Dim strColorSpace As String
    Dim strKids As String
    Dim strObject As String
    Dim strStream As String

    On Error GoTo ErrorHandler

    BuildPdfDocument = False
    strDocument = vbNullString

    If colPageContents Is Nothing Then
        Err.Raise 1001, METHOD_NAME, "Page contents are required."
    End If

    If colPageContents.Count <= 0 Then
        Err.Raise 1001, METHOD_NAME, "At least one page is required."
    End If

    If Len(strLogoHex) <= 1 Or lngLogoWidth <= 0 Or lngLogoHeight <= 0 Or lngLogoComponents <= 0 Then
        Err.Raise 1013, METHOD_NAME, "Valid JPEG logo data is required."
    End If

    lngPageCount = colPageContents.Count
    lngObjectCount = 5 + (lngPageCount * 2)
    ReDim arrOffsets(0 To lngObjectCount)

    For lngPage = 1 To colPageContents.Count
        lngPageObject = 6 + ((lngPage - 1) * 2)
        strKids = strKids & CStr(lngPageObject) & " 0 R "
    Next lngPage

    strDocument = PDF_HEADER & vbCrLf & "% Tradeversand key-value parser-safe PDF" & vbCrLf

    lngObject = 1
    arrOffsets(lngObject) = Len(strDocument)
    strObject = "1 0 obj" & vbCrLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbCrLf & "endobj" & vbCrLf
    strDocument = strDocument & strObject

    lngObject = 2
    arrOffsets(lngObject) = Len(strDocument)
    strObject = "2 0 obj" & vbCrLf & "<< /Type /Pages /Count " & CStr(colPageContents.Count) & " /Kids [" & Trim$(strKids) & "] >>" & vbCrLf & "endobj" & vbCrLf
    strDocument = strDocument & strObject

    lngObject = 3
    arrOffsets(lngObject) = Len(strDocument)
    strObject = "3 0 obj" & vbCrLf & "<< /Type /Font /Subtype /Type1 /BaseFont /Times-Bold " & "/Encoding /WinAnsiEncoding >>" & vbCrLf & "endobj" & vbCrLf
    strDocument = strDocument & strObject

    lngObject = 4
    arrOffsets(lngObject) = Len(strDocument)
    strObject = "4 0 obj" & vbCrLf & "<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman " & "/Encoding /WinAnsiEncoding >>" & vbCrLf & "endobj" & vbCrLf
    strDocument = strDocument & strObject

    Select Case lngLogoComponents
        Case 1
            strColorSpace = "/DeviceGray"
        Case 4
            strColorSpace = "/DeviceCMYK"
        Case Else
            strColorSpace = "/DeviceRGB"
    End Select

    lngObject = 5
    arrOffsets(lngObject) = Len(strDocument)
    strObject = "5 0 obj" & vbCrLf & "<< /Type /XObject /Subtype /Image /Width " & CStr(lngLogoWidth) & " /Height " & CStr(lngLogoHeight) & " /ColorSpace " & strColorSpace & " /BitsPerComponent 8 /Filter [/ASCIIHexDecode /DCTDecode] /Length " & CStr(Len(strLogoHex)) & " >>" & vbCrLf & "stream" & vbCrLf & strLogoHex & vbCrLf & "endstream" & vbCrLf & "endobj" & vbCrLf
    strDocument = strDocument & strObject

    For lngPage = 1 To colPageContents.Count
        lngPageObject = 6 + ((lngPage - 1) * 2)
        lngContentObject = lngPageObject + 1
        strStream = CStr(colPageContents(lngPage))

        arrOffsets(lngPageObject) = Len(strDocument)
        strObject = CStr(lngPageObject) & " 0 obj" & vbCrLf & "<< /Type /Page /Parent 2 0 R " & "/MediaBox [0 0 " & PdfNumber(PDF_PAGE_WIDTH) & " " & PdfNumber(PDF_PAGE_HEIGHT) & "] " & "/CropBox [0 0 " & PdfNumber(PDF_PAGE_WIDTH) & " " & PdfNumber(PDF_PAGE_HEIGHT) & "] " & "/Resources << /Font << /F0 3 0 R /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> " & "/Contents " & CStr(lngContentObject) & " 0 R >>" & vbCrLf & "endobj" & vbCrLf
        strDocument = strDocument & strObject

        arrOffsets(lngContentObject) = Len(strDocument)
        strObject = CStr(lngContentObject) & " 0 obj" & vbCrLf & "<< /Length " & CStr(Len(strStream)) & " >>" & vbCrLf & "stream" & vbCrLf & strStream & "endstream" & vbCrLf & "endobj" & vbCrLf
        strDocument = strDocument & strObject
    Next lngPage

    lngStartXref = Len(strDocument)
    strDocument = strDocument & "xref" & vbCrLf & "0 " & CStr(lngObjectCount + 1) & vbCrLf & "0000000000 65535 f " & vbCrLf

    For lngObject = 1 To lngObjectCount
        strDocument = strDocument & Format$(arrOffsets(lngObject), "0000000000") & " 00000 n " & vbCrLf
    Next lngObject

    strDocument = strDocument & "trailer" & vbCrLf & "<< /Size " & CStr(lngObjectCount + 1) & " /Root 1 0 R >>" & vbCrLf & "startxref" & vbCrLf & CStr(lngStartXref) & vbCrLf & "%%EOF" & vbCrLf

    BuildPdfDocument = True

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    BuildPdfDocument = False
    strDocument = vbNullString

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "pageCount;object", lngPageCount, lngObject
    errorManager.save
    Resume ExitFunction
End Function

Private Function ValidateTargetPath(ByVal targetPath As String) As Boolean
    Const METHOD_NAME As String = "ValidateTargetPath"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim fso As Object
    Dim strParentFolder As String

    On Error GoTo ErrorHandler

    ValidateTargetPath = False

    If Len(Trim$(targetPath)) = 0 Then
        Err.Raise 1020, METHOD_NAME, "The target PDF path is required."
    End If

    If LCase$(Right$(Trim$(targetPath), 4)) <> ".pdf" Then
        Err.Raise 1020, METHOD_NAME, "The target path must end with .pdf."
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    strParentFolder = CStr(fso.GetParentFolderName(targetPath))

    If Len(strParentFolder) = 0 Then
        Err.Raise 1020, METHOD_NAME, "The target parent folder could not be resolved."
    End If

    If Not fso.FolderExists(strParentFolder) Then
        Err.Raise 1022, METHOD_NAME, "The target parent folder does not exist."
    End If

    ValidateTargetPath = True

ExitFunction:
    Set fso = Nothing
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    ValidateTargetPath = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "targetPath", targetPath
    errorManager.save
    Resume ExitFunction
End Function

Private Function HasPdf14Header(ByVal targetPath As String) As Boolean
    Const METHOD_NAME As String = "HasPdf14Header"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim blnFileOpened As Boolean
    Dim lngFileNumber As Long
    Dim strHeader As String

    On Error GoTo ErrorHandler

    HasPdf14Header = False

    If Len(Dir$(targetPath)) = 0 Then GoTo ExitFunction

    lngFileNumber = FreeFile
    Open targetPath For Binary Access Read As #lngFileNumber
    blnFileOpened = True
    strHeader = Space$(8)
    Get #lngFileNumber, 1, strHeader
    Close #lngFileNumber
    blnFileOpened = False

    HasPdf14Header = (Left$(strHeader, Len(PDF_HEADER)) = PDF_HEADER)

ExitFunction:
    If blnFileOpened Then
        CloseFileHandle lngFileNumber
    End If
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    HasPdf14Header = False

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "targetPath", targetPath
    errorManager.save
    Resume ExitFunction
End Function

Private Function PdfNumber(ByVal dblValue As Double) As String
    Const METHOD_NAME As String = "PdfNumber"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    On Error GoTo ErrorHandler

    strValue = Replace$(Format$(dblValue, "0.00"), ",", ".")

    Do While InStr(1, strValue, ".", vbBinaryCompare) > 0 And Right$(strValue, 1) = "0"
        strValue = Left$(strValue, Len(strValue) - 1)
    Loop

    If Right$(strValue, 1) = "." Then
        strValue = Left$(strValue, Len(strValue) - 1)
    End If

    PdfNumber = strValue

ExitFunction:
    Exit Function

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    PdfNumber = "0"

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "value", dblValue
    errorManager.save
    Resume ExitFunction
End Function

'-------------------------------------------------------------------------------
' Description:   Closes a VBA file handle without using Resume Next suppression.
'-------------------------------------------------------------------------------
Private Sub CloseFileHandle(ByVal lngFileNumber As Long)
    Const METHOD_NAME As String = "CloseFileHandle"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    If lngFileNumber > 0 Then
        Close #lngFileNumber
    End If

ExitSub:
    Exit Sub

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "fileNumber", lngFileNumber
    errorManager.save
    Resume ExitSub
End Sub

'-------------------------------------------------------------------------------
' Description:   Deletes a generated file when present, with standard BAT error
'                handling and without suppressing runtime errors.
'-------------------------------------------------------------------------------
Private Sub DeleteFileIfPresent(ByVal targetPath As String)
    Const METHOD_NAME As String = "DeleteFileIfPresent"
    Dim errorManager As New ErrorManager
    Dim errDescription As String
    Dim errNumber As Long

    On Error GoTo ErrorHandler

    If Len(targetPath) = 0 Then GoTo ExitSub

    If Len(Dir$(targetPath)) > 0 Then
        Kill targetPath
    End If

ExitSub:
    Exit Sub

ErrorHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description

    errorManager.addError CLASS_NAME, METHOD_NAME, errNumber, errDescription, "targetPath", targetPath
    errorManager.save
    Resume ExitSub
End Sub

