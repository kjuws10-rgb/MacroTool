Attribute VB_Name = "SecureTransfer"
Option Explicit

' Secure, policy-bound value import for an administrator-approved workbook.
' This module never uses the clipboard, SendKeys, process manipulation, hooks,
' screen capture, OCR, temporary files, or an alternate read path after denial.

Private Const CONFIG_SHEET As String = "_SecureTransferConfig"
Private Const AUDIT_SHEET As String = "_SecureTransferAudit"
Private Const HOME_SHEET As String = "Secure Transfer"
Private Const CONFIG_SCHEMA_VERSION As String = "1"
Private Const CONFIG_STATE_SEALED As String = "SEALED"
Private Const CONFIG_APPROVED As String = "APPROVED"

Private Const DIALOG_FILE_PICKER As Long = 3
Private Const AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3
Private Const XLSX_FILE_FORMAT As Long = 51
Private Const STRICT_XLSX_FILE_FORMAT As Long = 61

Private Const ERR_CANCELLED_OFFSET As Long = 7101
Private Const ERR_POLICY_OFFSET As Long = 7102
Private Const ERR_DATA_OFFSET As Long = 7103

Public Function SecureTransfer_IsReady(Optional ByVal ShowResult As Boolean = False) As Boolean
    Dim cfg As Object

    On Error GoTo NotReady
    Set cfg = LoadAndValidateConfig()
    SecureTransfer_IsReady = True

    If ShowResult Then
        MsgBox "승인 설정과 감사 로그 구성이 정상입니다.", _
               vbInformation, "Secure Excel Transfer"
    End If
    Exit Function

NotReady:
    SecureTransfer_IsReady = False
    If ShowResult Then
        MsgBox PolicyStopMessage(Err.Description), _
               vbExclamation, "Secure Excel Transfer - 사용 불가"
    End If
End Function

Public Sub SecureTransfer_OnWorkbookOpen()
    Dim cfg As Object
    Dim homeSheet As Worksheet

    On Error GoTo NotReady
    Set cfg = LoadAndValidateConfig()
    Set homeSheet = TryGetWorksheet(ThisWorkbook, HOME_SHEET)

    If Not homeSheet Is Nothing Then
        homeSheet.Activate
        homeSheet.Range("A1").Select
    End If
    Exit Sub

NotReady:
    MsgBox PolicyStopMessage(Err.Description), _
           vbExclamation, "Secure Excel Transfer - 사용 불가"
End Sub

Public Sub SecureTransfer_ShowQuickGuide()
    MsgBox _
        "사용법은 세 단계입니다." & vbCrLf & vbCrLf & _
        "1. [파일 선택하고 가져오기] 버튼을 누르고 승인된 XLSX/CSV를 고릅니다." & vbCrLf & _
        "2. XLSX라면 가져올 범위를 마우스로 드래그한 뒤 [확인]을 누릅니다." & vbCrLf & _
        "3. 대상 문서에서 비어 있는 시작 셀 하나를 클릭한 뒤 [확인]을 누릅니다." & _
        vbCrLf & vbCrLf & _
        "CSV는 전체 파일을 가져오므로 2단계가 생략됩니다." & vbCrLf & _
        "기존 값이 있는 범위에는 덮어쓰지 않습니다.", _
        vbInformation, "Secure Excel Transfer - 빠른 사용법"
End Sub

Public Sub SecureTransfer_ShowApprovalStatus()
    Dim cfg As Object

    On Error GoTo NotReady
    Set cfg = LoadAndValidateConfig()

    MsgBox _
        "상태: 승인됨" & vbCrLf & _
        "승인자/티켓: " & CStr(cfg("ApprovedBy")) & vbCrLf & _
        "승인 만료일: " & CStr(cfg("ApprovalExpiry")) & vbCrLf & _
        "승인 폴더: " & CStr(cfg("AllowedRoot")), _
        vbInformation, "Secure Excel Transfer - 승인 상태"
    Exit Sub

NotReady:
    MsgBox PolicyStopMessage(Err.Description), _
           vbExclamation, "Secure Excel Transfer - 승인 확인 실패"
End Sub

Public Sub SecureTransfer_ImportApprovedFile()
    Dim cfg As Object
    Dim allowedRoot As String
    Dim sourcePath As String
    Dim sourceExtension As String
    Dim dataValues As Variant
    Dim importedRows As Long
    Dim importedColumns As Long
    Dim maxRows As Long
    Dim maxColumns As Long
    Dim maxCells As Long
    Dim maxFileMB As Long
    Dim csvCharset As String

    Dim destinationSheet As Worksheet
    Dim destinationStart As Range
    Dim destinationRange As Range
    Dim auditSheet As Worksheet
    Dim homeSheet As Worksheet
    Dim auditRow As Long
    Dim previousHomeResult As Variant
    Dim homeUpdated As Boolean
    Dim writeStarted As Boolean
    Dim committed As Boolean

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldAutomationSecurity As Long
    Dim oldStatusBar As Variant
    Dim appStateCaptured As Boolean

    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed

    Set cfg = LoadAndValidateConfig()
    allowedRoot = NormalizeExistingFolder(CStr(cfg("AllowedRoot")))
    maxRows = ReadPositiveLong(cfg, "MaxRows", 1, 1048576)
    maxColumns = ReadPositiveLong(cfg, "MaxColumns", 1, 16384)
    maxCells = ReadPositiveLong(cfg, "MaxCells", 1, 10000000)
    maxFileMB = ReadPositiveLong(cfg, "MaxFileMB", 1, 1024)
    csvCharset = LCase$(Trim$(CStr(cfg("CsvCharset"))))

    If ThisWorkbook.ReadOnly Then
        RaisePolicy "대상 통합문서가 읽기 전용입니다."
    End If

    If Not ThisWorkbook.Saved Then
        RaisePolicy "대상 통합문서에 저장되지 않은 변경 사항이 있습니다. 먼저 저장한 후 다시 실행하세요."
    End If

    sourcePath = PickApprovedSourceFile(allowedRoot)
    If Len(sourcePath) = 0 Then Exit Sub

    sourcePath = NormalizeExistingFile(sourcePath)
    EnsurePathWithinRoot sourcePath, allowedRoot, "원본 파일"

    If StrComp(sourcePath, NormalizeExistingFile(ThisWorkbook.FullName), vbTextCompare) = 0 Then
        RaisePolicy "원본과 대상 파일이 같을 수 없습니다."
    End If

    sourceExtension = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(sourcePath))
    If sourceExtension <> "xlsx" And sourceExtension <> "csv" Then
        RaisePolicy "허용된 확장자는 .xlsx와 .csv뿐입니다."
    End If

    If CDbl(FileLen(sourcePath)) > CDbl(maxFileMB) * 1024# * 1024# Then
        RaisePolicy "원본 파일이 관리자가 정한 최대 파일 크기를 초과합니다."
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    oldCalculation = Application.Calculation
    oldAutomationSecurity = Application.AutomationSecurity
    oldStatusBar = Application.StatusBar
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = AUTOMATION_SECURITY_FORCE_DISABLE
    Application.StatusBar = "승인된 파일을 읽는 중입니다..."

    If sourceExtension = "xlsx" Then
        dataValues = ReadApprovedXlsx(sourcePath, allowedRoot, _
                                      maxRows, maxColumns, maxCells, _
                                      importedRows, importedColumns)
    Else
        dataValues = ReadApprovedCsv(sourcePath, csvCharset, _
                                     maxRows, maxColumns, maxCells, _
                                     importedRows, importedColumns)
    End If

    Set destinationStart = PickDestinationStartCell()
    Set destinationSheet = destinationStart.Parent
    ValidateDestinationWorksheet destinationSheet

    If destinationStart.Row + importedRows - 1 > destinationSheet.Rows.Count Or _
       destinationStart.Column + importedColumns - 1 > destinationSheet.Columns.Count Then
        RaiseData "가져올 데이터가 대상 시트의 행 또는 열 한계를 벗어납니다."
    End If

    Set destinationRange = destinationStart.Resize(importedRows, importedColumns)
    ValidateDestinationRange destinationRange

    Application.StatusBar = "승인된 대상 범위에 값을 기록하는 중입니다..."
    writeStarted = True
    destinationRange.Value2 = dataValues

    If RangeContainsFormula(destinationRange) Then
        RaisePolicy "문자열이 수식으로 변환될 가능성이 감지되어 가져오기를 취소했습니다."
    End If

    Set auditSheet = ThisWorkbook.Worksheets(AUDIT_SHEET)
    auditRow = AppendAuditRecord(auditSheet, _
                                 FileNameOnly(sourcePath), _
                                 ThisWorkbook.Name, _
                                 Now, importedRows)

    Set homeSheet = TryGetWorksheet(ThisWorkbook, HOME_SHEET)
    If Not homeSheet Is Nothing Then
        On Error Resume Next
        previousHomeResult = homeSheet.Range("C18").Value2
        If Err.Number = 0 Then homeUpdated = True
        Err.Clear
        On Error GoTo Failed
    End If
    UpdateHomeLastResult FileNameOnly(sourcePath), importedRows, importedColumns

    Application.StatusBar = "대상 문서와 감사 로그를 저장하는 중입니다..."
    ThisWorkbook.Save
    committed = True

    RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, _
                            oldAskToUpdateLinks, oldCalculation, oldAutomationSecurity, _
                            oldStatusBar
    appStateCaptured = False

    ThisWorkbook.Activate
    destinationSheet.Activate
    Application.Goto destinationStart, True

    MsgBox Format$(importedRows, "#,##0") & "행 × " & _
           Format$(importedColumns, "#,##0") & _
           "열을 값으로 가져왔고 감사 로그와 함께 저장했습니다.", _
           vbInformation, "Secure Excel Transfer - 완료"
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description

    On Error Resume Next
    If writeStarted And Not committed Then
        destinationRange.ClearContents
    End If
    If auditRow > 0 And Not committed Then
        auditSheet.Cells(auditRow, 1).Resize(1, 4).ClearContents
    End If
    If homeUpdated And Not committed Then
        homeSheet.Range("C18").Value2 = previousHomeResult
    End If
    If (writeStarted Or auditRow > 0) And Not committed Then
        ThisWorkbook.Save
        If Err.Number <> 0 Then
            Err.Clear
            ThisWorkbook.Saved = True
        End If
    End If
    If appStateCaptured Then
        RestoreApplicationState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, _
                                oldAskToUpdateLinks, oldCalculation, oldAutomationSecurity, _
                                oldStatusBar
    End If
    On Error GoTo 0

    If failureNumber = vbObjectError + ERR_CANCELLED_OFFSET Then Exit Sub

    MsgBox PolicyStopMessage(failureDescription), _
           vbExclamation, "Secure Excel Transfer - 작업 중단"
End Sub

Private Function LoadAndValidateConfig() As Object
    Dim cfg As Object
    Dim configSheet As Worksheet
    Dim auditSheet As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim settingName As String
    Dim expectedChecksum As String
    Dim actualChecksum As String
    Dim allowedRoot As String
    Dim approvedWorkbook As String
    Dim approvalExpiry As Date
    Dim csvCharset As String

    On Error GoTo InvalidConfig

    If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then
        RaisePolicy "대상 파일은 .xlsm 형식이어야 합니다."
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        RaisePolicy "대상 통합문서를 먼저 승인된 위치에 저장해야 합니다."
    End If

    Set configSheet = ThisWorkbook.Worksheets(CONFIG_SHEET)
    Set auditSheet = ThisWorkbook.Worksheets(AUDIT_SHEET)

    If configSheet.Visible <> xlSheetVeryHidden Or Not configSheet.ProtectContents Then
        RaisePolicy "승인 설정 시트의 숨김 또는 보호 상태가 변경되었습니다."
    End If

    If auditSheet.Visible <> xlSheetVeryHidden Then
        RaisePolicy "감사 로그 시트의 숨김 상태가 변경되었습니다."
    End If

    If Not ThisWorkbook.ProtectStructure Then
        RaisePolicy "통합문서 구조 보호가 해제되었습니다."
    End If

    If CStr(auditSheet.Cells(1, 1).Value2) <> "SourceFile" Or _
       CStr(auditSheet.Cells(1, 2).Value2) <> "DestinationFile" Or _
       CStr(auditSheet.Cells(1, 3).Value2) <> "ProcessedAt" Or _
       CStr(auditSheet.Cells(1, 4).Value2) <> "RowCount" Then
        RaisePolicy "감사 로그 헤더가 변경되었습니다."
    End If

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare

    lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row
    For rowNumber = 3 To lastRow
        settingName = Trim$(CStr(configSheet.Cells(rowNumber, 1).Value2))
        If Len(settingName) > 0 Then
            If cfg.Exists(settingName) Then
                RaisePolicy "승인 설정에 중복 키가 있습니다: " & settingName
            End If
            cfg.Add settingName, configSheet.Cells(rowNumber, 2).Value2
        End If
    Next rowNumber

    RequireSetting cfg, "SchemaVersion"
    RequireSetting cfg, "DeploymentState"
    RequireSetting cfg, "ApprovalStatus"
    RequireSetting cfg, "AllowedRoot"
    RequireSetting cfg, "ApprovedWorkbookFullName"
    RequireSetting cfg, "ApprovedBy"
    RequireSetting cfg, "ApprovalExpiry"
    RequireSetting cfg, "MaxRows"
    RequireSetting cfg, "MaxColumns"
    RequireSetting cfg, "MaxCells"
    RequireSetting cfg, "MaxFileMB"
    RequireSetting cfg, "CsvCharset"
    RequireSetting cfg, "ConfigChecksum"

    If CStr(cfg("SchemaVersion")) <> CONFIG_SCHEMA_VERSION Then
        RaisePolicy "지원하지 않는 승인 설정 버전입니다."
    End If

    If UCase$(Trim$(CStr(cfg("DeploymentState")))) <> CONFIG_STATE_SEALED Then
        RaisePolicy "배포 상태가 봉인되지 않았습니다."
    End If

    If UCase$(Trim$(CStr(cfg("ApprovalStatus")))) <> CONFIG_APPROVED Then
        RaisePolicy "관리자 승인 상태가 아닙니다."
    End If

    If Len(Trim$(CStr(cfg("ApprovedBy")))) = 0 Then
        RaisePolicy "승인자 정보가 비어 있습니다."
    End If

    approvalExpiry = ParseIsoDate(CStr(cfg("ApprovalExpiry")))
    If approvalExpiry < Date Then
        RaisePolicy "관리자 승인 유효기간이 만료되었습니다."
    End If

    Call ReadPositiveLong(cfg, "MaxRows", 1, 1048576)
    Call ReadPositiveLong(cfg, "MaxColumns", 1, 16384)
    Call ReadPositiveLong(cfg, "MaxCells", 1, 10000000)
    Call ReadPositiveLong(cfg, "MaxFileMB", 1, 1024)

    csvCharset = LCase$(Trim$(CStr(cfg("CsvCharset"))))
    If csvCharset <> "utf-8" And _
       csvCharset <> "ks_c_5601-1987" And _
       csvCharset <> "windows-1252" Then
        RaisePolicy "허용되지 않은 CSV 문자 인코딩 설정입니다."
    End If

    expectedChecksum = UCase$(Trim$(CStr(cfg("ConfigChecksum"))))
    actualChecksum = ConfigChecksumHex(CanonicalConfig(cfg))
    If expectedChecksum <> actualChecksum Then
        RaisePolicy "승인 설정 무결성 확인에 실패했습니다."
    End If

    allowedRoot = NormalizeExistingFolder(CStr(cfg("AllowedRoot")))
    approvedWorkbook = NormalizeExistingFile(CStr(cfg("ApprovedWorkbookFullName")))
    EnsurePathWithinRoot approvedWorkbook, allowedRoot, "승인된 대상 파일"

    If StrComp(approvedWorkbook, NormalizeExistingFile(ThisWorkbook.FullName), vbTextCompare) <> 0 Then
        RaisePolicy "현재 통합문서의 경로 또는 파일명이 관리자 승인값과 다릅니다."
    End If

    RejectIfOfficeRightsManaged ThisWorkbook, "대상 통합문서"

    Set LoadAndValidateConfig = cfg
    Exit Function

InvalidConfig:
    If Err.Number >= vbObjectError + ERR_CANCELLED_OFFSET And _
       Err.Number <= vbObjectError + ERR_DATA_OFFSET Then
        Err.Raise Err.Number, "SecureTransfer.LoadAndValidateConfig", Err.Description
    Else
        RaisePolicy "승인 설정을 읽거나 검증할 수 없습니다: " & Err.Description
    End If
End Function

Private Function ReadApprovedXlsx(ByVal sourcePath As String, _
                                  ByVal allowedRoot As String, _
                                  ByVal maxRows As Long, _
                                  ByVal maxColumns As Long, _
                                  ByVal maxCells As Long, _
                                  ByRef rowCount As Long, _
                                  ByRef columnCount As Long) As Variant
    Dim sourceWorkbook As Workbook
    Dim sourceRange As Range
    Dim values As Variant
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed

    Set sourceWorkbook = Application.Workbooks.Open( _
        Filename:=sourcePath, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False, _
        Notify:=False, Local:=True)

    If sourceWorkbook Is ThisWorkbook Then
        RaisePolicy "원본과 대상 통합문서가 같습니다."
    End If

    If sourceWorkbook.FileFormat <> XLSX_FILE_FORMAT And _
       sourceWorkbook.FileFormat <> STRICT_XLSX_FILE_FORMAT Then
        RaisePolicy "파일 내용이 허용된 .xlsx 형식과 일치하지 않습니다."
    End If

    EnsurePathWithinRoot NormalizeExistingFile(sourceWorkbook.FullName), _
                         allowedRoot, "열린 원본 파일"

    If StrComp(NormalizeExistingFile(sourceWorkbook.FullName), _
               NormalizeExistingFile(sourcePath), vbTextCompare) <> 0 Then
        RaisePolicy "선택한 원본과 실제로 열린 파일의 경로가 다릅니다."
    End If

    RejectIfOfficeRightsManaged sourceWorkbook, "원본 통합문서"

    If sourceWorkbook.Connections.Count > 0 Then
        RaisePolicy "외부 데이터 연결이 포함된 .xlsx는 값 전용 사본으로 만든 후 승인을 받아야 합니다."
    End If

    Set sourceRange = PickSourceRange(sourceWorkbook)
    If sourceRange.Areas.Count <> 1 Then
        RaiseData "원본 범위는 하나의 연속된 범위여야 합니다."
    End If

    rowCount = sourceRange.Rows.Count
    columnCount = sourceRange.Columns.Count
    ValidateImportDimensions rowCount, columnCount, maxRows, maxColumns, maxCells

    values = sourceRange.Value2
    ReadApprovedXlsx = MakeSafeTwoDimensional(values, rowCount, columnCount)

    sourceWorkbook.Close SaveChanges:=False
    Set sourceWorkbook = Nothing
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise failureNumber, "SecureTransfer.ReadApprovedXlsx", failureDescription
End Function

Private Function ReadApprovedCsv(ByVal sourcePath As String, _
                                 ByVal csvCharset As String, _
                                 ByVal maxRows As Long, _
                                 ByVal maxColumns As Long, _
                                 ByVal maxCells As Long, _
                                 ByRef rowCount As Long, _
                                 ByRef columnCount As Long) As Variant
    Dim textContent As String
    Dim parsedRows As Collection
    Dim fields As Collection
    Dim fieldValue As String
    Dim rowValues As Variant
    Dim outputValues() As Variant
    Dim inQuotes As Boolean
    Dim quoteClosed As Boolean
    Dim rowHasToken As Boolean
    Dim currentCharacter As String
    Dim nextCharacter As String
    Dim position As Long
    Dim rowNumber As Long
    Dim columnNumber As Long

    textContent = ReadTextWithoutTempFile(sourcePath, csvCharset)
    If Len(textContent) > 0 Then
        If AscW(Left$(textContent, 1)) = &HFEFF Then
            textContent = Mid$(textContent, 2)
        End If
    End If

    Set parsedRows = New Collection
    Set fields = New Collection

    position = 1
    Do While position <= Len(textContent)
        currentCharacter = Mid$(textContent, position, 1)

        If inQuotes Then
            If currentCharacter = """" Then
                nextCharacter = vbNullString
                If position < Len(textContent) Then
                    nextCharacter = Mid$(textContent, position + 1, 1)
                End If

                If nextCharacter = """" Then
                    fieldValue = fieldValue & """"
                    position = position + 1
                Else
                    inQuotes = False
                    quoteClosed = True
                End If
            Else
                fieldValue = fieldValue & currentCharacter
            End If
        ElseIf quoteClosed Then
            Select Case currentCharacter
                Case ","
                    fields.Add fieldValue
                    fieldValue = vbNullString
                    quoteClosed = False
                    rowHasToken = True
                Case vbCr, vbLf
                    fields.Add fieldValue
                    fieldValue = vbNullString
                    FinishCsvRow fields, parsedRows, maxRows, maxColumns, columnCount
                    quoteClosed = False
                    rowHasToken = False
                    If currentCharacter = vbCr And position < Len(textContent) Then
                        If Mid$(textContent, position + 1, 1) = vbLf Then position = position + 1
                    End If
                Case " ", vbTab
                    ' Ignore whitespace after a closing quote.
                Case Else
                    RaiseData "CSV의 닫는 따옴표 뒤에 허용되지 않은 문자가 있습니다."
            End Select
        Else
            Select Case currentCharacter
                Case """"
                    If Len(fieldValue) <> 0 Then
                        RaiseData "CSV의 따옴표 형식이 올바르지 않습니다."
                    End If
                    inQuotes = True
                    rowHasToken = True
                Case ","
                    fields.Add fieldValue
                    fieldValue = vbNullString
                    rowHasToken = True
                Case vbCr, vbLf
                    If rowHasToken Or fields.Count > 0 Or Len(fieldValue) > 0 Then
                        fields.Add fieldValue
                        fieldValue = vbNullString
                        FinishCsvRow fields, parsedRows, maxRows, maxColumns, columnCount
                    End If
                    rowHasToken = False
                    If currentCharacter = vbCr And position < Len(textContent) Then
                        If Mid$(textContent, position + 1, 1) = vbLf Then position = position + 1
                    End If
                Case Else
                    fieldValue = fieldValue & currentCharacter
                    rowHasToken = True
            End Select
        End If

        position = position + 1
    Loop

    If inQuotes Then
        RaiseData "CSV의 따옴표가 닫히지 않았습니다."
    End If

    If rowHasToken Or fields.Count > 0 Or Len(fieldValue) > 0 Or quoteClosed Then
        fields.Add fieldValue
        FinishCsvRow fields, parsedRows, maxRows, maxColumns, columnCount
    End If

    rowCount = parsedRows.Count
    If rowCount = 0 Or columnCount = 0 Then
        RaiseData "CSV에 가져올 데이터가 없습니다."
    End If

    ValidateImportDimensions rowCount, columnCount, maxRows, maxColumns, maxCells
    ReDim outputValues(1 To rowCount, 1 To columnCount)

    For rowNumber = 1 To rowCount
        rowValues = parsedRows(rowNumber)
        For columnNumber = 1 To columnCount
            If columnNumber <= UBound(rowValues) Then
                outputValues(rowNumber, columnNumber) = SafeLiteralValue(rowValues(columnNumber))
            Else
                outputValues(rowNumber, columnNumber) = vbNullString
            End If
        Next columnNumber
    Next rowNumber

    ReadApprovedCsv = outputValues
End Function

Private Sub FinishCsvRow(ByRef fields As Collection, _
                         ByRef parsedRows As Collection, _
                         ByVal maxRows As Long, _
                         ByVal maxColumns As Long, _
                         ByRef widestColumnCount As Long)
    Dim rowValues() As Variant
    Dim columnNumber As Long

    If fields.Count > maxColumns Then
        RaiseData "CSV 열 수가 관리자가 정한 한도를 초과합니다."
    End If

    If parsedRows.Count + 1 > maxRows Then
        RaiseData "CSV 행 수가 관리자가 정한 한도를 초과합니다."
    End If

    ReDim rowValues(1 To fields.Count)
    For columnNumber = 1 To fields.Count
        rowValues(columnNumber) = CStr(fields(columnNumber))
    Next columnNumber

    parsedRows.Add rowValues
    If fields.Count > widestColumnCount Then widestColumnCount = fields.Count
    Set fields = New Collection
End Sub

Private Function ReadTextWithoutTempFile(ByVal filePath As String, _
                                         ByVal csvCharset As String) As String
    Dim stream As Object
    Dim failureDescription As String

    On Error GoTo Failed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = csvCharset
    stream.Open
    stream.LoadFromFile filePath
    ReadTextWithoutTempFile = stream.ReadText(-1)
    stream.Close
    Exit Function

Failed:
    failureDescription = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    RaisePolicy "CSV를 메모리에서 읽지 못했습니다. 인코딩과 접근 승인을 확인하세요: " & _
                failureDescription
End Function

Private Function MakeSafeTwoDimensional(ByVal sourceValues As Variant, _
                                        ByVal rowCount As Long, _
                                        ByVal columnCount As Long) As Variant
    Dim outputValues() As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long

    ReDim outputValues(1 To rowCount, 1 To columnCount)

    If rowCount = 1 And columnCount = 1 Then
        outputValues(1, 1) = SafeLiteralValue(sourceValues)
    Else
        For rowNumber = 1 To rowCount
            For columnNumber = 1 To columnCount
                outputValues(rowNumber, columnNumber) = _
                    SafeLiteralValue(sourceValues(rowNumber, columnNumber))
            Next columnNumber
        Next rowNumber
    End If

    MakeSafeTwoDimensional = outputValues
End Function

Private Function SafeLiteralValue(ByVal sourceValue As Variant) As Variant
    Dim textValue As String
    Dim probe As String

    If IsError(sourceValue) Then
        SafeLiteralValue = sourceValue
        Exit Function
    End If

    If IsNull(sourceValue) Or IsEmpty(sourceValue) Then
        SafeLiteralValue = vbNullString
        Exit Function
    End If

    If VarType(sourceValue) <> vbString Then
        SafeLiteralValue = sourceValue
        Exit Function
    End If

    textValue = CStr(sourceValue)
    probe = LTrimRiskCharacters(textValue)

    If Len(probe) > 0 Then
        If InStr(1, "=+-@", Left$(probe, 1), vbBinaryCompare) > 0 Then
            SafeLiteralValue = "'" & textValue
            Exit Function
        End If
    End If

    SafeLiteralValue = textValue
End Function

Private Function LTrimRiskCharacters(ByVal textValue As String) As String
    Dim position As Long
    Dim characterCode As Long

    position = 1
    Do While position <= Len(textValue)
        characterCode = AscW(Mid$(textValue, position, 1))
        If characterCode = 9 Or characterCode = 10 Or _
           characterCode = 13 Or characterCode = 32 Then
            position = position + 1
        Else
            Exit Do
        End If
    Loop

    LTrimRiskCharacters = Mid$(textValue, position)
End Function

Private Sub ValidateImportDimensions(ByVal rowCount As Long, _
                                     ByVal columnCount As Long, _
                                     ByVal maxRows As Long, _
                                     ByVal maxColumns As Long, _
                                     ByVal maxCells As Long)
    If rowCount < 1 Or columnCount < 1 Then
        RaiseData "가져올 범위가 비어 있습니다."
    End If

    If rowCount > maxRows Then
        RaisePolicy "가져올 행 수가 관리자가 정한 한도를 초과합니다."
    End If

    If columnCount > maxColumns Then
        RaisePolicy "가져올 열 수가 관리자가 정한 한도를 초과합니다."
    End If

    If CDbl(rowCount) * CDbl(columnCount) > CDbl(maxCells) Then
        RaisePolicy "가져올 전체 셀 수가 관리자가 정한 한도를 초과합니다."
    End If
End Sub

Private Sub ValidateDestinationWorksheet(ByVal targetSheet As Worksheet)
    If targetSheet.Parent Is Nothing Then RaiseData "대상 시트를 확인할 수 없습니다."

    If Not targetSheet.Parent Is ThisWorkbook Then
        RaisePolicy "대상 시트는 승인된 현재 통합문서에 있어야 합니다."
    End If

    If targetSheet.Name = CONFIG_SHEET Or _
       targetSheet.Name = AUDIT_SHEET Or _
       targetSheet.Name = HOME_SHEET Then
        RaisePolicy "시작 화면, 승인 설정 또는 감사 로그 시트에는 데이터를 쓸 수 없습니다."
    End If

    If targetSheet.Visible <> xlSheetVisible Then
        RaisePolicy "대상 시트는 표시 상태여야 합니다."
    End If

    If targetSheet.ProtectContents Then
        RaisePolicy "대상 시트가 보호되어 있습니다. 보안팀이 승인한 입력 시트를 사용하세요."
    End If
End Sub

Private Sub ValidateDestinationRange(ByVal targetRange As Range)
    Dim mergeState As Variant
    Dim listObject As ListObject

    mergeState = targetRange.MergeCells
    If IsNull(mergeState) Then
        RaiseData "대상 범위 일부에 병합 셀이 포함되어 있습니다."
    ElseIf CBool(mergeState) Then
        RaiseData "대상 범위에 병합 셀이 포함되어 있습니다."
    End If

    If Application.CountA(targetRange) > 0 Then
        RaisePolicy "기존 데이터를 덮어쓰지 않도록 대상 범위는 완전히 비어 있어야 합니다."
    End If

    For Each listObject In targetRange.Worksheet.ListObjects
        If Not Intersect(targetRange, listObject.Range) Is Nothing Then
            RaisePolicy "대상 범위가 Excel 표와 겹칩니다. 표 밖의 빈 범위를 사용하세요."
        End If
    Next listObject
End Sub

Private Function RangeContainsFormula(ByVal targetRange As Range) As Boolean
    Dim formulaCells As Range

    On Error Resume Next
    Set formulaCells = targetRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    RangeContainsFormula = Not formulaCells Is Nothing
End Function

Private Function AppendAuditRecord(ByVal auditSheet As Worksheet, _
                                   ByVal sourceFileName As String, _
                                   ByVal destinationFileName As String, _
                                   ByVal processedAt As Date, _
                                   ByVal importedRows As Long) As Long
    Dim nextRow As Long
    Dim recordValues(1 To 1, 1 To 4) As Variant

    If auditSheet.ProtectContents Then
        RaisePolicy "감사 로그가 쓰기 불가 상태입니다."
    End If

    nextRow = auditSheet.Cells(auditSheet.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    recordValues(1, 1) = sourceFileName
    recordValues(1, 2) = destinationFileName
    recordValues(1, 3) = processedAt
    recordValues(1, 4) = importedRows

    auditSheet.Cells(nextRow, 1).Resize(1, 4).Value2 = recordValues
    auditSheet.Cells(nextRow, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    auditSheet.Cells(nextRow, 4).NumberFormat = "#,##0"
    AppendAuditRecord = nextRow
End Function

Private Function PickApprovedSourceFile(ByVal allowedRoot As String) As String
    Dim picker As Object

    Set picker = Application.FileDialog(DIALOG_FILE_PICKER)
    With picker
        .Title = "승인된 루트 안의 XLSX 또는 CSV 선택"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "승인된 데이터 파일", "*.xlsx;*.csv"
        .InitialFileName = EnsureTrailingSlash(allowedRoot)

        If .Show <> -1 Then
            PickApprovedSourceFile = vbNullString
        Else
            PickApprovedSourceFile = CStr(.SelectedItems(1))
        End If
    End With
End Function

Private Function PickSourceRange(ByVal sourceWorkbook As Workbook) As Range
    Dim sourceSheet As Worksheet
    Dim defaultRange As Range
    Dim selectedRange As Range

    Set sourceSheet = FirstVisibleWorksheet(sourceWorkbook)
    Set defaultRange = EffectiveUsedRange(sourceSheet)

    Application.ScreenUpdating = True
    sourceWorkbook.Activate
    sourceSheet.Activate
    Application.Goto defaultRange, True

    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        Prompt:="가져올 범위를 마우스로 드래그한 뒤 [확인]을 누르세요." & _
                vbCrLf & "다른 시트도 직접 클릭해서 선택할 수 있습니다.", _
        Title:="1단계 - 원본 범위 선택", _
        Default:=defaultRange.Address(External:=True), Type:=8)
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = False

    If selectedRange Is Nothing Then RaiseCancelled

    If Not selectedRange.Parent.Parent Is sourceWorkbook Then
        RaisePolicy "원본 범위는 선택한 승인 파일 안에서만 고를 수 있습니다."
    End If

    If selectedRange.Areas.Count <> 1 Then
        RaiseData "원본 범위는 하나의 연속된 범위여야 합니다."
    End If

    Set PickSourceRange = selectedRange
End Function

Private Function PickDestinationStartCell() As Range
    Dim defaultSheet As Worksheet
    Dim selectedCell As Range

    Set defaultSheet = GetWorksheetStrict(ThisWorkbook, DefaultDestinationSheetName())

    Application.ScreenUpdating = True
    ThisWorkbook.Activate
    defaultSheet.Activate
    Application.Goto defaultSheet.Range("A1"), True

    On Error Resume Next
    Set selectedCell = Application.InputBox( _
        Prompt:="값을 넣을 비어 있는 시작 셀 하나를 클릭한 뒤 [확인]을 누르세요." & _
                vbCrLf & "선택한 셀을 왼쪽 위로 하여 데이터가 채워집니다.", _
        Title:="마지막 단계 - 대상 시작 셀 선택", _
        Default:=defaultSheet.Range("A1").Address(External:=True), Type:=8)
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = False

    If selectedCell Is Nothing Then RaiseCancelled

    If Not selectedCell.Parent.Parent Is ThisWorkbook Then
        RaisePolicy "대상 셀은 승인된 현재 통합문서 안에서만 고를 수 있습니다."
    End If

    If selectedCell.Cells.CountLarge <> 1 Then
        RaiseData "대상 시작 위치는 셀 하나만 선택해야 합니다."
    End If

    Set PickDestinationStartCell = selectedCell.Cells(1, 1)
End Function

Private Function FirstVisibleWorksheet(ByVal workbookObject As Workbook) As Worksheet
    Dim targetSheet As Worksheet

    For Each targetSheet In workbookObject.Worksheets
        If targetSheet.Visible = xlSheetVisible Then
            Set FirstVisibleWorksheet = targetSheet
            Exit Function
        End If
    Next targetSheet

    RaiseData "원본 파일에 표시 상태의 시트가 없습니다."
End Function

Private Function DefaultDestinationSheetName() As String
    Dim targetSheet As Worksheet

    If ActiveWorkbook Is ThisWorkbook Then
        If TypeName(ActiveSheet) = "Worksheet" Then
            Set targetSheet = ActiveSheet
            If targetSheet.Visible = xlSheetVisible And _
               targetSheet.Name <> CONFIG_SHEET And _
               targetSheet.Name <> AUDIT_SHEET And _
               targetSheet.Name <> HOME_SHEET Then
                DefaultDestinationSheetName = targetSheet.Name
                Exit Function
            End If
        End If
    End If

    For Each targetSheet In ThisWorkbook.Worksheets
        If targetSheet.Visible = xlSheetVisible And _
           targetSheet.Name <> CONFIG_SHEET And _
           targetSheet.Name <> AUDIT_SHEET And _
           targetSheet.Name <> HOME_SHEET Then
            DefaultDestinationSheetName = targetSheet.Name
            Exit Function
        End If
    Next targetSheet

    RaisePolicy "사용 가능한 표시 상태의 대상 시트가 없습니다."
End Function

Private Function GetWorksheetStrict(ByVal workbookObject As Workbook, _
                                    ByVal worksheetName As String) As Worksheet
    Dim targetSheet As Worksheet

    For Each targetSheet In workbookObject.Worksheets
        If StrComp(targetSheet.Name, worksheetName, vbBinaryCompare) = 0 Then
            Set GetWorksheetStrict = targetSheet
            Exit Function
        End If
    Next targetSheet

    RaiseData "시트를 찾을 수 없습니다: " & worksheetName
End Function

Private Function TryGetWorksheet(ByVal workbookObject As Workbook, _
                                 ByVal worksheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetWorksheet = workbookObject.Worksheets(worksheetName)
    On Error GoTo 0
End Function

Private Sub UpdateHomeLastResult(ByVal sourceFileName As String, _
                                 ByVal importedRows As Long, _
                                 ByVal importedColumns As Long)
    Dim homeSheet As Worksheet

    Set homeSheet = TryGetWorksheet(ThisWorkbook, HOME_SHEET)
    If homeSheet Is Nothing Then Exit Sub

    On Error Resume Next
    homeSheet.Range("C18").Value2 = _
        Format$(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
        sourceFileName & " | " & Format$(importedRows, "#,##0") & "행 × " & _
        Format$(importedColumns, "#,##0") & "열"
    On Error GoTo 0
End Sub

Private Function EffectiveUsedRange(ByVal targetSheet As Worksheet) As Range
    Dim lastRowCell As Range
    Dim lastColumnCell As Range

    Set lastRowCell = targetSheet.Cells.Find( _
        What:="*", After:=targetSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastRowCell Is Nothing Then
        RaiseData "선택한 원본 시트에 값이 없습니다."
    End If

    Set lastColumnCell = targetSheet.Cells.Find( _
        What:="*", After:=targetSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
        MatchCase:=False)

    Set EffectiveUsedRange = targetSheet.Range( _
        targetSheet.Cells(1, 1), _
        targetSheet.Cells(lastRowCell.Row, lastColumnCell.Column))
End Function

Private Sub RejectIfOfficeRightsManaged(ByVal workbookObject As Workbook, _
                                        ByVal roleDescription As String)
    Dim rightsManagementEnabled As Boolean

    On Error GoTo StatusUnknown
    rightsManagementEnabled = workbookObject.Permission.Enabled
    On Error GoTo 0

    If rightsManagementEnabled Then
        RaisePolicy roleDescription & _
                    "에 Microsoft 권한 관리가 적용되어 있어 자동 가져오기를 수행하지 않습니다."
    End If
    Exit Sub

StatusUnknown:
    RaisePolicy roleDescription & _
                "의 권한 관리 상태를 확인할 수 없어 안전하게 중단했습니다."
End Sub

Private Function NormalizeExistingFolder(ByVal folderPath As String) As String
    Dim fileSystem As Object
    Dim normalizedPath As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    normalizedPath = fileSystem.GetAbsolutePathName(Trim$(folderPath))
    If Not fileSystem.FolderExists(normalizedPath) Then
        RaisePolicy "승인된 루트 폴더를 찾을 수 없습니다."
    End If

    NormalizeExistingFolder = RemoveTrailingSlashExceptRoot(normalizedPath)
End Function

Private Function NormalizeExistingFile(ByVal filePath As String) As String
    Dim fileSystem As Object
    Dim normalizedPath As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    normalizedPath = fileSystem.GetAbsolutePathName(Trim$(filePath))
    If Not fileSystem.FileExists(normalizedPath) Then
        RaisePolicy "파일을 찾거나 접근할 수 없습니다: " & fileSystem.GetFileName(filePath)
    End If

    NormalizeExistingFile = normalizedPath
End Function

Private Sub EnsurePathWithinRoot(ByVal filePath As String, _
                                 ByVal allowedRoot As String, _
                                 ByVal roleDescription As String)
    Dim normalizedFile As String
    Dim normalizedRoot As String
    Dim rootPrefix As String

    normalizedFile = NormalizeExistingFile(filePath)
    normalizedRoot = NormalizeExistingFolder(allowedRoot)
    rootPrefix = EnsureTrailingSlash(normalizedRoot)

    If StrComp(Left$(normalizedFile, Len(rootPrefix)), rootPrefix, vbTextCompare) <> 0 Then
        RaisePolicy roleDescription & "이(가) 관리자가 승인한 루트 밖에 있습니다."
    End If
End Sub

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Private Function RemoveTrailingSlashExceptRoot(ByVal folderPath As String) As String
    If Len(folderPath) > 3 And Right$(folderPath, 1) = "\" Then
        RemoveTrailingSlashExceptRoot = Left$(folderPath, Len(folderPath) - 1)
    Else
        RemoveTrailingSlashExceptRoot = folderPath
    End If
End Function

Private Function FileNameOnly(ByVal filePath As String) As String
    FileNameOnly = CreateObject("Scripting.FileSystemObject").GetFileName(filePath)
End Function

Private Sub RequireSetting(ByVal cfg As Object, ByVal settingName As String)
    If Not cfg.Exists(settingName) Then
        RaisePolicy "필수 승인 설정이 없습니다: " & settingName
    End If
End Sub

Private Function ReadPositiveLong(ByVal cfg As Object, _
                                  ByVal settingName As String, _
                                  ByVal minimumValue As Long, _
                                  ByVal maximumValue As Long) As Long
    Dim numericValue As Double

    If Not IsNumeric(cfg(settingName)) Then
        RaisePolicy settingName & " 설정이 숫자가 아닙니다."
    End If

    numericValue = CDbl(cfg(settingName))
    If numericValue <> Fix(numericValue) Or _
       numericValue < minimumValue Or numericValue > maximumValue Then
        RaisePolicy settingName & " 설정이 허용 범위를 벗어났습니다."
    End If

    ReadPositiveLong = CLng(numericValue)
End Function

Private Function ParseIsoDate(ByVal isoText As String) As Date
    Dim parts As Variant
    Dim parsedDate As Date

    isoText = Trim$(isoText)
    parts = Split(isoText, "-")
    If UBound(parts) <> 2 Then RaisePolicy "승인 만료일 형식은 yyyy-mm-dd여야 합니다."
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        RaisePolicy "승인 만료일 형식은 yyyy-mm-dd여야 합니다."
    End If

    On Error GoTo InvalidDate
    parsedDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    On Error GoTo 0

    If Format$(parsedDate, "yyyy-mm-dd") <> isoText Then GoTo InvalidDate
    ParseIsoDate = parsedDate
    Exit Function

InvalidDate:
    RaisePolicy "승인 만료일이 올바르지 않습니다."
End Function

Private Function CanonicalConfig(ByVal cfg As Object) As String
    CanonicalConfig = _
        "SchemaVersion=" & CStr(cfg("SchemaVersion")) & vbLf & _
        "DeploymentState=" & CStr(cfg("DeploymentState")) & vbLf & _
        "ApprovalStatus=" & CStr(cfg("ApprovalStatus")) & vbLf & _
        "AllowedRoot=" & CStr(cfg("AllowedRoot")) & vbLf & _
        "ApprovedWorkbookFullName=" & CStr(cfg("ApprovedWorkbookFullName")) & vbLf & _
        "ApprovedBy=" & CStr(cfg("ApprovedBy")) & vbLf & _
        "ApprovalExpiry=" & CStr(cfg("ApprovalExpiry")) & vbLf & _
        "MaxRows=" & CStr(cfg("MaxRows")) & vbLf & _
        "MaxColumns=" & CStr(cfg("MaxColumns")) & vbLf & _
        "MaxCells=" & CStr(cfg("MaxCells")) & vbLf & _
        "MaxFileMB=" & CStr(cfg("MaxFileMB")) & vbLf & _
        "CsvCharset=" & CStr(cfg("CsvCharset"))
End Function

Private Function ConfigChecksumHex(ByVal textValue As String) As String
    Dim crc As Long
    Dim position As Long
    Dim characterCode As Long

    crc = &HFFFFFFFF
    For position = 1 To Len(textValue)
        characterCode = AscW(Mid$(textValue, position, 1))
        If characterCode < 0 Then characterCode = characterCode + 65536
        crc = UpdateCrc32Byte(crc, characterCode And &HFF&)
        crc = UpdateCrc32Byte(crc, (characterCode \ 256) And &HFF&)
    Next position

    ConfigChecksumHex = Right$("00000000" & Hex$(Not crc), 8)
End Function

Private Function UpdateCrc32Byte(ByVal crc As Long, ByVal byteValue As Long) As Long
    Dim bitNumber As Long

    crc = crc Xor byteValue
    For bitNumber = 1 To 8
        If (crc And 1) <> 0 Then
            crc = UnsignedShiftRightOne(crc) Xor &HEDB88320
        Else
            crc = UnsignedShiftRightOne(crc)
        End If
    Next bitNumber

    UpdateCrc32Byte = crc
End Function

Private Function UnsignedShiftRightOne(ByVal value As Long) As Long
    If (value And &H80000000) <> 0 Then
        UnsignedShiftRightOne = ((value And &H7FFFFFFF) \ 2) Or &H40000000
    Else
        UnsignedShiftRightOne = value \ 2
    End If
End Function

Private Sub RestoreApplicationState(ByVal screenUpdatingValue As Boolean, _
                                    ByVal enableEventsValue As Boolean, _
                                    ByVal displayAlertsValue As Boolean, _
                                    ByVal askToUpdateLinksValue As Boolean, _
                                    ByVal calculationValue As XlCalculation, _
                                    ByVal automationSecurityValue As Long, _
                                    ByVal statusBarValue As Variant)
    On Error Resume Next
    Application.ScreenUpdating = screenUpdatingValue
    Application.EnableEvents = enableEventsValue
    Application.DisplayAlerts = displayAlertsValue
    Application.AskToUpdateLinks = askToUpdateLinksValue
    Application.Calculation = calculationValue
    Application.AutomationSecurity = automationSecurityValue
    Application.StatusBar = statusBarValue
    On Error GoTo 0
End Sub

Private Function PolicyStopMessage(ByVal detail As String) As String
    PolicyStopMessage = _
        "작업을 안전하게 중단했습니다." & vbCrLf & vbCrLf & _
        detail & vbCrLf & vbCrLf & _
        "클립보드 또는 DRM/DLP 차단을 우회해 재시도하지 않습니다. " & _
        "업무상 이전이 필요하면 원본·대상 파일과 경로를 IT 보안팀에 제출하여 " & _
        "승인 또는 정책 예외를 요청하세요."
End Function

Private Sub RaiseCancelled()
    Err.Raise vbObjectError + ERR_CANCELLED_OFFSET, _
              "SecureTransfer", "사용자가 작업을 취소했습니다."
End Sub

Private Sub RaisePolicy(ByVal messageText As String)
    Err.Raise vbObjectError + ERR_POLICY_OFFSET, _
              "SecureTransfer.Policy", messageText
End Sub

Private Sub RaiseData(ByVal messageText As String)
    Err.Raise vbObjectError + ERR_DATA_OFFSET, _
              "SecureTransfer.Data", messageText
End Sub
