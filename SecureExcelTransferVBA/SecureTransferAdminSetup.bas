Attribute VB_Name = "SecureTransferAdminSetup"
Option Explicit

' ADMIN-ONLY DEPLOYMENT HELPER
' Import temporarily, run once on a clean deployment copy, then remove this
' entire module before the organization digitally signs and distributes it.

Private Const ADMIN_CONFIG_SHEET As String = "_SecureTransferConfig"
Private Const ADMIN_AUDIT_SHEET As String = "_SecureTransferAudit"
Private Const ADMIN_SCHEMA_VERSION As String = "1"
Private Const ADMIN_STATE_SEALED As String = "SEALED"
Private Const ADMIN_APPROVED As String = "APPROVED"
Private Const ADMIN_DIALOG_FOLDER_PICKER As Long = 4
Private Const ADMIN_ERR_CANCELLED_OFFSET As Long = 7201
Private Const ADMIN_ERR_SETUP_OFFSET As Long = 7202

Public Sub SecureTransfer_AdminInitialize()
    Dim allowedRoot As String
    Dim approvedWorkbook As String
    Dim approvedBy As String
    Dim approvalExpiry As String
    Dim maxRows As Long
    Dim maxColumns As Long
    Dim maxCells As Long
    Dim maxFileMB As Long
    Dim csvCharset As String
    Dim protectionPassword As String
    Dim passwordConfirmation As String
    Dim cfg As Object
    Dim configSheet As Worksheet
    Dim auditSheet As Worksheet
    Dim configCreated As Boolean
    Dim auditCreated As Boolean
    Dim structureProtected As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldScreenUpdating As Boolean
    Dim appStateCaptured As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed

    oldDisplayAlerts = Application.DisplayAlerts
    oldScreenUpdating = Application.ScreenUpdating
    appStateCaptured = True

    If MsgBox( _
        "이 작업은 IT/보안 관리자만 승인된 배포 사본에서 실행해야 합니다." & _
        vbCrLf & vbCrLf & _
        "초기화가 끝나면 이 관리자 모듈을 제거하고, VBA 프로젝트를 조직 인증서로 " & _
        "서명한 뒤 배포해야 합니다. 계속하시겠습니까?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Secure Excel Transfer - 관리자 초기화") <> vbYes Then
        Exit Sub
    End If

    If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then
        AdminRaiseSetup "대상 파일을 먼저 .xlsm 형식으로 저장하세요."
    End If

    If Len(ThisWorkbook.Path) = 0 Then
        AdminRaiseSetup "대상 파일을 먼저 승인 예정 위치에 저장하세요."
    End If

    If Not ThisWorkbook.Saved Then
        AdminRaiseSetup "저장되지 않은 변경 사항이 있습니다. 먼저 저장한 후 다시 실행하세요."
    End If

    If ThisWorkbook.ProtectStructure Then
        AdminRaiseSetup "통합문서 구조가 이미 보호되어 있습니다. 새 배포 사본에서 초기화하세요."
    End If

    If AdminSheetExists(ADMIN_CONFIG_SHEET) Or AdminSheetExists(ADMIN_AUDIT_SHEET) Then
        AdminRaiseSetup "기존 승인 설정 또는 감사 로그 시트가 있습니다. 덮어쓰지 않고 중단했습니다."
    End If

    AdminRejectIfOfficeRightsManaged

    allowedRoot = AdminSelectFolder(ThisWorkbook.Path)
    If Len(allowedRoot) = 0 Then AdminRaiseCancelled
    allowedRoot = AdminNormalizeExistingFolder(allowedRoot)
    approvedWorkbook = AdminNormalizeExistingFile(ThisWorkbook.FullName)

    If Not AdminPathIsWithinRoot(approvedWorkbook, allowedRoot) Then
        AdminRaiseSetup "현재 대상 통합문서가 선택한 승인 루트 안에 있지 않습니다."
    End If

    approvedBy = AdminPromptText( _
        "승인자 또는 승인 티켓 번호를 입력하세요.", _
        "Secure Excel Transfer - 승인자", vbNullString)

    approvalExpiry = AdminPromptText( _
        "승인 만료일을 yyyy-mm-dd 형식으로 입력하세요.", _
        "Secure Excel Transfer - 승인 만료일", _
        Format$(DateAdd("yyyy", 1, Date), "yyyy-mm-dd"))
    Call AdminParseIsoDate(approvalExpiry)
    If AdminParseIsoDate(approvalExpiry) < Date Then
        AdminRaiseSetup "승인 만료일은 오늘 이후여야 합니다."
    End If

    maxRows = AdminPromptLong( _
        "한 번에 허용할 최대 행 수를 입력하세요.", _
        "Secure Excel Transfer - 최대 행", 100000, 1, 1048576)

    maxColumns = AdminPromptLong( _
        "한 번에 허용할 최대 열 수를 입력하세요.", _
        "Secure Excel Transfer - 최대 열", 256, 1, 16384)

    maxCells = AdminPromptLong( _
        "한 번에 허용할 최대 전체 셀 수를 입력하세요.", _
        "Secure Excel Transfer - 최대 셀", 500000, 1, 10000000)

    maxFileMB = AdminPromptLong( _
        "허용할 최대 원본 파일 크기(MB)를 입력하세요.", _
        "Secure Excel Transfer - 최대 파일 크기", 25, 1, 1024)

    csvCharset = LCase$(AdminPromptText( _
        "CSV 인코딩을 입력하세요: utf-8, ks_c_5601-1987(CP949 계열), windows-1252", _
        "Secure Excel Transfer - CSV 인코딩", "utf-8"))

    If csvCharset <> "utf-8" And _
       csvCharset <> "ks_c_5601-1987" And _
       csvCharset <> "windows-1252" Then
        AdminRaiseSetup "지원하지 않는 CSV 인코딩입니다."
    End If

    protectionPassword = AdminPromptText( _
        "설정 시트와 통합문서 구조 보호용 암호를 입력하세요." & vbCrLf & _
        "입력 내용은 화면에 표시되므로 주변 노출에 주의하고 조직의 암호 저장소에 보관하세요.", _
        "Secure Excel Transfer - 보호 암호", vbNullString)

    If Len(protectionPassword) < 12 Then
        AdminRaiseSetup "보호 암호는 12자 이상이어야 합니다."
    End If

    passwordConfirmation = AdminPromptText( _
        "같은 보호 암호를 한 번 더 입력하세요.", _
        "Secure Excel Transfer - 보호 암호 확인", vbNullString)

    If protectionPassword <> passwordConfirmation Then
        AdminRaiseSetup "보호 암호 확인값이 일치하지 않습니다."
    End If

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare
    cfg.Add "SchemaVersion", ADMIN_SCHEMA_VERSION
    cfg.Add "DeploymentState", ADMIN_STATE_SEALED
    cfg.Add "ApprovalStatus", ADMIN_APPROVED
    cfg.Add "AllowedRoot", allowedRoot
    cfg.Add "ApprovedWorkbookFullName", approvedWorkbook
    cfg.Add "ApprovedBy", approvedBy
    cfg.Add "ApprovalExpiry", approvalExpiry
    cfg.Add "MaxRows", maxRows
    cfg.Add "MaxColumns", maxColumns
    cfg.Add "MaxCells", maxCells
    cfg.Add "MaxFileMB", maxFileMB
    cfg.Add "CsvCharset", csvCharset
    cfg.Add "ConfigChecksum", AdminConfigChecksumHex(AdminCanonicalConfig(cfg))

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set configSheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    configCreated = True
    configSheet.Name = ADMIN_CONFIG_SHEET

    Set auditSheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    auditCreated = True
    auditSheet.Name = ADMIN_AUDIT_SHEET

    AdminWriteConfiguration configSheet, cfg
    AdminPrepareAuditSheet auditSheet

    configSheet.Cells.Locked = True
    configSheet.Protect Password:=protectionPassword, _
                        DrawingObjects:=True, Contents:=True, Scenarios:=True
    configSheet.EnableSelection = xlNoSelection

    configSheet.Visible = xlSheetVeryHidden
    auditSheet.Visible = xlSheetVeryHidden

    ThisWorkbook.Protect Password:=protectionPassword, _
                         Structure:=True, Windows:=False
    structureProtected = True

    ThisWorkbook.Save

    If appStateCaptured Then
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
    End If
    protectionPassword = String$(Len(protectionPassword), vbNullChar)
    passwordConfirmation = vbNullString

    MsgBox _
        "관리자 초기화와 봉인이 완료되었습니다." & vbCrLf & vbCrLf & _
        "배포 전 필수 작업:" & vbCrLf & _
        "1) 이 SecureTransferAdminSetup 모듈을 VBA 프로젝트에서 제거" & vbCrLf & _
        "2) 조직 코드서명 인증서로 VBA 프로젝트 서명" & vbCrLf & _
        "3) 서명된 매크로만 허용하는 정책과 승인 폴더 ACL 적용" & vbCrLf & _
        "4) 파일을 닫았다가 다시 열어 SecureTransfer_IsReady 실행", _
        vbInformation, "Secure Excel Transfer - 초기화 완료"
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = False
    If structureProtected Then ThisWorkbook.Unprotect Password:=protectionPassword
    If auditCreated Then auditSheet.Delete
    If configCreated Then configSheet.Delete
    If configCreated Or auditCreated Or structureProtected Then ThisWorkbook.Save
    If appStateCaptured Then
        Application.DisplayAlerts = oldDisplayAlerts
        Application.ScreenUpdating = oldScreenUpdating
    End If
    protectionPassword = String$(Len(protectionPassword), vbNullChar)
    passwordConfirmation = vbNullString
    On Error GoTo 0

    If failureNumber = vbObjectError + ADMIN_ERR_CANCELLED_OFFSET Then Exit Sub

    MsgBox "관리자 초기화를 완료하지 못해 생성 항목을 되돌렸습니다." & _
           vbCrLf & vbCrLf & failureDescription, _
           vbExclamation, "Secure Excel Transfer - 초기화 실패"
End Sub

Private Sub AdminWriteConfiguration(ByVal configSheet As Worksheet, ByVal cfg As Object)
    Dim rowNumber As Long

    With configSheet
        .Cells.Clear
        .Range("A1:B1").Merge
        .Range("A1").Value2 = "Secure Excel Transfer Configuration"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A1").HorizontalAlignment = xlCenter

        .Cells(2, 1).Value2 = "Setting"
        .Cells(2, 2).Value2 = "Value"
        .Range("A2:B2").Font.Bold = True
        .Range("A2:B2").Interior.Color = RGB(217, 225, 242)

        rowNumber = 3
        AdminWriteSetting configSheet, rowNumber, "SchemaVersion", cfg("SchemaVersion")
        AdminWriteSetting configSheet, rowNumber, "DeploymentState", cfg("DeploymentState")
        AdminWriteSetting configSheet, rowNumber, "ApprovalStatus", cfg("ApprovalStatus")
        AdminWriteSetting configSheet, rowNumber, "AllowedRoot", cfg("AllowedRoot")
        AdminWriteSetting configSheet, rowNumber, "ApprovedWorkbookFullName", cfg("ApprovedWorkbookFullName")
        AdminWriteSetting configSheet, rowNumber, "ApprovedBy", cfg("ApprovedBy")
        AdminWriteSetting configSheet, rowNumber, "ApprovalExpiry", cfg("ApprovalExpiry")
        AdminWriteSetting configSheet, rowNumber, "MaxRows", cfg("MaxRows")
        AdminWriteSetting configSheet, rowNumber, "MaxColumns", cfg("MaxColumns")
        AdminWriteSetting configSheet, rowNumber, "MaxCells", cfg("MaxCells")
        AdminWriteSetting configSheet, rowNumber, "MaxFileMB", cfg("MaxFileMB")
        AdminWriteSetting configSheet, rowNumber, "CsvCharset", cfg("CsvCharset")
        AdminWriteSetting configSheet, rowNumber, "ConfigChecksum", cfg("ConfigChecksum")

        .Columns("A:B").AutoFit
        If .Columns("B").ColumnWidth > 90 Then .Columns("B").ColumnWidth = 90
        .Columns("B").WrapText = True
    End With
End Sub

Private Sub AdminWriteSetting(ByVal configSheet As Worksheet, _
                              ByRef rowNumber As Long, _
                              ByVal settingName As String, _
                              ByVal settingValue As Variant)
    configSheet.Cells(rowNumber, 1).Value2 = settingName
    configSheet.Cells(rowNumber, 2).Value2 = settingValue
    rowNumber = rowNumber + 1
End Sub

Private Sub AdminPrepareAuditSheet(ByVal auditSheet As Worksheet)
    With auditSheet
        .Cells.Clear
        .Cells(1, 1).Value2 = "SourceFile"
        .Cells(1, 2).Value2 = "DestinationFile"
        .Cells(1, 3).Value2 = "ProcessedAt"
        .Cells(1, 4).Value2 = "RowCount"
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(217, 225, 242)
        .Columns("A:B").ColumnWidth = 36
        .Columns("C").ColumnWidth = 22
        .Columns("D").ColumnWidth = 14
        .Columns("C").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Columns("D").NumberFormat = "#,##0"
    End With
End Sub

Private Function AdminSelectFolder(ByVal initialFolder As String) As String
    Dim picker As Object

    Set picker = Application.FileDialog(ADMIN_DIALOG_FOLDER_PICKER)
    With picker
        .Title = "보안팀이 승인한 루트 폴더 선택"
        .AllowMultiSelect = False
        .InitialFileName = initialFolder
        If .Show <> -1 Then
            AdminSelectFolder = vbNullString
        Else
            AdminSelectFolder = CStr(.SelectedItems(1))
        End If
    End With
End Function

Private Function AdminPromptText(ByVal promptText As String, _
                                 ByVal titleText As String, _
                                 ByVal defaultText As String) As String
    Dim response As Variant

    response = Application.InputBox(promptText, titleText, defaultText, Type:=2)
    If VarType(response) = vbBoolean And response = False Then AdminRaiseCancelled

    AdminPromptText = Trim$(CStr(response))
    If Len(AdminPromptText) = 0 Then AdminRaiseSetup "입력값이 비어 있습니다."
End Function

Private Function AdminPromptLong(ByVal promptText As String, _
                                 ByVal titleText As String, _
                                 ByVal defaultValue As Long, _
                                 ByVal minimumValue As Long, _
                                 ByVal maximumValue As Long) As Long
    Dim response As Variant
    Dim numericValue As Double

    response = Application.InputBox(promptText, titleText, defaultValue, Type:=1)
    If VarType(response) = vbBoolean And response = False Then AdminRaiseCancelled

    numericValue = CDbl(response)
    If numericValue <> Fix(numericValue) Or _
       numericValue < minimumValue Or numericValue > maximumValue Then
        AdminRaiseSetup "숫자가 허용 범위를 벗어났습니다: " & _
                        minimumValue & " ~ " & maximumValue
    End If

    AdminPromptLong = CLng(numericValue)
End Function

Private Function AdminSheetExists(ByVal sheetName As String) As Boolean
    Dim targetSheet As Worksheet

    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    AdminSheetExists = Not targetSheet Is Nothing
End Function

Private Sub AdminRejectIfOfficeRightsManaged()
    Dim rightsManagementEnabled As Boolean

    On Error GoTo StatusUnknown
    rightsManagementEnabled = ThisWorkbook.Permission.Enabled
    On Error GoTo 0

    If rightsManagementEnabled Then
        AdminRaiseSetup "Microsoft 권한 관리가 적용된 파일에는 자동 초기화를 수행하지 않습니다."
    End If
    Exit Sub

StatusUnknown:
    AdminRaiseSetup "대상 파일의 권한 관리 상태를 확인할 수 없습니다."
End Sub

Private Function AdminNormalizeExistingFolder(ByVal folderPath As String) As String
    Dim fileSystem As Object
    Dim normalizedPath As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    normalizedPath = fileSystem.GetAbsolutePathName(Trim$(folderPath))
    If Not fileSystem.FolderExists(normalizedPath) Then
        AdminRaiseSetup "선택한 승인 루트 폴더를 찾을 수 없습니다."
    End If

    If Len(normalizedPath) > 3 And Right$(normalizedPath, 1) = "\" Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If

    AdminNormalizeExistingFolder = normalizedPath
End Function

Private Function AdminNormalizeExistingFile(ByVal filePath As String) As String
    Dim fileSystem As Object
    Dim normalizedPath As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    normalizedPath = fileSystem.GetAbsolutePathName(Trim$(filePath))
    If Not fileSystem.FileExists(normalizedPath) Then
        AdminRaiseSetup "대상 파일을 찾을 수 없습니다."
    End If

    AdminNormalizeExistingFile = normalizedPath
End Function

Private Function AdminPathIsWithinRoot(ByVal filePath As String, _
                                       ByVal allowedRoot As String) As Boolean
    Dim rootPrefix As String

    allowedRoot = AdminNormalizeExistingFolder(allowedRoot)
    filePath = AdminNormalizeExistingFile(filePath)

    If Right$(allowedRoot, 1) = "\" Then
        rootPrefix = allowedRoot
    Else
        rootPrefix = allowedRoot & "\"
    End If

    AdminPathIsWithinRoot = _
        (StrComp(Left$(filePath, Len(rootPrefix)), rootPrefix, vbTextCompare) = 0)
End Function

Private Function AdminParseIsoDate(ByVal isoText As String) As Date
    Dim parts As Variant
    Dim parsedDate As Date

    isoText = Trim$(isoText)
    parts = Split(isoText, "-")
    If UBound(parts) <> 2 Then GoTo InvalidDate
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        GoTo InvalidDate
    End If

    On Error GoTo InvalidDate
    parsedDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    On Error GoTo 0

    If Format$(parsedDate, "yyyy-mm-dd") <> isoText Then GoTo InvalidDate
    AdminParseIsoDate = parsedDate
    Exit Function

InvalidDate:
    AdminRaiseSetup "날짜는 유효한 yyyy-mm-dd 형식이어야 합니다."
End Function

Private Function AdminCanonicalConfig(ByVal cfg As Object) As String
    AdminCanonicalConfig = _
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

Private Function AdminConfigChecksumHex(ByVal textValue As String) As String
    Dim crc As Long
    Dim position As Long
    Dim characterCode As Long

    crc = &HFFFFFFFF
    For position = 1 To Len(textValue)
        characterCode = AscW(Mid$(textValue, position, 1))
        If characterCode < 0 Then characterCode = characterCode + 65536
        crc = AdminUpdateCrc32Byte(crc, characterCode And &HFF&)
        crc = AdminUpdateCrc32Byte(crc, (characterCode \ 256) And &HFF&)
    Next position

    AdminConfigChecksumHex = Right$("00000000" & Hex$(Not crc), 8)
End Function

Private Function AdminUpdateCrc32Byte(ByVal crc As Long, ByVal byteValue As Long) As Long
    Dim bitNumber As Long

    crc = crc Xor byteValue
    For bitNumber = 1 To 8
        If (crc And 1) <> 0 Then
            crc = AdminUnsignedShiftRightOne(crc) Xor &HEDB88320
        Else
            crc = AdminUnsignedShiftRightOne(crc)
        End If
    Next bitNumber

    AdminUpdateCrc32Byte = crc
End Function

Private Function AdminUnsignedShiftRightOne(ByVal value As Long) As Long
    If (value And &H80000000) <> 0 Then
        AdminUnsignedShiftRightOne = ((value And &H7FFFFFFF) \ 2) Or &H40000000
    Else
        AdminUnsignedShiftRightOne = value \ 2
    End If
End Function

Private Sub AdminRaiseCancelled()
    Err.Raise vbObjectError + ADMIN_ERR_CANCELLED_OFFSET, _
              "SecureTransferAdminSetup", "관리자 초기화를 취소했습니다."
End Sub

Private Sub AdminRaiseSetup(ByVal messageText As String)
    Err.Raise vbObjectError + ADMIN_ERR_SETUP_OFFSET, _
              "SecureTransferAdminSetup", messageText
End Sub
