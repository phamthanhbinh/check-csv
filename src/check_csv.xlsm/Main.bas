Attribute VB_Name = "Main"
'Sheet
Public Const DATA_CHECK_SHEET = "�f�[�^�`�F�b�N�c�[��"
Public Const FILE_LIST_SHEET = "IF�t�@�C���ꗗ"
Public Const TABLE_DEFINITE_TEMPLATE = "�y�J������`�z�e���v���[�g"
'Range
Public Const DATA_CHECK_RANGE = "C6:M"
Public Const FILE_RULE_NAME_RANGE = "D5:D28"
Public Const DEFINITE_TABLE_FIRST_COL = 21
'Column Name "b��dash���O�f�[�^�`�F�b�N�c�[���iVer0.10�j"
Public Const COL_DATA_CHECK_CHECK_BOX = "B"
Public Const COL_DATA_CHECK_NO = "C"
Public Const COL_DATA_CHECK_FILE_NAME = "D"
Public Const COL_DATA_CHECK_FILE_NAME_PATTERN = "E"
Public Const COL_DATA_CHECK_MAX_QUANTITY_RECORD = "G"
Public Const COL_DATA_CHECK_MAX_FILE_SIZE = "H"
Public Const COL_DATA_CHECK_FILE_PATH = "I"
Public Const COL_DATA_CHECK_STATUS_CHECK = "K"
Public Const COL_DATA_CHECK_DATE = "L"
Public Const COL_DATA_CHECK_SAVE = "M"
'Column Name "IF�t�@�C���ꗗ"
Public Const COL_FILE_LIST_ROW_NUM = "B"
Public Const COL_FILE_LIST_FILE_NAME = "C"
Public Const COL_FILE_LIST_NAME_PATTERN = "D"
Public Const COL_FILE_LIST_FILE_TYPE = "E"
Public Const COL_FILE_LIST_DELIMITER = "F"
Public Const COL_FILE_LIST_ENCODING = "G"
Public Const COL_FILE_LIST_HEADER = "H"
Public Const COL_FILE_LIST_HEADER_START_ROW_DATA = "I"
Public Const COL_FILE_LIST_NEW_LINE = "J"
Public Const COL_FILE_LIST_FLAG_RECORD_SIZE = "L"
Public Const COL_FILE_LIST_MAX_QUANTITY_RECORD = "M"
Public Const COL_FILE_LIST_MAX_FILE_SIZE = "N"
'Column Name of "DefiniteTable"
Public Const COL_DEFINITE_TABLE_No = "B"
Public Const COL_DEFINITE_TABLE_NAME = "D"
Public Const COL_DEFINITE_TABLE_DATA_TYPE = "F"
Public Const COL_DEFINITE_TABLE_PRIMARY_KEY = "G"
Public Const COL_DEFINITE_TABLE_NOT_NULL = "H"
Public Const COL_DEFINITE_TABLE_DATE_FORMAT = "L"
'Message
Public Const CLEAR_SUCCESS_MSG = "�Ώۂ��N���A���܂����B"
Public Const CLEAR_CONFIRM_MSG = "�Ώۈꗗ�\�̓��e���N���A���܂����A��낵���ł��傤���B"
Public Const OVERVIEW_FILE_ERROR_MSG = "IF�t�@�C���ꗗ�V�[�g�̍s%{rowNum}�ɂ́u�t�@�C���T�v�v�܂��́u�t�@�C�������K���v�܂��́u�ő僌�R�[�h���v�܂��́u �ő�t�@�C���e�ʁv���܂���`����Ă��܂���B"
Public Const FILE_NOT_SELECT_MSG = "�t�@�C���p�X���܂��ݒ肳��ĂȂ��s������܂��B�`�F�b�N�������s�́u�I���v�{�^�����N���b�N���A�`�F�b�N�Ώۂ̃t�@�C���p�X��ݒ肵�Ă��������B"
Public Const ALL_CHECKBOX_NOT_CHECKED_MSG = "�`�F�b�N�Ώۃt�@�C���͂܂��ݒ肳��Ă��܂���B�`�F�b�N�������t�@�C���̍s���́����N���b�N���ă`�F�b�N���Ώۃt�@�C����I�����Ă��������B"
Public Const FILE_NOT_EXISTS_MSG = "%{fileName}�t�@�C���͑��݂��Ă��܂���B"
Public Const FILE_OVER_LIMIT_SIZE_MSG = "%{fileName}�F�ő�t�@�C���e�ʂ𒴂��Ă��܂��B"
Public Const NOTIFICATE_NOT_SELECT_FILE = "�I���{�^�����N���b�N���A�`�F�b�N�Ώۂ̃t�@�C���p�X��ݒ肵�Ă��������B"
Public Const ERROR_FILE_NAME_RULE_DUPLICATE = "%{fileName}���������݂��܂��B�t�@�C�������K���͈�ӂɂȂ�悤�ɐݒ肵�Ă��������B"
Public Const ERROR_FILE_EXTENSION = "%{fileName}�F�t�@�C���̊g���q�͒�`�ƈقȂ��Ă��܂��B"
Public Const ERROR_END_LINE_CHARACTER = "%{fileName}�F�`�F�b�N�Ώۍs�ɒ�`�ƈقȂ���s�������܂܂�Ă��܂��B"
Public Const ERROR_NAME_RULE_APHABLE = "%{fileName}�F�t�@�C������1�����ڂ����p�p���������̓A���_�[�o�[�̂ǂ��炩�Ŏn�܂�悤�ɐݒ肵�Ă��������B"
Public Const ERROR_NAME_RULE = "%{fileName}�F�t�@�C������IF�t�@�C���ꗗ�Œ�`���ꂽ�����K���ɓY���Ă��܂���B"
Public Const ERROR_BOM = "%{fileName}�FBOM�t���̃t�@�C���ł��B"
Public Const ERROR_ENDCODING = "%{fileName}�F�`�F�b�N�Ώۍs�ɒ�`�ƈقȂ镶���R�[�h���܂܂�Ă��܂��B"
Public Const ERROR_SEPERATED_CHARACTER = "%{fileName}�F�s%{row}�̃J�������̓J������`�V�[�g�ŋL�ڂ���Ă���J�������ƈ�v���Ă��܂���B�܂��A�`�F�b�N�Ώۍs�̂Ȃ��ɋ�؂蕶�����K�؂ł͂Ȃ��s���܂܂�Ă��܂��B"
Public Const ERROR_NOT_DATA_DEFINITE_TABLE = "%{fileName}�̃J������`�V�[�g����������`����Ă��܂���B" & vbNewLine & "�P�J�����ȏ�̏��𐳂�����`���Ă��������B"
Public Const ERROR_OVER_RECORD = "%{fileName}�F�t�@�C���̍ő僌�R�[�h�𒴂��Ă��܂��B"
Public Const CHECK_EXISTS_SHEET = "�Y���J������`�V�[�g�͊��ɑ��݂��Ă��܂��B"
Public Const ERROR_NOT_EXISTS_SHEET = "%{fileName}�̃J������`�V�[�g�͑��݂��Ă��܂���B" & vbNewLine & "�`�F�b�N�����s�́u�쐬�v�{�^�����N���b�N���A�J������`�V�[�g���쐬���Ă��������B"
Public Const ERROR_COLUMN_NOT_NULL = "%{fileName}�F�s%{row}��%{column}�J�����Ƀf�[�^��ݒ肵�Ă��������B"
Public Const ERROR_COLUMN_PRIMARY_KEY = "%{fileName}�F�s%{row}��%{column}�J�����ɂ͏d���f�[�^������܂��B"
Public Const ERROR_COLUMN_DATE_FORMAT = "%{fileName}�F�s%{row}��%{column}�J�����̒l�̃t�H�[�}�b�g�͒�`�ƈ�v���Ă��܂���B"
Public Const ERROR_DOUBLE_QUOTE = "%{fileName}�F�s%{row}�ɁA�J�������͋K���ƈقȂ�J���������݂��܂��B"
Public Const ERROR_DUPLICATE_NAME_PATTERN = "%{namePattern}���������݂��܂��B�t�@�C�������K���͈�ӂɂȂ�悤�ɐݒ肵�Ă��������B"
Public Const STATUS_PROCESSING = "���{��"
Public Const STATUS_PROCESS_COMPLETED = "�`�F�b�N����"
Public Const STATUS_PROCESS_COMPLETED_OK = "�����i����j"
Public Const STATUS_PROCESS_COMPLETED_NOK = "�����i�ُ킠��j"
Public Const STATUS_PROCESS_INIT_FILE = "�����{"
Public Const STATUS_PROCESS_STOP = "���f"

'Variant
Public dateColumns As Scripting.Dictionary
Public parseErrorRows As Scripting.Dictionary
Public currentCheckingRow As Integer
Public isNormal As Boolean
Public isForceStopNow As Boolean
'Checkbox check all handle
Sub chkAll_Click()
    Dim CB As CheckBox
    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name Then
            rowNum = Split(CB.Name, " ")(1)
            fileOverView = Trim(Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).Value)
            'Uncheck all
            If ActiveSheet.CheckBoxes("chkAll").Value <> 1 Then
                CB.Value = ActiveSheet.CheckBoxes("chkAll").Value
            'Check all
            ElseIf fileOverView <> "" Then
                CB.Value = ActiveSheet.CheckBoxes("chkAll").Value
            End If
      End If
    Next CB
End Sub

' Checkbox list handle
Sub Mixed_State()
    Dim CB As CheckBox
    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.Value <> ActiveSheet.CheckBoxes("chkAll").Value And ActiveSheet.CheckBoxes("chkAll").Value <> 2 Then
            ActiveSheet.CheckBoxes("chkAll").Value = 2
            Exit For
        Else
            ActiveSheet.CheckBoxes("chkAll").Value = CB.Value
        End If
    Next CB
End Sub

'Clear button handle
Sub btnClear_Click()
    clearContent
    MsgBox CLEAR_SUCCESS_MSG
End Sub

'Checking handle
Sub btnProcess_Click()
    Log.createFileLog
    'All input validation
    'Validate all checkbox are not checked
    checkboxValid = Validation.allCheckboxNotChecked()
    fileEmptyValid = True
    fileExistsValid = True
    fileNotExistsList = ""
    Set dateColumns = New Scripting.Dictionary
    Set parseErrorRows = New Scripting.Dictionary
    isForceStopNow = False
    
    If checkboxValid Then
        'Loop all checkbox to check file, if has error to call MessageBOX
        For Each CB In ActiveSheet.CheckBoxes
            DoEvents
            If isForceStopNow = False And CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.Value = 1 Then
                'Row number processing
                rowNum = Split(CB.Name, " ")(1)
                currentCheckingRow = rowNum
                '1.4 validate definiteSheet
                fileNameOverView = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).Value
                definiteSheetName = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowNum).Value
                If checkExistsSheet(definiteSheetName) = Fasle Then
                    MsgBox Replace(ERROR_NOT_EXISTS_SHEET, "%{fileName}", fileNameOverView)
                    Exit Sub
                End If
                'Get file path from list
                filePath = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).Value
                filePath = Trim(CStr(filePath))

                'Validate file not select
                If IsEmpty(filePath) Or filePath = "" Or filePath = NOTIFICATE_NOT_SELECT_FILE Then
                    fileEmptyValid = False
                    Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).Value = NOTIFICATE_NOT_SELECT_FILE
                    Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowNum).Font.Color = vbRed
                 Else
                    'Validate file does not exists
                    If file.exists(filePath) = False Then
                        fileExistsValid = False
                        fileNotExistsList = fileNotExistsList & vbCrLf & Replace(FILE_NOT_EXISTS_MSG, "%{fileName}", Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & rowNum).Value)
                    Else
                    End If
                End If
                currentCheckingRow = currentCheckingRow + 1
            ElseIf isForceStopNow = True Then Exit Sub
            End If
       Next CB

       '2 Validate Duplicate file name
       Dim fileRule As Worksheet
       Set fileRule = Common.fileListSheet()
       For Each E In fileRule.Range(FILE_RULE_NAME_RANGE)
        a = WorksheetFunction.CountIf(fileRule.Range(FILE_RULE_NAME_RANGE), E)
        If a >= 2 Then
            MsgBox Replace(ERROR_FILE_NAME_RULE_DUPLICATE, "%{fileName}", E)
            Exit Sub
        End If
       Next E

       If fileEmptyValid = False Then
            MsgBox FILE_NOT_SELECT_MSG
       ElseIf fileExistsValid = False Then
            MsgBox Replace(fileNotExistsList, vbCrLf, "", 1, 1)
       End If
    End If

    'Validate file content
    If checkboxValid And fileEmptyValid And fileExistsValid Then
       
        For Each CB In ActiveSheet.CheckBoxes
            Set dateColumns = New Scripting.Dictionary
            isNormal = True
           DoEvents
           If isForceStopNow = False And CB.Name <> ActiveSheet.CheckBoxes("chkAll").Name And CB.Value = 1 Then
               'Row number processing
                rowNum = Split(CB.Name, " ")(1)
                currentCheckingRow = rowNum
                Call updateStatusProcess(rowNum, STATUS_PROCESSING)

               'Get file path from list
                Dim dataCheck As Worksheet
                Set dataCheck = Common.dataCheckSheet()

                filePath = dataCheck.Range(COL_DATA_CHECK_FILE_PATH & rowNum).Value
                filePath = Trim(CStr(filePath))
                fileOverView = dataCheck.Range(COL_DATA_CHECK_FILE_NAME & rowNum).Value

                Dim fileList As Worksheet
                Set fileList = Common.fileListSheet()

                fileRowIndex = dataCheck.Range(COL_DATA_CHECK_SAVE & rowNum).Value
                limitSize = fileList.Range(COL_FILE_LIST_MAX_FILE_SIZE & fileRowIndex).Value
                fileNameRule = fileList.Range(COL_FILE_LIST_NAME_PATTERN & fileRowIndex).Value
                extensionFileList = fileList.Range(COL_FILE_LIST_FILE_TYPE & fileRowIndex).Value
                delimiter = fileList.Range(COL_FILE_LIST_DELIMITER & fileRowIndex).Value
                newLineDeclare = fileList.Range(COL_FILE_LIST_NEW_LINE & fileRowIndex).Value
                endcodingType = fileList.Range(COL_FILE_LIST_ENCODING & fileRowIndex).Value
                maxRecordRule = fileList.Range(COL_FILE_LIST_MAX_QUANTITY_RECORD & fileRowIndex).Value
                flagRecordSize = fileList.Range(COL_FILE_LIST_FLAG_RECORD_SIZE & fileRowIndex).Value
                isHeader = fileList.Range(COL_FILE_LIST_HEADER & fileRowIndex).Value
                startRowData = fileList.Range(COL_FILE_LIST_HEADER_START_ROW_DATA & fileRowIndex).Value
                If isHeader = "����" And IsNumeric(startRowData) Then
                    If startRowData < 1 Then
                        startRowData = 0
                    End If
                Else
                    startRowData = 0
                End If
                'B.2 Get quantity Column
                Dim nameSheeetDefiniteTable As String
                nameSheeetDefiniteTable = dataCheck.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowNum).Value
                Dim sheetOfDefiniteTable As Worksheet
                Set sheetOfDefiniteTable = Common.definiteSheet(nameSheeetDefiniteTable)
                lastRow = lastHasData(1, nameSheeetDefiniteTable, "D22:D500")
                quantityColumnTable = lastRow - DEFINITE_TABLE_FIRST_COL
                If quantityColumnTable < 0 Then
                    MsgBox Replace(ERROR_NOT_DATA_DEFINITE_TABLE, "%{fileName}", fileOverView)
                    GoTo NextIterationCB
                End If

                Dim lstColumnNotNull
                Dim lstColPrimaryKey
                lstColumnNotNull = ""
                lstColPrimaryKey = ""
                For i = (DEFINITE_TABLE_FIRST_COL + 1) To lastRow Step 1
                    no = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_No & i).Value
                    typeDefinite = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_DATA_TYPE & i).Value
                    isPrimaryKey = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_PRIMARY_KEY & i).Value
                    isNotNull = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NOT_NULL & i).Value
                    dateFormat = Trim(sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_DATE_FORMAT & i).Value)
                    If isPrimaryKey = "Yes" Then
                        lstColPrimaryKey = lstColPrimaryKey & " " & no
                    End If
                    If isNotNull = "Yes" Then
                        lstColumnNotNull = lstColumnNotNull & " " & no
                    End If
                    'get date columns for 11 validate
                    If dateFormat <> "" Then
                        dateColumns.Add no, dateFormat
                    End If
                Next
                lstColPrimaryKey = Split(Trim(lstColPrimaryKey), " ")
                lstColumnNotNull = Split(Trim(lstColumnNotNull), " ")

                'Validate Bom
                'IsValidateBom = validateBom(detectBOM(filePath), fileOverView)
                If detectBOM(filePath) = True Then
                    isNormal = False
                    Log.ERROR (Replace(ERROR_BOM, "%{fileName}", fileOverView))
                    Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_NOK)
                    GoTo NextIterationCB
                End If

                '3 Validate file size
                isNormal = isNormal And vaidateFileSize(filePath, limitSize, fileOverView, flagRecordSize)

                '4 Validate file extension
                isNormal = isNormal And validateExtenstion(filePath, extensionFileList, fileOverView)

                '5 Validate newLineCharacter
                isNormal = isNormal And validateNewLineCharacter(newLineCharacter(filePath, 0), newLineDeclare, fileOverView)

                '6 Validate rule's name
                fileNameRule = Split(fileNameRule, "<")(0)
                isNormal = isNormal And validateNameRule(filePath, fileNameRule, fileOverView)

                '7.1 Validate encode
                isNormal = isNormal And validateEncoding(encoding(filePath), endcodingType, fileOverView)

                '8 Validate sperated character
                newLineChar = file.newLineCharacter(filePath, 1)
                FileType = fileList.Range(COL_FILE_LIST_FILE_TYPE & fileRowIndex)
                If FileType = "tsv" Then
                 FileType = vbTab
                Else
                 FileType = ","
                End If
                
                Dim rawContent() As String
                rawContent = getRawContent(filePath, 3000)
                csvContent = readCSV(rawContent, FileType, 1000, quantityColumnTable, fileOverView, startRowData, currentCheckingRow)
                If IsNull(csvContent) Or IsEmpty(csvContent) Then
                    GoTo NextIterationCB
                End If

                isNormal = isNormal And checkValidateSeperatedCharacter(filePath, FileType, quantityColumnTable, fileOverView, startRowData)

                'Check Primary key
                For Each pkey In lstColPrimaryKey
                    rowNumCSV = 1
                    errorRowNumCSV = ""
                    errorColumnName = ""
                    For Each Row In csvContent
                        Count = 0
                        DoEvents
                        If isForceStopNow = True Then Exit Sub
                        For Each Row2 In csvContent
                            If Row(1, pkey) = Row2(1, pkey) Then
                                Count = Count + 1
                            End If
                        Next
                        If Count > 1 Then
                            isNormal = False
                            errorRowNumCSV = errorRowNumCSV & "�A" & rowNumCSV
                            errorColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + pkey).Value
                        End If
                        rowNumCSV = rowNumCSV + 1
                    Next
                    errorRowNumCSV = Replace(errorRowNumCSV, "�A", "", 1, 1)
                    If IsEmpty(errorRowNumCSV) = False And Trim(errorRowNumCSV) <> "" Then
                        Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_PRIMARY_KEY, "%{fileName}", fileOverView), "%{row}", errorRowNumCSV), "%{column}", errorColumnName))
                    End If
                Next pkey
                
                rowNumCSV = 1
                errorRowNumValidateDateCSV = ""
                For Each csvRow In csvContent
                    'Check NOT NULL
                    DoEvents
                    If isForceStopNow = True Then Exit Sub
                    errorColumnNameNotNull = ""
                    For Each n In lstColumnNotNull
                        n = CInt(n)
                        If csvRow(1, n) = "" Then
                            isNormal = False
                            ColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + n).Value
                            errorColumnNameNotNull = errorColumnNameNotNull & "�A" & ColumnName
                        End If
                    Next n
                    errorColumnNameNotNull = Replace(errorColumnNameNotNull, "�A", "", 1, 1)
                    If IsEmpty(errorColumnNameNotNull) = False And Trim(errorColumnNameNotNull) <> "" Then
                        Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_NOT_NULL, "%{fileName}", fileOverView), "%{row}", rowNumCSV), "%{column}", errorColumnNameNotNull))
                    End If
                    ' 11 validate
                    Dim key As Variant
                    Dim FormatString As String
                    Dim OriginalValue As String
                    Dim FormattedValue As String
                    For Each key In dateColumns.Keys
                        FormatString = dateColumns(key)
                        OriginalValue = csvRow(1, key)
                        ColumnName = sheetOfDefiniteTable.Range(COL_DEFINITE_TABLE_NAME & DEFINITE_TABLE_FIRST_COL + key).Value

                        If IsDate(OriginalValue) Then
                            FormattedValue = Format(OriginalValue, FormatString)
                            If FormattedValue <> OriginalValue Then
                                isNormal = False
                                errorRowNumValidateDateCSV = errorRowNumValidateDateCSV & "�A" & rowNumCSV
                            End If
                        Else
                            isNormal = False
                            errorRowNumValidateDateCSV = errorRowNumValidateDateCSV & "�A" & rowNumCSV
                        End If
                    Next key
                    If rowNumCSV > 1000 Then
                        Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_OK)
                        GoTo NextIterationCB
                    Else
                        rowNumCSV = rowNumCSV + 1
                    End If
                Next csvRow
                errorRowNumValidateDateCSV = Replace(errorRowNumValidateDateCSV, "�A", "", 1, 1)
                If IsEmpty(errorRowNumValidateDateCSV) = False And Trim(errorRowNumValidateDateCSV) <> "" Then
                    Log.ERROR (Replace(Replace(Replace(ERROR_COLUMN_DATE_FORMAT, "%{fileName}", fileOverView), "%{row}", errorRowNumValidateDateCSV), "%{column}", ColumnName))
                End If
                '9 Validate maxRecord
                fileRecordQuantity = getFileLine(filePath)
                isNormal = isNormal And validateMaxRecord(fileRecordQuantity, maxRecordRule, fileOverView, flagRecordSize)
                If isNormal = True Then
                    Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_OK)
                Else
                    Call updateStatusProcess(rowNum, STATUS_PROCESS_COMPLETED_NOK)
                End If
                currentCheckingRow = currentCheckingRow + 1
           ElseIf isForceStopNow = True Then Exit Sub
           End If
NextIterationCB:
        Next CB
    End If
End Sub

'Select file handle
Sub btnSelectFile_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

    .AllowMultiSelect = False
    .Title = "Please select the file to kill his non colored cells"
    .Filters.Add "CSV", "*.csv"
    .Filters.Add "TSV", "*.tsv"

    If .Show = True Then
        rowClicked = Split(ActiveSheet.Shapes(Application.Caller).Name)(1)
        txtFileName = .SelectedItems(1)
        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowClicked).Value = txtFileName
        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_PATH & rowClicked).Font.Color = vbBlack
    End If

    End With
End Sub

'Get conditions handle
Sub btnGetextractionList_Click()
    mbResult = MsgBox(CLEAR_CONFIRM_MSG, vbYesNo)

    If mbResult = vbYes Then
        clearContent

        Dim errorRows As String
        errorRows = ""

        lastRow = Common.lastRowFileListSheet()
        If lastRow > 104 Then
            lastRow = 104
        End If
        addIndex = 6
        errorIndex = 0

        Dim stringPattern As String
        stringPattern = ""
        Dim errorPatterns As String
        errorPatterns = ""

        For i = 5 To lastRow

            rowNum = Common.fileListSheet.Range(COL_FILE_LIST_ROW_NUM & i).Value
            name_pattern = Common.fileListSheet.Range(COL_FILE_LIST_NAME_PATTERN & i).Value
            fileOverView = Common.fileListSheet.Range(COL_FILE_LIST_FILE_NAME & i).Value
            maxRecord = Common.fileListSheet.Range(COL_FILE_LIST_MAX_QUANTITY_RECORD & i).Value
            maxFileSize = Common.fileListSheet.Range(COL_FILE_LIST_MAX_FILE_SIZE & i).Value
            flagRecordSize = Common.fileListSheet.Range(COL_FILE_LIST_FLAG_RECORD_SIZE & i).Value
            flagDuplicateNamePattern = False

            If IsEmpty(fileOverView) = False Then
                If (IsEmpty(maxRecord) Or maxRecord = 0 Or IsEmpty(maxFileSize) Or maxFileSize = 0) Then
                    If flagRecordSize = "�S��" Then
                        errorRows = errorRows & " " & rowNum
                    Else
                        arrayPattern = Split(stringPattern, ";")
                        For J = 0 To UBound(arrayPattern)
                            If arrayPattern(J) = Split(name_pattern, "<")(0) Then
                                errorPatterns = errorPatterns & " " & arrayPattern(J) & ":" & rowNum
                                flagDuplicateNamePattern = True
                                'Exit Sub
                            End If
                        Next J
                        If flagDuplicateNamePattern = False Then
                            stringPattern = stringPattern & ";" & Split(name_pattern, "<")(0)
                        End If

                        If Trim(errorPatterns) = "" Then
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_NO & addIndex).Value = rowNum
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & addIndex).Value = Split(name_pattern, "<")(0)
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & addIndex).Value = fileOverView
                            If flagRecordSize = "�S��" Then
                                Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_QUANTITY_RECORD & addIndex).Value = maxRecord
                                Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_FILE_SIZE & addIndex).Value = maxFileSize
                            End If
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_STATUS_CHECK & addIndex).Value = STATUS_PROCESS_INIT_FILE
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_SAVE & addIndex).Value = i
                        End If
                        addIndex = addIndex + 1
                    End If
                ElseIf IsEmpty(name_pattern) Then
                    errorRows = errorRows & " " & rowNum
                Else
                    arrayPattern = Split(stringPattern, ";")
                    For J = 0 To UBound(arrayPattern)
                        If arrayPattern(J) = Split(name_pattern, "<")(0) Then
                            errorPatterns = errorPatterns & " " & arrayPattern(J) & ":" & rowNum
                            flagDuplicateNamePattern = True
                            'Exit Sub
                        End If
                    Next J
                    If flagDuplicateNamePattern = False Then
                        stringPattern = stringPattern & ";" & Split(name_pattern, "<")(0)
                    End If
                    If Trim(errorPatterns) = "" Then
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_NO & addIndex).Value = rowNum
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & addIndex).Value = Split(name_pattern, "<")(0)
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME & addIndex).Value = fileOverView
                        If flagRecordSize = "�S��" Then
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_QUANTITY_RECORD & addIndex).Value = maxRecord
                            Common.dataCheckSheet.Range(COL_DATA_CHECK_MAX_FILE_SIZE & addIndex).Value = maxFileSize
                        End If
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_STATUS_CHECK & addIndex).Value = STATUS_PROCESS_INIT_FILE
                        Common.dataCheckSheet.Range(COL_DATA_CHECK_SAVE & addIndex).Value = i
                    End If
                    addIndex = addIndex + 1
                End If
            ElseIf (IsEmpty(fileOverView)) And checkRowEmpty(i) Then
                errorRows = errorRows & " " & rowNum
            End If
        Next i
        If IsEmpty(Trim(errorRows)) = False And Trim(errorRows) <> "" Then
            Validation.extractionEmptyFileSize (errorRows)
        Else
            Validation.duplicateNamePattern (errorPatterns)
        End If
    End If
End Sub

' Cancel checking handle
Sub btnCancel_Click()
    isForceStopNow = True
    Call updateStatusProcess(currentCheckingRow, STATUS_PROCESS_STOP)
End Sub

'Clear without confirmation
Private Sub clearContent()
    lastRow = Common.lastRowDataCheckSheet()
    If lastRow > 5 Then
        Common.dataCheckSheet.Range(DATA_CHECK_RANGE & lastRow).ClearContents
    End If
    For Each CB In ActiveSheet.CheckBoxes
      CB.Value = 0
    Next CB
End Sub

'Create Definite Sheet
Sub btnCreateDefiniteSheet_Click()
    rowClicked = Split(ActiveSheet.Shapes(Application.Caller).Name, " ")(1)
    sheetNamePattern = Common.dataCheckSheet.Range(COL_DATA_CHECK_FILE_NAME_PATTERN & rowClicked).Value
    sheetNameRule = Common.fileListSheet.Range(COL_FILE_LIST_NAME_PATTERN & rowClicked - 1).Value
    'Check exists sheets
    If sheetNamePattern = "" Then
        Exit Sub
    End If
    If checkExistsSheet(sheetNamePattern) Then
        MsgBox CHECK_EXISTS_SHEET
        Exit Sub
    End If
    'Copy sheet template
    Sheets(TABLE_DEFINITE_TEMPLATE).Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = sheetNamePattern
    Sheets(sheetNamePattern).Range("B2").Value = Split(sheetNameRule, "<")(0) & "�J������`�V�[�g"
End Sub

'Check exists sheet
Private Function checkExistsSheet(ByVal sheetName As String)
    For Each sh In ActiveWorkbook.Sheets
        If sh.Name = sheetName Then
            checkExistsSheet = True
            Exit Function
        End If
    Next sh
    checkExistsSheet = False
End Function
'Check row empty
Private Function checkRowEmpty(ByVal index As Integer)
    D = Common.fileListSheet.Range("D" & index).Value
    E = Common.fileListSheet.Range("E" & index).Value
    g = Common.fileListSheet.Range("G" & index).Value
    H = Common.fileListSheet.Range("H" & index).Value
    i = Common.fileListSheet.Range("i" & index).Value
    J = Common.fileListSheet.Range("J" & index).Value
    K = Common.fileListSheet.Range("K" & index).Value
    L = Common.fileListSheet.Range("L" & index).Value
    M = Common.fileListSheet.Range("M" & index).Value
    n = Common.fileListSheet.Range("n" & index).Value
    O = Common.fileListSheet.Range("O" & index).Value
    P = Common.fileListSheet.Range("P" & index).Value
    Q = Common.fileListSheet.Range("Q" & index).Value
    r = Common.fileListSheet.Range("r" & index).Value
    s = Common.fileListSheet.Range("s" & index).Value
    T = Common.fileListSheet.Range("T" & index).Value
    U = Common.fileListSheet.Range("U" & index).Value
    v = Common.fileListSheet.Range("v" & index).Value
    W = Common.fileListSheet.Range("W" & index).Value
    X = Common.fileListSheet.Range("X" & index).Value
    Y = Common.fileListSheet.Range("Y" & index).Value
    Z = Common.fileListSheet.Range("Z" & index).Value
    AA = Common.fileListSheet.Range("AA" & index).Value
    AB = Common.fileListSheet.Range("AB" & index).Value
    AC = Common.fileListSheet.Range("AC" & index).Value
    AD = Common.fileListSheet.Range("AD" & index).Value
    AE = Common.fileListSheet.Range("AE" & index).Value
    AF = Common.fileListSheet.Range("AF" & index).Value
    AG = Common.fileListSheet.Range("AG" & index).Value
    AH = Common.fileListSheet.Range("AH" & index).Value
    AI = Common.fileListSheet.Range("AI" & index).Value
    AJ = Common.fileListSheet.Range("AJ" & index).Value
    AK = Common.fileListSheet.Range("AK" & index).Value
    AL = Common.fileListSheet.Range("AL" & index).Value
    AM = Common.fileListSheet.Range("AM" & index).Value

    If IsEmpty(D) = False Or IsEmpty(E) = False Or IsEmpty(g) = False Or IsEmpty(H) = False Or IsEmpty(i) = False Or IsEmpty(J) = False Or IsEmpty(K) = False Or IsEmpty(L) = False Or IsEmpty(M) = False Or IsEmpty(n) = False Or IsEmpty(O) = False Or IsEmpty(P) = False Or IsEmpty(Q) = False Or IsEmpty(r) = False Or IsEmpty(s) = False Or IsEmpty(T) = False Or IsEmpty(U) = False Or IsEmpty(v) = False Then
        checkRowEmpty = True
        Exit Function
    ElseIf IsEmpty(AD) = False Or IsEmpty(AE) = False Or IsEmpty(AF) = False Or IsEmpty(AG) = False Or IsEmpty(AH) = False Or IsEmpty(AI) = False Or IsEmpty(AJ) = False Or IsEmpty(AK) = False Or IsEmpty(AL) = False Or IsEmpty(AM) = False Or IsEmpty(W) = False Or IsEmpty(X) = False Or IsEmpty(Y) = False Or IsEmpty(Z) = False Or IsEmpty(AA) = False Or IsEmpty(AB) = False Or IsEmpty(AC) = False Then
        checkRowEmpty = True
        Exit Function
    End If
    checkRowEmpty = False

End Function



'Read CSV by binary
'@param strLine                   Line to check
'@param strSeperatedChar   Seperated character
'@param intNoOfCol             Defined number of column of file
Function checkValidateSeperatedCharacter(ByVal fullFileName As String, ByVal strSeperatedChar As String, ByVal noOfCol As Integer, ByVal fileOverView As String, ByVal startRowData As Integer)

    Dim intUnit As Integer
    Dim my_string As String
    Dim vntLines As Variant

    intUnit = FreeFile
    Open fullFileName For Binary Access Read As #intUnit
    my_string = Input(size(fullFileName), intUnit)

    If InStr(my_string, vbCrLf) > 0 Then
        'Window file
        vntLines = Split(my_string, vbCrLf)
        vntLines = getProcessLine(vntLines, startRowData)

    ElseIf InStr(my_string, vbCr) > 0 Then
        'MAC file
        vntLines = Split(my_string, vbCr)
        vntLines = getProcessLine(vntLines, startRowData)

    Else
        'Unix file
        vntLines = Split(my_string, vbLf)
        vntLines = getProcessLine(vntLines, startRowData)

    End If
    Close intUnit
    checkValidateSeperatedCharacter = True
    'Checking
    For i = (LBound(vntLines) + startRowData) To UBound(vntLines)
        If checkSeperatedCharacter(vntLines(i), strSeperatedChar, noOfCol) = False Then
            checkValidateSeperatedCharacter = False
            Log.ERROR (Replace(Replace(ERROR_SEPERATED_CHARACTER, "%{fileName}", fileOverView), "%{row}", i + 1 + startRowData))
        End If
    Next
End Function

' Read and parse CSV file
' Return array 2d if parse success, otherwise return Null
Function readCSV(ByRef arrayData() As String, ByVal separater As String, ByVal limitLine As Long, ByVal columnCount As Integer, ByVal fileOverView As String, ByVal startRow As Integer, ByVal currentRowFile As Integer) As Variant
    Dim textData As String
    Dim textRow As String
    
    Dim csv As Variant
    Dim csvArray As Variant
    Dim errorTmp As String
    Dim lineIndex As Integer

    ReDim csvArray(1 To limitLine) As Variant
    arrayIndex = 1
    lineIndex = 1
    IsValid = True
    csv = Null
    SetCSVUtilsAnyErrorIsFatal (True)

    If columnCount = 1 Then
        parttern = "^\""([^\""]|\""\"")*\""$"

    ElseIf columnCount = 2 Then
        parttern = "^\""([^\""]|\""\"")*\""" & separater & "\""([^\""]|\""\"")*\""$"
    Else
        parttern = "^\""([^\""]|\""\"")*\""" & separater

        For i = 1 To columnCount - 2
            parttern = parttern & "\""([^\""]|\""\"")*\""" & separater
        Next i
        parttern = parttern & "\""([^\""]|\""\"")*\""$"
    End If

    Dim regEx As New RegExp

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = parttern
    End With

    For lineIndex = 1 To UBound(arrayData)
        textRow = arrayData(lineIndex)
        If textData = "" Then
            textData = textRow
        Else
            textData = textData & textRow
        End If
        
        'try parse CSV
        csv = ParseCSVToArray(textData, separater, fileOverView, arrayIndex)
        
        'parse csv error
        If IsNull(csv) Or UBound(csv) = -1 Then
            errorTmp = arrayIndex
        'parse ok
        Else
            'Validate quote
            If lineIndex > startRow Then
                If validateQuote(textData, regEx) = False Then
                    Log.ERROR (Replace(Replace(ERROR_DOUBLE_QUOTE, "%{fileName}", fileOverView), "%{row}", lineIndex))
                End If
            End If
            'parse csv success and column matched
            If UBound(csv, 2) = columnCount Then
                If lineIndex > startRow Then
                    csvArray(arrayIndex) = csv
                    arrayIndex = arrayIndex + 1
                End If
                textData = ""
                errorTmp = ""
            'parse csv ok but column not match
            ElseIf UBound(csv, 2) <> columnCount Then
                If parseErrorRows.exists(lineIndex) = False Then
                    parseErrorRows.Add lineIndex, lineIndex
                End If
                errorTmp = ""
                textData = ""
            End If
        End If
        
    Next lineIndex

    'check has error
    If errorTmp <> "" Then
        If parseErrorRows.exists(errorTmp) = False Then
            parseErrorRows.Add errorTmp, errorTmp
        End If
        errorTmp = ""
    End If

    'Log if has error
    If textData <> "" Or parseErrorRows.Count > 0 Then
        IsValid = False
        For Each key In parseErrorRows.Keys
            Log.ERROR (Replace(Replace(ERROR_SEPERATED_CHARACTER, "%{fileName}", fileOverView), "%{row}", parseErrorRows(key) + startRow))
        Next
        Call updateStatusProcess(currentRowFile, STATUS_PROCESS_COMPLETED_NOK)
    End If

    'trim result array
    If IsValid = True Then
        Dim csvTmp As Variant
        If arrayIndex > 1 Then
            If (arrayIndex < limitLine) Then
                ReDim csvTmp(1 To arrayIndex - 1) As Variant
                For i = 1 To arrayIndex - 1
                    csvTmp(i) = csvArray(i)
                Next i
            Else
                csvTmp = csvArray
            End If
        End If
        'return value
        readCSV = csvTmp
        'free memory
        'Erase csvTmp
    Else
        readCSV = Null
        Call updateStatusProcess(currentRowFile, STATUS_PROCESS_COMPLETED_NOK)
    End If

    'free memory
    Erase csvArray
    parseErrorRows.RemoveAll
End Function

Sub updateStatusProcess(ByVal rowNum As Integer, ByVal statusProcess As String)
    If ActiveSheet.CheckBoxes("Check " & rowNum).Value = 1 Then
        Common.dataCheckSheet.Range(COL_DATA_CHECK_STATUS_CHECK & rowNum).Value = statusProcess
        Common.dataCheckSheet.Range(COL_DATA_CHECK_DATE & rowNum).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
End Sub

