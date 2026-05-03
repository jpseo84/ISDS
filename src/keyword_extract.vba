Sub KeywordMatcher()
    Dim wsData As Worksheet, wsKeys As Worksheet
    Dim KeyArr As Variant, DataArr As Variant, ResultArr As Variant
    Dim lastRowData As Long, lastRowKeys As Long
    Dim i As Long, j As Long, startRow As Long, endRow As Long
    Dim currentBatchSize As Long
    Dim matchFound As Boolean
    Dim regEx As Object '일본어 및 중국어 문자 제거용 정규식 객체
    Dim cleanData As String, cleanKey As String
    
    ' =======================================================
    ' 배치 사이즈를 설정합니다
    ' =======================================================
    Dim batchSize As Long
    batchSize = 100 ' 메모리 오류 대비 - 100행씩 처리
    ' =======================================================
    
    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsKeys = ThisWorkbook.Sheets("Keyword")
    
    ' Optimize Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 정규식 객체 생성 - CJK 문자 제거용
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    ' 중국어 및 일본어 문자를 포함한 CJK 범위를 제거하기 위해 유니코드 범위를 사용
    regEx.Pattern = "[\u3040-\u309F\u30A0-\u30FF\u3400-\u4DBF\u4E00-\u9FFF]" 
    
    ' 마지막 행 계산
    lastRowData = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    lastRowKeys = wsKeys.Cells(wsKeys.Rows.Count, "A").End(xlUp).Row
    
    ' 키워드 로드
    KeyArr = wsKeys.Range("A1:C" & lastRowKeys).Value
    
    ' 배치 사이즈에 맞춰 데이터 처리
    For startRow = 1 To lastRowData Step batchSize
        
        endRow = startRow + batchSize - 1
        If endRow > lastRowData Then endRow = lastRowData
        
        currentBatchSize = endRow - startRow + 1
        
        ' 상태 표시줄에 진행 상황 업데이트
        Application.StatusBar = "Processing batch: rows " & Format(startRow, "#,##0") & " to " & Format(endRow, "#,##0") & " of " & Format(lastRowData, "#,##0")
        DoEvents 
        
        ' Step1: LOAD - 현재 배치의 데이터를 배열로 로드
        DataArr = wsData.Range("C" & startRow & ":C" & endRow).Value
        
        ' 특수 케이스: 배치 사이즈가 1인 경우에도 배열로 처리하기 위해 강제로 2차원 배열로 변환
        If currentBatchSize = 1 Then
            ReDim DataArr(1 To 1, 1 To 1)
            DataArr(1, 1) = wsData.Range("C" & startRow).Value
        End If
        
        ' Step2: MATCH - 각 데이터 문자열에 대해 키워드 배열과 비교하여 일치 여부 확인
        ReDim ResultArr(1 To currentBatchSize, 1 To 3)
        
        For i = 1 To currentBatchSize
            matchFound = False
            
            If Trim(CStr(DataArr(i, 1))) <> "" Then
                
                ' 데이터 문자열에서 CJK 문자 제거 및 소문자 변환
                cleanData = regEx.Replace(CStr(DataArr(i, 1)), "")
                cleanData = LCase(cleanData)
                
                For j = 1 To UBound(KeyArr)
                    If Trim(CStr(KeyArr(j, 1))) <> "" Then
                        
                        ' 키워드를 소문자로 변환하여 비교
                        cleanKey = LCase(CStr(KeyArr(j, 1)))
                        
                        ' 데이터 문자열에 키워드가 포함되어 있는지 확인 (대소문자 구분)
                        If InStr(1, cleanData, cleanKey, vbBinaryCompare) > 0 Then
                            ' Record the successful match
                            ResultArr(i, 1) = KeyArr(j, 1) ' Keyword
                            ResultArr(i, 2) = KeyArr(j, 2) ' Category 1
                            ResultArr(i, 3) = KeyArr(j, 3) ' Category 2
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next j
            End If
            
            ' 키워드가 일치하지 않는 경우(일치 키워드 없는 경우) "N/A"로 표시
            If Not matchFound Then
                ResultArr(i, 1) = "N/A"
                ResultArr(i, 2) = ""
                ResultArr(i, 3) = ""
            End If
        Next i
        
        ' Step3: WRITE - 결과를 원래 데이터 시트의 D, E, F 열에 기록
        wsData.Range("D" & startRow).Resize(currentBatchSize, 3).Value = ResultArr
        
        ' Step4: CLEANUP - 메모리 해제 및 다음 배치 준비
        Erase DataArr
        Erase ResultArr
        
    Next startRow
    
    '  최종 정리 및 사용자 알림
    Set regEx = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Keyword Matching Process Completed", vbInformation
End Sub