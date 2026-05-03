Sub EnhancedKeywordMatcher()
    Dim wsData As Worksheet, wsKeys As Worksheet
    Dim DataArr As Variant, KeyArr As Variant, ResultArr As Variant
    Dim lastRowData As Long, lastRowKeys As Long
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    
    ' 워크시트는 다음 2개로 가정: "Data" 및 "Keyword"
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsKeys = ThisWorkbook.Sheets("Keyword")
    
    ' 성능 향상을 위해 화면 업데이트와 자동 계산을 일시적으로 비활성화
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 처리할 데이터와 키워드의 마지막 행을 Column C와 A에서 각각 찾기
    lastRowData = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    lastRowKeys = wsKeys.Cells(wsKeys.Rows.Count, "A").End(xlUp).Row
    
    ' 처리할 데이터 (Data 워크시트의 Column C) 전체를 메모리에 로드
    DataArr = wsData.Range("C1:C" & lastRowData).Value
    
    ' 키워드와 카테고리 (Keyword 워크시트의 A, B, C) 전체를 메모리에 로드
    KeyArr = wsKeys.Range("A1:C" & lastRowKeys).Value
    
    ' 출력 결과를 저장할 배열을 초기화 (DataArr와 같은 행 수, 3개의 열: Keyword, Category 1, Category 2)
    ReDim ResultArr(1 To UBound(DataArr), 1 To 3)
    
    ' Data 워크시트의 각 셀을 반복하면서 키워드와 비교
    For i = 1 To UBound(DataArr)
        matchFound = False
        
        ' 실제 데이터가 있는 셀인지 확인 (공백이 아닌 경우에만 처리)
        If Trim(CStr(DataArr(i, 1))) <> "" Then
            ' 총 키워드 목록을 반복하면서 현재 데이터 텍스트에 키워드가 포함되어 있는지 확인
            For j = 1 To UBound(KeyArr)
                
                ' 비어 있지 않은 키워드인지 확인
                If Trim(CStr(KeyArr(j, 1))) <> "" Then
                    If InStr(1, CStr(DataArr(i, 1)), CStr(KeyArr(j, 1)), vbTextCompare) > 0 Then
                        ResultArr(i, 1) = KeyArr(j, 1) ' Keyword > Column D
                        ResultArr(i, 2) = KeyArr(j, 2) ' Category 1 > Column E
                        ResultArr(i, 3) = KeyArr(j, 3) ' Category 2 > Column F
                        matchFound = True
                        Exit For
                    End If
                End If
                
            Next j
        End If
        
        ' 만약 일치하는 키워드가 발견되지 않았다면, Column D에 "N/A"를 출력하고 E와 F는 빈칸으로 남겨둠
        If Not matchFound Then
            ResultArr(i, 1) = "N/A" ' N/A - Column D
            ResultArr(i, 2) = ""
            ResultArr(i, 3) = ""
        End If
    Next i
    
    ' 모든 데이터를 처리한 후, 결과를 Data 워크시트의 Column D, E, F에 한 번에 출력
    wsData.Range("D1").Resize(UBound(ResultArr), 3).Value = ResultArr
    
    ' 기본 화면 업데이트 및 계산 설정 복원
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Enhanced Matching Complete!", vbInformation
End Sub