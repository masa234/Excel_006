Public Const DATA_SHEET_NAME = "データ"
Public Const DATA_LOAD_FAILED = "データのロードに失敗しました。"
Public Const CSV_OUTPUT_FAILED = "CSVの出力に失敗しました。"
Public Const PROCCESS_COMPLETE = "処理が完了しました。"
Public Const CSV_FILE_NAME = "sample.csv"
Public Const CONFIRM = "確認"


'【概要】配列内に存在するか
Public Function InArr(ByVal arrSearch As Variant, _
                    ByVal strSearch As String) As Boolean
On Error GoTo InArr_Err

    Dim lngArrIdx As Long
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrSearch)
        '配列の要素番号が検索値の場合、True
        If arrSearch(lngArrIdx) = strSearch Then
            InArr = True
        End If
    Next lngArrIdx
    
InArr_Err:

InArr_Exit:
End Function


'【概要】配列をCSV出力
Public Function ArrToCSV(ByVal arrOutput As Variant, _
                        ByVal strCSVOutputPath As String) As Boolean
On Error GoTo ArrToCSV_Err

    Dim lngFreeFile As Long
    Dim lngArrIdx As Long
    Dim strValue As String
    
    ArrToCSV = False
    
    'フリーファイルを取得する
    lngFreeFile = FreeFile

    'CSVファイルを開く（書き込みモード）
    Open strCSVOutputPath For Output As #lngFreeFile

    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrOutput)
        '値
        strValue = arrOutput(lngArrIdx)
        '1行出力
        Print #lngFreeFile, strValue
    Next lngArrIdx
    
    ArrToCSV = True
    
ArrToCSV_Err:

ArrToCSV_Exit:
    'CSVファイルを閉じる
    Close #lngFreeFile
End Function
