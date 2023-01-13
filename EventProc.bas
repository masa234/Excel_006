
 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    Dim arrMissing() As Variant
    
    '漏れを配列で取得
    arrMissing = GetMissingInRow(DATA_SHEET_NAME)
    
    '配列をCSV出力
    If ArrToCSV(arrMissing, ThisWorkbook.Path & "\" & CSV_FILE_NAME) = False Then
        Call MsgBox(CSV_OUTPUT_FAILED, vbInformation, CONFIRM)
    End If
    
    '完了メッセージ
    Call MsgBox(PROCCESS_COMPLETE, vbInformation, CONFIRM)
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
End Sub
