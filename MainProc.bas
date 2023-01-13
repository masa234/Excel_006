
'【概要】親番号に抜けがあるかどうか
Public Function IsMissing(ByVal arrParentNumber As Variant, _
                        ByVal strParentNumber As String) As Boolean
On Error GoTo IsMissing_Err

    IsMissing = False
    
    Dim strParentNumberMinusOne As String
    
    '親番号-1を取得
    strParentNumberMinusOne = GetParentNumberMinusOne(strParentNumber)
    
    '親番号配列に親番号-1が存在するかどうか
    If InArr(arrParentNumber, strParentNumberMinusOne) = False Then
        '存在しない場合、True(漏れあり)
        IsMissing = True
    End If
    
    IsMissing = True
    
IsMissing_Err:

IsMissing_Exit:

End Function


'【概要】親番号を取得する
Public Function GetParentNumber(ByVal strNumber As String) As String
On Error GoTo GetParentNumber_Err

    '末尾2文字目が"_"の場合
    If Right(strNumber, 2) = "_" Then
        '末尾2文字目以外を受け取る
        GetParentNumber = Left(strNumber, Len(strNumber) - 2)
    Else
        GetParentNumber = strNumber
    End If
    
GetParentNumber_Err:

GetParentNumber_Exit:
End Function


'【概要】親番号-1を取得する
Public Function GetParentNumberMinusOne(ByVal strParentNumber As String) As String
On Error GoTo GetParentNumberMinusOne_Err

    Dim strRightTwoString As String

    '末尾2文字
    strRightTwoString = Right(strParentNumber, 2)
    
    '末尾2桁をLong型に変換して、-1する。
    GetParentNumberMinusOne = Left(strParentNumber, Len(strParentNumber) - 2) & CStr(CLng(strRightTwoString) - 1)
    
GetParentNumberMinusOne_Err:

GetParentNumberMinusOne_Exit:
End Function


'【概要】行の中にある漏れを配列で取得する
Public Function GetMissingInRow(ByVal strSheetName As String) As Variant
On Error GoTo GetMissingInRow_Err
    
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim lngArrParentNumberIdx As Long
    Dim lngArrRetIdx As Long
    Dim strNumber As String
    Dim strParentNumber As String
    Dim strLatestParentNumber As String
    Dim strParentNumberMinusOne As String
    Dim strBeforeParentNumber As String
    Dim arrParentNumber() As Variant
    Dim arrRet() As Variant
    
    '最終行取得
    lngLastRow = ThisWorkbook.Worksheets(strSheetName).Cells(1, 1).End(xlDown).Row
    
    '配列の要素番号初期化
    lngArrParentNumberIdx = 0
    lngArrRetIdx = 0
    
    '最終行まで繰り返す
    For lngCurrentRow = 1 To lngLastRow
        '番号取得
        strNumber = ThisWorkbook.Worksheets(strSheetName).Cells(lngCurrentRow, 1).Value
        '親番号取得
        strParentNumber = GetParentNumber(strNumber)
        '初回
        If lngCurrentRow = 1 Then
            '最新親番号を更新
            strLatestParentNumber = strParentNumber
            '次の行へ
            GoTo nextRow
        End If
        '親番号が最新の親番号と等しくない場合、
        If strParentNumber <> strLatestParentNumber Then
            '配列再宣言
            ReDim Preserve arrParentNumber(lngArrParentNumberIdx)
            '親番号を配列に格納
            arrParentNumber(lngArrParentNumberIdx) = strParentNumber
            '配列の要素番号を1つ進める
            lngArrParentNumberIdx = lngArrParentNumberIdx + 1
            '漏れがあるかどうか調べる
            If IsMissing(arrParentNumber, strParentNumber) = True Then
                '漏れがある場合
                '前の親番号
                strBeforeParentNumber = ThisWorkbook.Worksheets(strSheetName).Cells(lngCurrentRow - 1, 1).Value
                '親番号-1を取得
                strParentNumberMinusOne = GetParentNumberMinusOne(strParentNumber)
                '前の親番号と親番号-1が等しくない間、繰り返す
                Do While strBeforeParentNumber <> strParentNumberMinusOne
                    '配列再宣言
                    ReDim Preserve arrRet(lngArrRetIdx)
                    '親番号を配列に格納
                    arrRet(lngArrRetIdx) = strParentNumberMinusOne
                    '配列の要素番号を1つ進める
                    lngArrRetIdx = lngArrRetIdx + 1
                    '親番号-1を取得
                    strParentNumberMinusOne = GetParentNumberMinusOne(strParentNumberMinusOne)
                Loop
            End If
        End If
nextRow:
    Next lngCurrentRow
                    
    GetMissingInRow = arrRet
    
GetMissingInRow_Err:

GetMissingInRow_Exit:
End Function
