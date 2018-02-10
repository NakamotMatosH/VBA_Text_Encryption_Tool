Attribute VB_Name = "Encryption"
'##
'# Made by N
'
'
'
Public Const UAl = "0123456789~!@#$%^&*()-=+[]{}\|/?,.><`HIJKLMNOPTUXYVQABCDEFGRSWZ"
Public Const LAl = "0123456789~!@#$%^&*()-=+[]{}\|/?,.><`wvqponzyutsrhgfedcbxmlkjia"
Public Const NAl = "012345678987654321012345678987654321012345678987654321123456789"

Public Function Encry(ByVal var As Variant) As String

    Dim tmp As String: tmp = CStr(var)
    Dim tmp2() As String: ReDim tmp2(1 To Len(tmp))

    For i = 1 To Len(tmp)
        tmp2(i) = Mid(var, i, 1)
    Next

    Dim EncResult As String

    For i = LBound(tmp2) To UBound(tmp2)

        If Len(CStr(Hex(Asc(tmp2(i))))) > 3 Then
            If Application.WorksheetFunction.IsEven(Right(now(), 1)) = True Then
                tmp2(i) = X1e(tmp2(i))
            Else
                tmp2(i) = X2e(tmp2(i))
            End If
        Else
            If Application.WorksheetFunction.IsEven(Right(now(), 1)) = True Then
                tmp2(i) = X3e(tmp2(i))
            Else
                tmp2(i) = X4e(tmp2(i))
            End If
        End If

        EncResult = EncResult + tmp2(i)
    Next

    var = Empty
    tmp = Empty
    Encry = EncResult

End Function

Function X1e(ByVal CharStr As String) As String

    Dim k As String
    k = Hex(Asc(CharStr))
    Dim n As String
    n = Mid(k, 4, 1) + Mid(k, 2, 1) + Mid(k, 1, 1) + Mid(k, 3, 1)
    n = "X1" + n
    X1e = n

End Function

Function X2e(ByVal CharStr As String) As String

    Dim k As String
    k = Hex(Asc(CharStr))
    Dim n As String
    n = Mid(k, 3, 1) + Mid(k, 4, 1) + Mid(k, 1, 1) + Mid(k, 2, 1)
    n = "X2" + n
    X2e = n

End Function

Function X3e(ByVal CharStr As String) As String

    Dim k As String
    k = Hex(Asc(CharStr))
    Dim K1 As String, K2 As String
    K1 = CStr(Left(k, 1)): K2 = CStr(Right(k, 1))
    K1 = Hex(Asc(K1)): K2 = Hex(Asc(K2))
    k = K1 + K2

    Dim n As String
    n = Mid(k, 1, 1) + Mid(k, 3, 1) + Mid(k, 2, 1) + Mid(k, 4, 1)
    n = "X3" + n
    X3e = n

End Function

Function X4e(ByVal CharStr As String) As String

    Dim k As String
    k = Hex(Asc(CharStr))
    Dim K1 As String, K2 As String
    K1 = CStr(Left(k, 1)): K2 = CStr(Right(k, 1))
    K1 = Hex(Asc(K1)): K2 = Hex(Asc(K2))
    k = K1 + K2

    Dim n As String
    n = Mid(k, 3, 1) + Mid(k, 4, 1) + Mid(k, 1, 1) + Mid(k, 2, 1)
    n = "X4" + n
    X4e = n

End Function

Function DeCry(ByVal str As String) As String

    Dim k As Variant
    k = Split(str, "X")
    Dim res As String

    For i = LBound(k) To UBound(k)

        If Left(k(i), 1) = "1" Then
            res = res + X1d(k(i))
        ElseIf Left(k(i), 1) = "2" Then
            res = res + X2d(k(i))
        ElseIf Left(k(i), 1) = "3" Then
            res = res + X3d(k(i))
        ElseIf Left(k(i), 1) = "4" Then
            res = res + X4d(k(i))
        Else
        End If
    
    Next

DeCry = res
res = Empty

End Function

Function X1d(ByVal CharStr As String) As String

    Dim k As String: Dim n As String
    k = Right(CharStr, 4)
    n = Mid(k, 3, 1) + Mid(k, 2, 1) + Mid(k, 4, 1) + Mid(k, 1, 1)
    If MRD() = True Then
        X1d = Chr(WorksheetFunction.Hex2Dec(n))
    Else
        X1d = Left(Application.Username, 3)
    End If

End Function

Function X2d(ByVal CharStr As String) As String

    Dim k As String: Dim n As String
    k = Right(CharStr, 4)
    n = Mid(k, 3, 1) + Mid(k, 4, 1) + Mid(k, 1, 1) + Mid(k, 2, 1)
    If MRD() = True Then
        X2d = Chr(WorksheetFunction.Hex2Dec(n))
    Else
        X2d = Left(Application.Username, 1)
    End If

End Function

Function X3d(ByVal CharStr As String) As String

    Dim k As String: Dim n As String: Dim n1 As String, n2 As String, n1d As String, n2d As String
    k = Right(CharStr, 4)
    n = Mid(k, 1, 1) + Mid(k, 3, 1) + Mid(k, 2, 1) + Mid(k, 4, 1)
    n1 = Left(n, 2): n2 = Right(n, 2)
    n1d = Chr(WorksheetFunction.Hex2Dec(n1))
    n2d = Chr(WorksheetFunction.Hex2Dec(n2))

    If MRD() = True Then
        X3d = Chr(WorksheetFunction.Hex2Dec(n1d + n2d))
    Else
        X3d = Left(Application.Username, 2)
    End If

End Function

Function X4d(ByVal CharStr As String) As String

    Dim k As String: Dim n As String: Dim n1 As String, n2 As String, n1d As String, n2d As String
    k = Right(CharStr, 4)
    n = Mid(k, 3, 1) + Mid(k, 4, 1) + Mid(k, 1, 1) + Mid(k, 2, 1)
    n1 = Left(n, 2): n2 = Right(n, 2)
    n1d = Chr(WorksheetFunction.Hex2Dec(n1))
    n2d = Chr(WorksheetFunction.Hex2Dec(n2))

    If MRD() = True Then
        X4d = Chr(WorksheetFunction.Hex2Dec(n1d + n2d))
    Else
        X4d = Left(Application.Username, 8)
    End If

End Function

Function KeyVal(ByVal str As String) As Long

    For i = 1 To Len(str)
        If Asc(Mid(str, 1)) < 0 Then
            k = k - Asc(Mid(str, 1))
        Else
            k = k + Asc(Mid(str, 1))
        End If
    Next

    KeyVal = k

End Function

Function MRD() As Boolean

    MRD = true

End Function

Sub RealEnc(ByVal str As String)
End Sub

Function RollOver(ByVal k As String, ByVal Ts As String, ByVal T As Boolean) As String

    Dim tbl1 As String
    tbl1 = UAl + UAl + LAl + LAl

    Dim tbl2 As String
    tbl2 = RevStr(UAl) + RevStr(UAl) + RevStr(LAl) + RevStr(LAl)

    Dim Lv As Long, Kv As Long, Kt As Long
    Lv = AZtoNum(Ts)

    If T = True Then
        Kv = InStr(1, tbl1, k, vbTextCompare)
        Kt = Mid(tbl1, Kv, 10)
    Else
        Kv = InStr(1, tbl2, k, vbTextCompare)
        Kt = Mid(tbl2, Kv, 10)
    End If
    
    RollOver = Mid(Kt, Lv + 1, 1)

End Function

Function RollOver2(ByVal k As String, ByVal Ts As String, ByVal T As Boolean, _
                   ByVal rv As Long) As String
    Dim tbl1 As String
    tbl1 = UAl + UAl + UAl + LAl + LAl + LAl
    tbl1 = RevolveStr(tbl1, rv)

    Dim tbl2 As String
    tbl2 = RevStr(UAl) + RevStr(UAl) + RevStr(UAl) + RevStr(LAl) + RevStr(LAl) + RevStr(LAl)
    tbl2 = RevolveStr(tbl2, rv)

    Dim Lv As Long, Kv As Long, Kt As String
    Lv = AZtoNum(Ts)

    If T = True Then
        Kv = InStr(1, tbl1, k, vbTextCompare)
        Kt = Mid(tbl1, Kv, 10)
    Else
        Kv = InStr(1, tbl2, k, vbTextCompare)
        Kt = Mid(tbl2, Kv, 10)
    End If
    
    RollOver2 = Mid(Kt, Lv + 1, 1)

End Function

Function RODoc(ByVal D1 As String, ByVal D2 As String, ByVal T As Boolean) As String
    
    Dim i As Long, tmp As String
    For i = 1 To Len(D1)
        tmp = tmp + RollOver(Mid(D1, i, 1), Mid(D2, i, 1), T)
    Next
    RODoc = tmp

End Function

Function RODoc2(ByVal D1 As String, ByVal D2 As String, ByVal T As Boolean, _
               ByVal rv As Long) As String

    Dim i As Long, tmp As String
    For i = 1 To Len(D1)
        tmp = tmp + RollOver2(Mid(D1, i, 1), Mid(D2, i, 1), T, rv)
    Next
    
    RODoc2 = tmp

End Function

Function AZtoNum(ByVal Alpha As String) As Long
    Dim tl As Long
    If InStr(1, UAl, Alpha) > 0 Then
        tl = InStr(1, UAl, Alpha)
    ElseIf InStr(1, LAl, Alpha) > 0 Then
        tl = InStr(1, LAl, Alpha)
    Else
     tl = 0
    End If

    AZtoNum = Mid(NAl, tl, 1)
End Function

Function RevStr(ByVal strs As String) As String
    Dim tk As String
        For i = 1 To Len(strs)
        tk = Mid(strs, i, 1) + tk
        Next
        RevStr = tk
End Function

Function Encode(ByVal RawText As String, ByVal pk As String) As String

 
    If RawText = vbNullString Then
        Encode = vbNullString
        Exit Function
    Else
    End If

    Dim PKlongstr As String
    Dim PKBlock As String, PKmulti As Long

    Dim tmpenc As String
    tmpenc = Encry(RawText)

    Dim tmprv As Long
    tmprv = CLng(Right(CStr(Len(tmpenc) + LngtoRevolve(pk)), 2))
    pk = RevolveStr(pk, tmprv)
    PKBlock = pk + RevStr(pk)
    PKmulti = CLng(Len(tmpenc) / Len(PKBlock) + 1)
    PKlongstr = Application.WorksheetFunction.Rept(PKBlock, PKmulti)


    Dim tmpenc2 As String
    tmpenc2 = RODoc2(tmpenc, PKlongstr, True, tmprv)

    Encode = tmpenc2

End Function

Function Decode(ByVal EncText As String, ByVal pk As String) As String

    Dim tmprv As Long
    tmprv = CLng(Right(CStr(Len(EncText) + LngtoRevolve(pk)), 2))
    pk = RevolveStr(pk, tmprv)

    If EncText = vbNullString Then
        Decode = vbNullString
        Exit Function
    Else
    End If

    Dim PKlongstr As String
    Dim PKBlock As String, PKmulti As Long

    PKBlock = pk + RevStr(pk)
    PKmulti = CLng(Len(EncText) / Len(PKBlock) + 1)
    PKlongstr = Application.WorksheetFunction.Rept(PKBlock, PKmulti)

    Dim tmpenc As String
    tmpenc = RODoc2(EncText, PKlongstr, False, tmprv)

    Dim tmpenc2 As String
    tmpenc2 = DeCry(tmpenc)

    Decode = tmpenc2

End Function

Function MasterSet() As String

    MasterSet = "UAL=" + UAl + "||" + "LAL=" + LAl + "||" + "NAL=" + NAl

End Function

Function LngtoRevolve(ByVal k As String) As Long

    Dim Klv As Long
    Dim i As Long

    For i = 1 To Len(k)
        Klv = AZtoNum(Mid(k, i, 1)) + Klv
    Next

    If Application.WorksheetFunction.IsEven(Klv) Then
        LngtoRevolve = CLng(RevStr(CStr(Klv)))
    ElseIf Application.WorksheetFunction.IsEven(Len(CStr(Klv))) = True Then
        LngtoRevolve = CLng(RevStr(CStr(Klv))) * 3
    Else
        LngtoRevolve = CLng(RevStr(CStr(Klv))) * 5
    End If

    LngtoRevolve = Klv * 2

End Function

Function RevolveStr(ByVal k As String, ByVal rv As Long) As String

    Dim tmp As String
    Dim lt As Long
    tmp = k
    lt = Len(tmp) + 1

    For i = 1 To rv
    
        tmp = Right(tmp, lt) + Left(tmp, 1)
    
    Next

    RevolveStr = tmp

End Function

Function EncZip(ByVal encstr As String) As String

    Dim k() As String
    k = Split(encstr, "X")

    Dim sdc As Object, cnt As Long
    Set sdc = CreateObject("Scripting.Dictionary")
    cnt = 1
    Dim i As Long

    '키값 마이닝
    For i = 1 To UBound(k)
        If sdc.Exists(k(i)) Then
        Else
            sdc.Add k(i), cnt
            cnt = cnt + 1
        End If
    Next

    Dim K2() As Variant: K2() = sdc.Items
    Dim K3() As Variant: K3() = sdc.Keys
    
    Dim K4 As String
    
    '키값으로 압축
    For i = 1 To UBound(k)
        K4 = K4 + CStr(sdc.item(k(i))) + "|"
    Next

    'Value - Index 구분자 "M"삽입
    K4 = Left(K4, Len(K4) - 1) + "M"

    'Index Keyword Save
    Dim T As String
    For i = 0 To UBound(K3)
        T = T + K3(i) + "X"
    Next

    'Reverse Table
    T = RevStr(T)
    K4 = RevStr(K4)

    '결합
    K4 = T + K4

    K4 = ReplaceFB(K4, 1) '압축

    EncZip = K4
End Function

Function DecZip(ByVal encstr As String) As String

    On Error GoTo erroc

    '압축풀기
    encstr = ReplaceFB(encstr, -1)

    'M단 기점으로 앞뒤분리
    Dim K1 As String, K2 As String
    K1 = Left(encstr, InStr(1, encstr, "M") - 1)
    K2 = Mid(encstr, InStr(1, encstr, "M") + 1, Len(encstr))

    '분리된거 다시 Reverse
    K1 = RevStr(K1)
    K2 = RevStr(K2)

    'Split용으로 포맷팅
    K1 = Right(K1, 1) + Left(K1, Len(K1) - 1)
    K2 = "|" + K2

    Dim T1 As Variant, T2 As Variant
    T1 = Split(K1, "X")
    T2 = Split(K2, "|")

    '다시 재조립
    Dim T3 As String, i As Long
    For i = 1 To UBound(T2)
        T3 = T3 + "X" + T1(CLng(T2(i)))
    Next

    DecZip = T3

Exit Function

erroc:
    DecZip = "Err"

End Function

Function ReplaceFB(ByVal str As String, ByVal opt As Long) As String
 '압축
 'opt=1 forward(enc), -1 Backward(dec)
    On Error GoTo erroc
    Const maxarr As Long = 100

    Dim data(1 To 100, 1 To 2) As String
    Dim datanum As Long
    datanum = 1

    Dim i As Long, i2 As Long, K1 As Long, K2 As Long

    Call InDatas(data, datanum, "11", "G")
    Call InDatas(data, datanum, "22", "H")
    Call InDatas(data, datanum, "33", "I")
    Call InDatas(data, datanum, "44", "J")
    Call InDatas(data, datanum, "55", "K")

    Call InDatas(data, datanum, "66", "L")
    Call InDatas(data, datanum, "88", "N")
    Call InDatas(data, datanum, "99", "O")
    Call InDatas(data, datanum, "12", "P")
    Call InDatas(data, datanum, "23", "Q")

    Call InDatas(data, datanum, "34", "R")
    Call InDatas(data, datanum, "45", "S")
    Call InDatas(data, datanum, "56", "T")
    Call InDatas(data, datanum, "67", "U")
    Call InDatas(data, datanum, "78", "V")
    
    Call InDatas(data, datanum, "89", "W")
    Call InDatas(data, datanum, "13", "g")
    Call InDatas(data, datanum, "24", "h")
    Call InDatas(data, datanum, "35", "i")
    Call InDatas(data, datanum, "46", "j")

    Call InDatas(data, datanum, "57", "k")
    Call InDatas(data, datanum, "68", "l")
    Call InDatas(data, datanum, "79", "m")
    Call InDatas(data, datanum, "80", "n")
    Call InDatas(data, datanum, "91", "o")

    Call InDatas(data, datanum, "|10", "p")
    Call InDatas(data, datanum, "|11", "q")
    Call InDatas(data, datanum, "|1", "r")
    Call InDatas(data, datanum, "|2", "s")
    Call InDatas(data, datanum, "|3", "t")
    
    Call InDatas(data, datanum, "|4", "u")
    Call InDatas(data, datanum, "|5", "v")
    Call InDatas(data, datanum, "|6", "w")
    Call InDatas(data, datanum, "|7", "x")
    Call InDatas(data, datanum, "|8", "y")

    Call InDatas(data, datanum, "|9", "z")
    Call InDatas(data, datanum, "X1", "~")
    Call InDatas(data, datanum, "X2", "`")
    Call InDatas(data, datanum, "X3", "!")
    Call InDatas(data, datanum, "X4", "@")
    
    Call InDatas(data, datanum, "X5", "#")
    Call InDatas(data, datanum, "X6", "$")
    Call InDatas(data, datanum, "X7", "%")
    Call InDatas(data, datanum, "X8", "^")
    Call InDatas(data, datanum, "X9", "&")

    Call InDatas(data, datanum, "XJ", "*")
    Call InDatas(data, datanum, "XK", "(")
    Call InDatas(data, datanum, "XL", ")")
    Call InDatas(data, datanum, "XN", "-")
    Call InDatas(data, datanum, "XO", "_")

    Call InDatas(data, datanum, "XP", "+")
    Call InDatas(data, datanum, "XQ", "=")
    Call InDatas(data, datanum, "XS", "|")
    Call InDatas(data, datanum, "XT", "\")
    Call InDatas(data, datanum, "XU", "{")

    Call InDatas(data, datanum, "XV", "}")
    Call InDatas(data, datanum, "XW", "[")
    Call InDatas(data, datanum, "XZ", "]")

    If opt = 1 Then i = 1: i2 = datanum - 1: K1 = 1: K2 = 2
    If opt = -1 Then i = datanum - 1: i2 = 1: K1 = 2: K2 = 1

    For j = i To i2 Step opt
        str = Replace(str, data(j, K1), data(j, K2))
    Next

    ReplaceFB = str
    Exit Function

erroc:
    ReplaceFB = "Err"

End Function

Sub InDatas(ByRef arr As Variant, ByRef arrnum As Long, befstr As String, aftstr As String)

    arr(arrnum, 1) = befstr
    arr(arrnum, 2) = aftstr
    arrnum = arrnum + 1

End Sub

Function PublicKey(ByVal pk As String)


 PublicKey = EncZip(Encry(Encode(pk, "1")))
End Function

Function Encode_By_PBKEY(ByVal str As String, ByVal pubk As String) As String
    '공개키를 가지고 암호화처리(복호화는 공개키로 불가하며, 공개키에 대응하는 개인키로만 복호화 가능)
    Encode_By_PBKEY = Encode(str, Decode(DeCry(DecZip(pubk)), "1"))
End Function
