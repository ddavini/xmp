VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnigma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Enigma Cipher
' By: David Midkiff (mznull@earthlink.net)
'
' Here is the classic enigma cipher that the Germans used in WW2.
' A team of Polish mathematicians broke it and this cipher has gone
' down in the history books as a great classic of cryptography and
' cryptanlysis.
'
' Good accounts of the cipher can be found at:
' http://www.codesandciphers.org.uk/virtualbp/poles/poles.htm

Private I As Integer, J As Integer, K As Integer
Private WheelIndex1 As Integer, WheelIndex2 As Integer, WheelIndex3 As Integer
Private WheelIndex4 As Integer, WheelIndex5 As Integer, iTotalChar As Long
Private lProcessCount As Long, Wheel2String As String, Wheel3String As String
Private Wheel4String As String, Wheel5String As String, Wheel6String As String
Private Wheel7String As String, Wheel8String As String, Wheel9String As String
Private Wheel10String As String, strCurrentChar As String
    
Private Wheel1() As Variant
Private Wheel2() As Variant
Private Wheel3() As Variant
Private Wheel4() As Variant
Private Wheel5() As Variant
Private Wheel6() As Variant
Private Wheel7() As Variant
Private Wheel8() As Variant
Private Wheel9() As Variant
Private Wheel10() As Variant
Private MyValue As Integer
Private strW1 As String
Private strW2 As String
Private strW3 As String
Private strW4 As String
Private strW5 As String
Private strW6 As String
Private strW7 As String
Private strW8 As String
Private strW9 As String
Private strW10 As String
Private strMessage As String
Private strOutput As String
Private strDecryptOutput As String

Public Function EnigmaDecrypt(ByVal EncryptedText As String) As String
    Dim D As Integer
    strW1 = ">v=6h</*}B|TaJ?0{9 s._3(~%#5zI1Cy+&)ZRx-M:YVjcQbmU[wor,2SOe!K87gAf`]N4\pHqEGPldWuDk@n';L^i$FtX"
    strW2 = "ETHt?$8~FL:pg>YmB<0!O6GJCc\iQ2 j#rV3kRKN.-l}h%e&;`Pa4@^]o5A_79y{x*)WbS+dqsfz,wUZn[1M|DvI='X(u/"
    strW3 = "9?fazRNjLS]5rM 3)}d_|,gvF><D'eE=o^420pYOJ7C;[k:&hbmtP\~8cyZq(%WTHK6@+$uB!#sA.-UXQwlnV`Ix/1*iG{"
    strW4 = "c({W=yJ^[vuXlaLGpNn$r>6,TZ#3mRe%+UzESA.h5s\q_!1}o@:4fM|0-'DKkHOwCY2)IP7d/xF`b;8?~&t Bi]9<Qj*Vg"
    strW5 = "[^#!mF.iWK9J0a|LfPpuB)AR'@kX_;/r(t}d*gq3CMVe<O,n7+h=Tv1?j~w\>bIGZ& 6s28S]-zc{Hl5U:4QN%$Do`EYxy"
    strW6 = "<(~c0POH%8t{[Fb>ow,BN_GM&^Isq:' 4prf+Y$S3viZW;!?dlK/mLC6ygXA=)5Q2-]J}TUVx*D7\Ra1#jE9@zhn`|uke."
    strW7 = "R8?u.*|AFbf&-0~ZOYy`S>J)C =%Vc,(:]TEKMBIoaxvG/Uz$kX+sgH3j<i{[4p2L9#'qd7PW}t6h_l@;w!n^er5mQND\1"
    strW8 = "7l]6zJHgjo*2[u%UW(Rp1Q?5S=`c$)x3's0-I<bk~^h!aq+LN_{/dy,Fe@M:tG r;Ow}\E8P&.BC#nfY|9XDv4Z>VimTKA"
    strW9 = "b`M}Ik)sA{=#!0f(/PU\RlxWoY3imOt*,w&vpD<EX:^e25d]h9T;L8J4FZru$1S>KQn' _|c%a+[Hz67q~CBNGy@g.j?V-"
    strW10 = " {A}G0S59|eXgD7:x'/c1M#=?!$tNY]InzUhC_wRqO2PF)H<.(T+`,v-34Lp%jZk8[l\sE^ayB6mr>bdW*K@Q;oJVfu~&i"
    lProcessCount = 0
    iTotalChar = Len(strW1)
    
    ReDim Wheel1(1 To iTotalChar) As Variant
    ReDim Wheel2(1 To iTotalChar) As Variant
    ReDim Wheel3(1 To iTotalChar) As Variant
    ReDim Wheel4(1 To iTotalChar) As Variant
    ReDim Wheel5(1 To iTotalChar) As Variant
    ReDim Wheel6(1 To iTotalChar) As Variant
    ReDim Wheel7(1 To iTotalChar) As Variant
    ReDim Wheel8(1 To iTotalChar) As Variant
    ReDim Wheel9(1 To iTotalChar) As Variant
    ReDim Wheel10(1 To iTotalChar) As Variant
    
    For I = 1 To iTotalChar Step 1
        Wheel1(I) = Mid$(strW1, I, 1)
        Wheel2(I) = Mid$(strW2, I, 1)
        Wheel3(I) = Mid$(strW3, I, 1)
        Wheel4(I) = Mid$(strW4, I, 1)
        Wheel5(I) = Mid$(strW5, I, 1)
        Wheel6(I) = Mid$(strW6, I, 1)
        Wheel7(I) = Mid$(strW7, I, 1)
        Wheel8(I) = Mid$(strW8, I, 1)
        Wheel9(I) = Mid$(strW9, I, 1)
        Wheel10(I) = Mid$(strW10, I, 1)
    Next
    strOutput = EncryptedText
    If Len(strOutput) <= 94 Then
        For D = 1 To Len(strOutput) Step 1
            EncryptRotateWheels
        Next
    Else
        For D = 1 To Len(strOutput) Mod iTotalChar Step 1
            EncryptRotateWheels
        Next
    End If
    lProcessCount = 0
    For J = Len(strOutput) To 1 Step -1
        DecryptRotateWheels
        strCurrentChar = Mid$(strOutput, J, 1)
        If strCurrentChar <> Chr$(13) And strCurrentChar <> Chr$(10) And strCurrentChar <> Chr$(34) Then
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel10(K) Then WheelIndex5 = K
            Next
            strCurrentChar = Wheel9(WheelIndex5)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel8(K) Then WheelIndex4 = K
            Next
            strCurrentChar = Wheel7(WheelIndex4)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel6(K) Then WheelIndex3 = K
            Next
            strCurrentChar = Wheel5(WheelIndex3)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel4(K) Then WheelIndex2 = K
            Next
            strCurrentChar = Wheel3(WheelIndex2)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel2(K) Then WheelIndex1 = K
            Next
            strCurrentChar = Wheel1(WheelIndex1)
            strDecryptOutput = strCurrentChar & strDecryptOutput
        Else
            If strCurrentChar = Chr$(34) Then strDecryptOutput = strCurrentChar & strDecryptOutput
            If strCurrentChar = Chr$(13) Then strDecryptOutput = strCurrentChar & strDecryptOutput
            If strCurrentChar = Chr$(10) Then strDecryptOutput = strCurrentChar & strDecryptOutput
        End If
        lProcessCount = lProcessCount + 1
        DoEvents
    Next
    EnigmaDecrypt = strDecryptOutput
    strMessage = ""
    strOutput = ""
    strDecryptOutput = ""
End Function


Private Function DecryptRotateWheels()
    Dim K As Integer, strTempHold As String
    strTempHold = Wheel10(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel10(K) = Wheel10(K + 1)
    Next
    Wheel10(iTotalChar) = strTempHold
    strTempHold = Wheel9(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel9(K) = Wheel9(K - 1)
    Next
    Wheel9(1) = strTempHold
    strTempHold = Wheel8(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel8(K) = Wheel8(K + 1)
    Next
    Wheel8(iTotalChar) = strTempHold
    strTempHold = Wheel7(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel7(K) = Wheel7(K - 1)
    Next
    Wheel7(1) = strTempHold
    strTempHold = Wheel6(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel6(K) = Wheel6(K + 1)
    Next
    Wheel6(iTotalChar) = strTempHold
    strTempHold = Wheel5(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel5(K) = Wheel5(K - 1)
    Next
    Wheel5(1) = strTempHold
    strTempHold = Wheel4(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel4(K) = Wheel4(K + 1)
    Next
    Wheel4(iTotalChar) = strTempHold
    strTempHold = Wheel3(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel3(K) = Wheel3(K - 1)
    Next
    Wheel3(1) = strTempHold
    strTempHold = Wheel2(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel2(K) = Wheel2(K + 1)
    Next
    Wheel2(iTotalChar) = strTempHold
    strTempHold = Wheel1(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel1(K) = Wheel1(K - 1)
    Next
    Wheel1(1) = strTempHold
End Function

Public Function EnigmaEncrypt(ByVal Text As String) As String
    lProcessCount = 0
    strW1 = ">v=6h</*}B|TaJ?0{9 s._3(~%#5zI1Cy+&)ZRx-M:YVjcQbmU[wor,2SOe!K87gAf`]N4\pHqEGPldWuDk@n';L^i$FtX"
    strW2 = "ETHt?$8~FL:pg>YmB<0!O6GJCc\iQ2 j#rV3kRKN.-l}h%e&;`Pa4@^]o5A_79y{x*)WbS+dqsfz,wUZn[1M|DvI='X(u/"
    strW3 = "9?fazRNjLS]5rM 3)}d_|,gvF><D'eE=o^420pYOJ7C;[k:&hbmtP\~8cyZq(%WTHK6@+$uB!#sA.-UXQwlnV`Ix/1*iG{"
    strW4 = "c({W=yJ^[vuXlaLGpNn$r>6,TZ#3mRe%+UzESA.h5s\q_!1}o@:4fM|0-'DKkHOwCY2)IP7d/xF`b;8?~&t Bi]9<Qj*Vg"
    strW5 = "[^#!mF.iWK9J0a|LfPpuB)AR'@kX_;/r(t}d*gq3CMVe<O,n7+h=Tv1?j~w\>bIGZ& 6s28S]-zc{Hl5U:4QN%$Do`EYxy"
    strW6 = "<(~c0POH%8t{[Fb>ow,BN_GM&^Isq:' 4prf+Y$S3viZW;!?dlK/mLC6ygXA=)5Q2-]J}TUVx*D7\Ra1#jE9@zhn`|uke."
    strW7 = "R8?u.*|AFbf&-0~ZOYy`S>J)C =%Vc,(:]TEKMBIoaxvG/Uz$kX+sgH3j<i{[4p2L9#'qd7PW}t6h_l@;w!n^er5mQND\1"
    strW8 = "7l]6zJHgjo*2[u%UW(Rp1Q?5S=`c$)x3's0-I<bk~^h!aq+LN_{/dy,Fe@M:tG r;Ow}\E8P&.BC#nfY|9XDv4Z>VimTKA"
    strW9 = "b`M}Ik)sA{=#!0f(/PU\RlxWoY3imOt*,w&vpD<EX:^e25d]h9T;L8J4FZru$1S>KQn' _|c%a+[Hz67q~CBNGy@g.j?V-"
    strW10 = " {A}G0S59|eXgD7:x'/c1M#=?!$tNY]InzUhC_wRqO2PF)H<.(T+`,v-34Lp%jZk8[l\sE^ayB6mr>bdW*K@Q;oJVfu~&i"
    iTotalChar = Len(strW1)

    ReDim Wheel1(1 To iTotalChar) As Variant
    ReDim Wheel2(1 To iTotalChar) As Variant
    ReDim Wheel3(1 To iTotalChar) As Variant
    ReDim Wheel4(1 To iTotalChar) As Variant
    ReDim Wheel5(1 To iTotalChar) As Variant
    ReDim Wheel6(1 To iTotalChar) As Variant
    ReDim Wheel7(1 To iTotalChar) As Variant
    ReDim Wheel8(1 To iTotalChar) As Variant
    ReDim Wheel9(1 To iTotalChar) As Variant
    ReDim Wheel10(1 To iTotalChar) As Variant

    For I = 1 To iTotalChar Step 1
        Wheel1(I) = Mid$(strW1, I, 1)
        Wheel2(I) = Mid$(strW2, I, 1)
        Wheel3(I) = Mid$(strW3, I, 1)
        Wheel4(I) = Mid$(strW4, I, 1)
        Wheel5(I) = Mid$(strW5, I, 1)
        Wheel6(I) = Mid$(strW6, I, 1)
        Wheel7(I) = Mid$(strW7, I, 1)
        Wheel8(I) = Mid$(strW8, I, 1)
        Wheel9(I) = Mid$(strW9, I, 1)
        Wheel10(I) = Mid$(strW10, I, 1)
    Next
    strMessage = Text

    For J = 1 To Len(strMessage) Step 1
        strCurrentChar = Mid$(strMessage, J, 1)
        If strCurrentChar <> Chr$(13) And strCurrentChar <> Chr$(10) And strCurrentChar <> Chr$(34) Then
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel1(K) Then WheelIndex1 = K
            Next
            strCurrentChar = Wheel2(WheelIndex1)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel3(K) Then WheelIndex2 = K
            Next
            strCurrentChar = Wheel4(WheelIndex2)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel5(K) Then WheelIndex3 = K
            Next
            strCurrentChar = Wheel6(WheelIndex3)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel7(K) Then WheelIndex4 = K
            Next
            strCurrentChar = Wheel8(WheelIndex4)
            For K = 1 To iTotalChar Step 1
                If strCurrentChar = Wheel9(K) Then WheelIndex5 = K
            Next
            strCurrentChar = Wheel10(WheelIndex5)
            strOutput = strOutput & strCurrentChar
        Else
            If strCurrentChar = Chr$(34) Then strOutput = strOutput & strCurrentChar
            If strCurrentChar = Chr$(13) Then strOutput = strOutput & strCurrentChar
            If strCurrentChar = Chr$(10) Then strOutput = strOutput & strCurrentChar
        End If
        EncryptRotateWheels
        lProcessCount = lProcessCount + 1
        DoEvents
    Next
    EnigmaEncrypt = strOutput
    strMessage = ""
    strOutput = ""
End Function


Private Function EncryptRotateWheels()
    Dim K As Integer, strTempHold As String
    strTempHold = Wheel1(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel1(K) = Wheel1(K + 1)
    Next
    Wheel1(iTotalChar) = strTempHold
    strTempHold = Wheel2(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel2(K) = Wheel2(K - 1)
    Next
    Wheel2(1) = strTempHold
    strTempHold = Wheel3(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel3(K) = Wheel3(K + 1)
    Next
    Wheel3(iTotalChar) = strTempHold
    strTempHold = Wheel4(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel4(K) = Wheel4(K - 1)
    Next
    Wheel4(1) = strTempHold
    strTempHold = Wheel5(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel5(K) = Wheel5(K + 1)
    Next
    Wheel5(iTotalChar) = strTempHold
    strTempHold = Wheel6(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel6(K) = Wheel6(K - 1)
    Next
    Wheel6(1) = strTempHold
    strTempHold = Wheel7(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel7(K) = Wheel7(K + 1)
    Next
    Wheel7(iTotalChar) = strTempHold
    strTempHold = Wheel8(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel8(K) = Wheel8(K - 1)
    Next
    Wheel8(1) = strTempHold
    strTempHold = Wheel9(1)
    For K = 1 To (iTotalChar - 1) Step 1
        Wheel9(K) = Wheel9(K + 1)
    Next
    Wheel9(iTotalChar) = strTempHold
    strTempHold = Wheel10(iTotalChar)
    For K = iTotalChar To 2 Step -1
        Wheel10(K) = Wheel10(K - 1)
    Next
    Wheel10(1) = strTempHold
End Function


Private Sub GenerateRandomWheels()
    Dim I As Integer, J As Integer, bDupe As Boolean
    Randomize

    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel2(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel2(J) = "" Then
                        bDupe = False
                        Wheel2(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel3(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel3(J) = "" Then
                        bDupe = False
                        Wheel3(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel4(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel4(J) = "" Then
                        bDupe = False
                        Wheel4(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel5(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel5(J) = "" Then
                        bDupe = False
                        Wheel5(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel6(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel6(J) = "" Then
                        bDupe = False
                        Wheel6(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel7(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel7(J) = "" Then
                        bDupe = False
                        Wheel7(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel8(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel8(J) = "" Then
                        bDupe = False
                        Wheel8(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel9(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel9(J) = "" Then
                        bDupe = False
                        Wheel9(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)
            For J = 1 To I Step 1
                If Wheel10(J) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel10(J) = "" Then
                        bDupe = False
                        Wheel10(J) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For I = 1 To iTotalChar Step 1
        Wheel2String = Wheel2String & CStr(Wheel1(Wheel2(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel3String = Wheel3String & CStr(Wheel1(Wheel3(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel4String = Wheel4String & CStr(Wheel1(Wheel4(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel5String = Wheel5String & CStr(Wheel1(Wheel5(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel6String = Wheel6String & CStr(Wheel1(Wheel6(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel7String = Wheel7String & CStr(Wheel1(Wheel7(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel8String = Wheel8String & CStr(Wheel1(Wheel8(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel9String = Wheel9String & CStr(Wheel1(Wheel9(I)))
    Next
    For I = 1 To iTotalChar Step 1
        Wheel10String = Wheel10String & CStr(Wheel1(Wheel10(I)))
    Next
End Sub
