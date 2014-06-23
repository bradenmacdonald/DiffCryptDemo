' Differential Cryptanalysis Example
' Copyright (c) 2011 Braden MacDonald
' zlib license


Public Function BitXOR(x As Long, y As Long) As Long
    BitXOR = x Xor y
End Function

Public Function DoSubst(x As Integer, sbox As Range) As Integer
    DoSubst = CInt("&H" & Application.HLookup(Hex(x), sbox, 2))
End Function

Public Function CountXORoutputsOf(x As Integer, y As Integer, sbox As Range)
    
    Dim count_equal As Integer
    count_equal = 0
    
    Dim do_sbox(0 To 15) As Long
    Dim i As Integer
    For i = 0 To 15
        do_sbox(i) = CLng("&H" & sbox.Item(2, i + 1).Value)
    Next i
    
    Dim out1 As Integer
    Dim out2 As Integer
    
    For i = 0 To 15
        out1 = do_sbox(i)
        out2 = do_sbox(x Xor i)
        If (out1 Xor out2) = y Then
            count_equal = count_equal + 1
        End If
    Next i
    
    CountXORoutputsOf = count_equal
End Function

Public Function Dec2BinHex(x As Integer)
    Dec2BinHex = Application.Dec2Bin(x, 4) & " (" & Hex(x) & ")"
End Function

Public Function BinHex2Dec(x As String)
    BinHex2Dec = Application.Bin2Dec(Mid(x, 1, 4))
End Function



Public Function Encrypt(plaintext As Long, sbox As Range, perm As Range, key1 As Long, key2 As Long, key3 As Long, key4 As Long, key5 As Long)
    Dim w As Long
    Dim u As Long
    Dim v As String
    Dim v4 As Long
        
    Dim do_sbox(0 To 15) As Long
    For i = 0 To 15
        do_sbox(i) = CLng("&H" & sbox.Item(2, i + 1).Value)
    Next i
    
    
    
    ' Round 1:
    u = plaintext Xor key1
    ' Apply S-Boxes to each 4-bit block, then store the result as a 16-digit binary string:
    v = Application.Dec2Bin(do_sbox((u And 61440) / 4096) * 16 Or do_sbox((u And 3840) / 256), 8) & Application.Dec2Bin(do_sbox((u And 240) / 16) * 16 Or do_sbox(u And 15), 8)
    w = 0
    For i = 1 To 16
        w = w + (2 ^ (15 - perm.Item(2, i).Value)) * Int(Mid(v, i, 1))
    Next i
    
    'h = Right("0000" & Hex(w), 4)
    'MsgBox "After Round 1 it is " & h & "(" & Application.Hex2Bin(Left(h, 2), 8) & " " & Application.Hex2Bin(Right(h, 2), 8) & ")"
    
    ' Round 2:
    u = w Xor key2
    v = Application.Dec2Bin(do_sbox((u And 61440) / 4096) * 16 Or do_sbox((u And 3840) / 256), 8) & Application.Dec2Bin(do_sbox((u And 240) / 16) * 16 Or do_sbox(u And 15), 8)
    w = 0
    For i = 1 To 16
        w = w + (2 ^ (15 - perm.Item(2, i).Value)) * Int(Mid(v, i, 1))
    Next i
    
    'h = Right("0000" & Hex(w), 4)
    'MsgBox "After Round 2 it is " & h & "(" & Application.Hex2Bin(Left(h, 2), 8) & " " & Application.Hex2Bin(Right(h, 2), 8) & ")"
    
    ' Round 3:
    u = w Xor key3
    v = Application.Dec2Bin(do_sbox((u And 61440) / 4096) * 16 Or do_sbox((u And 3840) / 256), 8) & Application.Dec2Bin(do_sbox((u And 240) / 16) * 16 Or do_sbox(u And 15), 8)
    w = 0
    For i = 1 To 16
        w = w + (2 ^ (15 - perm.Item(2, i).Value)) * Int(Mid(v, i, 1))
    Next i
    
    'h = Right("0000" & Hex(w), 4)
    'MsgBox "After Round 3 it is " & h & "(" & Application.Hex2Bin(Left(h, 2), 8) & " " & Application.Hex2Bin(Right(h, 2), 8) & ")"
    
    ' Final Round
    u = w Xor key4
    v4 = do_sbox((u And 61440) / 4096) * 4096 Or do_sbox((u And 3840) / 256) * 256 Or do_sbox((u And 240) / 16) * 16 Or do_sbox(u And 15)
    w = v4 Xor key5

    'h = Right("0000" & Hex(w), 4)
    'MsgBox "Finally, y is " & h & "(" & Application.Hex2Bin(Left(h, 2), 8) & " " & Application.Hex2Bin(Right(h, 2), 8) & ")"

    Encrypt = w
    
End Function


Private Sub SplitToFourBytes(num As Long, ByRef output() As Integer)
    h = Right("0000" & Hex(num), 4)
    output(0) = CInt("&H" & Mid(h, 1, 1))
    output(1) = CInt("&H" & Mid(h, 2, 1))
    output(2) = CInt("&H" & Mid(h, 3, 1))
    output(3) = CInt("&H" & Mid(h, 4, 1))
End Sub


Public Function DiffCryptGuessKey(sbox As Range, x1x2y1y2_data As Range, expected_xor As Long)
    
  
    Dim data As Variant
    data = x1x2y1y2_data.Value ' The range given should be n rows x 4 columns
    
    Dim RowCount As Integer
    RowCount = x1x2y1y2_data.Rows.Count
    
    Dim i As Integer
    
    ' The number of times each key has matched:
    Dim Count(0 To 2 ^ 16 - 1) As Long
    
    xor_check = data(1, 1) Xor data(1, 2)
    
    ' Compute the inverse of the S-Box
    Dim inverted_sbox(0 To 15) As Long
    For i = 0 To 15
        For j = 1 To 15
            If Application.Hex2Dec(sbox.Item(2, j)) = i Then
                inverted_sbox(i) = j - 1
                Exit For
            End If
        Next j
    Next i
    
    Dim expected_u4_parts(0 To 3)
    expected_u4_parts(0) = (expected_xor And 61440) / 4096
    expected_u4_parts(1) = (expected_xor And 3840) / 256
    expected_u4_parts(2) = (expected_xor And 240) / 16
    expected_u4_parts(3) = (expected_xor And 15)

    Dim RightPair As Boolean
    Dim IgnoreMask As Long
    If expected_u4_parts(0) = 0 Then
        IgnoreMask = IgnoreMask Or 15 * 2 ^ 12
    End If
    If expected_u4_parts(1) = 0 Then
        IgnoreMask = IgnoreMask Or 15 * 2 ^ 8
    End If
    If expected_u4_parts(2) = 0 Then
        IgnoreMask = IgnoreMask Or 15 * 2 ^ 4
    End If
    If expected_u4_parts(3) = 0 Then
        IgnoreMask = IgnoreMask Or 15 * 2 ^ 0
    End If
    
    
    Dim y1 As Long, y2 As Long
    
    Dim KeysChecked As Long
    KeysChecked = 0
    Dim UsableRows As Long
    UsableRows = 0
    
    For r = 1 To RowCount
        
        If Not ((data(r, 1) Xor data(r, 2)) = xor_check) Then
            DiffCryptGuessKey = "Invalid input - found x1 XOR x2 that is different (Row " & RowCount + 1 & " should be " & xor_check & " but is " & (data(r, 1) Xor data(r, 2)) & ")"
            Exit Function
        End If
        
        y1 = data(r, 3)
        y2 = data(r, 4)
        
        If Not (y1 And IgnoreMask) = (y2 And IgnoreMask) Then
            ' This cannot be a right pair, because the parts of u4 in IgnoreMask that are generated back from y1 y2 will only XOR to zero if those parts are exactly the same
        Else
            For k = 0 To 2 ^ 16 - 1
                ' Don't bother checking keys that vary only in the bits set in the IgnoreMask:
                If (k And IgnoreMask) = 0 Then
                    v1 = y1 Xor k
                    v2 = y2 Xor k
        
                    ' Split v1 and v2 up into 4-bit blocks, and apply the inverted permutation to each 4-bit block:
                    u1 = (inverted_sbox((v1 And 61440) / 4096) * 4096) Or (inverted_sbox((v1 And 3840) / 256) * 256) Or (inverted_sbox((v1 And 240) / 16) * 16) Or inverted_sbox(v1 And 15)
                    u2 = (inverted_sbox((v2 And 61440) / 4096) * 4096) Or (inverted_sbox((v2 And 3840) / 256) * 256) Or (inverted_sbox((v2 And 240) / 16) * 16) Or inverted_sbox(v2 And 15)
                    
                    If (u1 Xor u2) = expected_xor Then
                        Count(k) = Count(k) + 1
                    End If
                    KeysChecked = KeysChecked + 1
                End If
            Next k
            UsableRows = UsableRows + 1
        End If
        
    Next r

    Dim guessedKey As Long
    guessedKey = 0
    For k = 1 To 2 ^ 16 - 1
        If Count(k) > Count(guessedKey) Then
            guessedKey = k
        End If
    Next k
    
    If Count(guessedKey) = 0 Then
        DiffCryptGuessKey = "No likely key was found. Checked " & KeysChecked & " keys in " & UsableRows & " of " & RowCount & " rows."
    Else
        Dim key_formatted As String
        key_formatted = Right("0000" & Hex(guessedKey), 4)
        For i = 0 To 3
            If IgnoreMask And 2 ^ (4 * (3 - i)) Then
                key_formatted = Left(key_formatted, i) & "?" & Right(key_formatted, 3 - i)
            End If
        Next i
        DiffCryptGuessKey = "K5 is likely " & key_formatted & " based on " & UsableRows & " rows of (x,x*,y,y*) (" & (RowCount - UsableRows) & " out of " & RowCount & " rows unused). It has a count of " & Count(guessedKey) & ", i.e. occurrence of " & Round(Count(guessedKey) / RowCount * 100, 2) & "%."
    End If
    
End Function
