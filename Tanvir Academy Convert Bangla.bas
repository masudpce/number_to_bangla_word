Function CurrencyToBangla(ByVal MyNumber)
Dim Temp
         Dim Rupees, Paise
         Dim DecimalPlace, Count
         ReDim Place(9) As String
         Place(2) = " nvRvi "
         Place(3) = " jvL "
         Place(4) = " †KvwU "
      '   Place(5) = " Hundred Core "
         ' Convert MyNumber to a string, trimming extra spaces.
         MyNumber = Trim(Str(MyNumber))
         ' Find decimal place.
         DecimalPlace = InStr(MyNumber, ".")
         ' If we find decimal place...
         If DecimalPlace > 0 Then
            ' Convert Paise
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paise = ConvertTens(Temp)
            ' Strip off Paise from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
         End If
         Count = 1
         
         Do While MyNumber <> ""
                 If Count = 1 Then
                
                   Temp = ConvertHundreds(Right(MyNumber, 3))
                     
                    If Temp <> "" Then Rupees = Temp & Place(Count) & Rupees
                    If Len(MyNumber) > 3 Then
                       ' Remove last 3 converted digits from MyNumber.
                       MyNumber = Left(MyNumber, Len(MyNumber) - 3)
                    Else
                       MyNumber = ""
                    End If
                    Count = Count + 1
                 Else
                 ' Convert last 3 digits of MyNumber to Bangla Rupees.
                 If Len(MyNumber) = 1 Then
                 Temp = ConvertDigit(MyNumber)
                 Else
                 Temp = ConvertTens(Right(MyNumber, 2))
                 End If
                    If Temp <> "" Then Rupees = Temp & Place(Count) & Rupees
                    If Len(MyNumber) >= 3 Then
                       ' Remove last 3 converted digits from MyNumber.
                       MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                       MyNumber = ""
                    End If
                    Count = Count + 1
                    End If
         Loop
         ' Clean up Rupees.
         Select Case Rupees
            Case ""
               Rupees = ""
            Case Else
               Rupees = Rupees & " UvKv"
         End Select
         ' Clean up Paise.
         Select Case Paise
            Case ""
               Paise = ""
            Case Else
               Paise = " Ges " & Paise & " cqmv"
         End Select
         CurrencyToBangla = Rupees & Paise
End Function
Private Function ConvertHundreds(ByVal MyNumber)
Dim Result As String
         ' Exit if there is nothing to convert.
         If Val(MyNumber) = 0 Then Exit Function
         ' Append leading zeros to number.
         MyNumber = Right("000" & MyNumber, 3)
         ' Do we have a hundreds place digit to convert?
         If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & "kZ "
         End If
         ' Do we have a tens place digit to convert?
         If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
         Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
         End If
         ConvertHundreds = Trim(Result)
End Function
Private Function ConvertTens(ByVal MyTens)
Dim Result As String
         ' Is value between 10 and 19?
         If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
               Case 1: Result = "GK"
               Case 10: Result = "`k"
               Case 11: Result = "GMvi"
               Case 12: Result = "evi"
               Case 13: Result = "†Zi"
               Case 14: Result = "†PŠÏ"
               Case 15: Result = "c‡bi"
               Case 16: Result = "†lvj"
               Case 17: Result = "m‡Zi"
               Case 18: Result = "AvVvi"
               Case 19: Result = "Ewbk"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 2 Then
            Select Case Val(MyTens)
               Case 2: Result = "`yB"
               Case 20: Result = "wek"
               Case 21: Result = "GKzk"
               Case 22: Result = "evBk"
               Case 23: Result = "†ZBk"
               Case 24: Result = "PweŸk"
               Case 25: Result = "cuwPk"
               Case 26: Result = "QvweŸk"
               Case 27: Result = "mvZvk"
               Case 28: Result = "AvVvk"
               Case 29: Result = "EbwÎk"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 3 Then
            Select Case Val(MyTens)
               Case 3: Result = "wZb"
               Case 30: Result = "wÎk"
               Case 31: Result = "GKwÎk"
               Case 32: Result = "ewÎk"
               Case 33: Result = "†ZwÎk"
               Case 34: Result = "†PŠwÎk"
               Case 35: Result = "cuqwÎk"
               Case 36: Result = "QwÎk"
               Case 37: Result = "mvuBwÎk"
               Case 38: Result = "AvUwÎk"
               Case 39: Result = "EbPwjøk"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 4 Then
            Select Case Val(MyTens)
               Case 4: Result = "Pvi"
               Case 40: Result = "Pwjøk"
               Case 41: Result = "GKPwjøk"
               Case 42: Result = "weqvwjøk"
               Case 43: Result = "†ZZvwjøk"
               Case 44: Result = "Pzqvwjøk"
               Case 45: Result = "cuqZvwjøk"
               Case 46: Result = "†QPwjøk"
               Case 47: Result = "mvZPwjøk"
               Case 48: Result = "AvUPwjøk"
               Case 49: Result = "EbcÂvk"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 5 Then
            Select Case Val(MyTens)
               Case 5: Result = "cvuP"
               Case 50: Result = "cÂvk"
               Case 51: Result = "GKvbœ"
               Case 52: Result = "evqvbœ"
               Case 53: Result = "wZàvbœ"
               Case 54: Result = "Pzqvbœ"
               Case 55: Result = "cÂvbœ"
               Case 56: Result = "Qvàvbœ"
               Case 57: Result = "mvZvbœ"
               Case 58: Result = "AvUvbœ"
               Case 59: Result = "EblvU"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 6 Then
            Select Case Val(MyTens)
               Case 6: Result = "Qq"
               Case 60: Result = "lvU"
               Case 61: Result = "GKlwÆ"
               Case 62: Result = "evlwÆ"
               Case 63: Result = "†ZlwÆ"
               Case 64: Result = "†PŠlwÆ"
               Case 65: Result = "cuqlwÆ"
               Case 66: Result = "†QlwÆ"
               Case 67: Result = "mvZlwÆ"
               Case 68: Result = "AvUlwÆ"
               Case 69: Result = "EbmËi"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 7 Then
            Select Case Val(MyTens)
               Case 7: Result = "mvZ"
               Case 70: Result = "mËi"
               Case 71: Result = "GKvËi"
               Case 72: Result = "evnvËi"
               Case 73: Result = "wZqvËi"
               Case 74: Result = "PzqvËi"
               Case 75: Result = "cuPvËi"
               Case 76: Result = "wQqvËi"
               Case 77: Result = "mvZvËi"
               Case 78: Result = "AvUvËi"
               Case 79: Result = "EbAvwk "
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 8 Then
            Select Case Val(MyTens)
               Case 8: Result = "AvU"
               Case 80: Result = "Avwk"
               Case 81: Result = "GKvwk"
               Case 82: Result = "weivwk"
               Case 83: Result = "wZivwk"
               Case 84: Result = "Pzivwk"
               Case 85: Result = "cuPvwk"
               Case 86: Result = "wQqvwk"
               Case 87: Result = "mvZvwk"
               Case 88: Result = "AvUvwk"
               Case 89: Result = "EbbeŸB"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 9 Then
            Select Case Val(MyTens)
               Case 9: Result = "bq"
               Case 90: Result = "beŸB"
               Case 91: Result = "GKvbeŸB"
               Case 92: Result = "weivbeŸB"
               Case 93: Result = "wZivbeŸB"
               Case 94: Result = "PzivbeŸB"
               Case 95: Result = "cuPvbeŸB"
               Case 96: Result = "wQqvbeŸB"
               Case 97: Result = "mvZvbeŸB"
               Case 98: Result = "AvUvbeŸB"
               Case 99: Result = "wbivbeŸB"
               Case Else
            End Select
         Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select
            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If
         ConvertTens = Result
End Function
Private Function ConvertDigit(ByVal MyDigit)
Select Case Val(MyDigit)
            Case 1: ConvertDigit = "GK"
            Case 2: ConvertDigit = "`yB"
            Case 3: ConvertDigit = "wZb"
            Case 4: ConvertDigit = "Pvi"
            Case 5: ConvertDigit = "cvuP"
            Case 6: ConvertDigit = "Qq"
            Case 7: ConvertDigit = "mvZ"
            Case 8: ConvertDigit = "AvU"
            Case 9: ConvertDigit = "bq"
            Case Else: ConvertDigit = ""
         End Select
End Function
