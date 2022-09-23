Attribute VB_Name = "Module1"
'Attribute VB_Name = "Module2"
' ****  Author          : Abdulla Al Mamun
' ****  Tittle          : Converting Currency(Bengali System) to Words
' ****  Copyright Owner : Mamun Academy
' ****  Description     : This utility converts currencies in Bengali numbering system to words.
' ****  Limitations     : Converts only upto 10,00,00,000( Ten Crores)

Function ConvertCurrencyToBangla(ByVal MyNumber)
Dim Temp
         Dim Rupees, Paise
         Dim DecimalPlace, Count
         ReDim Place(9) As String
         Place(2) = " nvRvi "
         Place(3) = " jvL "
         Place(4) = " �KvwU "
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
         ConvertCurrencyToBangla = Rupees & Paise
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
               Case 13: Result = "�Zi"
               Case 14: Result = "�P��"
               Case 15: Result = "c�bi"
               Case 16: Result = "�lvj"
               Case 17: Result = "m�Zi"
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
               Case 23: Result = "�ZBk"
               Case 24: Result = "Pwe�k"
               Case 25: Result = "cuwPk"
               Case 26: Result = "Qvwe�k"
               Case 27: Result = "mvZvk"
               Case 28: Result = "AvVvk"
               Case 29: Result = "Ebw�k"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 3 Then
            Select Case Val(MyTens)
               Case 3: Result = "wZb"
               Case 30: Result = "w�k"
               Case 31: Result = "GKw�k"
               Case 32: Result = "ew�k"
               Case 33: Result = "�Zw�k"
               Case 34: Result = "�P�w�k"
               Case 35: Result = "cuqw�k"
               Case 36: Result = "Qw�k"
               Case 37: Result = "mvuBw�k"
               Case 38: Result = "AvUw�k"
               Case 39: Result = "EbPwj�k"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 4 Then
            Select Case Val(MyTens)
               Case 4: Result = "Pvi"
               Case 40: Result = "Pwj�k"
               Case 41: Result = "GKPwj�k"
               Case 42: Result = "weqvwj�k"
               Case 43: Result = "�ZZvwj�k"
               Case 44: Result = "Pzqvwj�k"
               Case 45: Result = "cuqZvwj�k"
               Case 46: Result = "�QPwj�k"
               Case 47: Result = "mvZPwj�k"
               Case 48: Result = "AvUPwj�k"
               Case 49: Result = "Ebc�vk"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 5 Then
            Select Case Val(MyTens)
               Case 5: Result = "cvuP"
               Case 50: Result = "c�vk"
               Case 51: Result = "GKvb�"
               Case 52: Result = "evqvb�"
               Case 53: Result = "wZ�vb�"
               Case 54: Result = "Pzqvb�"
               Case 55: Result = "c�vb�"
               Case 56: Result = "Qv�vb�"
               Case 57: Result = "mvZvb�"
               Case 58: Result = "AvUvb�"
               Case 59: Result = "EblvU"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 6 Then
            Select Case Val(MyTens)
               Case 6: Result = "Qq"
               Case 60: Result = "lvU"
               Case 61: Result = "GKlw�"
               Case 62: Result = "evlw�"
               Case 63: Result = "�Zlw�"
               Case 64: Result = "�P�lw�"
               Case 65: Result = "cuqlw�"
               Case 66: Result = "�Qlw�"
               Case 67: Result = "mvZlw�"
               Case 68: Result = "AvUlw�"
               Case 69: Result = "Ebm�i"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 7 Then
            Select Case Val(MyTens)
               Case 7: Result = "mvZ"
               Case 70: Result = "m�i"
               Case 71: Result = "GKv�i"
               Case 72: Result = "evnv�i"
               Case 73: Result = "wZqv�i"
               Case 74: Result = "Pzqv�i"
               Case 75: Result = "cuPv�i"
               Case 76: Result = "wQqv�i"
               Case 77: Result = "mvZv�i"
               Case 78: Result = "AvUv�i"
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
               Case 89: Result = "Ebbe�B"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 9 Then
            Select Case Val(MyTens)
               Case 9: Result = "bq"
               Case 90: Result = "be�B"
               Case 91: Result = "GKvbe�B"
               Case 92: Result = "weivbe�B"
               Case 93: Result = "wZivbe�B"
               Case 94: Result = "Pzivbe�B"
               Case 95: Result = "cuPvbe�B"
               Case 96: Result = "wQqvbe�B"
               Case 97: Result = "mvZvbe�B"
               Case 98: Result = "AvUvbe�B"
               Case 99: Result = "wbivbe�B"
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







