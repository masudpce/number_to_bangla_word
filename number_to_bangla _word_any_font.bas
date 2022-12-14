'Attribute VB_Name = "VBA Access Excel - Number to Bangla word Converter"
' ****  Author          : Masud Ahmed (masudahmed.cuet@gmail.com)
' ****  Tittle          : Converting Currency(Bengali System) to Words without old-font problem.
' ****  Copyright Owner : Masud Ahmed (masudahmed.cuet@gmail.com)
' ****  Description     : This utility converts currencies in Bengali numbering system to words. 
' *********************** It can work with any modern font. Does not depend on sutonnyMJ like fonts.
' ****  Limitations     : Don't Know
' =======================================
' based on:
' ****  Author          : Abdulla Al Mamun
' ****  Tittle          : Converting Currency(Bengali System) to Words
' ****  Copyright Owner : Mamun Academy
' ****  Description     : This utility converts currencies in Bengali numbering system to words.
' ****  Limitations     : Converts only upto 10,00,00,000( Ten Crores). Depends on sutonnyMJ like fonts.
' ========================================
' Use instruction: Copy all the code and paste into new blank module . 
' Or drag and drop this file(has .bas extension in file name).
' Save the module in vba editor with any name.
' Use function>>>    ConvertCurrencyToBangla(YourNumber)   <<<<in any textbox you want.





Function ConvertCurrencyToBangla(ByVal MyNumber)
Dim Temp
         Dim Rupees, Paise
         Dim DecimalPlace, Count
         ReDim Place(9) As String
         Place(2) = " হাজার "
         Place(3) = " লক্ষ "
         Place(4) = " কোটি "
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
               Rupees = Rupees & " টাকা "
         End Select
         ' Clean up Paise.
         Select Case Paise
            Case ""
               Paise = ""
            Case Else
               Paise = " এবং " & Paise & " পয়সা "
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
            Result = ConvertDigit(Left(MyNumber, 1)) & " শত "
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
               Case 1: Result = "এক"
               Case 10: Result = "দশ"
               Case 11: Result = "এগারো"
               Case 12: Result = "বার"
               Case 13: Result = "তের"
               Case 14: Result = "চৌদ্দ"
               Case 15: Result = "পনের"
               Case 16: Result = "ষোল"
               Case 17: Result = "সতের"
               Case 18: Result = "আঠার"
               Case 19: Result = "উনিশ"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 2 Then
            Select Case Val(MyTens)
               Case 2: Result = "দুই"
               Case 20: Result = "বিশ"
               Case 21: Result = "একুশ"
               Case 22: Result = "বাইশ"
               Case 23: Result = "তেইশ"
               Case 24: Result = "চব্বিশ"
               Case 25: Result = "পঁচিশ"
               Case 26: Result = "ছাব্বিশ"
               Case 27: Result = "সাতাশ"
               Case 28: Result = "আঠাশ"
               Case 29: Result = "ঊনত্রিশ"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 3 Then
            Select Case Val(MyTens)
               Case 3: Result = "তিন"
               Case 30: Result = "ত্রিশ"
               Case 31: Result = "একত্রিশ"
               Case 32: Result = "বত্রিশ"
               Case 33: Result = "তেত্রিশ"
               Case 34: Result = "চৌত্রিশ"
               Case 35: Result = "পঁয়ত্রিশ"
               Case 36: Result = "ছত্রিশ"
               Case 37: Result = "সাঁইত্রিশ"
               Case 38: Result = "আটত্রিশ"
               Case 39: Result = "ঊনচল্লিশ"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 4 Then
            Select Case Val(MyTens)
               Case 4: Result = "চার"
               Case 40: Result = "চল্লিশ"
               Case 41: Result = "একচল্লিশ"
               Case 42: Result = "বিয়াল্লিশ"
               Case 43: Result = "তেতাল্লিশ"
               Case 44: Result = "চুয়াল্লিশ"
               Case 45: Result = "পঁয়তাল্লিশ"
               Case 46: Result = "ছেচল্লিশ"
               Case 47: Result = "সাতচল্লিশ"
               Case 48: Result = "আটচল্লিশ"
               Case 49: Result = "ঊনপঞ্চাশ"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 5 Then
            Select Case Val(MyTens)
               Case 5: Result = "পাঁচ"
               Case 50: Result = "পঞ্চাশ"
               Case 51: Result = "একান্ন"
               Case 52: Result = "বায়ান্ন"
               Case 53: Result = "তিপ্পান্ন"
               Case 54: Result = "চুয়ান্ন"
               Case 55: Result = "পঞ্চান্ন"
               Case 56: Result = "ছাপ্পান্ন"
               Case 57: Result = "সাতান্ন"
               Case 58: Result = "আটান্ন"
               Case 59: Result = "ঊনষাট"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 6 Then
            Select Case Val(MyTens)
               Case 6: Result = "ছয়"
               Case 60: Result = "ষাট"
               Case 61: Result = "একষট্টি"
               Case 62: Result = "বাষট্টি"
               Case 63: Result = "তেষট্টি"
               Case 64: Result = "চৌষট্টি"
               Case 65: Result = "পঁয়ষট্টি"
               Case 66: Result = "ছেষট্টি"
               Case 67: Result = "সাতষট্টি"
               Case 68: Result = "আটষট্টি"
               Case 69: Result = "ঊনসত্তর"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 7 Then
            Select Case Val(MyTens)
               Case 7: Result = "সাত"
               Case 70: Result = "সত্তর"
               Case 71: Result = "একাত্তর"
               Case 72: Result = "বাহাত্তর"
               Case 73: Result = "তিয়াত্তর"
               Case 74: Result = "চুয়াত্তর"
               Case 75: Result = "পঁচাত্তর"
               Case 76: Result = "ছিয়াত্তর"
               Case 77: Result = "সাতাত্তর"
               Case 78: Result = "আটাত্তর"
               Case 79: Result = "ঊনআশি"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 8 Then
            Select Case Val(MyTens)
               Case 8: Result = "আট"
               Case 80: Result = "আশি"
               Case 81: Result = "একাশি"
               Case 82: Result = "বিরাশি"
               Case 83: Result = "তিরাশি"
               Case 84: Result = "চুরাশি"
               Case 85: Result = "পঁচাশি"
               Case 86: Result = "ছিয়াশি"
               Case 87: Result = "সাতাশি"
               Case 88: Result = "আটাশি"
               Case 89: Result = "ঊননব্বই"
               Case Else
            End Select
        ElseIf Val(Left(MyTens, 1)) = 9 Then
            Select Case Val(MyTens)
               Case 9: Result = "নয়"
               Case 90: Result = "নব্বই"
               Case 91: Result = "একানব্বই"
               Case 92: Result = "বিরানব্বই"
               Case 93: Result = "তিরানব্বই"
               Case 94: Result = "চুরানব্বই"
               Case 95: Result = "পঁচানব্বই"
               Case 96: Result = "ছিয়ানব্বই"
               Case 97: Result = "সাতানব্বই"
               Case 98: Result = "আটানব্বই"
               Case 99: Result = "নিরানব্বই"
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
            Case 1: ConvertDigit = "এক"
            Case 2: ConvertDigit = "দুই"
            Case 3: ConvertDigit = "তিন"
            Case 4: ConvertDigit = "চার"
            Case 5: ConvertDigit = "পাঁচ"
            Case 6: ConvertDigit = "ছয়"
            Case 7: ConvertDigit = "সাত"
            Case 8: ConvertDigit = "আট"
            Case 9: ConvertDigit = "নয়"
            Case Else: ConvertDigit = ""
         End Select
End Function

