Attribute VB_Name = "modCCPreAuth"
'This is a custom return type to be used as the return
'data type for the function ValidateCreditCardNumber
Public Type CreditCardStats
    IsValidNumber As Boolean
    CreditCardCo As Integer
    CheckSum As String
    TotalSum As String
End Type

'In order to save some time, we declare a public const
'consisting of 20 zeros.
'This is done so that left padding of the inversed credit
'credit card number can be done with ease.
Public Const CreditCardNulls = "00000000000000000000"


Public Function ValidateCreditCardNumber(CreditCardNumber As String) As CreditCardStats
'Credit card pre-authorization function
'This is where it all happens.
    Dim CCNumTemp() As Byte
    Dim CCNumVal() As Byte
    CCNumVal() = CreditCardNumber
    CreditCardNumber = ""
    For X = 0 To UBound(CCNumVal)
        Select Case IsNumeric(Chr(CCNumVal(X)))
            Case False
            Case True
                CreditCardNumber = CreditCardNumber & Chr(CCNumVal(X))
        End Select
    Next X
    If CreditCardNumber = "" Then CreditCardNumber = "0"
    Select Case Len(CreditCardNumber)
        Case 13
            Select Case Left(CreditCardNumber, 1)
                Case "4"
                    ValidateCreditCardNumber.CreditCardCo = 2
                Case Else
                    ValidateCreditCardNumber.CreditCardCo = 0
            End Select
        Case 14
            Select Case Left(CreditCardNumber, 2)
                Case "30"
                    Select Case Left(CreditCardNumber, 3)
                        Case "300"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case "301"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case "302"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case "303"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case "304"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case "305"
                            ValidateCreditCardNumber.CreditCardCo = 4
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case "36"
                    ValidateCreditCardNumber.CreditCardCo = 4
                Case "38"
                    ValidateCreditCardNumber.CreditCardCo = 4
                Case Else
                    ValidateCreditCardNumber.CreditCardCo = 0
            End Select
        Case 15
            Select Case Left(CreditCardNumber, 2)
                Case "34"
                    ValidateCreditCardNumber.CreditCardCo = 3
                Case "37"
                    ValidateCreditCardNumber.CreditCardCo = 3
                Case "20"
                    Select Case Left(CreditCardNumber, 4)
                        Case "2014"
                            ValidateCreditCardNumber.CreditCardCo = 6
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case "21"
                    Select Case Left(CreditCardNumber, 4)
                        Case "2149"
                            ValidateCreditCardNumber.CreditCardCo = 6
                        Case "2131"
                            ValidateCreditCardNumber.CreditCardCo = 7
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case "18"
                    Select Case Left(CreditCardNumber, 4)
                        Case "1800"
                            ValidateCreditCardNumber.CreditCardCo = 7
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case Else
                    ValidateCreditCardNumber.CreditCardCo = 0
            End Select
        Case 16
            Select Case Left(CreditCardNumber, 1)
                Case "3"
                    ValidateCreditCardNumber.CreditCardCo = 7
                Case "4"
                    ValidateCreditCardNumber.CreditCardCo = 2
                Case "5"
                    Select Case Left(CreditCardNumber, 2)
                        Case "51"
                            ValidateCreditCardNumber.CreditCardCo = 1
                        Case "52"
                            ValidateCreditCardNumber.CreditCardCo = 1
                        Case "53"
                            ValidateCreditCardNumber.CreditCardCo = 1
                        Case "54"
                            ValidateCreditCardNumber.CreditCardCo = 1
                        Case "55"
                            ValidateCreditCardNumber.CreditCardCo = 1
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case "6"
                    Select Case Left(CreditCardNumber, 4)
                        Case "6011"
                            ValidateCreditCardNumber.CreditCardCo = 5
                        Case Else
                            ValidateCreditCardNumber.CreditCardCo = 0
                    End Select
                Case Else
                    ValidateCreditCardNumber.CreditCardCo = 0
            End Select
        Case Else
            ValidateCreditCardNumber.CreditCardCo = 0
    End Select
    Select Case ValidateCreditCardNumber.CreditCardCo
        Case 0
        Case Else
            'LUHN Formula (Mod 10)
            Dim MultiplyByTwo As Boolean
            Dim CCNumTempNum As Integer
            For X = 1 To Len(CreditCardNumber)
                Select Case MultiplyByTwo
                    Case False
                        MultiplyByTwo = True
                        CCNumTempNum = Mid(Right(CreditCardNumber, X), 1, 1)
                    Case True
                        MultiplyByTwo = False
                        CCNumTempNum = (Mid(Right(CreditCardNumber, X), 1, 1) * 2)
                End Select
                Select Case CCNumTempNum
                    Case 9
                    Case 8
                    Case 7
                    Case 6
                    Case 5
                    Case 4
                    Case 3
                    Case 2
                    Case 1
                    Case 0
                    Case Else
                        CCNumTempNum = (Int(CCNumTempNum / 10) + (CCNumTempNum Mod 10))
                End Select
                ValidateCreditCardNumber.CheckSum = CCNumTempNum & ValidateCreditCardNumber.CheckSum
                ValidateCreditCardNumber.TotalSum = IIf(ValidateCreditCardNumber.TotalSum = "", "0", ValidateCreditCardNumber.TotalSum) + CCNumTempNum
                DoEvents
            Next X
            Select Case (ValidateCreditCardNumber.TotalSum Mod 10)
                Case 0
                    ValidateCreditCardNumber.IsValidNumber = True
                Case Else
                    ValidateCreditCardNumber.IsValidNumber = False
                    ValidateCreditCardNumber.CreditCardCo = 0
            End Select
    End Select
End Function

