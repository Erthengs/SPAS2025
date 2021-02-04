Module Functions

    Public Function GetDouble(ByVal doublestring As String) As Double
        Dim retval As Double
        'Dim sep As String = CurrentCulture.NumberFormat.NumberDecimalSeparator

        Double.TryParse(doublestring, retval)
        Return retval
    End Function

    Function Tbx2Dec(ByVal tbx As String)
        Dim Number As Decimal
        If Decimal.TryParse(tbx, Number) Then tbx = Format(Number, "#,##0.00") Else tbx = 0
        Return tbx
    End Function
    Function Tbx2Int(ByVal tbx As String)
        Dim Number As Integer
        If Decimal.TryParse(tbx, Number) Then tbx = Format(Number, "#,##0") Else tbx = 0
        Return tbx
    End Function

    Function IBANcheck(IBAN)
        'If IBANcheck(Main.Tbx_Relatie_IBAN_08.Value) <> 1 Then ErrMsg = ErrMsg & "- IBAN-nummer is niet correct"
        Dim myLengte As Integer, myIBANnew As String, I As Integer, myIBANcheck, letter As String

        myLengte = Len(IBAN)
        myIBANnew = Right(IBAN, myLengte - 4) & Left(IBAN, 4)
        myIBANcheck = Asc(Mid(myIBANnew, 1, 1).ToUpper) - 55


        For I = 2 To myLengte
            letter = Mid(myIBANnew, I, 1)
            If Not IsNumeric(letter) Then letter = Asc(Mid(myIBANnew, I, 1).ToUpper) - 55
            myIBANcheck &= letter
        Next

        'IBANcheck = myIBANcheck 'testcommand
        IBANcheck = Modulo(myIBANcheck, 97)
        Return IBANcheck



    End Function

    Function IBAN_Length(ByVal countrycode As String)

        Select Case countrycode
            Case "BE"
                Return 18
            Case "DK" Or "NL" Or "FI"
                Return 18
            Case Else
                Return 0
        End Select

    End Function
    Function Tabs(ByVal n As Integer)
        Dim Tab As String = ""
        For i = 1 To n
            Tab &= "    "
        Next
        Return Tab



    End Function
    Function Modulo(strGetal As String, intRest As Integer)

#If VBA7 And Win64 Then
    Dim intC As LongPtr
#Else
        Dim intC As Long
#End If

        Dim I As Integer
        For I = Len(strGetal) To 1 Step -1
            intC = intC + (Mid(strGetal, I, 1) *
                TienTallen(Len(strGetal) - I, intRest)) Mod intRest
        Next
        intC = intC Mod intRest
        Modulo = intC

    End Function
    Function TienTallen(intMacht As Integer, intRest As Integer)

        Select Case intMacht
            Case Is <= 8
                TienTallen = 10 ^ intMacht Mod intRest
            Case Is <= 16
                TienTallen = TienTallen(8, intRest) *
                    (10 ^ (intMacht - 8) Mod intRest) Mod intRest
            Case Is <= 24
                TienTallen = TienTallen(16, intRest) *
                    (10 ^ (intMacht - 16) Mod intRest) Mod intRest
            Case Is <= 32
                TienTallen = TienTallen(24, intRest) *
                    (10 ^ (intMacht - 24) Mod intRest) Mod intRest
            Case Is <= 40
                TienTallen = TienTallen(32, intRest) *
                    (10 ^ (intMacht - 32) Mod intRest) Mod intRest
            Case Is <= 48
                TienTallen = TienTallen(40, intRest) *
                    (10 ^ (intMacht - 40) Mod intRest) Mod intRest
            Case Else
                MsgBox("Getal is te lang")
                TienTallen = ""
        End Select


    End Function


End Module
