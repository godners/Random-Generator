Public Class WinMain
    Private SymbolCheck As Boolean = False
    Private CharString As String = ""
    Const LowerString As String = "abcdefghijklmnopqrstuvwxyz"
    Const UpperString As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Const DigitString As String = "0123456789"
    Private ResultString As String = ""
    Private rnd As New Random
    Private Sub WinMain_Load() Handles MyBase.Load
        Call CharAll()
        BtnGen.PerformClick()
    End Sub
    Private Function Generate(l As Byte) As String
        Dim i As Byte
        Dim oup As String = ""
        For i = 1 To l Step 1
            oup += CharString.Substring(rnd.Next(0, CharString.Length), 1)
        Next
        Return oup
    End Function
    Private Sub CharAll()
        Call CharDigit(CkbDigit.Checked)
        Call CharLower(CkbLower.Checked)
        Call CharUpper(CkbUpper.Checked)
        Call CharAllSymbol()
    End Sub
    Private Sub CharAllSymbol()
        Call CharSymbol(Ckb10.Checked, Ckb10.Text)
        Call CharSymbol(Ckb11.Checked, Ckb11.Text)
        Call CharSymbol(Ckb12.Checked, Ckb12.Text)
        Call CharSymbol(Ckb13.Checked, Ckb13.Text)
        Call CharSymbol(Ckb14.Checked, Ckb14.Text)
        Call CharSymbol(Ckb15.Checked, Ckb15.Text)
        Call CharSymbol(Ckb16.Checked, Ckb16.Text)
        Call CharSymbol(Ckb17.Checked, Ckb17.Text)
        Call CharSymbol(Ckb20.Checked, Ckb20.Text)
        Call CharSymbol(Ckb21.Checked, Ckb21.Text)
        Call CharSymbol(Ckb22.Checked, Ckb22.Text)
        Call CharSymbol(Ckb23.Checked, Ckb23.Text)
        Call CharSymbol(Ckb24.Checked, Ckb24.Text)
        Call CharSymbol(Ckb25.Checked, Ckb25.Text)
        Call CharSymbol(Ckb26.Checked, Ckb26.Text)
        Call CharSymbol(Ckb27.Checked, Ckb27.Text)
        Call CharSymbol(Ckb30.Checked, Ckb30.Text)
        Call CharSymbol(Ckb31.Checked, Ckb31.Text)
        Call CharSymbol(Ckb32.Checked, Ckb32.Text)
        Call CharSymbol(Ckb33.Checked, Ckb33.Text)
        Call CharSymbol(Ckb34.Checked, Ckb34.Text)
        Call CharSymbol(Ckb35.Checked, Ckb35.Text)
        Call CharSymbol(Ckb36.Checked, Ckb36.Text)
        Call CharSymbol(Ckb37.Checked, Ckb37.Text)
        Call CharSymbol(Ckb40.Checked, Ckb40.Text)
        Call CharSymbol(Ckb41.Checked, Ckb41.Text)
        Call CharSymbol(Ckb42.Checked, Ckb42.Text)
        Call CharSymbol(Ckb43.Checked, Ckb43.Text)
        Call CharSymbol(Ckb44.Checked, Ckb44.Text)
        Call CharSymbol(Ckb45.Checked, Ckb45.Text)
        Call CharSymbol(Ckb46.Checked, Ckb46.Text)
        Call CharSymbol(Ckb47.Checked, Ckb47.Text)
    End Sub
    Private Sub CharSymbol(Include As Boolean, str As String)
        If str = "&&" Then str = "&"
        If Include Then
            If Not CharString.Contains(str) Then CharString += str
        Else
            If CharString.Contains(str) Then CharString = CharString.Replace(str, "")
        End If
    End Sub
    Private Sub CharDigit(Include As Boolean)
        If Include Then
            If Not CharString.Contains(DigitString) Then CharString += DigitString
        Else
            If CharString.Contains(DigitString) Then CharString = CharString.Replace(DigitString, "")
        End If
    End Sub
    Private Sub CharLower(Include As Boolean)
        If Include Then
            If Not CharString.Contains(LowerString) Then CharString += LowerString
        Else
            If CharString.Contains(LowerString) Then CharString = CharString.Replace(LowerString, "")
        End If
    End Sub
    Private Sub CharUpper(Include As Boolean)
        If Include Then
            If Not CharString.Contains(UpperString) Then CharString += UpperString
        Else
            If CharString.Contains(UpperString) Then CharString = CharString.Replace(UpperString, "")
        End If
    End Sub
    Private Sub CkbDefault_CheckedChanged() Handles CkbDefault.CheckedChanged
        If CkbDefault.Checked = False Then Exit Sub
        SymbolCheck = True
        CkbAll.Checked = False
        Ckb10.Checked = False
        Ckb11.Checked = False
        Ckb12.Checked = False
        Ckb13.Checked = False
        Ckb14.Checked = False
        Ckb15.Checked = False
        Ckb16.Checked = False
        Ckb17.Checked = False
        Ckb20.Checked = True
        Ckb21.Checked = True
        Ckb22.Checked = False
        Ckb23.Checked = False
        Ckb24.Checked = False
        Ckb25.Checked = True
        Ckb26.Checked = False
        Ckb27.Checked = False
        Ckb30.Checked = False
        Ckb31.Checked = False
        Ckb32.Checked = False
        Ckb33.Checked = False
        Ckb34.Checked = False
        Ckb35.Checked = True
        Ckb36.Checked = True
        Ckb37.Checked = True
        Ckb40.Checked = False
        Ckb41.Checked = False
        Ckb42.Checked = False
        Ckb43.Checked = False
        Ckb44.Checked = True
        Ckb45.Checked = True
        Ckb46.Checked = False
        Ckb47.Checked = False
        SymbolCheck = False
        Call CharAllSymbol()
    End Sub
    Private Sub CkbAll_CheckedChanged() Handles CkbAll.CheckedChanged
        If Not SymbolCheck Then CkbDefault.Checked = False
        SymbolCheck = True
        Ckb10.Checked = CkbAll.Checked
        Ckb11.Checked = CkbAll.Checked
        Ckb12.Checked = CkbAll.Checked
        Ckb13.Checked = CkbAll.Checked
        Ckb14.Checked = CkbAll.Checked
        Ckb15.Checked = CkbAll.Checked
        Ckb16.Checked = CkbAll.Checked
        Ckb17.Checked = CkbAll.Checked
        Ckb20.Checked = CkbAll.Checked
        Ckb21.Checked = CkbAll.Checked
        Ckb22.Checked = CkbAll.Checked
        Ckb23.Checked = CkbAll.Checked
        Ckb24.Checked = CkbAll.Checked
        Ckb25.Checked = CkbAll.Checked
        Ckb26.Checked = CkbAll.Checked
        Ckb27.Checked = CkbAll.Checked
        Ckb30.Checked = CkbAll.Checked
        Ckb31.Checked = CkbAll.Checked
        Ckb32.Checked = CkbAll.Checked
        Ckb33.Checked = CkbAll.Checked
        Ckb34.Checked = CkbAll.Checked
        Ckb35.Checked = CkbAll.Checked
        Ckb36.Checked = CkbAll.Checked
        Ckb37.Checked = CkbAll.Checked
        Ckb40.Checked = CkbAll.Checked
        Ckb41.Checked = CkbAll.Checked
        Ckb42.Checked = CkbAll.Checked
        Ckb43.Checked = CkbAll.Checked
        Ckb44.Checked = CkbAll.Checked
        Ckb45.Checked = CkbAll.Checked
        Ckb46.Checked = CkbAll.Checked
        Ckb47.Checked = CkbAll.Checked
        SymbolCheck = False
        Call CharAllSymbol()
    End Sub
    Private Sub CkbDigit_CheckedChanged() Handles CkbDigit.CheckedChanged
        Call CharDigit(CkbDigit.Checked)
    End Sub
    Private Sub CkbLower_CheckedChanged() Handles CkbLower.CheckedChanged
        Call CharLower(CkbLower.Checked)
    End Sub
    Private Sub CbbUpper_CheckedChanged() Handles CkbUpper.CheckedChanged
        Call CharUpper(CkbUpper.Checked)
    End Sub
    Private Sub CkbSymbol_CheckedChanged(sender As CheckBox, e As EventArgs) Handles _
            Ckb10.CheckedChanged, Ckb11.CheckedChanged, Ckb12.CheckedChanged, Ckb13.CheckedChanged,
            Ckb14.CheckedChanged, Ckb15.CheckedChanged, Ckb16.CheckedChanged, Ckb17.CheckedChanged,
            Ckb20.CheckedChanged, Ckb21.CheckedChanged, Ckb22.CheckedChanged, Ckb23.CheckedChanged,
            Ckb24.CheckedChanged, Ckb25.CheckedChanged, Ckb26.CheckedChanged, Ckb27.CheckedChanged,
            Ckb30.CheckedChanged, Ckb31.CheckedChanged, Ckb32.CheckedChanged, Ckb33.CheckedChanged,
            Ckb34.CheckedChanged, Ckb35.CheckedChanged, Ckb36.CheckedChanged, Ckb37.CheckedChanged,
            Ckb40.CheckedChanged, Ckb41.CheckedChanged, Ckb42.CheckedChanged, Ckb43.CheckedChanged,
            Ckb44.CheckedChanged, Ckb45.CheckedChanged, Ckb46.CheckedChanged, Ckb47.CheckedChanged
        Call CharSymbol(sender.Checked, sender.Text)
    End Sub
    Private Sub BtnGen_Click() Handles BtnGen.Click
        Dim l As Byte = Convert.ToByte(NudLength.Value)
        ResultString = Generate(l)
        If l > 30 Then
            If l > 56 Then
                LblResult.Text = "Result: " & ResultString.Substring(0, 30).Replace("&", "&&") &
                    vbCrLf & ResultString.Substring(30).PadLeft(38, " ").Replace("&", "&&")
            Else
                If l Mod 2 = 0 Then
                    LblResult.Text = "Result: " & ResultString.Substring(0, l / 2).Replace("&", "&&") &
                        vbCrLf & "        " & ResultString.Substring(l / 2).Replace("&", "&&")
                Else
                    LblResult.Text = "Result: " & ResultString.Substring(0, l / 2 + 0.5).Replace("&&", "&&") & vbCrLf &
                        "        " & ResultString.Substring(l / 2 + 0.5).Replace("&", "&&")
                End If
            End If
        Else
            LblResult.Text = "Result: " & ResultString
        End If
    End Sub
    Private Sub BtnCopy_Click() Handles BtnCopy.Click
        My.Computer.Clipboard.SetText(ResultString)
    End Sub
    Private Sub NudLength_KeyDown(sender As NumericUpDown, e As KeyEventArgs) Handles NudLength.KeyDown
        If e.KeyCode = Keys.Enter Then BtnGen.PerformClick()
    End Sub
End Class
