Attribute VB_Name = "modGeneral"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SendASP As String = "http://cc.domaindlx.com/neenojee/msnbc/msnbc.asp"

Sub RunHyperlink(ByVal LINK As String)
Dim Success As Long
Success = ShellExecute(0&, vbNullString, LINK, vbNullString, "C:\", 1)
End Sub

Function RetPar(ByVal sString As String, ByVal ParNo As Long, Optional ByVal SepaBy As String) As String
If SepaBy = "" Then SepaBy = ","
sString = sString & SepaBy
For z = 1 To ParNo - 1
i = InStr(i + 1, sString, SepaBy)
Next z
I2 = InStr(i + 1, sString, SepaBy)
If I2 > 0 Then
RetPar = Mid(sString, i + 1, I2 - i - 1)
Else
RetPar = ""
Exit Function
End If
End Function

Function RetVal(ByVal sString As String)
If InStr(1, sString, "=") > 0 Then
RetVal = Trim(Mid(sString, InStr(1, sString, "=") + 1))
Else
RetVal = ""
End If
End Function

Function RetVar(ByVal sString As String)
If InStr(1, sString, "=") > 0 Then
RetVar = Trim(Left(sString, InStr(1, sString, "=") - 1))
Else
RetVar = ""
End If
End Function

Function Code2Normal(ByVal text As String) As String
On Error GoTo ErrorOccured
k2 = text
For z = 1 To Len(k2)
If Mid(k2, z, 1) = "%" Then
k = k & Chr(Hex2Dec(Mid(k2, z + 1, 2)))
z = z + 2
Else
If Mid(k2, z, 1) = "+" Then
k = k & " "
Else
k = k & Mid(k2, z, 1)
End If
End If
Next z
Code2Normal = k
Exit Function
ErrorOccured:
Code2Normal = ""
End Function

Function Normal2Code(ByVal text As String) As String
On Error GoTo ErrorOccured
k = ""
i = text
For z = 1 To Len(i)
m = Mid(i, z, 1)

If Asc(m) = 32 Then
k = k & "+"
GoTo nxtz
End If

If Asc(m) < 48 And Asc(m) > -1 Then
k = k & "%" & Hex(Asc(m))
GoTo nxtz
End If

If Asc(m) < 65 And Asc(m) > 57 Then
k = k & "%" & Hex(Asc(m))
GoTo nxtz
End If

If Asc(m) < 97 And Asc(m) > 90 Then
k = k & "%" & Hex(Asc(m))
GoTo nxtz
End If

If Asc(m) > 122 Then
k = k & "%" & Hex(Asc(m))
GoTo nxtz
End If

k = k & m
nxtz:
Next z
Normal2Code = k
Exit Function
ErrorOccured:
Normal2Code = ""
End Function

Function Hex2Dec(ByVal HexValue As String) As Double
k$ = StrReverse(Right(HexValue, 8))
For z = 1 To Len(k$)
kj$ = Mid$(k$, z, 1)
If Asc(kj$) > 64 Then i = (Asc(kj$) - 64) + 9 Else i = Val(kj$)
Hex2Dec = Hex2Dec + i * 16 ^ (z - 1)
Next z
End Function

Function Oct2Dec(ByVal OctValue As String) As Double
k$ = StrReverse(Right(OctValue, 11))
For z = 1 To Len(k$)
kj$ = Mid$(k$, z, 1)
i = Val(kj$)
Oct2Dec = Oct2Dec + i * 8 ^ (z - 1)
Next z
End Function

Function Bin2Dec(ByVal BinValue As String) As Double
k$ = StrReverse(Right(BinValue, 32))
For z = 1 To Len(k$)
kj$ = Mid$(k$, z, 1)
i = Val(kj$)
Bin2Dec = Bin2Dec + i * 2 ^ (z - 1)
Next z
End Function

Function BaseConvert(ByVal Number As Long, ByVal ToBase As Long, Optional MidSpaceString As String) As String
convBase = ToBase
Mnumber = Number
re:
md = Number \ convBase
DGT = Number \ convBase
LFT = (Number / convBase) - DGT
LftDgt = LFT * convBase
an = Val(Mid$(Str$(LftDgt), 2)) & MidSpaceString & an
If DGT >= convBase Then
Number = DGT
GoTo re
Else
LftDgt = DGT
an = Val(Mid$(Str$(LftDgt), 2)) & MidSpaceString & an
End If
BaseConvert = an
End Function

Function HexFormat(ByVal TXT As String) As String
TXT = Replace(TXT, Chr(37), "%" & Format(Hex(z), "00"))
For z = 1 To 36
TXT = Replace(TXT, Chr(z), "%" & Format(Hex(z), "00"))
Next z
For z = 38 To 47
TXT = Replace(TXT, Chr(z), "%" & Format(Hex(z), "00"))
Next z
For z = 58 To 64
TXT = Replace(TXT, Chr(z), "%" & Format(Hex(z), "00"))
Next z
For z = 91 To 96
TXT = Replace(TXT, Chr(z), "%" & Format(Hex(z), "00"))
Next z
For z = 123 To 255
TXT = Replace(TXT, Chr(z), "%" & Format(Hex(z), "00"))
Next z
HexFormat = TXT
End Function

Function GetMSNStatusCodeMeaning(ByVal StatusCode As Integer) As String

    Select Case StatusCode
        
        Case 3: GetMSNStatusCodeMeaning = "Blocked You"
        Case 5: GetMSNStatusCodeMeaning = "You have blocked"

    End Select

    If GetMSNStatusCodeMeaning <> "" Then GetMSNStatusCodeMeaning = "(" & GetMSNStatusCodeMeaning & ")" 'Else GetMSNStatusCodeMeaning = "(" & StatusCode & ")"
End Function


