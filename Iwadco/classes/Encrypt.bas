Attribute VB_Name = "Encrypt"
'--- DECRYPT VARIABLES
Dim crypt2, D_enc As Integer
Dim D_enc_result As String

Public Function Decrypt(Dec As String)
Rnd -1
Randomize Asc(Mid(Dec, 1, 1))
    For crypt2 = 2 To Len(Dec)
        D_enc = Asc(Mid(Dec, crypt2, 1)) - Int(Rnd * 255)
        D_enc_result = D_enc_result & Chr(D_enc And 255)
    Next
Decrypt = ""
Decrypt = D_enc_result
End Function

