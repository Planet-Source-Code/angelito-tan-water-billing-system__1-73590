Attribute VB_Name = "Module1"
Option Explicit
'--------------------------------------------------------------------------------------------------------  |
'    Author: Angelito S. Tan       ( Wednesday, May, 23, 2007 )                                                                        |
'    E-mail me @: diamond_shoe1234@yahoo.com                                                               |
'    Description: Encrypt and decrypt a string to ascii and ascii to string                                |
'                  this type of encryption is seeding ..                                                   |
'                  as u run the program evertym u press the encrypt button it produce different character  |
'                  combination                                                                             |
'                  u can try this by putting this different character below                                |
'                  try putting this Encrypted text into text 2 the result will be in the text3              |
'                       j∏~kåu«fwè7A†græ±¨y:íT7t}} << First Encrypted text                                |
'                       6uvxÄ¢ƒàE∞ªGHÆôº«°ªØclç9Z•y << Second Encrypted text                               |
'                                                                                                          |
'-----------------------------------------------------------------------------------------------------------
            'Find a bug and solve it !!
            'I find one :))
            
'--- ENCRYPT VARIABLES
Dim seed&, seed2, seed1, crypt1, E_enc As Integer
Dim E_enc_result As String

'--- DECRYPT VARIABLES
Dim crypt2, D_enc As Integer
Dim D_enc_result As String

Public Function ENCRYPT(ByVal Enc As String)
seed2 = Timer
seed& = Timer

Rnd -1
seed1 = (Mid(Timer, Len(seed2), 2) * 2) + Mid(seed, Len(seed), 2)
Randomize seed1
E_enc_result = Chr(seed1 And 255)
    For crypt1 = 1 To Len(Enc)
            E_enc = Asc(Mid(Enc, crypt1, 1)) + Int(Rnd * 255)
            E_enc_result = E_enc_result & Chr(E_enc And 255)
    Next
ENCRYPT = E_enc_result
End Function

Public Function Decrypt(Dec As String)
Rnd -1
Randomize Asc(Mid(Dec, 1, 1))
    For crypt2 = 2 To Len(Dec)
        D_enc = Asc(Mid(Dec, crypt2, 1)) - Int(Rnd * 255)
        D_enc_result = D_enc_result & Chr(D_enc And 255)
    Next
Decrypt = D_enc_result
End Function

