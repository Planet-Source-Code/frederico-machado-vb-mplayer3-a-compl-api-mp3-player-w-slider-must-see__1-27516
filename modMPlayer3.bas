Attribute VB_Name = "modMPlayer3"
Public Const AliasName = "audiomp3"
Public strFilePath As String

Public Sub ReadMP3Header(sPassFileName As String)
Dim z, i
Dim BinaryString As String
Dim byteArray(4) As Byte    'array that store first four bytes
Dim bin As String           'string that store binary number converted from readed bytes
Dim BinString As String     'containing binary string
Dim DecString As Integer  'containing decimal extracted from BinString
'''''''''''''''end of declarations'''''''

Open sPassFileName For Binary Access Read As #1  'open file #1 for read
   For z = 1 To 4                           'step through four bytes
   Get #1, z, byteArray(z)                  'store every(z)byte  in array position z
   Next z                                   'back for next byte
 Close #1                                   'close file
 bin = ""                                   'reset and build the desired binary number in this string
   For z = 1 To 4                           'convert all bytes to binary
     For i = 0 To 7 Step 1                  'Here comes the decimal=>binary conversion
         If byteArray(z) And (2 ^ i) Then   'Use the logical "AND" operator.
            bin = bin + "1"
            Else
            bin = bin + "0"
         End If
         Next i                             'End of binary conversion
Next z
BinaryString = bin
'''''''''check MP3HeaderInfo.Frequency''''
DecString = 0
BinString = Mid(bin, 19, 2)         'take 19 to 21
For i = 1 To Len(BinString)         'convert to decimal
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblKhz = 44
  Case 1
    frmMain.lblKhz = 32
  Case 2
    frmMain.lblKhz = 48
  Case 3
End Select
''''check MP3HeaderInfo.Mode''''
DecString = 0
BinString = Mid(bin, 31, 2)
For i = 1 To Len(BinString)
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblmode = "stereo"
  Case 1
    frmMain.lblmode = "stereo"
  Case 2
    frmMain.lblmode = "stereo"
  Case 3
    frmMain.lblmode = "mono"
End Select
'''''check MP3HeaderInfo.Bitrate''''
DecString = 0
BinString = Mid(bin, 21, 4)
For i = 1 To Len(BinString)
  If Mid(BinString, i, 1) = 1 Then
    DecString = DecString + 2 ^ (Len(BinString) - i)
  End If
Next i
Select Case DecString
  Case 0
    frmMain.lblbitrate = 0
  Case 1
    frmMain.lblbitrate = 112
  Case 2
    frmMain.lblbitrate = 56
  Case 3
    frmMain.lblbitrate = 224
  Case 4
    frmMain.lblbitrate = 40
  Case 5
    frmMain.lblbitrate = 160
  Case 6
    frmMain.lblbitrate = 80
  Case 7
    frmMain.lblbitrate = 320
  Case 8
    frmMain.lblbitrate = 32
  Case 9
    frmMain.lblbitrate = 128
  Case 10
    frmMain.lblbitrate = 64
  Case 11
    frmMain.lblbitrate = 256
  Case 12
    frmMain.lblbitrate = 48
  Case 13
    frmMain.lblbitrate = 192
  Case 14
    frmMain.lblbitrate = 96
  Case 15
    frmMain.lblbitrate = 0
End Select
End Sub
