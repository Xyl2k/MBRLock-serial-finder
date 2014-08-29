VERSION 5.00
Begin VB.Form Frm_serial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MBRLocker Builder v0.2: Serial finder"
   ClientHeight    =   4845
   ClientLeft      =   135
   ClientTop       =   3720
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Frm.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Drop your infected MBR Dump"
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Frm_serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MBRLocker Builder v0.2 Generic serial retriever
'Contact: xylitol@malwareint.com
'Add just one Textbox with name: Text1
'Then just Drop your infected MBR Dump

Option Explicit
Dim XorKey, strSubstr, buffer, sHold, sText, sFileExtention, TotalChaine, Path, ipos, sOutput As String
Dim iLocationOfString1, iLocationOfString2, iCompareStyle As Long
Dim a, b, i As Integer

Private Sub Form_Load()
Me.OLEDropMode = 1  'Manual Ole Drop Mode
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, _
                              Effect As Long, _
                              Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
  With Data
  sText = ""
    If .GetFormat(vbCFFiles) Then
      If .Files.Count = 1 Then
        'Only one filename was dragged.  Retrieve it.
        Path = .Files(1)  'Note numeral 1.
        'Pull text file data into our textbox.
          Open Path For Input As #1
            Do Until EOF(1)
              Line Input #1, sHold
              sText = sText & sHold & vbCrLf
            Loop
          Close #1
          Text1.Text = sText
      If .Files.Count > 1 Then
      Exit Sub ' Dude, wtf are you doing here ?
      End If
    End If
  End If
End With

ipos = InStr(1, Text1, "EnTeR c0d3:") 'Check 1
If ipos = "0" Then
Text1.Text = "Error: This is not a MBR or it's not infected"
Exit Sub
Else
ipos = InStr(1, Text1, "wExE") 'Check 2
If ipos = "0" Then
Text1.Text = "Error: This is not a MBR or it's not infected"
Exit Sub
Else
TotalChaine = ""
buffer = ""
Text1 = funcParseStringFromString2String(Text1, "EnTeR c0d3:", "wExE") 'Parsing
Text1 = Replace(Text1, " º€¹¸»Íê8Š‹}¬Š€ò", "") 'Remove useless shit (generic)
XorKey = Left$(Text1, 1)      'XOR Key
XorKey = Asc(XorKey) 'Get the Dec
strSubstr = Left$(Text1, 15) 'Hardcoded remove
Text1 = Replace(Text1, strSubstr, "") 'Remove hardcoded
For a = 1 To Len(Text1)           'Let's loop until the end
b = Asc(Mid(Text1, a, 1)) 'Grab one dec char
TotalChaine = XorKey Xor b 'unxor it
buffer = buffer & Chr(TotalChaine) 'make it char and add it to the final serial
Next
Text1 = buffer 'Final result
End If
End If
End Sub

Function funcParseStringFromString2String(sSourceString, sString1 As String, sString2 As String, Optional fCaseCaseInsensitive As Boolean = False) As String
 If fCaseCaseInsensitive Then
iCompareStyle = vbTextCompare
 Else
iCompareStyle = vbBinaryCompare
 End If
 
 sOutput = sSourceString
 iLocationOfString1 = InStr(1, sOutput, sString1, iCompareStyle)
 iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
 If iLocationOfString1 = 0 And iLocationOfString2 = 0 Then
'nothing found
sOutput = ""
 Else
If Len(sString1) = 0 And Len(sString2) = 0 Then
 'do nothing
ElseIf Len(sString1) = 0 Then
 If iLocationOfString2 <> 0 Then
sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
 End If
ElseIf Len(sString2) = 0 Then
 sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
Else
 'cut off begining
 If iLocationOfString1 <> 0 Then
sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
 End If
 'take off the end part
 iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
 If iLocationOfString2 <> 0 Then
sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
 End If
End If
 End If
 funcParseStringFromString2String = sOutput
End Function
