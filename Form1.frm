VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text --> DOS Executable"
   ClientHeight    =   3525
   ClientLeft      =   2475
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Compile to DOS Executable"
      Height          =   825
      Left            =   1350
      TabIndex        =   1
      Top             =   2475
      Width           =   4380
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   945
      Width           =   6180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form1.frx":0010
      Height          =   915
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   6540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function ChrStr(sAscVals As String) As String
' Like the CHR() function, but accepts a string argument
' that contains ASCII values and returns a string

Dim vTmp As Variant
Dim idx As Integer

vTmp = Split(sAscVals, ",")
For idx = LBound(vTmp) To UBound(vTmp)
        vTmp(idx) = Chr$(vTmp(idx))
Next idx

ChrStr = Join(vTmp, "")
End Function


Function Chrs(ParamArray vAscValArray() As Variant) As String
' Like the CHR() function, but accepts many arguments
' and returns a string

Dim idx As Integer
ReDim vTmp(LBound(vAscValArray) To UBound(vAscValArray))

For idx = LBound(vAscValArray) To UBound(vAscValArray)
        vTmp(idx) = Chr$(vAscValArray(idx))
Next idx

Chrs = Join(vTmp, "")
End Function



Private Sub Command1_Click()
Dim a$

a = Chrs(180, 9, 186, 11, 1, 205, 33, 180, 76, 205, 33)
a = a & CStr(Text1.Text) & "$"

Open "coco.com" For Binary As #1
Put #1, , a
Close

End Sub


