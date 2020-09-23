VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Color"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"richtextboxcolor.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "You select the text then you press ""Color"" then you pick your color."
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is for newbies, vote if u like
'By Mr.PoO
Private Sub Command1_Click()
'this brings up the color dialog box
CommonDialog1.ShowColor
'when you select the text, this is the code which will change it into
'color
RichTextBox1.SelColor = CommonDialog1.Color
End Sub
