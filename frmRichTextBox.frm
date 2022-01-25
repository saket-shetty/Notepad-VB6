VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmRichTextBox 
   Caption         =   "Untitled - Notepad"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H0000FF00&
      Caption         =   "Color"
      Height          =   360
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   990
   End
   Begin VB.ComboBox ComFontSize 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
   Begin VB.ComboBox comFonts 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Frame FraCustomTools 
      Caption         =   "CustomTools"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   3495
      Begin VB.CommandButton cmdU 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdI 
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   990
      End
      Begin VB.CommandButton cmdB 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   990
      End
   End
   Begin RichTextLib.RichTextBox rtText 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRichTextBox.frx":0000
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu MnuNewWindow 
         Caption         =   "New Window"
      End
      Begin VB.Menu Open 
         Caption         =   "Open..."
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu Seperate 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu MnuFormal 
      Caption         =   "F&ormat"
      Begin VB.Menu MnuWrap 
         Caption         =   "Word Wrap"
      End
      Begin VB.Menu MnuFont 
         Caption         =   "Font..."
      End
   End
End
Attribute VB_Name = "frmRichTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdB_Click()
    rtText.SelBold = True
End Sub

Private Sub cmdColor_Click()
    CommonDialog1.ShowColor
    rtText.SelColor = CommonDialog1.Color
End Sub

Private Sub cmdI_Click()
    rtText.SelItalic = True
End Sub

Private Sub cmdU_Click()
    rtText.SelUnderline = True
End Sub

Private Sub comFonts_Click()
    rtText.SelFontName = comFonts.Text
End Sub

Private Sub ComFontSize_Click()
    rtText.SelFontSize = ComFontSize.Text
End Sub

Private Sub Form_Load()
    
    For i = 0 To Screen.FontCount
'        comFonts.AddItem
        comFonts.AddItem Screen.Fonts(i)
    Next
    
    For j = 1 To 38
        ComFontSize.AddItem j
        
    Next
    
End Sub

Private Sub MnuNewWindow_Click()
    CommonDialog1.ShowOpen
End Sub

Private Sub MnuPrint_Click()
    CommonDialog1.ShowPrinter
End Sub

Private Sub MnuSaveAs_Click()
    CommonDialog1.ShowSave
End Sub

Private Sub New_Click()
    CommonDialog1.ShowOpen
End Sub

Private Sub Open_Click()
    CommonDialog1.ShowOpen
End Sub

Private Sub rtText_Change()
    If rtText.Text = "" Or rtText.Text = Empty Then
        frmRichTextBox.Caption = "Untitled - Notepad"
    Else
        frmRichTextBox.Caption = "*Untitled - Notepad"
    End If
End Sub

Private Sub Save_Click()
    CommonDialog1.ShowSave
End Sub
