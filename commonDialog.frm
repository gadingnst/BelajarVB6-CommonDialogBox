VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Common Dialog"
   ClientHeight    =   3315
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Menu menuDemo 
      Caption         =   "Menu"
      Begin VB.Menu menuOpen 
         Caption         =   "File Open"
      End
      Begin VB.Menu menuSave 
         Caption         =   "File Save"
      End
      Begin VB.Menu menuFont 
         Caption         =   "Font"
      End
      Begin VB.Menu menuColor 
         Caption         =   "Color"
      End
      Begin VB.Menu menuPrint 
         Caption         =   "Print Setup"
      End
      Begin VB.Menu menuSpc 
         Caption         =   ""
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearScr()
    Cls
    With Label1
        .BackColor = &H8000000F
        .Caption ""
        .Visible = False
    End With
End Sub

Private Sub Form_Load()

End Sub

Private Sub menuColor_Click()
    ClearScr
    CommonDialog11.ShowColor
    Label1.Visible = True
    Label1.BackColor = CommonDialog1.Color
    Print "Color"
    Print
    Print "Anda Memilih kode warna hexa : " + Hex$(CommonDialog1.Color)
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

Private Sub menuFont_Click()
    Dim ukuranFont As Long
    ClearScr
    CommonDialog1.Flags = &H1& Or &H100&
    CommonDialog1.Action = 4
    With Label1.Font
        .Name = CommonDialog1.FontName
        .Bold = CommonDialog1.FontBold
        .Italic = CommonDialog1.FontItalic
        .Underline = CommonDialog1.FontUnderline
    End With
    Label1.ForeColor = CommonDialog1.Color
    Label1.FontStrikethru = CommonDialog1.FontStrikethru
    ukuranFont = CommonDialog1.FontSize
    If ukuranFont > 48 Then
        Label1.Font.Size = 48
    Else
        Label1.Font.Size = CommonDialog1.FontSize
    End If
    Print "Font"
    Print
    Print "Anda Memilih FOnt : " + CommonDialog1.FontName + ", " + Str(ukuranFont)
    Label1.Visible = True
    Label1.Caption = "Contoh"
End Sub

Private Sub menuOpen_Click()
    ClearScr
    With CommonDialog1
        .Filter = "All Files (*.*)|*.*"
        .DialogTitle = "Buka File"
        .ShowOpen
    End With
    Print "File Open"
    Print
    Print "Anda Membuka FIle : " + CommonDialog1.FileName
End Sub

Private Sub menuPrint_Click()
    ClearScr
    CommonDialog1.ShowPrinter
    Print "Print Setup"
    Print
    Print "Anda Telah mengatur Setup Printer"
End Sub

Private Sub menuSave_Click()
    ClearScr
    With CommonDialog1
        .Filter = "All Files (*.*)|*.*"
        .DialogTitle = "Simpan File"
        .ShowSave
    End With
    Print "File Save"
    Print
    Print "Anda Menyimpan FIle : " + CommonDialog1.FileName
End Sub
