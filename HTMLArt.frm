VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTMLArt"
   ClientHeight    =   5985
   ClientLeft      =   1635
   ClientTop       =   1545
   ClientWidth     =   8715
   Icon            =   "HTMLArt.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   8715
   Begin VB.PictureBox Picture2 
      ForeColor       =   &H00FF8040&
      Height          =   200
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   8415
      TabIndex        =   17
      Top             =   5760
      Width           =   8475
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "About HTMLArt"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   8040
      Picture         =   "HTMLArt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open text file"
      Top             =   1560
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   8040
      Picture         =   "HTMLArt.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open Image file"
      Top             =   600
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Make text upper case"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Convert text to all upper case?"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "File name for your creation"
      Top             =   4680
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      ToolTipText     =   "Title for your creation"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Text            =   "120"
      ToolTipText     =   "Y size"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Text            =   "180"
      ToolTipText     =   "X size"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make HTMLArt"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Create and save your HTMLArt"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawMode        =   6  'Mask Pen Not
      Height          =   5555
      Left            =   120
      ScaleHeight     =   5490
      ScaleWidth      =   4500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4555
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(none)"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      ToolTipText     =   "Selected Text file"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(none)"
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      ToolTipText     =   "Selected Image file"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Save file name:"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "HTML title"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "HTML width       HTML hight"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Text file:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Image file:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''
'''' HTMLArt                             ''''
'''' by Paul Bahlawan March 13, 2001     ''''
'''' Update: Sep 8, 2003
'''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private Type RGBType
    R As Byte
    G As Byte
    b As Byte
    Filler As Byte
End Type
Private Type RGBLongType
    clr As Long
End Type
Dim iFile As String
Dim tFile As String

'Main sub to create ART!
Private Sub Command1_Click()
Dim x As Long, y As Long, z As String, pchar As Byte
Dim oldColor As String

If Text4.Text = "" Or Label2.Caption = "(none)" Or tFile = "" Then
    MsgBox "You must select a valid image AND a valid text file before this command is allowed", vbExclamation, "HTMLArt - Error"
    Exit Sub
End If
If Val(Text1.Text) < 5 Or Val(Text2.Text) < 5 Then
    MsgBox "There seems to be a problem with the HTML width and/or hight settings", vbExclamation, "HTMLArt - Error"
    Exit Sub
End If
Open tFile For Binary Access Read As #2
If LOF(2) < 2 Then
    Close #2
    MsgBox "The selected text file is invalid", vbExclamation, "HTMLArt - Error"
    Exit Sub
End If

Open App.Path & "\" & Text4.Text For Output As #1
Print #1, "<HTML>"
Print #1, "<TITLE>" & Text3.Text & "</TITLE>"
Print #1, "<!-- This image was generated with HTMLArt by Paul Bahlawan -->"
Print #1, "<BODY BGCOLOR=""#000000"">"
Print #1, "<BASEFONT SIZE=""1""><PRE>"
    
    For y = 0 To Picture1.ScaleHeight Step Picture1.ScaleHeight / Val(Text2.Text)
        Picture2.Line (0, 0)-(y * (Picture2.ScaleWidth / Picture1.ScaleHeight), Picture2.ScaleHeight), , BF
        For x = 0 To (Picture1.ScaleWidth - 5) Step Int(Picture1.ScaleWidth / Val(Text1.Text))
            z = Process(x, y)
            If oldColor <> z Then  'same color as last time, then no need to put it again
                If x + y <> 0 Then Print #1, "</FONT>";
                Print #1, "<FONT COLOR=" & Chr$(34) & "#" & z & Chr$(34) & ">";
                oldColor = z
            End If
            Print #1, getChar;
        Next x
        Print #1, Chr$(13);
    Next y

Print #1, "</FONT></PRE></BODY></HTML>"
Close #2
Close #1
Picture2.Cls
End Sub

Function Process(x As Long, y As Long) As String
Process = ColorToHTML(Picture1.Point(x, y))
End Function

'Convert VB color to HTML color
Function ColorToHTML(clr As Long) As String
    Dim R As RGBType, RL As RGBLongType
    RL.clr = clr
    LSet R = RL
    Dim aR As String, aG As String, aB As String
    With R
        aR = Trim$(Hex$(.R)): If Len(aR) = 1 Then aR = "0" & aR
        aG = Trim$(Hex$(.G)): If Len(aG) = 1 Then aG = "0" & aG
        aB = Trim$(Hex$(.b)): If Len(aB) = 1 Then aB = "0" & aB
    End With
    ColorToHTML = aR & aG & aB
End Function

'Retrieve next byte from text file & convert special characters to HTML
Function getChar() As String
Dim a As String
a = "X"
Do
    Get #2, , a
    If EOF(2) Then Get #2, 1, a  'When we reach the end... start again!
    If a = Chr$(34) Then a = "&quot;": Exit Do
    If a = Chr$(38) Then a = "&amp;": Exit Do
    If a = Chr$(60) Then a = "&lt;": Exit Do
    If a = Chr$(62) Then a = "&gt;": Exit Do
    If a >= Chr$(33) And a <= Chr$(126) Then Exit Do
Loop
If Check1.Value Then a = UCase(a)
getChar = a
End Function

'Select Image file
Private Sub Command2_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg"
If iFile <> "" Then
    CommonDialog1.InitDir = iFile
    CommonDialog1.FileName = Label2.Caption
Else
    CommonDialog1.InitDir = ""
    CommonDialog1.FileName = ""
End If
CommonDialog1.ShowOpen
Label2.Caption = CommonDialog1.FileTitle
Call ScalePicture(Picture1, CommonDialog1.FileName)
Text3.Text = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
iFile = CommonDialog1.FileName
  
ErrHandler:
End Sub

'Select TEXT file
Private Sub Command3_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
CommonDialog1.Filter = "Text (*.txt)|*.txt"
If tFile <> "" Then
    CommonDialog1.InitDir = tFile
    CommonDialog1.FileName = Label3.Caption
Else
    CommonDialog1.InitDir = ""
    CommonDialog1.FileName = ""
End If
CommonDialog1.ShowOpen
Label3.Caption = CommonDialog1.FileTitle
tFile = CommonDialog1.FileName
  
ErrHandler:
End Sub

'ABOUT
Private Sub Command4_Click()
MsgBox "HTMLArt by Paul Bahlawan - March 13, 2001" & Chr$(13) & "Version " & App.Major & "." & App.Minor & " (build" & App.Revision & ")", vbInformation, "HTMLArt - About"
End Sub

'Guess save file name
Private Sub Text3_Change()
Text4.Text = Text3.Text & ".html"
End Sub

'Resize image to best fit
'This sub: Thanks to Frank Adam on comp.lang.basic.visual.misc for his assistance.
Private Sub ScalePicture(pb As PictureBox, strPicturePath As String)
Const xScale As Single = 0.566941199004022
Const yScale As Single = 0.566909725788231
Dim ScaleFactor As Single
Dim pic As StdPicture
Set pic = LoadPicture(strPicturePath)
If pic.Width >= pic.Height Then
    Picture1.Width = 4555
    With pb
        ScaleFactor = .Width / (pic.Width * xScale)
        .Height = pic.Height * yScale * ScaleFactor
        .PaintPicture pic, 0, 0, .Width, .Height, , , , , vbSrcCopy
    End With
Else
    Picture1.Height = 5555
    With pb
        ScaleFactor = .Height / (pic.Height * xScale)
        .Width = pic.Width * yScale * ScaleFactor
        .PaintPicture pic, 0, 0, .Width, .Height, , , , , vbSrcCopy
    End With
End If
Set pic = Nothing
End Sub

