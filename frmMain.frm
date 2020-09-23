VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0071584E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RGB Example"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00947061&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   3855
      TabIndex        =   6
      Top             =   120
      Width           =   3885
      Begin VB.HScrollBar Red 
         Height          =   255
         LargeChange     =   10
         Left            =   480
         Max             =   255
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.HScrollBar Green 
         Height          =   255
         LargeChange     =   10
         Left            =   480
         Max             =   255
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
      Begin VB.HScrollBar Blue 
         Height          =   255
         LargeChange     =   10
         Left            =   480
         Max             =   255
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox RText 
         Appearance      =   0  'Flat
         BackColor       =   &H00947061&
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox GText 
         Appearance      =   0  'Flat
         BackColor       =   &H00947061&
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3120
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox BText 
         Appearance      =   0  'Flat
         BackColor       =   &H00947061&
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox HexT 
         Appearance      =   0  'Flat
         BackColor       =   &H00947061&
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "#000000"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox VisB 
         Appearance      =   0  'Flat
         BackColor       =   &H00947061&
         ForeColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Text            =   "&H000000&"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Color 'Picker'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "G"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hex:"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "VB:"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label convHex 
         BackStyle       =   0  'Transparent
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   1590
         Width           =   735
      End
      Begin VB.Label ConvVB 
         BackStyle       =   0  'Transparent
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label copyHex 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1590
         Width           =   495
      End
      Begin VB.Label copyVB 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1950
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00947061&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2265
      ScaleWidth      =   3855
      TabIndex        =   1
      Top             =   3600
      Width           =   3885
      Begin VB.TextBox Text1 
         BackColor       =   &H00947061&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFC0C0&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmMain.frx":0000
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox pictureBox2 
      Appearance      =   0  'Flat
      BackColor       =   &H00947061&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   825
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   2640
      Width           =   3885
      Begin VB.PictureBox clrBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3615
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Color Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RGB Example
' Code created by Siebe Tolsma
'
' If you plan to use this code, please put my name somewhere in
' your app (e.g. in your About).
'
' This program is well commented so I hope you guys learn
' something from this one :)
'
' Comments and bugs can be mailed to: siebe-tolsma@home.nl

Private Sub convHex_Click()
    'In this nice little sub we convert a Hex color to an VB color and RGB.
    'Watch and learn ;) The VB color is easy since the Hex layout is the same,
    'the only thing different is the way it is 'presented'.
    
    If HexT.Text <> "" And HexT.Text <> "#" And HexT.Text <> Space(Len(HexT.Text)) Then
    
        'Hex to VB color..
        If Left$(HexT.Text, 1) <> "#" Then HexT.Text = "#" & HexT.Text
        VisB.Text = "&H" & Mid$(HexT.Text, 6, 2) & Mid$(HexT.Text, 4, 2) & Mid$(HexT.Text, 2, 2) & "&"
        
        'Decoding the hex to decimal isn't hard, you just have to know HOW...
        Dim R As Integer
        Dim G As Integer
        Dim B As Integer
        
        R = Val("&H" & Mid$(HexT.Text, 2, 2))
        G = Val("&H" & Mid$(HexT.Text, 4, 2))
        B = Val("&H" & Mid$(HexT.Text, 6, 2))
        
        Red.Value = R
        Green.Value = G
        Blue.Value = B
        
        RText.Text = R
        GText.Text = G
        BText.Text = B
        
        clrBox.BackColor = RGB(R, G, B)
    
    End If
    
End Sub

Private Sub ConvVB_Click()
    'In here we will convert the VB code in the textbox to an proper RGB and Hex code.

    If VisB.Text <> "" And VisB.Text <> "#" And VisB.Text <> Space(Len(VisB.Text)) Then
        
        'First the Hex code...
        If Left$(VisB.Text, 2) <> "&H" Then VisB.Text = "&H" & HexT.Text
        If Right$(VisB.Text, 1) <> "&" Then VisB.Text = VisB.Text & "&"
        HexT.Text = "#" & Mid$(VisB.Text, 7, 2) & Mid$(VisB.Text, 5, 2) & Mid$(VisB.Text, 3, 2)
        
        'Now just use the same method as from the Hex convert sub... :)
        Dim R As Integer
        Dim G As Integer
        Dim B As Integer
        
        R = Val("&H" & Mid$(HexT.Text, 2, 2))
        G = Val("&H" & Mid$(HexT.Text, 4, 2))
        B = Val("&H" & Mid$(HexT.Text, 6, 2))
        
        Red.Value = R
        Green.Value = G
        Blue.Value = B
        
        RText.Text = R
        GText.Text = G
        BText.Text = B
        
        clrBox.BackColor = RGB(R, G, B)
    
    End If
    
End Sub

Private Sub copyHex_Click()
    'Copy the hex code to the clipboard using 'Clipboard.Settext'
    'Also notify the user...
    Clipboard.SetText (HexT.Text)
    MsgBox "Hex code was succesfully copied to your clipboard!", vbOKOnly + vbInformation, "Copied"
End Sub

Private Sub copyVB_Click()
    'Copy the VB code to the clipboard using 'Clipboard.Settext'
    'Also notife the user...
    Clipboard.SetText (VisB.Text)
    MsgBox "VB color code was succesfully copied to your clipboard!", vbOKOnly + vbInformation, "Copied"
End Sub

Private Sub Green_Change()
    'Call the sub
    Call UpdateRGB
End Sub

Private Sub Blue_Change()
    'Call the sub
    Call UpdateRGB
End Sub

Private Sub Red_Change()
    'Call the sub
    Call UpdateRGB
End Sub

Private Sub UpdateRGB()
    Dim Color As Long
    Dim HexColor As String
    
    'This is the Sub in which the Hex and VB code get updated. Also
    'we will update the window background if bgSetColor has been checked
    'First fill the boxes
    RText.Text = Red.Value
    GText.Text = Green.Value
    BText.Text = Blue.Value
       
    'Let us color the picturebox!
    'Use the function 'RGB' (standard in VB) to create a colorcode
    'which VB can understand so we can color the clrBox...
    Color = RGB(Red.Value, Green.Value, Blue.Value)
    clrBox.BackColor = Color
    
    'Now here we calculate the Hex color. It's pretty neat,
    'if I may say so :) Oh and if you use this code, note that
    'the HEX code isn't hex(red) hex(green) hex(blue). You have to
    'swapp Green and Blue to get a proper HTML HEX color!
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    R = (Color And &HFF&) 'Red
    B = (Color And &HFF00&) / &H100& 'Blue
    G = (Color And &HFF0000) / &H10000 'Green
    HexT.Text = "#" & Right("0" & Hex(R), 2) & Right("0" & Hex(B), 2) & Right("0" & Hex(G), 2)
    
    'VB colors.... are about the same as Hex, only Hex is the other way around.
    'So we have to swap the colors to get a good code.
    VisB.Text = "&H" & Right("0" & Hex(G), 2) & Right("0" & Hex(B), 2) & Right("0" & Hex(R), 2) & "&"
    
End Sub


