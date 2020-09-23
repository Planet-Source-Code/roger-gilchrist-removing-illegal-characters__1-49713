VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "StripIllegalChars Demo"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMaxLength 
      Caption         =   "Max Length"
      Height          =   855
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   2415
      Begin VB.PictureBox picCFXPBugFix1 
         BorderStyle     =   0  'None
         Height          =   605
         Left            =   100
         ScaleHeight     =   600
         ScaleWidth      =   2220
         TabIndex        =   6
         Top             =   175
         Width           =   2215
         Begin VB.OptionButton optLength 
            Caption         =   "255(Variables/Filename)"
            Height          =   195
            Index           =   1
            Left            =   20
            TabIndex        =   8
            Top             =   280
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optLength 
            Caption         =   "40 (controls)"
            Height          =   195
            Index           =   0
            Left            =   20
            TabIndex        =   7
            Top             =   40
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraSubstituteChar 
      Caption         =   "Substitute Char"
      Height          =   855
      Left            =   3180
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      Begin VB.PictureBox picCFXPBugFix2 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   100
         ScaleHeight     =   480
         ScaleWidth      =   1500
         TabIndex        =   9
         Top             =   175
         Width           =   1495
         Begin VB.OptionButton optSubChar 
            Caption         =   "none"
            Height          =   195
            Index           =   0
            Left            =   20
            TabIndex        =   11
            Top             =   40
            Width           =   1455
         End
         Begin VB.OptionButton optSubChar 
            Caption         =   "Underscores"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdStripPunct 
      Caption         =   "Do It"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtOutput 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "IllegalChar.frx":0000
      Top             =   1920
      Width           =   8055
   End
   Begin VB.TextBox txtInput 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label lblDescription 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description"
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStripPunct_Click()

  Dim strSub As String

  Select Case True
   Case optSubChar(0)
    strSub = vbNullString
   Case optSubChar(1)
    strSub = "_"
  End Select
  txtOutput.Text = StripIllegalChars(txtInput.Text, strSub, optLength(0))

End Sub

Private Sub Form_Load()

  lblDescription.Caption = "This is a simple demo of a few routines I use to generate legal VB control and variable names. It could also be used to check that file names are legal(although 'As Is' it deletes the tilde '~' of short filenames (but who still uses those :) )." & vbNewLine & _
                          "the rules are:" & vbNewLine & _
                          "1. No punctuation except underscore '_'." & vbNewLine & _
                          "2. No numerals or underscores at start of name." & vbNewLine & _
                          "3. Max length for controls = 40 and for variables and filenames(without extention) = 255 (not recommended for variables)" & vbNewLine & _
                          vbNewLine & _
                          "For readability reasons the code also makes the name ProperCase." & vbNewLine & _
                          "See code for explanation of how to overcome problem of using Replace to remove multiples of the same character." & vbNewLine & _
                          "Check out the functions RStrip and LStrip(not used in this prog) which are like RTrim and LTrim for the rest of the characters." & vbNewLine & _
                          "Also check out function IsAlphaIntl for a quick and internationalised letter tester." & vbNewLine & _
                          "None of this is original to me (I've used them so long I'm not sure where I got them( probably VBPJ)), except for the way they are combining in StripIllegalChars."


  txtInput.Text = "___ 123 this string wouldn't make a god name! This is because of all the ~!@#$%^&*()_+=-{}|[]\:"";'<>?,./ punctuation ( see rules below ). These numbers and underscores are OK 1_2_324 unlike ones at start of string."

End Sub


':)Roja's VB Code Fixer V1.1.49 (8/11/2003 3:24:37 PM) 1 + 27 = 28 Lines Thanks Ulli for inspiration and lots of code.

