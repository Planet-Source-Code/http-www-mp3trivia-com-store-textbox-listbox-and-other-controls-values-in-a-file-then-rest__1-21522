VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Controls to save data for"
      Height          =   3525
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   9525
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   3060
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Top             =   2640
         Width           =   2325
      End
      Begin VB.TextBox Text2 
         Height          =   1395
         Left            =   6330
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form1.frx":0000
         Top             =   2070
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option Button3"
         Height          =   315
         Left            =   3210
         TabIndex        =   10
         Top             =   3090
         Width           =   1545
      End
      Begin VB.ListBox List1 
         Height          =   840
         ItemData        =   "Form1.frx":0068
         Left            =   3150
         List            =   "Form1.frx":007B
         TabIndex        =   8
         Top             =   1440
         Width           =   3105
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":00BA
         Left            =   6330
         List            =   "Form1.frx":00C7
         TabIndex        =   7
         Text            =   "ComboBox with text to save"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option Button2"
         Height          =   375
         Left            =   3210
         TabIndex        =   6
         Top             =   2700
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option Button1"
         Height          =   315
         Left            =   3210
         TabIndex        =   5
         Top             =   2370
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check Box to save State of"
         Height          =   465
         Left            =   150
         TabIndex        =   4
         Top             =   2190
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Text            =   "TextBox with Text to Save"
         Top             =   1440
         Width           =   2865
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":00E3
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   9345
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load This Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6390
      TabIndex        =   1
      Top             =   3690
      Width           =   3195
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save this form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   3750
      Width           =   2745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Sub cmdSave_Click()
 Module1.SaveFormState Form1
End Sub

Private Sub Command1_Click()
 Module1.LoadFormState Form1
End Sub

