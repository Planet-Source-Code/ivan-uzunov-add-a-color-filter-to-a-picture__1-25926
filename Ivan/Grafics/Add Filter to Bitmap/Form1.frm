VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Yellow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Black and White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1995
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   122
      TabIndex        =   0
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Now restart the program and add next filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvert_Click(Index As Integer)
    Call AddFilter(picSource, Index)
    Label1.Visible = True
End Sub
