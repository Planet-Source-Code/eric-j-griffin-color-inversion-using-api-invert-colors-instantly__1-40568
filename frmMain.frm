VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Invert using API  by Eric J. Griffin"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4710
      Width           =   945
   End
   Begin VB.PictureBox picInvert 
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   60
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   60
      Width           =   6060
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##################################################################
'##
'##  Color Inversion with API
'##     by Eric J. Griffin
'##        eric@coderpost.com
'##        http://www.coderpost.com
'##
'##  This code is very short but does the trick much faster than
'##  using vb's inversion methods. Hope you like!
'##
'##################################################################


Private Sub cmdInvert_Click()
    Dim rval&, rctArea As RECT
    
    '## First, set the area you want to invert.
    '## We want to invert everything in the picture box so...
    rval& = SetRect(rctArea, 0, 0, picInvert.Width, picInvert.Height)
    
    '## Second, call the InvertRect function that will do our inversion
    rval& = InvertRect(picInvert.hdc, rctArea)
End Sub
