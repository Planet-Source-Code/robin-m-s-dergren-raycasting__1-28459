VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RayCasting"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDummy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1785
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   4
      Top             =   3240
      Width           =   390
      Visible         =   0   'False
   End
   Begin VB.PictureBox picMAP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   120
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   3150
      Width           =   1500
      Visible         =   0   'False
   End
   Begin VB.PictureBox picView 
      Height          =   3000
      Left            =   120
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Top             =   105
      Width           =   4800
   End
   Begin VB.PictureBox picBB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   345
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   -120
      Width           =   4800
      Visible         =   0   'False
   End
   Begin VB.PictureBox picMAP_BB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   585
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   3
      Top             =   3300
      Width           =   1500
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
