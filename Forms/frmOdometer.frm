VERSION 5.00
Begin VB.Form frmOdometer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odometer"
   ClientHeight    =   852
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   2532
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOdometer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   StartUpPosition =   2  'CenterScreen
   Begin prjOdometer.Odometer odmCounter 
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2064
      _ExtentX        =   3641
      _ExtentY        =   656
      Digits          =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrCounter 
      Interval        =   500
      Left            =   1920
      Top             =   720
   End
End
Attribute VB_Name = "frmOdometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   Show
   DoEvents
   odmCounter.Value = 140530400

End Sub

Private Sub tmrCounter_Timer()

Static intValue As Integer

   With odmCounter
      If .Value = 140530400 Then
         intValue = 1
         
      ElseIf .Value = 140530415 Then
         intValue = -1
      End If
      
      .Value = .Value + Val(intValue)
   End With

End Sub

