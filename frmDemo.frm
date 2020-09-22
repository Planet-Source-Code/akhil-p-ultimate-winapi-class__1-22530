VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SilverSoft WINAPI Class Demonstrator by Akhil"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send E-Mail to Akhil P (akhiljayaraj@hotmail.com)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2025
      Width           =   5040
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   2640
      TabIndex        =   1
      Top             =   135
      Width           =   2520
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find all files in C:\"
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   2355
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsWinApi As CWinAPI 'Declare a class variable

Private Sub cmdFind_Click()
Dim FileCount As Integer, DirCount As Integer
clsWinApi.FindFilesAPI "c:\", "*.*", FileCount, DirCount, List1.hWnd, False, True
MsgBox "No. of Files Found: " & FileCount, , "WINAPI Class Demonstrator"
End Sub

Private Sub Command1_Click()
clsWinApi.SendEmail "akhiljayaraj@hotmail.com", Me.hWnd
End Sub

Private Sub Form_Load()
Set clsWinApi = New CWinAPI 'This makes a new instance of the class
MsgBox "This program just demonstrates two functions that can be done with the CWinApi class and how to use it. You will be able to use other functions with no difficulties. The CWinApi class will be very useful for any programmer.", vbApplicationModal + vbInformation + vbOKOnly, "WINAPI class Demonstrator"
End Sub
