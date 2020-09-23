VERSION 5.00
Begin VB.Form frmBitMapMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "test"
      Begin VB.Menu mnuTest2 
         Caption         =   "Test2"
      End
   End
End
Attribute VB_Name = "frmBitMapMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MenuInfo As MenuBitmap



Private Sub Form_Load()
    Set MenuInfo = New MenuBitmap
    
    MenuInfo.AddBitmap Me.hwnd, 0, 0, App.Path & "\" & "Open.bmp"
    MenuInfo.AddBitmap Me.hwnd, 0, 1, App.Path & "\" & "Close.bmp"
    MenuInfo.AddBitmap Me.hwnd, 1, 0, App.Path & "\" & "1stboot.bmp"
    

End Sub
