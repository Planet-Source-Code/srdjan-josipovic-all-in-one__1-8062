VERSION 5.00
Begin VB.Form frmMenu1 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graphic"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   2520
      Picture         =   "Form2.frx":48526
      ScaleHeight     =   2370
      ScaleWidth      =   1470
      TabIndex        =   0
      Top             =   4440
      Width           =   1470
   End
   Begin VB.Image imgYield 
      Height          =   240
      Left            =   5280
      Picture         =   "Form2.frx":4948B
      Top             =   6360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgStop 
      Height          =   240
      Left            =   4320
      Picture         =   "Form2.frx":495CD
      Top             =   6360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgDelete 
      Height          =   240
      Left            =   6240
      Picture         =   "Form2.frx":4970F
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgCaution 
      Height          =   240
      Left            =   4800
      Picture         =   "Form2.frx":49851
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   240
      Left            =   5760
      Picture         =   "Form2.frx":49993
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFirst 
      Caption         =   "&First"
      Begin VB.Menu mnuSecond 
         Caption         =   "Second"
      End
      Begin VB.Menu mnuThird 
         Caption         =   "Third"
         Begin VB.Menu mnu1 
            Caption         =   "&1"
         End
         Begin VB.Menu mnu2 
            Caption         =   "&2"
         End
         Begin VB.Menu mnu3 
            Caption         =   "&3"
         End
      End
   End
End
Attribute VB_Name = "frmMenu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Private Sub Form_Load()
    ' Set the menu bitmaps.
    SetMenuBitmap Me, Array(0, 0), imgExit.Picture
    SetMenuBitmap Me, Array(1, 0), imgDelete.Picture
    SetMenuBitmap Me, Array(1, 1, 0), imgStop.Picture
    SetMenuBitmap Me, Array(1, 1, 1), imgYield.Picture
    SetMenuBitmap Me, Array(1, 1, 2), imgCaution.Picture
End Sub
' Put a bitmap in a menu item.
Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    ' Get the menu handle.
    menu_handle = GetMenu(frm.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    ' Initialize the menu information.
    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With

    ' Assign the picture.
    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub

