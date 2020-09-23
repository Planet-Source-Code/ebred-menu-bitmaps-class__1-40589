VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" _
    (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
    ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const IMAGE_BITMAP = 0
Private Const LR_LOADFROMFILE = &H10
Private Const MF_BYPOSITION = &H400&

Private m_lngMenuHwnd As Long
Private m_SubMenuHandle As Long
Private m_BitmapHandle As Long
Private lngRet As Long


'to pic the fisrt submenu in the second menu then
'SubMenuLocation should be passed = 0 and MenuNumber = 1
'for each menu all submenus start at 0
Public Sub AddBitmap(FormHandle As Long, MenuNumber As Long, _
    SubMenuLocation As Long, BitmapLocation As String)

    'stores a handle for the bitmap.  also resizes it in pixels ie 14,13
    'ive found these sizes look best in the menu.  Ive also changed the _
    image sizes themselves to 14 by 13 pixels
    m_BitmapHandle = LoadImage(0, BitmapLocation, IMAGE_BITMAP, 14, 13, _
        LR_LOADFROMFILE)
   
    'Gets handles of the Menus which are contained in Form1
    m_lngMenuHwnd = GetMenu(FormHandle)

    'Gets handles of sub-menus within the Menu.
    m_SubMenuHandle = GetSubMenu(m_lngMenuHwnd, MenuNumber)
    
    'Enters the appropriate bitmaps to the assigned locations
    lngRet = SetMenuItemBitmaps(m_SubMenuHandle, SubMenuLocation, MF_BYPOSITION, m_BitmapHandle, m_BitmapHandle)

End Sub








