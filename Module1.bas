Attribute VB_Name = "modMisc"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Exist As Boolean
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&
Public installdate As String, grphx As Integer
Public dlcolor As Long, ulcolor As Long, bgcolor As Long
Public ULtot As Long, DLtot As Long, what As String, conspeed As Long, ontop As Boolean
Public Sub StayOnTop(Frm As Form, x As Long, y As Long, width As Long, height As Long)
    setontop = SetWindowPos(Frm.hWnd, -1, x / 15, y / 15, width / 15, height / 15, flags)
End Sub
Public Sub removefromtop(Frm As Form, x As Long, y As Long, width As Long, height As Long)
    setontop = SetWindowPos(Frm.hWnd, -2, x / 15, y / 15, width / 15, height / 15, flags)
End Sub
Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
    Dim curstyle As Long
    Dim newstyle As Long
    curstyle = GetWindowLong(NumberText.hWnd, GWL_STYLE)
    If Flag Then
        curstyle = curstyle Or ES_NUMBER
    Else
        curstyle = curstyle And (Not ES_NUMBER)
    End If
    newstyle = SetWindowLong(NumberText.hWnd, GWL_STYLE, curstyle)
    NumberText.refresh
End Sub

Public Sub SetStringValue(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim i As Long
    i = RegCreateKey(Hkey, strPath, keyhand)
    i = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    i = RegCloseKey(keyhand)
End Sub
Public Sub DelStringValue(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim i As Long
    i = RegOpenKey(Hkey, strPath, keyhand)
    i = RegDeleteValue(keyhand, strValue)
    i = RegCloseKey(keyhand)
End Sub
