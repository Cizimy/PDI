Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modWindowsAPI"

' ======================
' Windows API宣言
' ======================

' --- INIファイル操作 ---
Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long

' --- ファイル操作 ---
Public Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long

Public Declare PtrSafe Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" ( _
    ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

' --- タイマー操作 ---
Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

' --- ミューテックス操作 ---
Public Declare PtrSafe Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
    ByVal lpMutexAttributes As LongPtr, ByVal bInitialOwner As Long, _
    ByVal lpName As String) As LongPtr

Public Declare PtrSafe Function ReleaseMutex Lib "kernel32" ( _
    ByVal hMutex As LongPtr) As Long

Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Public Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

' --- 暗号化操作 ---
Public Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As LongPtr, ByVal pszContainer As String, _
    ByVal pszProvider As String, ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByVal Algid As Long, _
    ByVal hKey As LongPtr, ByVal dwFlags As Long, _
    ByRef phHash As LongPtr) As Long

Public Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr, ByRef pbData As Any, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr, ByVal dwParam As Long, _
    ByRef pbData As Any, ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Public Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr) As Long

Public Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long

' --- パフォーマンスカウンター ---
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency) As Long

Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As Currency) As Long

Public Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" ( _
    ByVal Process As LongPtr, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
    ByVal cb As Long) As Long

Public Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr

' ======================
' 定数定義
' ======================

' --- ファイル属性 ---
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const INVALID_FILE_ATTRIBUTES As Long = -1

' --- タイマー関連 ---
Public Const INFINITE As Long = -1
Public Const WAIT_OBJECT_0 As Long = 0

' --- 暗号化関連 ---
Public Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const PROV_RSA_FULL As Long = 1
Public Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Public Const CALG_SHA_256 As Long = &H800C
Public Const HP_HASHVAL As Long = 2
Public Const HP_HASHSIZE As Long = 4

' ======================
' 型定義
' ======================
Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Currency
    WorkingSetSize As Currency
    QuotaPeakPagedPoolUsage As Currency
    QuotaPagedPoolUsage As Currency
    QuotaPeakNonPagedPoolUsage As Currency
    QuotaNonPagedPoolUsage As Currency
    PagefileUsage As Currency
    PeakPagefileUsage As Currency
End Type

' ======================
' エラーコードマッピング
' ======================
Public Function MapWindowsErrorToAppError(ByVal windowsError As Long) As ErrorCode
    Select Case windowsError
        ' ファイル操作エラー
        Case 2, 3 ' ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND
            MapWindowsErrorToAppError = ErrFileNotFound
        Case 5 ' ERROR_ACCESS_DENIED
            MapWindowsErrorToAppError = ErrFileAccessDenied
        Case 32 ' ERROR_SHARING_VIOLATION
            MapWindowsErrorToAppError = ErrFileAccessDenied
            
        ' メモリ関連エラー
        Case 8, 14 ' ERROR_NOT_ENOUGH_MEMORY, ERROR_OUTOFMEMORY
            MapWindowsErrorToAppError = ErrSystemOutOfMemory
            
        ' その他のシステムエラー
        Case Else
            MapWindowsErrorToAppError = ErrUnexpected
    End Select
End Function

' ======================
' ユーティリティ関数
' ======================
Public Function GetLastWindowsError() As Long
    #If Win64 Then
        GetLastWindowsError = CreateObject("WScript.Shell").Environment("PROCESS")("ERROR_CODE")
    #Else
        GetLastWindowsError = Err.LastDllError
    #End If
End Function

Public Function IsValidHandle(ByVal handle As LongPtr) As Boolean
    IsValidHandle = (handle <> 0)
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function TestAPIAvailability() As Boolean
        Dim result As Boolean
        result = True
        
        ' 基本的なAPI機能のテスト
        Dim counter As Currency
        result = result And (QueryPerformanceCounter(counter) <> 0)
        
        ' ファイル操作APIのテスト
        Dim attr As Long
        attr = GetFileAttributes("C:\")
        result = result And (attr <> INVALID_FILE_ATTRIBUTES)
        
        TestAPIAvailability = result
    End Function
#End If