Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modWindowsAPI"

' ======================
' Windows API宣言（レガシーサポート用）
' ======================
' 従来のAPI宣言は維持しますが、新規コードでは非推奨です。
' 代わりにインターフェースベースの実装を使用してください。
#If LegacySupport Then
    ' --- INIファイル操作 ---
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long, ByVal lpFileName As String) As Long
    ' ... (その他のAPI宣言)
#End If

' ======================
' プライベート変数
' ======================
Private mConverter As ModWindowsAPIConverter
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mConverter = New ModWindowsAPIConverter
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    Set mConverter = Nothing
    mIsInitialized = False
End Sub

' ======================
' パブリック関数
' ======================

' --- ミューテックス操作 ---
Public Function CreateMutex(ByVal lpMutexAttributes As LongPtr, _
                          ByVal bInitialOwner As Long, _
                          ByVal lpName As String) As LongPtr
    InitializeIfNeeded
    
    Dim mutex As IMutex
    Set mutex = mConverter.Mutex
    
    If mutex.CreateMutex(bInitialOwner <> 0, lpName) Then
        CreateMutex = GetHandleFromMutex(mutex)
    End If
End Function

Public Function ReleaseMutex(ByVal hMutex As LongPtr) As Long
    InitializeIfNeeded
    
    Dim mutex As IMutex
    Set mutex = mConverter.Mutex
    
    ReleaseMutex = IIf(mutex.ReleaseMutex(), 1, 0)
End Function

Public Function WaitForSingleObject(ByVal hHandle As LongPtr, _
                                  ByVal dwMilliseconds As Long) As Long
    InitializeIfNeeded
    
    Dim mutex As IMutex
    Set mutex = mConverter.Mutex
    
    WaitForSingleObject = IIf(mutex.WaitForSingleObject(dwMilliseconds), 0, &HFFFFFFFF)
End Function

' --- 暗号化操作 ---
Public Function CryptAcquireContext(ByRef phProv As LongPtr, _
                                  ByVal pszContainer As String, _
                                  ByVal pszProvider As String, _
                                  ByVal dwProvType As Long, _
                                  ByVal dwFlags As Long) As Long
    InitializeIfNeeded
    
    Dim crypto As ICryptography
    Set crypto = mConverter.Crypto
    
    CryptAcquireContext = IIf(crypto.CryptAcquireContext(pszContainer, pszProvider, dwProvType, dwFlags), 1, 0)
End Function

' ... (他の暗号化関数も同様にインターフェース経由に変更)

' --- INIファイル操作 ---
Public Function GetPrivateProfileString(ByVal lpApplicationName As String, _
                                      ByVal lpKeyName As Any, _
                                      ByVal lpDefault As String, _
                                      ByVal lpReturnedString As String, _
                                      ByVal nSize As Long, _
                                      ByVal lpFileName As String) As Long
    InitializeIfNeeded
    
    Dim iniFile As IIniFile
    Set iniFile = mConverter.IniFile
    
    Dim result As String
    result = iniFile.GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpFileName)
    
    If Len(result) > 0 Then
        If Len(result) > nSize - 1 Then result = Left$(result, nSize - 1)
        Mid$(lpReturnedString, 1, Len(result)) = result
        GetPrivateProfileString = Len(result)
    End If
End Function

' --- パフォーマンスカウンター ---
Public Function QueryPerformanceCounter(ByRef lpPerformanceCount As Currency) As Long
    InitializeIfNeeded
    
    Dim perfCounter As IPerformanceCounter
    Set perfCounter = mConverter.PerformanceCounter
    
    QueryPerformanceCounter = IIf(perfCounter.QueryPerformanceCounter(lpPerformanceCount), 1, 0)
End Function

Public Function QueryPerformanceFrequency(ByRef lpFrequency As Currency) As Long
    InitializeIfNeeded
    
    Dim perfCounter As IPerformanceCounter
    Set perfCounter = mConverter.PerformanceCounter
    
    QueryPerformanceFrequency = IIf(perfCounter.QueryPerformanceFrequency(lpFrequency), 1, 0)
End Function

' --- スリープ操作 ---
Public Sub Sleep(ByVal dwMilliseconds As Long)
    InitializeIfNeeded
    
    Dim sleeper As ISleep
    Set sleeper = mConverter.Sleep
    
    sleeper.Sleep dwMilliseconds
End Sub

' ======================
' プライベート関数
' ======================
Private Sub InitializeIfNeeded()
    If Not mIsInitialized Then InitializeModule
End Sub

Private Function GetHandleFromMutex(ByVal mutex As IMutex) As LongPtr
    ' 実装クラス固有のハンドル取得
    If TypeOf mutex Is MutexImpl Then
        GetHandleFromMutex = DirectCast(mutex, MutexImpl).GetMutexHandle()
    End If
End Function

' ======================
' エラー処理
' ======================
Public Function MapWindowsErrorToAppError(ByVal windowsError As Long) As ErrorCode
    Select Case windowsError
        Case 2, 3 ' ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND
            MapWindowsErrorToAppError = ErrFileNotFound
        Case 5 ' ERROR_ACCESS_DENIED
            MapWindowsErrorToAppError = ErrFileAccessDenied
        Case 32 ' ERROR_SHARING_VIOLATION
            MapWindowsErrorToAppError = ErrFileAccessDenied
        Case 8, 14 ' ERROR_NOT_ENOUGH_MEMORY, ERROR_OUTOFMEMORY
            MapWindowsErrorToAppError = ErrSystemOutOfMemory
        Case Else
            MapWindowsErrorToAppError = ErrUnexpected
    End Select
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function TestAPIAvailability() As Boolean
        InitializeIfNeeded
        
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
    
    Public Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
#End If