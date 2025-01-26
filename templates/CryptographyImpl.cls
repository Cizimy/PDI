VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CryptographyImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICryptography

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "CryptographyImpl"

' ======================
' メンバ変数
' ======================
Private mCryptoProvider As LongPtr
Private mHash As LongPtr
Private mKey As LongPtr
Private mIsInitialized As Boolean
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mCryptoProvider = 0
    mHash = 0
    mKey = 0
    Set mPerformanceMonitor = New clsPerformanceMonitor
    Set mLock = New clsLock
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    CleanupResources
    Set mPerformanceMonitor = Nothing
    Set mLock = Nothing
    mIsInitialized = False
End Sub

' ======================
' ICryptography実装
' ======================
Private Function ICryptography_CryptAcquireContext(ByVal container As String, _
                                                 ByVal provider As String, _
                                                 ByVal provType As Long, _
                                                 ByVal flags As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptAcquireContext"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' 既存のプロバイダをクリーンアップ
    If mCryptoProvider <> 0 Then
        modWindowsAPI.CryptReleaseContext mCryptoProvider, 0
        mCryptoProvider = 0
    End If
    
    ' 新しいプロバイダを取得
    Dim result As Boolean
    result = (modWindowsAPI.CryptAcquireContext(mCryptoProvider, container, provider, provType, flags) <> 0)
    
    ICryptography_CryptAcquireContext = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptAcquireContext"
    Exit Function

ErrorHandler:
    LogError "暗号化プロバイダの取得中にエラーが発生しました: " & Err.Description
    ICryptography_CryptAcquireContext = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptCreateHash(ByVal algorithm As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mCryptoProvider = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptCreateHash"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' 既存のハッシュをクリーンアップ
    If mHash <> 0 Then
        modWindowsAPI.CryptDestroyHash mHash
        mHash = 0
    End If
    
    ' 新しいハッシュを作成
    Dim result As Boolean
    result = (modWindowsAPI.CryptCreateHash(mCryptoProvider, algorithm, 0, 0, mHash) <> 0)
    
    ICryptography_CryptCreateHash = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptCreateHash"
    Exit Function

ErrorHandler:
    LogError "ハッシュの作成中にエラーが発生しました: " & Err.Description
    ICryptography_CryptCreateHash = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptHashData(ByRef data As Any, _
                                           ByVal dataLen As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mHash = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptHashData"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptHashData(mHash, data, dataLen, 0) <> 0)
    
    ICryptography_CryptHashData = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptHashData"
    Exit Function

ErrorHandler:
    LogError "データのハッシュ化中にエラーが発生しました: " & Err.Description
    ICryptography_CryptHashData = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptDeriveKey(ByVal algorithm As Long, _
                                            ByVal flags As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mHash = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptDeriveKey"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' 既存の鍵をクリーンアップ
    If mKey <> 0 Then
        modWindowsAPI.CryptDestroyKey mKey
        mKey = 0
    End If
    
    ' 新しい鍵を生成
    Dim result As Boolean
    result = (modWindowsAPI.CryptDeriveKey(mCryptoProvider, algorithm, mHash, flags, mKey) <> 0)
    
    ICryptography_CryptDeriveKey = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptDeriveKey"
    Exit Function

ErrorHandler:
    LogError "鍵の生成中にエラーが発生しました: " & Err.Description
    ICryptography_CryptDeriveKey = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptEncrypt(ByRef data As Any, _
                                          ByRef dataLen As Long, _
                                          ByVal bufLen As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mKey = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptEncrypt"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptEncrypt(mKey, 0, True, 0, data, dataLen, bufLen) <> 0)
    
    ICryptography_CryptEncrypt = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptEncrypt"
    Exit Function

ErrorHandler:
    LogError "データの暗号化中にエラーが発生しました: " & Err.Description
    ICryptography_CryptEncrypt = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptDecrypt(ByRef data As Any, _
                                          ByRef dataLen As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mKey = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptDecrypt"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptDecrypt(mKey, 0, True, 0, data, dataLen) <> 0)
    
    ICryptography_CryptDecrypt = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptDecrypt"
    Exit Function

ErrorHandler:
    LogError "データの復号化中にエラーが発生しました: " & Err.Description
    ICryptography_CryptDecrypt = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptDestroyKey() As Boolean
    If Not mIsInitialized Then Exit Function
    If mKey = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptDestroyKey"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptDestroyKey(mKey) <> 0)
    If result Then mKey = 0
    
    ICryptography_CryptDestroyKey = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptDestroyKey"
    Exit Function

ErrorHandler:
    LogError "鍵の破棄中にエラーが発生しました: " & Err.Description
    ICryptography_CryptDestroyKey = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptDestroyHash() As Boolean
    If Not mIsInitialized Then Exit Function
    If mHash = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptDestroyHash"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptDestroyHash(mHash) <> 0)
    If result Then mHash = 0
    
    ICryptography_CryptDestroyHash = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptDestroyHash"
    Exit Function

ErrorHandler:
    LogError "ハッシュの破棄中にエラーが発生しました: " & Err.Description
    ICryptography_CryptDestroyHash = False
    Resume CleanUp
End Function

Private Function ICryptography_CryptReleaseContext() As Boolean
    If Not mIsInitialized Then Exit Function
    If mCryptoProvider = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CryptReleaseContext"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    result = (modWindowsAPI.CryptReleaseContext(mCryptoProvider, 0) <> 0)
    If result Then mCryptoProvider = 0
    
    ICryptography_CryptReleaseContext = result
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "CryptReleaseContext"
    Exit Function

ErrorHandler:
    LogError "暗号化コンテキストの解放中にエラーが発生しました: " & Err.Description
    ICryptography_CryptReleaseContext = False
    Resume CleanUp
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub CleanupResources()
    If mKey <> 0 Then
        modWindowsAPI.CryptDestroyKey mKey
        mKey = 0
    End If
    
    If mHash <> 0 Then
        modWindowsAPI.CryptDestroyHash mHash
        mHash = 0
    End If
    
    If mCryptoProvider <> 0 Then
        modWindowsAPI.CryptReleaseContext mCryptoProvider, 0
        mCryptoProvider = 0
    End If
End Sub

Private Sub LogError(ByVal message As String)
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrUnexpected
        .Description = message
        .Category = ECSecurity
        .Source = MODULE_NAME
        .ProcedureName = "Unknown"  ' スタックトレースから取得可能
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = mIsInitialized And _
                       Not (mPerformanceMonitor Is Nothing) And _
                       Not (mLock Is Nothing)
    End Function
    
    Public Sub ForceCleanup()
        CleanupResources
    End Sub
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
#End If