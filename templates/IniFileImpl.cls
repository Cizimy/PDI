VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFileImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IIniFile

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "IniFileImpl"

' ======================
' 定数定義
' ======================
Private Const MAX_BUFFER_SIZE As Long = 32767 ' Win32 APIの制限
Private Const DEFAULT_BUFFER_SIZE As Long = 1024

' ======================
' メンバ変数
' ======================
Private mIsInitialized As Boolean
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mPerformanceMonitor = New clsPerformanceMonitor
    Set mLock = New clsLock
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mPerformanceMonitor = Nothing
    Set mLock = Nothing
    mIsInitialized = False
End Sub

' ======================
' IIniFile実装
' ======================
Private Function IIniFile_GetPrivateProfileString(ByVal section As String, _
                                                ByVal key As String, _
                                                ByVal defaultValue As String, _
                                                ByVal filePath As String) As String
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "GetPrivateProfileString"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim buffer As String
    Dim result As Long
    Dim retryCount As Long
    Const MAX_RETRIES As Long = 3
    
    ' 徐々にバッファサイズを大きくしながら値の取得を試みる
    Dim bufferSize As Long
    bufferSize = DEFAULT_BUFFER_SIZE
    
    Do
        buffer = String$(bufferSize, 0)
        result = modWindowsAPI.GetPrivateProfileString(section, key, defaultValue, _
                                                     buffer, bufferSize, filePath)
        
        ' バッファが不足している場合は拡張して再試行
        If result = bufferSize - 1 Then
            bufferSize = bufferSize * 2
            If bufferSize > MAX_BUFFER_SIZE Then
                LogError "バッファサイズが最大値を超えました。"
                IIniFile_GetPrivateProfileString = defaultValue
                GoTo CleanUp
            End If
        Else
            Exit Do
        End If
        
        retryCount = retryCount + 1
    Loop While retryCount < MAX_RETRIES
    
    If result > 0 Then
        IIniFile_GetPrivateProfileString = Left$(buffer, result)
    Else
        IIniFile_GetPrivateProfileString = defaultValue
    End If
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "GetPrivateProfileString"
    Exit Function

ErrorHandler:
    LogError "INIファイルからの読み取り中にエラーが発生しました: " & Err.Description
    IIniFile_GetPrivateProfileString = defaultValue
    Resume CleanUp
End Function

Private Function IIniFile_WritePrivateProfileString(ByVal section As String, _
                                                  ByVal key As String, _
                                                  ByVal value As String, _
                                                  ByVal filePath As String) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "WritePrivateProfileString"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim result As Long
    result = modWindowsAPI.WritePrivateProfileString(section, key, value, filePath)
    
    IIniFile_WritePrivateProfileString = (result <> 0)
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "WritePrivateProfileString"
    Exit Function

ErrorHandler:
    LogError "INIファイルへの書き込み中にエラーが発生しました: " & Err.Description
    IIniFile_WritePrivateProfileString = False
    Resume CleanUp
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LogError(ByVal message As String)
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrUnexpected
        .Description = message
        .Category = ECFileIO
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
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    Public Function TestBufferHandling(ByVal testValue As String) As String
        ' 大きな値の読み書きをテスト
        Const TEST_SECTION As String = "TestSection"
        Const TEST_KEY As String = "TestKey"
        Const TEST_FILE As String = "test.ini"
        
        ' 書き込み
        IIniFile_WritePrivateProfileString TEST_SECTION, TEST_KEY, testValue, TEST_FILE
        
        ' 読み込み
        TestBufferHandling = IIniFile_GetPrivateProfileString(TEST_SECTION, TEST_KEY, "", TEST_FILE)
    End Function
#End If