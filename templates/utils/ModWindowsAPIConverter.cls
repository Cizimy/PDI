VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModWindowsAPIConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' modWindowsAPIの関数呼び出しを新しいインターフェースベースの実装に変換するヘルパークラス
''' </summary>
''' <remarks>
''' このクラスは、従来のmodWindowsAPI関数呼び出しを新しいインターフェースベースの
''' 実装に橋渡しする機能を提供します。既存のコードとの互換性を維持しながら、
''' 新しいインターフェースベースの実装への段階的な移行を支援します。
''' </remarks>

Private mMutex As IMutex
Private mCrypto As ICryptography
Private mIniFile As IIniFile
Private mPerformanceCounter As IPerformanceCounter
Private mEventLog As IEventLog
Private mSleep As ISleep

' エラー情報の保持用
Private mLastError As Long
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLock = New clsLock
    
    ' 各インターフェースの実装クラスをインスタンス化
    Set mMutex = New MutexImpl
    Set mCrypto = New CryptographyImpl
    Set mIniFile = New IniFileImpl
    Set mPerformanceCounter = New PerformanceCounterImpl
    Set mEventLog = New EventLogImpl
    Set mSleep = New SleepImpl
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mMutex = Nothing
    Set mCrypto = Nothing
    Set mIniFile = Nothing
    Set mPerformanceCounter = Nothing
    Set mEventLog = Nothing
    Set mSleep = Nothing
End Sub

' ======================
' パブリックメソッド
' ======================

''' <summary>
''' 最後に発生したWin32エラーコードを取得します
''' </summary>
Public Property Get LastError() As Long
    LastError = mLastError
End Property

''' <summary>
''' ミューテックスインターフェースを取得します
''' </summary>
Public Property Get Mutex() As IMutex
    Set Mutex = mMutex
End Property

''' <summary>
''' 暗号化インターフェースを取得します
''' </summary>
Public Property Get Crypto() As ICryptography
    Set Crypto = mCrypto
End Property

''' <summary>
''' INIファイルインターフェースを取得します
''' </summary>
Public Property Get IniFile() As IIniFile
    Set IniFile = mIniFile
End Property

''' <summary>
''' パフォーマンスカウンターインターフェースを取得します
''' </summary>
Public Property Get PerformanceCounter() As IPerformanceCounter
    Set PerformanceCounter = mPerformanceCounter
End Property

''' <summary>
''' イベントログインターフェースを取得します
''' </summary>
Public Property Get EventLog() As IEventLog
    Set EventLog = mEventLog
End Property

''' <summary>
''' スリープインターフェースを取得します
''' </summary>
Public Property Get Sleep() As ISleep
    Set Sleep = mSleep
End Property

' ======================
' エラー処理
' ======================
Private Sub UpdateLastError()
    #If Win64 Then
        mLastError = CLng(CreateObject("WScript.Shell").Environment("PROCESS")("ERROR_CODE"))
    #Else
        mLastError = Err.LastDllError
    #End If
End Sub

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Sub ResetInterfaces()
        Class_Terminate
        Class_Initialize
    End Sub
    
    Public Function ValidateInterfaces() As Boolean
        ValidateInterfaces = Not (mMutex Is Nothing Or _
                                mCrypto Is Nothing Or _
                                mIniFile Is Nothing Or _
                                mPerformanceCounter Is Nothing Or _
                                mEventLog Is Nothing Or _
                                mSleep Is Nothing)
    End Function
#End If