VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' アプリケーション全体の設定を管理するインターフェース
''' </summary>

''' <summary>
''' アプリケーション設定値を取得します
''' </summary>
''' <param name="settingName">設定名</param>
''' <returns>設定値</returns>
Public Function GetSetting(ByVal settingName As String) As Variant
End Function

''' <summary>
''' アプリケーション設定値を設定します
''' </summary>
''' <param name="settingName">設定名</param>
''' <param name="settingValue">設定値</param>
Public Sub SetSetting(ByVal settingName As String, ByVal settingValue As Variant)
End Sub

''' <summary>
''' 設定ファイルから設定を読み込みます
''' </summary>
''' <param name="filePath">設定ファイルのパス</param>
Public Sub LoadSettings(ByVal filePath As String)
End Sub

''' <summary>
''' 設定ファイルに設定を保存します
''' </summary>
''' <param name="filePath">設定ファイルのパス</param>
Public Sub SaveSettings(ByVal filePath As String)
End Sub

' ======================
' プロパティ
' ======================

''' <summary>
''' ログファイルのパス
''' </summary>
Public Property Get LogFilePath() As String
End Property

Public Property Let LogFilePath(ByVal value As String)
End Property

''' <summary>
''' ログレベル
''' </summary>
Public Property Get LogLevel() As String
End Property

Public Property Let LogLevel(ByVal value As String)
End Property

''' <summary>
''' 最大再試行回数
''' </summary>
Public Property Get MaxRetryCount() As Long
End Property

Public Property Let MaxRetryCount(ByVal value As Long)
End Property

''' <summary>
''' スタックトレースの最大深さ
''' </summary>
Public Property Get MaxStackDepth() As Long
End Property

Public Property Let MaxStackDepth(ByVal value As Long)
End Property

''' <summary>
''' タイマー間隔
''' </summary>
Public Property Get TimerInterval() As Long
End Property

Public Property Let TimerInterval(ByVal value As Long)
End Property

''' <summary>
''' イベントソース名
''' </summary>
Public Property Get LogEventSource() As String
End Property

Public Property Let LogEventSource(ByVal value As String)
End Property

''' <summary>
''' パフォーマンス監視の有効/無効
''' </summary>
Public Property Get PerformanceMonitoringEnabled() As Boolean
End Property

Public Property Let PerformanceMonitoringEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' 診断機能の有効/無効
''' </summary>
Public Property Get DiagnosticsEnabled() As Boolean
End Property

Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' 自動保存の有効/無効
''' </summary>
Public Property Get AutoSave() As Boolean
End Property

Public Property Let AutoSave(ByVal value As Boolean)
End Property