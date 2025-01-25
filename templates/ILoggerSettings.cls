VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILoggerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' ロガー設定インターフェース
' ======================

''' <summary>
''' ログレベルを取得します
''' </summary>
Public Property Get LogLevel() As LogLevelEnum
End Property

''' <summary>
''' ログ出力先を取得します
''' </summary>
Public Property Get LogDestination() As ILogDestination
End Property

''' <summary>
''' ログファイルパスを取得します
''' </summary>
Public Property Get LogFilePath() As String
End Property

''' <summary>
''' ログテーブル名を取得します
''' </summary>
Public Property Get LogTableName() As String
End Property

''' <summary>
''' ログイベントソースを取得します
''' </summary>
Public Property Get LogEventSource() As String
End Property

''' <summary>
''' タイマー間隔を取得します
''' </summary>
Public Property Get TimerInterval() As Long
End Property

''' <summary>
''' データベース接続文字列を取得します
''' </summary>
Public Property Get ConnectionString() As String
End Property