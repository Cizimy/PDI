VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' イベントログの種類定義
Public Enum EventLogType
    EVENTLOG_SUCCESS = 0
    EVENTLOG_ERROR_TYPE = 1
    EVENTLOG_WARNING_TYPE = 2
    EVENTLOG_INFORMATION_TYPE = 4
    EVENTLOG_AUDIT_SUCCESS = 8
    EVENTLOG_AUDIT_FAILURE = 16
End Enum

''' <summary>
''' Windowsイベントログ操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、Windowsイベントログへのメッセージ書き込み機能を提供します。
''' システムの監視、診断、トラブルシューティングに使用されます。
''' 
''' 注意点：
''' - イベントログへの書き込みには適切な権限が必要です
''' - イベントソースは事前に登録が必要です
''' - メッセージは簡潔で意味のある内容にすべきです
''' - 機密情報の書き込みは避けるべきです
''' </remarks>

''' <summary>
''' イベントログにメッセージを書き込みます
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="message">メッセージ</param>
''' <param name="eventType">イベントの種類</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - sourceは事前にRegisterEventSourceで登録されている必要があります
''' - messageは32,766文字以内である必要があります
''' - イベントログが満杯の場合は失敗する可能性があります
''' - アプリケーションログに書き込まれます
''' </remarks>
Public Function WriteToEventLog(ByVal source As String, _
                              ByVal message As String, _
                              ByVal eventType As EventLogType) As Boolean
End Function

''' <summary>
''' イベントソースの存在を確認します
''' </summary>
''' <param name="source">イベントソース名</param>
''' <returns>存在する場合True、存在しない場合False</returns>
''' <remarks>
''' - レジストリキーの読み取り権限が必要です
''' - システム管理者権限がない場合は失敗する可能性があります
''' </remarks>
Public Function EventSourceExists(ByVal source As String) As Boolean
End Function