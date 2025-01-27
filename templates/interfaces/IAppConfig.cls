VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("アプリケーション設定を管理するインターフェース")
'@Remarks("設定の読み書き、暗号化、バリデーション、バックアップ/復元などの機能を提供")

'@Description("設定値を取得する")
'@Param("settingName: 設定名")
'@Param("options: 取得時のオプション（SettingOptions）")
'@Return("設定値（Variant型）")
Public Function GetSetting(ByVal settingName As String, Optional ByVal options As SettingOptions = soNone) As Variant
End Function

'@Description("設定値を設定する")
'@Param("settingName: 設定名")
'@Param("settingValue: 設定値")
'@Param("options: 設定時のオプション（SettingOptions）")
Public Sub SetSetting(ByVal settingName As String, ByVal settingValue As Variant, Optional ByVal options As SettingOptions = soNone)
End Sub

'@Description("設定ファイルから設定を読み込む")
'@Param("filePath: 設定ファイルのパス")
'@Param("options: 読み込み時のオプション（LoadOptions）")
Public Sub LoadSettings(ByVal filePath As String, Optional ByVal options As LoadOptions = loNone)
End Sub

'@Description("設定をファイルに保存する")
'@Param("filePath: 保存先ファイルパス")
'@Param("options: 保存時のオプション（SaveOptions）")
Public Sub SaveSettings(ByVal filePath As String, Optional ByVal options As SaveOptions = soNone)
End Sub

'@Description("設定の暗号化キーを設定する")
'@Param("encryptionKey: 暗号化キー")
'@Param("options: 暗号化オプション（EncryptionOptions）")
Public Sub SetEncryptionKey(ByVal encryptionKey As String, Optional ByVal options As EncryptionOptions = eoNone)
End Sub

'@Description("現在の環境を設定する")
'@Param("environment: 環境名（dev/test/prod等）")
'@Param("options: 環境設定オプション（EnvironmentOptions）")
Public Sub SetEnvironment(ByVal environment As String, Optional ByVal options As EnvironmentOptions = enoNone)
End Sub

'@Description("設定値を検証する")
'@Param("options: 検証オプション（ValidationOptions）")
'@Return("検証結果（ValidationResult型）")
Public Function ValidateSettings(Optional ByVal options As ValidationOptions = voNone) As ValidationResult
End Function

'@Description("設定の変更履歴を取得する")
'@Param("settingName: 設定名（省略可）")
'@Param("options: 履歴取得オプション（HistoryOptions）")
'@Return("設定変更履歴のコレクション")
Public Function GetSettingHistory(Optional ByVal settingName As String, Optional ByVal options As HistoryOptions = hoNone) As Collection
End Function

'@Description("設定のバックアップを作成する")
'@Param("backupPath: バックアップファイルのパス")
'@Return("バックアップの成功可否")
Public Function CreateBackup(ByVal backupPath As String) As Boolean
End Function

'@Description("バックアップから設定を復元する")
'@Param("backupPath: バックアップファイルのパス")
'@Param("options: 復元オプション（RestoreOptions）")
'@Return("復元の成功可否")
Public Function RestoreFromBackup(ByVal backupPath As String, Optional ByVal options As RestoreOptions = roNone) As Boolean
End Function

'@Description("パフォーマンスメトリクスを取得する")
'@Return("パフォーマンスメトリクスのコレクション")
Public Function GetPerformanceMetrics() As Collection
End Function

' プロパティ定義
'@Description("ログファイルパス")
Public Property Get LogFilePath() As String
End Property
Public Property Let LogFilePath(ByVal value As String)
End Property

'@Description("ログレベル")
Public Property Get LogLevel() As LogLevel
End Property
Public Property Let LogLevel(ByVal value As LogLevel)
End Property

'@Description("最大リトライ回数")
Public Property Get MaxRetryCount() As Long
End Property
Public Property Let MaxRetryCount(ByVal value As Long)
End Property

'@Description("最大スタック深度")
Public Property Get MaxStackDepth() As Long
End Property
Public Property Let MaxStackDepth(ByVal value As Long)
End Property

'@Description("タイマー間隔（ミリ秒）")
Public Property Get TimerInterval() As Long
End Property
Public Property Let TimerInterval(ByVal value As Long)
End Property

'@Description("ログイベントソース")
Public Property Get LogEventSource() As String
End Property
Public Property Let LogEventSource(ByVal value As String)
End Property

'@Description("パフォーマンス監視の有効/無効")
Public Property Get PerformanceMonitoringEnabled() As Boolean
End Property
Public Property Let PerformanceMonitoringEnabled(ByVal value As Boolean)
End Property

'@Description("診断機能の有効/無効")
Public Property Get DiagnosticsEnabled() As Boolean
End Property
Public Property Let DiagnosticsEnabled(ByVal value As Boolean)
End Property

'@Description("自動保存の有効/無効")
Public Property Get AutoSave() As Boolean
End Property
Public Property Let AutoSave(ByVal value As Boolean)
End Property

'@Description("現在の環境名")
Public Property Get CurrentEnvironment() As String
End Property

'@Description("暗号化キーのバージョン")
Public Property Get EncryptionKeyVersion() As String
End Property