VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConnectionStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("データベース接続文字列の安全な生成と管理を提供するインターフェース")

' 定数定義
Private Const MAX_SERVER_LENGTH As Long = 255      ' サーバー名の最大長
Private Const MAX_DATABASE_LENGTH As Long = 255    ' データベース名の最大長
Private Const MAX_USERNAME_LENGTH As Long = 128    ' ユーザー名の最大長
Private Const MAX_PASSWORD_LENGTH As Long = 128    ' パスワードの最大長
Private Const DEFAULT_TIMEOUT As Long = 30         ' デフォルトのタイムアウト（秒）
Private Const MAX_RETRY_COUNT As Long = 3          ' 最大リトライ回数
Private Const CACHE_DURATION_MS As Long = 300000   ' キャッシュの有効期間（5分）

' イベント定義
Public Event ConnectionStringBuilt(ByVal connectionString As String, ByVal isEncrypted As Boolean)
Public Event ParameterChanged(ByVal parameterName As String, ByVal oldValue As String, ByVal newValue As String)
Public Event ValidationError(ByVal parameterName As String, ByVal errorMessage As String)
Public Event SecurityAlert(ByVal alertType As String, ByVal details As String)
Public Event ConnectionTested(ByVal success As Boolean, ByVal errorMessage As String)
Public Event EncryptionKeyRotated(ByVal keyVersion As String)
Public Event ConfigurationLoaded(ByVal source As String, ByVal success As Boolean)

''' <summary>
''' データベース接続文字列を生成します
''' </summary>
''' <param name="options">生成オプション（オプション）</param>
''' <returns>生成された接続文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証
''' - セキュリティチェック
''' - 文字列の整形
''' - 暗号化処理
''' </remarks>
Public Function BuildConnectionString(Optional ByVal options As ConnectionStringOptions) As String
End Function

''' <summary>
''' 接続文字列をパースします
''' </summary>
''' <param name="connectionString">接続文字列</param>
''' <param name="options">パースオプション（オプション）</param>
Public Sub ParseConnectionString(ByVal connectionString As String, Optional ByVal options As ParseOptions)
End Sub

''' <summary>
''' 接続文字列を暗号化します
''' </summary>
''' <param name="connectionString">接続文字列</param>
''' <param name="options">暗号化オプション（オプション）</param>
''' <returns>暗号化された接続文字列</returns>
Public Function EncryptConnectionString(ByVal connectionString As String, Optional ByVal options As EncryptionOptions) As String
End Function

''' <summary>
''' 接続文字列を復号化します
''' </summary>
''' <param name="encryptedString">暗号化された接続文字列</param>
''' <param name="options">復号化オプション（オプション）</param>
''' <returns>復号化された接続文字列</returns>
Public Function DecryptConnectionString(ByVal encryptedString As String, Optional ByVal options As DecryptionOptions) As String
End Function

''' <summary>
''' 接続をテストします
''' </summary>
''' <param name="options">テストオプション（オプション）</param>
''' <returns>テスト結果</returns>
Public Function TestConnection(Optional ByVal options As TestOptions) As TestResult
End Function

''' <summary>
''' 設定を検証します
''' </summary>
''' <param name="options">検証オプション（オプション）</param>
''' <returns>検証結果</returns>
Public Function ValidateConfiguration(Optional ByVal options As ValidationOptions) As ValidationResult
End Function

' プロパティ定義

''' <summary>
''' サーバー名
''' </summary>
Public Property Get Server() As String
End Property

Public Property Let Server(ByVal value As String)
End Property

''' <summary>
''' データベース名
''' </summary>
Public Property Get Database() As String
End Property

Public Property Let Database(ByVal value As String)
End Property

''' <summary>
''' ユーザー名
''' </summary>
Public Property Get UserName() As String
End Property

Public Property Let UserName(ByVal value As String)
End Property

''' <summary>
''' パスワード
''' </summary>
Public Property Get Password() As String
End Property

Public Property Let Password(ByVal value As String)
End Property

''' <summary>
''' 追加のパラメータ
''' </summary>
Public Property Get AdditionalParameters() As String
End Property

Public Property Let AdditionalParameters(ByVal value As String)
End Property

''' <summary>
''' タイムアウト値（秒）
''' </summary>
Public Property Get Timeout() As Long
End Property

Public Property Let Timeout(ByVal value As Long)
End Property

''' <summary>
''' 暗号化が有効かどうか
''' </summary>
Public Property Get IsEncryptionEnabled() As Boolean
End Property

Public Property Let IsEncryptionEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' 設定を保存します
''' </summary>
''' <param name="filePath">保存先のファイルパス</param>
''' <param name="options">保存オプション（オプション）</param>
Public Sub SaveConfiguration(ByVal filePath As String, Optional ByVal options As SaveOptions)
End Sub

''' <summary>
''' 設定を読み込みます
''' </summary>
''' <param name="filePath">読み込むファイルのパス</param>
''' <param name="options">読み込みオプション（オプション）</param>
Public Sub LoadConfiguration(ByVal filePath As String, Optional ByVal options As LoadOptions)
End Sub

''' <summary>
''' 接続文字列ビルダーの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub