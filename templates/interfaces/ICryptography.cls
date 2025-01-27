VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("暗号化操作と鍵管理を提供し、セキュアな暗号化機能を実装するインターフェース")

' 定数定義
Private Const MAX_KEY_SIZE As Long = 4096          ' 最大鍵サイズ（ビット）
Private Const MIN_KEY_SIZE As Long = 128           ' 最小鍵サイズ（ビット）
Private Const DEFAULT_ITERATION_COUNT As Long = 10000 ' デフォルトの反復回数
Private Const MAX_SALT_SIZE As Long = 32           ' 最大ソルトサイズ（バイト）
Private Const KEY_ROTATION_INTERVAL As Long = 7776000000 ' 鍵のローテーション間隔（90日）
Private Const CACHE_DURATION_MS As Long = 300000   ' キャッシュの有効期間（5分）

' イベント定義
Public Event KeyGenerated(ByVal keyId As String, ByVal keySize As Long)
Public Event KeyRotated(ByVal oldKeyId As String, ByVal newKeyId As String)
Public Event EncryptionCompleted(ByVal dataSize As Long, ByVal executionTime As Double)
Public Event DecryptionCompleted(ByVal dataSize As Long, ByVal executionTime As Double)
Public Event SecurityAlert(ByVal alertType As String, ByVal severity As Long)
Public Event AlgorithmChanged(ByVal oldAlgorithm As String, ByVal newAlgorithm As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event AuditLog(ByVal operation As String, ByVal details As String)

''' <summary>
''' 暗号化コンテキストを初期化します
''' </summary>
''' <param name="options">初期化オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - プロバイダーの検証
''' - アルゴリズムの互換性確認
''' - リソースの確保
''' - セキュリティ要件の検証
''' </remarks>
Public Function Initialize(Optional ByVal options As CryptoOptions) As Boolean
End Function

''' <summary>
''' データを暗号化します
''' </summary>
''' <param name="data">暗号化するデータ</param>
''' <param name="options">暗号化オプション（オプション）</param>
''' <returns>暗号化されたデータ</returns>
''' <remarks>
''' エラー処理要件：
''' - データの検証
''' - 鍵の有効性確認
''' - メモリ使用量の監視
''' - パフォーマンスの最適化
''' </remarks>
Public Function Encrypt(ByRef data As Variant, Optional ByVal options As EncryptionOptions) As Variant
End Function

''' <summary>
''' データを復号化します
''' </summary>
''' <param name="data">復号化するデータ</param>
''' <param name="options">復号化オプション（オプション）</param>
''' <returns>復号化されたデータ</returns>
Public Function Decrypt(ByRef data As Variant, Optional ByVal options As DecryptionOptions) As Variant
End Function

''' <summary>
''' ハッシュ値を計算します
''' </summary>
''' <param name="data">ハッシュ化するデータ</param>
''' <param name="algorithm">ハッシュアルゴリズム（オプション）</param>
''' <returns>ハッシュ値</returns>
Public Function ComputeHash(ByRef data As Variant, Optional ByVal algorithm As String) As String
End Function

''' <summary>
''' 鍵を生成します
''' </summary>
''' <param name="keySize">鍵サイズ（ビット）</param>
''' <param name="options">生成オプション（オプション）</param>
''' <returns>鍵ID</returns>
Public Function GenerateKey(ByVal keySize As Long, Optional ByVal options As KeyGenerationOptions) As String
End Function

''' <summary>
''' 鍵を導出します
''' </summary>
''' <param name="password">パスワード</param>
''' <param name="salt">ソルト</param>
''' <param name="options">導出オプション（オプション）</param>
''' <returns>導出された鍵</returns>
Public Function DeriveKey(ByVal password As String, ByVal salt As Variant, Optional ByVal options As KeyDerivationOptions) As Variant
End Function

''' <summary>
''' 鍵をローテーションします
''' </summary>
''' <param name="keyId">現在の鍵ID</param>
''' <param name="options">ローテーションオプション（オプション）</param>
''' <returns>新しい鍵ID</returns>
Public Function RotateKey(ByVal keyId As String, Optional ByVal options As KeyRotationOptions) As String
End Function

''' <summary>
''' 鍵を検証します
''' </summary>
''' <param name="keyId">鍵ID</param>
''' <returns>有効な場合True</returns>
Public Function ValidateKey(ByVal keyId As String) As Boolean
End Function

''' <summary>
''' セキュリティ監査を実行します
''' </summary>
''' <param name="options">監査オプション（オプション）</param>
''' <returns>監査結果</returns>
Public Function PerformSecurityAudit(Optional ByVal options As AuditOptions) As SecurityAuditResult
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' 暗号化アルゴリズムを設定します
''' </summary>
''' <param name="algorithm">アルゴリズム名</param>
''' <param name="options">アルゴリズムオプション（オプション）</param>
Public Sub SetAlgorithm(ByVal algorithm As String, Optional ByVal options As AlgorithmOptions)
End Sub

''' <summary>
''' 現在の暗号化アルゴリズムを取得します
''' </summary>
''' <returns>アルゴリズム名</returns>
Public Function GetCurrentAlgorithm() As String
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub

''' <summary>
''' 暗号化操作の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function