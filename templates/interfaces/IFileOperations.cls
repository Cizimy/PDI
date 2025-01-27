VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("高レベルなファイル操作とメタデータ管理を提供するインターフェース")

' 定数定義
Private Const MAX_FILE_SIZE As Long = 2147483648   ' 最大ファイルサイズ（2GB）
Private Const MAX_PATH_LENGTH As Long = 260        ' 最大パス長
Private Const BUFFER_SIZE As Long = 8192           ' バッファサイズ（8KB）
Private Const MAX_RETRY_COUNT As Long = 3          ' 最大リトライ回数
Private Const LOCK_TIMEOUT_MS As Long = 5000       ' ロックタイムアウト（5秒）
Private Const CLEANUP_INTERVAL_MS As Long = 300000 ' クリーンアップ間隔（5分）

' イベント定義
Public Event FileOperationStarted(ByVal operationType As String, ByVal filePath As String)
Public Event FileOperationCompleted(ByVal operationType As String, ByVal filePath As String, ByVal success As Boolean)
Public Event SecurityViolation(ByVal operationType As String, ByVal details As String)
Public Event ProgressUpdated(ByVal operationType As String, ByVal progress As Double)
Public Event MetadataChanged(ByVal filePath As String, ByVal metadataType As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal details As String)

' テキストファイル操作
''' <summary>
''' テキストファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="encoding">文字エンコーディング（オプション）</param>
''' <param name="options">読み込みオプション（オプション）</param>
''' <returns>ファイルの内容</returns>
Public Function ReadTextFile(ByVal filePath As String, _
                           Optional ByVal encoding As String, _
                           Optional ByVal options As TextReadOptions) As String
End Function

''' <summary>
''' テキストファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="content">書き込む内容</param>
''' <param name="options">書き込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function WriteTextFile(ByVal filePath As String, _
                            ByVal content As String, _
                            Optional ByVal options As TextWriteOptions) As Boolean
End Function

' バイナリファイル操作
''' <summary>
''' バイナリファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">読み込みオプション（オプション）</param>
''' <returns>ファイルのバイトデータ</returns>
Public Function ReadBinaryFile(ByVal filePath As String, _
                             Optional ByVal options As BinaryReadOptions) As Byte()
End Function

''' <summary>
''' バイナリファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="data">書き込むバイトデータ</param>
''' <param name="options">書き込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function WriteBinaryFile(ByVal filePath As String, _
                              ByRef data() As Byte, _
                              Optional ByVal options As BinaryWriteOptions) As Boolean
End Function

' ファイルシステム操作
''' <summary>
''' ファイルの存在を確認します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">確認オプション（オプション）</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
Public Function FileExists(ByVal filePath As String, _
                         Optional ByVal options As FileCheckOptions) As Boolean
End Function

''' <summary>
''' フォルダの存在を確認します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <param name="options">確認オプション（オプション）</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
Public Function FolderExists(ByVal folderPath As String, _
                           Optional ByVal options As FolderCheckOptions) As Boolean
End Function

''' <summary>
''' フォルダを作成します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <param name="options">作成オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CreateFolder(ByVal folderPath As String, _
                           Optional ByVal options As FolderCreateOptions) As Boolean
End Function

''' <summary>
''' ファイルを削除します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">削除オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function DeleteFile(ByVal filePath As String, _
                         Optional ByVal options As FileDeleteOptions) As Boolean
End Function

''' <summary>
''' フォルダを削除します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <param name="options">削除オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function DeleteFolder(ByVal folderPath As String, _
                           Optional ByVal options As FolderDeleteOptions) As Boolean
End Function

' メタデータ操作
''' <summary>
''' ファイルのメタデータを取得します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>メタデータ情報</returns>
Public Function GetMetadata(ByVal filePath As String) As FileMetadata
End Function

''' <summary>
''' ファイルのメタデータを設定します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="metadata">メタデータ情報</param>
''' <returns>成功時True、失敗時False</returns>
Public Function SetMetadata(ByVal filePath As String, _
                          ByVal metadata As FileMetadata) As Boolean
End Function

' セキュリティ操作
''' <summary>
''' ファイルのアクセス権を設定します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="permissions">アクセス権設定</param>
''' <returns>成功時True、失敗時False</returns>
Public Function SetPermissions(ByVal filePath As String, _
                             ByVal permissions As FilePermissions) As Boolean
End Function

''' <summary>
''' ファイルを暗号化します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">暗号化オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function EncryptFile(ByVal filePath As String, _
                          Optional ByVal options As EncryptionOptions) As Boolean
End Function

' パフォーマンス監視
''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' ファイル操作の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub