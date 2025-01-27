VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("ファイル入出力操作とストリーム処理を提供するインターフェース")

' 定数定義
Private Const MAX_BUFFER_SIZE As Long = 8192       ' 最大バッファサイズ（8KB）
Private Const MAX_PATH_LENGTH As Long = 260        ' 最大パス長
Private Const MAX_RETRY_COUNT As Long = 3          ' 最大リトライ回数
Private Const LOCK_TIMEOUT_MS As Long = 5000       ' ロックタイムアウト（5秒）
Private Const DEFAULT_ENCODING As String = "UTF-8"  ' デフォルトエンコーディング
Private Const CLEANUP_INTERVAL_MS As Long = 300000 ' クリーンアップ間隔（5分）

' イベント定義
Public Event FileOpened(ByVal filePath As String, ByVal mode As FileMode)
Public Event FileClosed(ByVal filePath As String)
Public Event FileCreated(ByVal filePath As String)
Public Event FileDeleted(ByVal filePath As String)
Public Event FileMoved(ByVal sourcePath As String, ByVal destinationPath As String)
Public Event FileCopied(ByVal sourcePath As String, ByVal destinationPath As String)
Public Event SecurityAlert(ByVal alertType As String, ByVal details As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' ファイルの存在を確認します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">確認オプション（オプション）</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - パスの検証
''' - アクセス権の確認
''' - セキュリティチェック
''' - タイムアウト処理
''' </remarks>
Public Function FileExists(ByVal filePath As String, Optional ByVal options As FileCheckOptions) As Boolean
End Function

''' <summary>
''' フォルダの存在を確認します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <param name="options">確認オプション（オプション）</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
Public Function FolderExists(ByVal folderPath As String, Optional ByVal options As FolderCheckOptions) As Boolean
End Function

''' <summary>
''' フォルダを作成します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <param name="options">作成オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CreateFolder(ByVal folderPath As String, Optional ByVal options As FolderCreateOptions) As Boolean
End Function

''' <summary>
''' ファイルを削除します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">削除オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function DeleteFile(ByVal filePath As String, Optional ByVal options As FileDeleteOptions) As Boolean
End Function

''' <summary>
''' ファイルをコピーします
''' </summary>
''' <param name="sourceFilePath">コピー元ファイルパス</param>
''' <param name="destinationFilePath">コピー先ファイルパス</param>
''' <param name="options">コピーオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CopyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal options As FileCopyOptions) As Boolean
End Function

''' <summary>
''' ファイルを移動します
''' </summary>
''' <param name="sourceFilePath">移動元ファイルパス</param>
''' <param name="destinationFilePath">移動先ファイルパス</param>
''' <param name="options">移動オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function MoveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String, Optional ByVal options As FileMoveOptions) As Boolean
End Function

''' <summary>
''' ファイルを開きます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="mode">オープンモード</param>
''' <param name="options">オープンオプション（オプション）</param>
''' <returns>ファイルストリーム</returns>
Public Function OpenFile(ByVal filePath As String, ByVal mode As FileMode, Optional ByVal options As FileOpenOptions) As IFileStream
End Function

''' <summary>
''' ファイルを閉じます
''' </summary>
''' <param name="fileStream">ファイルストリーム</param>
''' <param name="options">クローズオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CloseFile(ByVal fileStream As IFileStream, Optional ByVal options As FileCloseOptions) As Boolean
End Function

''' <summary>
''' ファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">読み込みオプション（オプション）</param>
''' <returns>ファイルの内容</returns>
Public Function ReadFile(ByVal filePath As String, Optional ByVal options As FileReadOptions) As String
End Function

''' <summary>
''' ファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="content">書き込む内容</param>
''' <param name="options">書き込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function WriteFile(ByVal filePath As String, ByVal content As String, Optional ByVal options As FileWriteOptions) As Boolean
End Function

''' <summary>
''' ファイル情報を取得します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>ファイル情報</returns>
Public Function GetFileInfo(ByVal filePath As String) As FileInfo
End Function

''' <summary>
''' ファイルをロックします
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">ロックオプション（オプション）</param>
''' <returns>ロックトークン</returns>
Public Function LockFile(ByVal filePath As String, Optional ByVal options As FileLockOptions) As String
End Function

''' <summary>
''' ファイルのロックを解除します
''' </summary>
''' <param name="lockToken">ロックトークン</param>
''' <returns>成功時True、失敗時False</returns>
Public Function UnlockFile(ByVal lockToken As String) As Boolean
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' ファイルI/Oの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub