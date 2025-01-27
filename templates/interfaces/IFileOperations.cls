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

' ======================
' ファイル操作インターフェース
' ======================

''' <remarks>
''' このインターフェースを実装するクラスは、以下のエラー処理を実装する必要があります：
'''
''' 1. 共通のエラー処理要件
'''    - ファイルパスの検証（存在確認、アクセス権限、パス長など）
'''    - ファイルロックの適切な管理
'''    - リソースの確実な解放（ファイルハンドルなど）
'''    - エラー情報の適切なログ記録
'''    - エラーの種類に応じた適切なエラーコードの設定
'''
''' 2. エラー発生時の戻り値
'''    - Boolean型を返すメソッド：失敗時はFalseを返す
'''    - String型を返すメソッド：エラー時は空文字列を返す
'''    - バイト配列を返すメソッド：エラー時は空配列を返す
'''
''' 3. エラーの種類と対応
'''    - ErrFileNotFound: ファイルが存在しない
'''    - ErrFileAccessDenied: アクセス権限がない
'''    - ErrFileInvalidFormat: ファイル形式が不正
'''    - その他のシステムエラー
''' </remarks>

' テキストファイル操作
''' <summary>
''' テキストファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="encoding">文字エンコーディング（オプション）</param>
''' <returns>ファイルの内容、エラー時は空文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - ファイルの存在確認
''' - エンコーディングの検証
''' - ファイルロックの確認
''' - メモリ不足への対応
''' </remarks>
Public Function ReadTextFile(ByVal filePath As String, _
                           Optional ByVal encoding As String) As String
End Function

''' <summary>
''' テキストファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="content">書き込む内容</param>
''' <param name="append">追記モード（オプション）</param>
''' <param name="encoding">文字エンコーディング（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 書き込み権限の確認
''' - ディスク容量の確認
''' - 既存ファイルのバックアップ
''' - 書き込み失敗時の復旧処理
''' </remarks>
Public Function WriteTextFile(ByVal filePath As String, _
                            ByVal content As String, _
                            Optional ByVal append As Boolean = False, _
                            Optional ByVal encoding As String) As Boolean
End Function

' バイナリファイル操作
''' <summary>
''' バイナリファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>ファイルのバイトデータ、エラー時は空配列</returns>
''' <remarks>
''' エラー処理要件：
''' - ファイルサイズの検証
''' - メモリ使用量の監視
''' - 破損ファイルの検出
''' </remarks>
Public Function ReadBinaryFile(ByVal filePath As String) As Byte()
End Function

''' <summary>
''' バイナリファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="data">書き込むバイトデータ</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - データの整合性チェック
''' - 部分書き込みの防止
''' - 書き込み失敗時のロールバック
''' </remarks>
Public Function WriteBinaryFile(ByVal filePath As String, _
                              ByRef data() As Byte) As Boolean
End Function

' ファイルシステム操作
''' <summary>
''' ファイルの存在を確認します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - パスの妥当性検証
''' - アクセス権限の確認
''' - ネットワークパスの考慮
''' </remarks>
Public Function FileExists(ByVal filePath As String) As Boolean
End Function

''' <summary>
''' フォルダの存在を確認します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - パスの妥当性検証
''' - アクセス権限の確認
''' - シンボリックリンクの考慮
''' </remarks>
Public Function FolderExists(ByVal folderPath As String) As Boolean
End Function

''' <summary>
''' フォルダを作成します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 親フォルダの存在確認
''' - 作成権限の確認
''' - 既存フォルダの処理
''' </remarks>
Public Function CreateFolder(ByVal folderPath As String) As Boolean
End Function

''' <summary>
''' ファイルを削除します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 削除権限の確認
''' - ファイルロックの確認
''' - 読み取り専用属性の処理
''' </remarks>
Public Function DeleteFile(ByVal filePath As String) As Boolean
End Function

''' <summary>
''' フォルダを削除します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - フォルダの空確認
''' - 削除権限の確認
''' - サブフォルダの処理
''' </remarks>
Public Function DeleteFolder(ByVal folderPath As String) As Boolean
End Function

' パス操作
''' <summary>
''' 相対パスから絶対パスを取得します
''' </summary>
''' <param name="relativePath">相対パス</param>
''' <param name="basePath">基準パス（オプション）</param>
''' <returns>絶対パス、エラー時は空文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - パスの妥当性検証
''' - 循環参照の検出
''' - パス長の制限チェック
''' </remarks>
Public Function GetAbsolutePath(ByVal relativePath As String, _
                              Optional ByVal basePath As String) As String
End Function