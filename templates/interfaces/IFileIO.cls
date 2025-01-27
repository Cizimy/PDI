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

' ======================
' ファイルI/Oインターフェース
' ======================

''' <summary>
''' ファイルの存在を確認します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
Public Function FileExists(ByVal filePath As String) As Boolean
End Function

''' <summary>
''' フォルダの存在を確認します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <returns>存在する場合True、それ以外はFalse</returns>
Public Function FolderExists(ByVal folderPath As String) As Boolean
End Function

''' <summary>
''' フォルダを作成します
''' </summary>
''' <param name="folderPath">フォルダパス</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CreateFolder(ByVal folderPath As String) As Boolean
End Function

''' <summary>
''' ファイルを削除します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>成功時True、失敗時False</returns>
Public Function DeleteFile(ByVal filePath As String) As Boolean
End Function

''' <summary>
''' ファイルをコピーします
''' </summary>
''' <param name="sourceFilePath">コピー元ファイルパス</param>
''' <param name="destinationFilePath">コピー先ファイルパス</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CopyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
End Function

''' <summary>
''' ファイルを移動します
''' </summary>
''' <param name="sourceFilePath">移動元ファイルパス</param>
''' <param name="destinationFilePath">移動先ファイルパス</param>
''' <returns>成功時True、失敗時False</returns>
Public Function MoveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
End Function

''' <summary>
''' ファイルを開きます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="mode">オープンモード（"Read", "Write", "Append"）</param>
''' <returns>ファイルオブジェクト、エラー時はNothing</returns>
Public Function OpenFile(ByVal filePath As String, ByVal mode As String) As Object
End Function

''' <summary>
''' ファイルを閉じます
''' </summary>
''' <param name="file">ファイルオブジェクト</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CloseFile(ByVal file As Object) As Boolean
End Function