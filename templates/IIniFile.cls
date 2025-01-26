VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' INIファイル操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、Windows APIを使用したINIファイルの読み書き操作を提供します。
''' アプリケーション設定の永続化や構成管理に使用されます。
''' </remarks>

''' <summary>
''' INIファイルから値を読み取ります
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="key">キー名</param>
''' <param name="defaultValue">デフォルト値</param>
''' <param name="filePath">INIファイルのパス</param>
''' <returns>読み取った文字列。キーが存在しない場合はデフォルト値</returns>
''' <remarks>
''' - セクションやキーが存在しない場合、デフォルト値が返されます
''' - バッファサイズを超える値は切り捨てられます
''' - エラー発生時はデフォルト値が返されます
''' </remarks>
Public Function GetPrivateProfileString(ByVal section As String, _
                                      ByVal key As String, _
                                      ByVal defaultValue As String, _
                                      ByVal filePath As String) As String
End Function

''' <summary>
''' INIファイルに値を書き込みます
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="key">キー名</param>
''' <param name="value">書き込む値</param>
''' <param name="filePath">INIファイルのパス</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - セクションやキーが存在しない場合は新規作成されます
''' - ファイルが存在しない場合は新規作成されます
''' - 既存の値は上書きされます
''' - valueにNullを指定するとキーが削除されます
''' - sectionにNullを指定するとセクションが削除されます
''' </remarks>
Public Function WritePrivateProfileString(ByVal section As String, _
                                        ByVal key As String, _
                                        ByVal value As String, _
                                        ByVal filePath As String) As Boolean
End Function