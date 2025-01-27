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

'@Description("INIファイルの読み書きと高度な設定管理を提供するインターフェース")

' 定数定義
Private Const MAX_SECTION_LENGTH As Long = 255     ' セクション名の最大長
Private Const MAX_KEY_LENGTH As Long = 255        ' キー名の最大長
Private Const MAX_VALUE_LENGTH As Long = 32768    ' 値の最大長
Private Const MAX_LINE_LENGTH As Long = 32768     ' 行の最大長
Private Const CACHE_DURATION_MS As Long = 300000  ' キャッシュの有効期間（5分）
Private Const MAX_BACKUP_COUNT As Long = 5        ' 最大バックアップ数
Private Const BUFFER_SIZE As Long = 8192          ' バッファサイズ（8KB）

' イベント定義
Public Event ValueChanged(ByVal section As String, ByVal key As String, ByVal oldValue As String, ByVal newValue As String)
Public Event SectionAdded(ByVal section As String)
Public Event SectionRemoved(ByVal section As String)
Public Event FileLoaded(ByVal filePath As String, ByVal success As Boolean)
Public Event FileSaved(ByVal filePath As String, ByVal success As Boolean)
Public Event CacheUpdated(ByVal section As String, ByVal key As String)
Public Event BackupCreated(ByVal backupPath As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' INIファイルから値を読み取ります
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="key">キー名</param>
''' <param name="defaultValue">デフォルト値（オプション）</param>
''' <param name="options">読み取りオプション（オプション）</param>
''' <returns>読み取った値</returns>
Public Function GetValue(ByVal section As String, _
                        ByVal key As String, _
                        Optional ByVal defaultValue As Variant, _
                        Optional ByVal options As ReadOptions) As Variant
End Function

''' <summary>
''' INIファイルに値を書き込みます
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="key">キー名</param>
''' <param name="value">書き込む値</param>
''' <param name="options">書き込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function SetValue(ByVal section As String, _
                        ByVal key As String, _
                        ByVal value As Variant, _
                        Optional ByVal options As WriteOptions) As Boolean
End Function

''' <summary>
''' セクションを取得します
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>セクション内のキーと値のコレクション</returns>
Public Function GetSection(ByVal section As String, _
                         Optional ByVal options As SectionOptions) As Collection
End Function

''' <summary>
''' すべてのセクションを取得します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>セクション名のコレクション</returns>
Public Function GetSections(Optional ByVal options As SectionOptions) As Collection
End Function

''' <summary>
''' セクションを削除します
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="options">削除オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function RemoveSection(ByVal section As String, _
                            Optional ByVal options As RemoveOptions) As Boolean
End Function

''' <summary>
''' キーを削除します
''' </summary>
''' <param name="section">セクション名</param>
''' <param name="key">キー名</param>
''' <param name="options">削除オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function RemoveKey(ByVal section As String, _
                         ByVal key As String, _
                         Optional ByVal options As RemoveOptions) As Boolean
End Function

''' <summary>
''' INIファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">読み込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function LoadFile(ByVal filePath As String, _
                        Optional ByVal options As LoadOptions) As Boolean
End Function

''' <summary>
''' INIファイルを保存します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">保存オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function SaveFile(ByVal filePath As String, _
                        Optional ByVal options As SaveOptions) As Boolean
End Function

''' <summary>
''' バックアップを作成します
''' </summary>
''' <param name="options">バックアップオプション（オプション）</param>
''' <returns>バックアップファイルのパス</returns>
Public Function CreateBackup(Optional ByVal options As BackupOptions) As String
End Function

''' <summary>
''' バックアップから復元します
''' </summary>
''' <param name="backupPath">バックアップファイルのパス</param>
''' <param name="options">復元オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function RestoreFromBackup(ByVal backupPath As String, _
                                Optional ByVal options As RestoreOptions) As Boolean
End Function

''' <summary>
''' キャッシュをクリアします
''' </summary>
Public Sub ClearCache()
End Sub

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' INIファイルの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub