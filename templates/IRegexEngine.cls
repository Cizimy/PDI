VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRegexEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' 正規表現エンジンインターフェース
' ======================

''' <summary>
''' 正規表現パターンに基づいて文字列を検証します
''' </summary>
''' <param name="pattern">正規表現パターン</param>
''' <param name="text">検証する文字列</param>
''' <param name="ignoreCase">大文字小文字を区別しない場合はTrue</param>
''' <returns>パターンにマッチする場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - パターンの妥当性確認
''' - メモリ使用量の監視
''' - タイムアウト処理
''' - 無限ループの防止
''' 
''' 実装時の注意点：
''' 1. パターンの構文エラーは適切に処理すること
''' 2. 大規模な入力に対する性能を考慮すること
''' 3. セキュリティ上の脆弱性に注意すること
''' </remarks>
Public Function Test(ByVal pattern As String, ByVal text As String, Optional ByVal ignoreCase As Boolean = True) As Boolean
End Function

''' <summary>
''' 正規表現パターンに基づいて文字列を置換します
''' </summary>
''' <param name="pattern">正規表現パターン</param>
''' <param name="text">対象の文字列</param>
''' <param name="replacement">置換後の文字列</param>
''' <param name="ignoreCase">大文字小文字を区別しない場合はTrue</param>
''' <returns>置換後の文字列</returns>
Public Function Replace(ByVal pattern As String, ByVal text As String, ByVal replacement As String, Optional ByVal ignoreCase As Boolean = True) As String
End Function

''' <summary>
''' 正規表現パターンに基づいて文字列からマッチする部分を抽出します
''' </summary>
''' <param name="pattern">正規表現パターン</param>
''' <param name="text">対象の文字列</param>
''' <param name="ignoreCase">大文字小文字を区別しない場合はTrue</param>
''' <returns>マッチした文字列の配列</returns>
Public Function Match(ByVal pattern As String, ByVal text As String, Optional ByVal ignoreCase As Boolean = True) As Variant
End Function