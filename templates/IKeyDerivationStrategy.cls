VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IKeyDerivationStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' キー導出戦略のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、パスワードからキーを導出するための戦略パターンを定義します。
''' 異なるキー導出アルゴリズムを実装するクラスは、このインターフェースを実装する必要があります。
''' </remarks>

''' <summary>
''' パスワードからキーを導出します
''' </summary>
''' <param name="password">パスワード文字列</param>
''' <param name="salt">ソルト値（バイト配列）</param>
''' <param name="iterations">反復回数</param>
''' <returns>導出されたキー（バイト配列）</returns>
Public Function DeriveKey(ByVal password As String, _
                         ByRef salt As Variant, _
                         ByVal iterations As Long) As Byte()
End Function