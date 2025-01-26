VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' ロック操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、リソースの排他制御を提供します。
''' クリティカルセクションの保護やリソースの共有アクセスの制御に使用されます。
''' </remarks>

''' <summary>
''' ロックを取得します
''' </summary>
''' <param name="timeout">タイムアウト時間（ミリ秒）。-1は無限待機</param>
''' <returns>ロックの取得に成功した場合はTrue、タイムアウトした場合はFalse</returns>
''' <remarks>
''' - タイムアウトに-1を指定すると無限に待機します
''' - タイムアウトした場合はFalseを返します
''' - 既に所有している場合は即座にTrueを返します
''' </remarks>
Public Function Acquire(Optional ByVal timeout As Long = -1) As Boolean
End Function

''' <summary>
''' ロックを解放します
''' </summary>
''' <remarks>
''' - 他のスレッドがロックを待機している場合、そのスレッドに所有権が移ります
''' - 所有していないロックを解放しようとするとエラーになります
''' </remarks>
Public Sub Release()
End Sub