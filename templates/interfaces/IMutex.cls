VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' ミューテックス操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、スレッド間の同期とリソースの排他制御を提供します。
''' クリティカルセクションの保護やリソースの共有アクセスの制御に使用されます。
''' </remarks>

''' <summary>
''' ミューテックスを作成します
''' </summary>
''' <param name="initialOwner">初期所有者フラグ</param>
''' <param name="name">ミューテックスの名前</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - NULLまたは空の名前は許可されます
''' - 既存のミューテックスが存在する場合はそのハンドルを返します
''' - エラー処理はIErrorHandlerインターフェースを通じて行われます
''' </remarks>
Public Function CreateMutex(ByVal initialOwner As Boolean, ByVal name As String) As Boolean
End Function

''' <summary>
''' ミューテックスを解放します
''' </summary>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - 他のスレッドがミューテックスを待機している場合、そのスレッドに所有権が移ります
''' - 所有していないミューテックスを解放しようとするとエラーになります
''' </remarks>
Public Function ReleaseMutex() As Boolean
End Function

''' <summary>
''' ミューテックスの所有権を取得します
''' </summary>
''' <param name="timeoutMilliseconds">タイムアウト時間（ミリ秒）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - タイムアウトに-1を指定すると無限に待機します
''' - タイムアウトした場合はFalseを返します
''' - 既に所有している場合は即座にTrueを返します
''' </remarks>  
Public Function WaitForSingleObject(ByVal timeoutMilliseconds As Long) As Boolean
End Function