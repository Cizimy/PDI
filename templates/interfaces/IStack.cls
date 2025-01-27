VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("スタックデータ構造を管理し、LIFOアクセスとイテレーション機能を提供するインターフェース")

' 定数定義
Private Const DEFAULT_CAPACITY As Long = 1000      ' デフォルトの容量
Private Const MAX_CAPACITY As Long = 1000000       ' 最大容量
Private Const GROWTH_FACTOR As Double = 1.5        ' 成長係数
Private Const SHRINK_THRESHOLD As Double = 0.25    ' 縮小閾値
Private Const MAX_BATCH_SIZE As Long = 1000        ' 最大バッチサイズ
Private Const CLEANUP_INTERVAL_MS As Long = 60000  ' クリーンアップ間隔（1分）

' イベント定義
Public Event ItemPushed(ByVal item As Variant, ByVal newCount As Long)
Public Event ItemPopped(ByVal item As Variant, ByVal newCount As Long)
Public Event StackCleared(ByVal oldCount As Long)
Public Event CapacityChanged(ByVal oldCapacity As Long, ByVal newCapacity As Long)
Public Event ThresholdReached(ByVal currentCount As Long, ByVal capacity As Long)
Public Event OperationFailed(ByVal operation As String, ByVal reason As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' スタックに要素を追加します
''' </summary>
''' <param name="item">追加する要素</param>
''' <param name="options">追加オプション（オプション）</param>
''' <returns>追加成功時True、スタックが最大数を超えた場合False</returns>
''' <remarks>
''' エラー処理要件：
''' - アイテムの検証
''' - 容量の確認
''' - メモリ使用量の監視
''' - 型の互換性確認
''' </remarks>
Public Function Push(ByVal item As Variant, Optional ByVal options As PushOptions) As Boolean
End Function

''' <summary>
''' スタックから最後の要素を削除して返します
''' </summary>
''' <param name="options">取り出しオプション（オプション）</param>
''' <returns>削除した要素</returns>
''' <remarks>
''' エラー処理要件：
''' - 空スタックの処理
''' - メモリ解放の確認
''' - 型の一貫性確認
''' </remarks>
Public Function Pop(Optional ByVal options As PopOptions) As Variant
End Function

''' <summary>
''' 複数の要素をバッチでスタックに追加します
''' </summary>
''' <param name="items">追加する要素のコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>追加された要素数</returns>
Public Function PushBatch(ByVal items As Collection, Optional ByVal options As BatchOptions) As Long
End Function

''' <summary>
''' 指定した数の要素をバッチで取り出します
''' </summary>
''' <param name="count">取り出す要素数</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>取り出した要素のコレクション</returns>
Public Function PopBatch(ByVal count As Long, Optional ByVal options As BatchOptions) As Collection
End Function

''' <summary>
''' スタックの最後の要素を削除せずに取得します
''' </summary>
''' <param name="options">参照オプション（オプション）</param>
''' <returns>最後の要素</returns>
Public Function Peek(Optional ByVal options As PeekOptions) As Variant
End Function

''' <summary>
''' スタック内の指定位置の要素を参照します
''' </summary>
''' <param name="index">インデックス</param>
''' <returns>指定位置の要素</returns>
Public Function PeekAt(ByVal index As Long) As Variant
End Function

''' <summary>
''' スタックが空かどうかを確認します
''' </summary>
''' <returns>スタックが空の場合True</returns>
Public Function IsEmpty() As Boolean
End Function

''' <summary>
''' スタックが満杯かどうかを確認します
''' </summary>
''' <returns>スタックが満杯の場合True</returns>
Public Function IsFull() As Boolean
End Function

''' <summary>
''' スタック内の要素数を取得します
''' </summary>
Public Property Get Count() As Long
End Property

''' <summary>
''' スタックの容量を取得または設定します
''' </summary>
Public Property Get Capacity() As Long
End Property

Public Property Let Capacity(ByVal value As Long)
End Property

''' <summary>
''' スタック内の要素を検索します
''' </summary>
''' <param name="predicate">検索条件</param>
''' <returns>条件に一致する要素のコレクション</returns>
Public Function Search(ByVal predicate As String) As Collection
End Function

''' <summary>
''' スタックをクリアします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
Public Sub Clear(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' スタックの内容を配列にコピーします
''' </summary>
''' <returns>スタックの内容を含む配列</returns>
Public Function ToArray() As Variant()
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' スタックの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub