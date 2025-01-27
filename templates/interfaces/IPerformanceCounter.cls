VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPerformanceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' パフォーマンスカウンター操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、Windows APIを使用した高分解能タイマーの操作を提供します。
''' パフォーマンス測定やベンチマーク、タイミング制御に使用されます。
''' </remarks>

''' <summary>
''' パフォーマンスカウンターの現在値を取得します
''' </summary>
''' <param name="performanceCount">カウンター値を格納する変数</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - カウンター値の解釈にはQueryPerformanceFrequencyの値が必要です
''' - 高分解能タイマーが利用できない環境では失敗する可能性があります
''' - マルチプロセッサシステムでは、プロセッサ間でカウンター値が異なる場合があります
''' </remarks>
Public Function QueryPerformanceCounter(ByRef performanceCount As Currency) As Boolean
End Function

''' <summary>
''' パフォーマンスカウンターの周波数を取得します
''' </summary>
''' <param name="frequency">周波数を格納する変数</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' - 周波数は1秒あたりのカウント数を表します
''' - システムの起動中は一定です
''' - 高分解能タイマーが利用できない環境では失敗する可能性があります
''' - この値を使用してカウンター値を時間に変換します
''' </remarks>
Public Function QueryPerformanceFrequency(ByRef frequency As Currency) As Boolean
End Function