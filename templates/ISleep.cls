VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' スリープ操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、プロセスの実行を一時停止する機能を提供します。
''' タイミング制御やスロットリング、リトライ処理などに使用されます。
''' 
''' 注意点：
''' - UIスレッドでの長時間のスリープは避けるべきです
''' - 正確な時間制御が必要な場合はPerformanceCounterの使用を検討してください
''' - テスト時にはモック実装に置き換えることで、待ち時間をスキップできます
''' </remarks>

''' <summary>
''' 指定された時間だけ実行を停止します
''' </summary>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <remarks>
''' - 実際のスリープ時間は指定時間より長くなる可能性があります
''' - Windows のタイマー分解能に依存します（通常約15ms）
''' - スリープ中はCPUリソースを消費しません
''' - 0を指定すると、タイムスライスを放棄してすぐに復帰します
''' </remarks>
Public Sub Sleep(ByVal milliseconds As Long)
End Sub