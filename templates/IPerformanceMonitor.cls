VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("パフォーマンス監視機能を提供するインターフェース")

'==========================
' インターフェースメソッド
'==========================

''' <summary>
''' 計測を開始します
''' </summary>
''' <param name="measurementName">計測名</param>
Public Sub Start(ByVal measurementName As String)
End Sub

''' <summary>
''' 計測を終了します
''' </summary>
''' <param name="measurementName">計測名</param>
Public Sub Stop(ByVal measurementName As String)
End Sub

''' <summary>
''' 指定した計測の結果を取得します
''' </summary>
''' <param name="measurementName">計測名</param>
''' <returns>計測結果（ミリ秒）</returns>
Public Function GetMeasurement(ByVal measurementName As String) As Double
End Function

''' <summary>
''' すべての計測結果を取得します
''' </summary>
''' <returns>計測結果の配列</returns>
Public Function GetAllMeasurements() As Variant
End Function

''' <summary>
''' すべての計測結果をクリアします
''' </summary>
Public Sub Clear()
End Sub

''' <summary>
''' パフォーマンス監視が有効かどうかを取得します
''' </summary>
''' <returns>有効な場合True</returns>
Public Property Get IsEnabled() As Boolean
End Property

''' <summary>
''' パフォーマンス監視を有効にします
''' </summary>
Public Sub Enable()
End Sub

''' <summary>
''' パフォーマンス監視を無効にします
''' </summary>
Public Sub Disable()
End Sub