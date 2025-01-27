VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("ログレベルを定義する列挙型")
Public Enum LogLevel
    llTrace = 0    ' 最も詳細なデバッグ情報
    llDebug = 1    ' デバッグ情報
    llInfo = 2     ' 一般的な情報
    llWarning = 3  ' 警告
    llError = 4    ' エラー
    llFatal = 5    ' 致命的なエラー
End Enum

'@Description("エラー通知レベルを定義する列挙型")
Public Enum ErrorNotificationLevel
    enlNone = 0      ' 通知なし
    enlSilent = 1    ' サイレント（ログのみ）
    enlInfo = 2      ' 情報
    enlWarning = 3   ' 警告
    enlError = 4     ' エラー
    enlCritical = 5  ' 重大
End Enum

'@Description("バリデーション優先度を定義する列挙型")
Public Enum ValidationPriority
    vpLowest = 0    ' 最低優先度
    vpLow = 1       ' 低優先度
    vpNormal = 2    ' 通常優先度
    vpHigh = 3      ' 高優先度
    vpHighest = 4   ' 最高優先度
End Enum