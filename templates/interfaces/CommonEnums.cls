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

'@Description("システム全体で使用される共通の列挙型を定義するクラス")

'@Description("ログレベルを定義する列挙型")
'@Remarks("ログの重要度を表し、フィルタリングやハンドリングの基準として使用")
Public Enum LogLevel
    llTrace = 0    ' 最も詳細なデバッグ情報（開発時のみ使用）
    llDebug = 1    ' デバッグ情報（開発・テスト環境で使用）
    llInfo = 2     ' 一般的な情報（通常の操作ログ）
    llWarning = 3  ' 警告（潜在的な問題）
    llError = 4    ' エラー（回復可能な問題）
    llFatal = 5    ' 致命的なエラー（システム停止レベルの問題）
End Enum

'@Description("エラー通知レベルを定義する列挙型")
'@Remarks("エラーの重大度に応じた通知方法の選択に使用")
Public Enum ErrorNotificationLevel
    enlNone = 0      ' 通知なし（ログも記録しない）
    enlSilent = 1    ' サイレント（ログのみ記録）
    enlInfo = 2      ' 情報（ユーザーに情報として通知）
    enlWarning = 3   ' 警告（ユーザーに警告として通知）
    enlError = 4     ' エラー（ユーザーにエラーとして通知）
    enlCritical = 5  ' 重大（即時通知が必要な重大な問題）
End Enum

'@Description("バリデーション優先度を定義する列挙型")
'@Remarks("バリデーションルールの実行順序の制御に使用")
Public Enum ValidationPriority
    vpLowest = 0    ' 最低優先度（他のすべての検証の後に実行）
    vpLow = 1       ' 低優先度（基本的な検証の後に実行）
    vpNormal = 2    ' 通常優先度（標準的な実行順序）
    vpHigh = 3      ' 高優先度（基本的な検証の前に実行）
    vpHighest = 4   ' 最高優先度（最初に実行される検証）
End Enum

'@Description("設定オプションを定義する列挙型")
'@Remarks("設定値の取得・設定時のオプションを指定")
Public Enum SettingOptions
    soNone = 0              ' オプションなし
    soEncrypt = 1           ' 設定値を暗号化する
    soValidate = 2          ' 設定値を検証する
    soLogHistory = 4        ' 設定変更履歴を記録する
    soCache = 8             ' 設定値をキャッシュする
    soEnvironmentSpecific = 16 ' 環境固有の設定として扱う
End Enum

'@Description("読み込みオプションを定義する列挙型")
'@Remarks("設定ファイルからの読み込み時のオプションを指定")
Public Enum LoadOptions
    loNone = 0              ' オプションなし
    loDecrypt = 1           ' 設定値を復号化する
    loValidate = 2          ' 設定値を検証する
    loApplyEnvironment = 4  ' 環境固有の設定を適用する
    loMerge = 8             ' 既存の設定とマージする
End Enum

'@Description("保存オプションを定義する列挙型")
'@Remarks("設定ファイルへの保存時のオプションを指定")
Public Enum SaveOptions
    soNone = 0              ' オプションなし
    soEncrypt = 1           ' 設定値を暗号化する
    soBackup = 2            ' 保存前にバックアップを作成する
    soFormat = 4            ' 設定ファイルを整形して保存する
    soAtomic = 8            ' アトミックな書き込みを行う
End Enum

'@Description("暗号化オプションを定義する列挙型")
'@Remarks("暗号化操作時のオプションを指定")
Public Enum EncryptionOptions
    eoNone = 0              ' オプションなし
    eoRotateKey = 1         ' 新しいキーで暗号化する
    eoReEncrypt = 2         ' 既存の設定を新しいキーで再暗号化する
    eoSetKeyVersion = 4     ' キーバージョンを設定する
End Enum

'@Description("環境オプションを定義する列挙型")
'@Remarks("環境設定時のオプションを指定")
Public Enum EnvironmentOptions
    enoNone = 0             ' オプションなし
    enoReload = 1           ' 環境変更時に設定を再読み込みする
    enoValidate = 2         ' 環境変更時に設定を検証する
    enoNotify = 4           ' 環境変更を通知する
End Enum

'@Description("検証オプションを定義する列挙型")
'@Remarks("設定値の検証時のオプションを指定")
Public Enum ValidationOptions
    voNone = 0              ' オプションなし
    voFull = 1              ' すべての設定を検証する
    voChangedOnly = 2       ' 変更された設定のみ検証する
    voEnvironment = 4       ' 環境固有の設定を検証する
    voDataType = 8          ' データ型の検証を行う
    voRange = 16            ' 値の範囲を検証する
    voDependency = 32       ' 依存関係を検証する
End Enum

'@Description("履歴オプションを定義する列挙型")
'@Remarks("設定履歴の取得時のオプションを指定")
Public Enum HistoryOptions
    hoNone = 0              ' オプションなし
    hoLimit = 1             ' 取得する履歴の数を制限する
    hoSort = 2              ' 履歴をソートする
    hoFilter = 4            ' 履歴をフィルタリングする
End Enum

'@Description("復元オプションを定義する列挙型")
'@Remarks("バックアップからの復元時のオプションを指定")
Public Enum RestoreOptions
    roNone = 0              ' オプションなし
    roValidate = 1          ' 復元前に設定を検証する
    roMerge = 2             ' 既存の設定とマージする
    roNotify = 4            ' 復元を通知する
End Enum

'@Description("バリデーション状態を定義する列挙型")
'@Remarks("バリデーションの現在の実行状態を表す")
Public Enum ValidationState
    vsNotStarted = 0   ' 未開始
    vsPending = 1      ' 保留中
    vsRunning = 2      ' 実行中
    vsPaused = 3       ' 一時停止
    vsCompleted = 4    ' 完了
    vsFailed = 5       ' 失敗
    vsCancelled = 6    ' キャンセル
End Enum

'@Description("パフォーマンスメトリクスの種類を定義する列挙型")
'@Remarks("システムの各種パフォーマンス指標を表す")
Public Enum MetricType
    mtResponseTime = 0   ' 応答時間
    mtThroughput = 1    ' スループット
    mtErrorRate = 2     ' エラー率
    mtResourceUsage = 3  ' リソース使用率
    mtQueueLength = 4   ' キュー長
    mtLatency = 5       ' レイテンシ
End Enum