VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー関連の列挙型と定数を定義するクラス")

'@Description("システムで発生する可能性のあるエラーコードを定義する列挙型")
Public Enum ErrorCode
    ecNoError = 0
    ecInvalidInput = 1      ' 不正な入力
    ecDatabaseError = 2     ' データベースエラー
    ecNetworkError = 3      ' ネットワークエラー
    ecFileNotFound = 4      ' ファイルが見つからない
    ecFileAccessDenied = 5  ' ファイルアクセス拒否
    ecOutOfMemory = 6      ' メモリ不足
    ecTimeout = 7          ' タイムアウト
    ecDivisionByZero = 8   ' ゼロ除算
    ecInvalidOperation = 9  ' 不正な操作
    ecNotImplemented = 10   ' 未実装
    ecNotSupported = 11     ' サポートされていない
    ecInvalidState = 12     ' 不正な状態
    ecValidationError = 13  ' 検証エラー
    ecConfigurationError = 14 ' 設定エラー
    ecSecurityError = 15    ' セキュリティエラー
    ecSystemError = 16      ' システムエラー
    ecBusinessError = 17    ' ビジネスロジックエラー
    ecCommunicationError = 18 ' 通信エラー
    ecResourceError = 19    ' リソースエラー
    ecUnknownError = 999    ' 不明なエラー
End Enum

'@Description("エラーの重大度を定義する列挙型")
Public Enum ErrorSeverity
    esInfo = 0      ' 情報
    esWarning = 1   ' 警告
    esError = 2     ' エラー
    esCritical = 3  ' 重大
    esFatal = 4     ' 致命的
End Enum

'@Description("エラーの状態を定義する列挙型")
Public Enum ErrorStatus
    esNew = 0           ' 新規
    esPending = 1       ' 保留中
    esInProgress = 2    ' 処理中
    esRetrying = 3      ' リトライ中
    esResolved = 4      ' 解決済み
    esFailed = 5        ' 失敗
    esIgnored = 6       ' 無視
    esCancelled = 7     ' キャンセル
End Enum

'@Description("エラーの分類を定義する列挙型")
Public Enum ErrorCategory
    ecInputError = 0         ' 入力エラー
    ecSystemError = 1        ' システムエラー
    ecDatabaseError = 2      ' データベースエラー
    ecNetworkError = 3       ' ネットワークエラー
    ecFileIOError = 4        ' ファイル入出力エラー
    ecLogicError = 5         ' ロジックエラー
    ecSecurityError = 6      ' セキュリティエラー
    ecUserInterfaceError = 7 ' ユーザーインターフェースエラー
    ecConfigurationError = 8  ' 設定エラー
    ecExternalSystemError = 9 ' 外部システムエラー
    ecUnknownCategory = 999  ' 不明なカテゴリ
End Enum

'@Description("エラーコードの説明を取得します")
Public Function GetErrorCodeDescription(ByVal errorCode As ErrorCode) As String
    Select Case errorCode
        Case ecNoError: GetErrorCodeDescription = "エラーなし"
        Case ecInvalidInput: GetErrorCodeDescription = "不正な入力"
        Case ecDatabaseError: GetErrorCodeDescription = "データベースエラー"
        Case ecNetworkError: GetErrorCodeDescription = "ネットワークエラー"
        Case ecFileNotFound: GetErrorCodeDescription = "ファイルが見つからない"
        Case ecFileAccessDenied: GetErrorCodeDescription = "ファイルアクセス拒否"
        Case ecOutOfMemory: GetErrorCodeDescription = "メモリ不足"
        Case ecTimeout: GetErrorCodeDescription = "タイムアウト"
        Case ecDivisionByZero: GetErrorCodeDescription = "ゼロ除算"
        Case ecInvalidOperation: GetErrorCodeDescription = "不正な操作"
        Case ecNotImplemented: GetErrorCodeDescription = "未実装"
        Case ecNotSupported: GetErrorCodeDescription = "サポートされていない"
        Case ecInvalidState: GetErrorCodeDescription = "不正な状態"
        Case ecValidationError: GetErrorCodeDescription = "検証エラー"
        Case ecConfigurationError: GetErrorCodeDescription = "設定エラー"
        Case ecSecurityError: GetErrorCodeDescription = "セキュリティエラー"
        Case ecSystemError: GetErrorCodeDescription = "システムエラー"
        Case ecBusinessError: GetErrorCodeDescription = "ビジネスロジックエラー"
        Case ecCommunicationError: GetErrorCodeDescription = "通信エラー"
        Case ecResourceError: GetErrorCodeDescription = "リソースエラー"
        Case ecUnknownError: GetErrorCodeDescription = "不明なエラー"
        Case Else: GetErrorCodeDescription = "未定義のエラー"
    End Select
End Function

'@Description("エラーの重大度の説明を取得します")
Public Function GetSeverityDescription(ByVal severity As ErrorSeverity) As String
    Select Case severity
        Case esInfo: GetSeverityDescription = "情報"
        Case esWarning: GetSeverityDescription = "警告"
        Case esError: GetSeverityDescription = "エラー"
        Case esCritical: GetSeverityDescription = "重大"
        Case esFatal: GetSeverityDescription = "致命的"
        Case Else: GetSeverityDescription = "不明な重大度"
    End Select
End Function

'@Description("エラーの状態の説明を取得します")
Public Function GetStatusDescription(ByVal status As ErrorStatus) As String
    Select Case status
        Case esNew: GetStatusDescription = "新規"
        Case esPending: GetStatusDescription = "保留中"
        Case esInProgress: GetStatusDescription = "処理中"
        Case esRetrying: GetStatusDescription = "リトライ中"
        Case esResolved: GetStatusDescription = "解決済み"
        Case esFailed: GetStatusDescription = "失敗"
        Case esIgnored: GetStatusDescription = "無視"
        Case esCancelled: GetStatusDescription = "キャンセル"
        Case Else: GetStatusDescription = "不明な状態"
    End Select
End Function

'@Description("エラーのカテゴリの説明を取得します")
Public Function GetCategoryDescription(ByVal category As ErrorCategory) As String
    Select Case category
        Case ecInputError: GetCategoryDescription = "入力エラー"
        Case ecSystemError: GetCategoryDescription = "システムエラー"
        Case ecDatabaseError: GetCategoryDescription = "データベースエラー"
        Case ecNetworkError: GetCategoryDescription = "ネットワークエラー"
        Case ecFileIOError: GetCategoryDescription = "ファイル入出力エラー"
        Case ecLogicError: GetCategoryDescription = "ロジックエラー"
        Case ecSecurityError: GetCategoryDescription = "セキュリティエラー"
        Case ecUserInterfaceError: GetCategoryDescription = "ユーザーインターフェースエラー"
        Case ecConfigurationError: GetCategoryDescription = "設定エラー"
        Case ecExternalSystemError: GetCategoryDescription = "外部システムエラー"
        Case ecUnknownCategory: GetCategoryDescription = "不明なカテゴリ"
        Case Else: GetCategoryDescription = "未定義のカテゴリ"
    End Select
End Function