Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modErrorCodes"

' ======================
' エラーコード定義
' ======================
Public Enum ErrorCodeCategory
    ECGeneral = 1000    ' 一般エラー
    ECFileIO = 2000     ' ファイル操作エラー
    ECDatabase = 3000   ' データベースエラー
    ECNetwork = 4000    ' ネットワークエラー
    ECSystem = 5000     ' システムエラー
    ECSecurity = 6000   ' セキュリティエラー
End Enum

Public Enum ErrorCode
    ' 一般エラー (1000-1999)
    ErrUnexpected = vbObjectError + 1000             ' 予期せぬエラー
    ErrInvalidInput = vbObjectError + 1001           ' 無効な入力
    
    ' ファイル操作エラー (2000-2999)
    ErrFileNotFound = vbObjectError + 2000           ' ファイルが見つからない
    ErrFileInvalidFormat = vbObjectError + 2001      ' ファイル形式エラー
    ErrFileAccessDenied = vbObjectError + 2002       ' アクセス拒否
    
    ' データベースエラー (3000-3999)
    ErrDbConnectionFailed = vbObjectError + 3000     ' データベース接続エラー
    ErrDbQueryFailed = vbObjectError + 3001         ' データベースクエリエラー
    
    ' ネットワークエラー (4000-4999)
    ErrNetworkError = vbObjectError + 4000          ' ネットワークエラー
    ErrNetworkTimeout = vbObjectError + 4001        ' タイムアウト
    
    ' システムエラー (5000-5999)
    ErrSystemOutOfMemory = vbObjectError + 5000     ' メモリ不足
    ErrSystemResourceUnavailable = vbObjectError + 5001 ' リソース利用不可
    
    ' セキュリティエラー (6000-6999)
    ErrSecurityAccessDenied = vbObjectError + 6000  ' セキュリティアクセス拒否
    ErrSecurityInvalidCredentials = vbObjectError + 6001 ' 無効な認証情報
    
    ' 暗号化エラー (7000-7099)
    ErrCryptoProviderInitFailed = vbObjectError + 7000  ' 暗号化プロバイダーの初期化失敗
    ErrCryptoNotInitialized = vbObjectError + 7001      ' 暗号化プロバイダー未初期化
    ErrCryptoKeyNotSpecified = vbObjectError + 7002     ' 暗号化キー未指定
    ErrCryptoHashCreateFailed = vbObjectError + 7003    ' ハッシュオブジェクト作成失敗
    ErrCryptoHashDataFailed = vbObjectError + 7004      ' データハッシュ化失敗
    ErrCryptoKeyDerivationFailed = vbObjectError + 7005 ' キー生成失敗
    ErrCryptoEncryptionFailed = vbObjectError + 7006    ' 暗号化失敗
    ErrCryptoDecryptionFailed = vbObjectError + 7007    ' 復号化失敗
    ErrCryptoInvalidKeyLength = vbObjectError + 7008    ' 無効なキー長
    
    ' ロック関連エラー (7100-7199)
    ErrLockMutexCreateFailed = vbObjectError + 7100     ' Mutexの作成失敗
    ErrLockAcquireFailed = vbObjectError + 7101         ' ロックの取得失敗
    ErrLockReleaseFailed = vbObjectError + 7102         ' ロックの解放失敗
    
    ' スタックトレース関連エラー (7200-7299)
    ErrStackTracePushFailed = vbObjectError + 7200      ' スタックトレースへの追加失敗
    ErrStackTracePopFailed = vbObjectError + 7201       ' スタックトレースからの取得失敗
    ErrStackTraceGetFailed = vbObjectError + 7202       ' スタックトレース文字列の生成失敗
    
    ' イベントログ関連エラー (7300-7399)
    ErrEventLogCreateFailed = vbObjectError + 7300      ' イベントログオブジェクトの作成失敗
    ErrEventLogWriteFailed = vbObjectError + 7301       ' イベントログへの書き込み失敗
    ErrEventLogSourceInvalid = vbObjectError + 7302     ' イベントログソースが無効
End Enum

' ======================
' エラーカテゴリ取得
' ======================
Public Function GetErrorCategory(ByVal errCode As ErrorCode) As ErrorCodeCategory
    If errCode >= ECGeneral And errCode < ECFileIO Then
        GetErrorCategory = ECGeneral
    ElseIf errCode >= ECFileIO And errCode < ECDatabase Then
        GetErrorCategory = ECFileIO
    ElseIf errCode >= ECDatabase And errCode < ECNetwork Then
        GetErrorCategory = ECDatabase
    ElseIf errCode >= ECNetwork And errCode < ECSystem Then
        GetErrorCategory = ECNetwork
    ElseIf errCode >= ECSystem And errCode < ECSecurity Then
        GetErrorCategory = ECSystem
    ElseIf errCode >= ECSecurity Then
        GetErrorCategory = ECSecurity
    End If
End Function