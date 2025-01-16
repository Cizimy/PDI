Attribute VB_Name = "modErrorHandler"
Option Explicit

'*******************************************************************************
' モジュール: modErrorHandler
' 目的：     エラーハンドリングの中央管理システム
' 作成日：   2025/01/17
'*******************************************************************************

' エラー種別を定義
Public Enum ErrorType
    etDatabase = 1
    etFileSystem = 2
    etValidation = 3
    etSecurity = 4
    etBusiness = 5
    etUI = 6
End Enum

' エラー情報を格納する型
Private Type ErrorInfo
    Number As Long
    Description As String
    Source As String
    ErrorType As ErrorType
    Timestamp As Date
    AdditionalInfo As String
End Type

' エラーログの設定
Private Const ERROR_LOG_PATH As String = "errors.log"
Private Const MAX_LOG_SIZE As Long = 1048576 ' 1MB

'*******************************************************************************
' 目的：    エラーを処理し、ログに記録する
' 引数：    errNumber - エラー番号
'           errDesc - エラーの説明
'           errSource - エラーの発生源
'           errType - エラーの種別
'           additionalInfo - 追加情報（オプション）
' 戻り値：  なし
'*******************************************************************************
Public Sub HandleError(ByVal errNumber As Long, _
                      ByVal errDesc As String, _
                      ByVal errSource As String, _
                      ByVal errType As ErrorType, _
                      Optional ByVal additionalInfo As String = "")
    
    Dim errorInfo As ErrorInfo
    
    ' エラー情報を設定
    With errorInfo
        .Number = errNumber
        .Description = errDesc
        .Source = errSource
        .ErrorType = errType
        .Timestamp = Now
        .AdditionalInfo = additionalInfo
    End With
    
    ' エラーをログに記録
    LogError errorInfo
    
    ' エラー種別に応じた処理
    Select Case errType
        Case etDatabase
            HandleDatabaseError errorInfo
        Case etFileSystem
            HandleFileSystemError errorInfo
        Case etSecurity
            HandleSecurityError errorInfo
        Case Else
            HandleGeneralError errorInfo
    End Select
End Sub

'*******************************************************************************
' 目的：    エラー情報をログファイルに記録する
' 引数：    errorInfo - エラー情報
' 戻り値：  なし
'*******************************************************************************
Private Sub LogError(ByRef errorInfo As ErrorInfo)
    On Error Resume Next
    
    Dim fileNum As Integer
    Dim logEntry As String
    
    ' ログエントリの作成
    logEntry = Format(errorInfo.Timestamp, "yyyy-mm-dd hh:mm:ss") & vbTab & _
               errorInfo.Number & vbTab & _
               errorInfo.Source & vbTab & _
               GetErrorTypeName(errorInfo.ErrorType) & vbTab & _
               errorInfo.Description & vbTab & _
               errorInfo.AdditionalInfo
               
    ' ログファイルへの書き込み
    fileNum = FreeFile
    Open ERROR_LOG_PATH For Append As fileNum
    Print #fileNum, logEntry
    Close fileNum
End Sub

'*******************************************************************************
' 目的：    エラー種別の名称を取得する
' 引数：    errType - エラー種別
' 戻り値：  エラー種別の文字列表現
'*******************************************************************************
Private Function GetErrorTypeName(ByVal errType As ErrorType) As String
    Select Case errType
        Case etDatabase
            GetErrorTypeName = "Database"
        Case etFileSystem
            GetErrorTypeName = "FileSystem"
        Case etValidation
            GetErrorTypeName = "Validation"
        Case etSecurity
            GetErrorTypeName = "Security"
        Case etBusiness
            GetErrorTypeName = "Business"
        Case etUI
            GetErrorTypeName = "UI"
        Case Else
            GetErrorTypeName = "Unknown"
    End Select
End Function

'*******************************************************************************
' 目的：    データベースエラーの特別処理
' 引数：    errorInfo - エラー情報
' 戻り値：  なし
'*******************************************************************************
Private Sub HandleDatabaseError(ByRef errorInfo As ErrorInfo)
    ' データベース特有のエラー処理をここに実装
    MsgBox "データベースエラーが発生しました。" & vbNewLine & _
           "エラー: " & errorInfo.Description, _
           vbCritical + vbOKOnly, _
           "データベースエラー"
End Sub

'*******************************************************************************
' 目的：    ファイルシステムエラーの特別処理
' 引数：    errorInfo - エラー情報
' 戻り値：  なし
'*******************************************************************************
Private Sub HandleFileSystemError(ByRef errorInfo As ErrorInfo)
    ' ファイルシステム特有のエラー処理をここに実装
    MsgBox "ファイルシステムエラーが発生しました。" & vbNewLine & _
           "エラー: " & errorInfo.Description, _
           vbCritical + vbOKOnly, _
           "ファイルシステムエラー"
End Sub

'*******************************************************************************
' 目的：    セキュリティエラーの特別処理
' 引数：    errorInfo - エラー情報
' 戻り値：  なし
'*******************************************************************************
Private Sub HandleSecurityError(ByRef errorInfo As ErrorInfo)
    ' セキュリティ特有のエラー処理をここに実装
    MsgBox "セキュリティエラーが発生しました。" & vbNewLine & _
           "エラー: " & errorInfo.Description, _
           vbCritical + vbOKOnly, _
           "セキュリティエラー"
End Sub

'*******************************************************************************
' 目的：    一般的なエラーの処理
' 引数：    errorInfo - エラー情報
' 戻り値：  なし
'*******************************************************************************
Private Sub HandleGeneralError(ByRef errorInfo As ErrorInfo)
    ' 一般的なエラー処理をここに実装
    MsgBox "エラーが発生しました。" & vbNewLine & _
           "エラー: " & errorInfo.Description, _
           vbCritical + vbOKOnly, _
           "エラー"
End Sub
