VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileNotFoundErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileNotFoundErrorHandler"

' ======================
' プライベート変数
' ======================
Private mLock As ILock
Private mLogger As ILogger
Private mEmergencyLogger As IEmergencyLogger
Private mUserNotifier As IUserNotifier
Private mFileOperations As IFileOperations

Implements IErrorHandler

' =====================
' コンストラクタ
' =====================
Public Function Create( _
    ByVal lock As ILock, _
    ByVal logger As ILogger, _
    ByVal emergencyLogger As IEmergencyLogger, _
    ByVal userNotifier As IUserNotifier, _
    ByVal fileOperations As IFileOperations _
) As FileNotFoundErrorHandler
    
    Dim instance As New FileNotFoundErrorHandler
    
    ' 依存性の注入
    With instance
        If lock Is Nothing Then
            Err.Raise 5, "FileNotFoundErrorHandler.Create", "ILock is required"
        End If
        Set .Lock = lock
        
        If logger Is Nothing Then
            Err.Raise 5, "FileNotFoundErrorHandler.Create", "ILogger is required"
        End If
        Set .Logger = logger
        
        If emergencyLogger Is Nothing Then
            Err.Raise 5, "FileNotFoundErrorHandler.Create", "IEmergencyLogger is required"
        End If
        Set .EmergencyLogger = emergencyLogger
        
        If userNotifier Is Nothing Then
            Err.Raise 5, "FileNotFoundErrorHandler.Create", "IUserNotifier is required"
        End If
        Set .UserNotifier = userNotifier
        
        If fileOperations Is Nothing Then
            Err.Raise 5, "FileNotFoundErrorHandler.Create", "IFileOperations is required"
        End If
        Set .FileOperations = fileOperations
    End With
    
    Set Create = instance
End Function

' プロパティ
Public Property Set Lock(ByVal value As ILock)
    Set mLock = value
End Property

Public Property Set Logger(ByVal value As ILogger)
    Set mLogger = value
End Property

Public Property Set EmergencyLogger(ByVal value As IEmergencyLogger)
    Set mEmergencyLogger = value
End Property

Public Property Set UserNotifier(ByVal value As IUserNotifier)
    Set mUserNotifier = value
End Property

Public Property Set FileOperations(ByVal value As IFileOperations)
    Set mFileOperations = value
End Property

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo) As Boolean
    Dim proceed As Boolean
    proceed = False  ' デフォルトでは処理を中断
    
    mLock.AcquireLock
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' エラー情報の補完
    If errorDetail.HasAdditionalInfo("FilePath") Then
        Dim filePath As String
        filePath = errorDetail.GetAdditionalInfo("FilePath")
        
        ' 代替ファイルパスの確認
        Dim alternativePath As String
        alternativePath = FindAlternativeFile(filePath)
        
        If alternativePath <> "" Then
            errorDetail.AddAdditionalInfo "AlternativeFilePath", alternativePath
            errorDetail.RecoveryAttempted = True
            errorDetail.RecoverySuccessful = True
            proceed = True
        Else
            ' バックアップファイルの確認
            Dim backupPath As String
            backupPath = CheckBackupFile(filePath)
            
            If backupPath <> "" Then
                errorDetail.AddAdditionalInfo "BackupFilePath", backupPath
                errorDetail.RecoveryAttempted = True
                errorDetail.RecoverySuccessful = True
                proceed = True
            End If
        End If
    End If
    
    mLock.ReleaseLock
    
    ' ユーザーに通知
    NotifyUser errorDetail, IIf(proceed, vbInformation, vbExclamation)
    
    IErrorHandler_HandleError = proceed
End Function

' ======================
' プライベートメソッド
' ======================
Private Function FindAlternativeFile(ByVal originalPath As String) As String
    On Error GoTo ErrorHandler
    
    ' 1. 異なる拡張子のファイルを確認
    Dim extensions() As String
    extensions = Array(".bak", ".tmp", ".old")
    
    Dim basePath As String
    basePath = Left(originalPath, InStrRev(originalPath, ".") - 1)
    
    Dim i As Long
    For i = LBound(extensions) To UBound(extensions)
        Dim altPath As String
        altPath = basePath & extensions(i)
        
        If mFileOperations.FileExists(altPath) Then
            FindAlternativeFile = altPath
            Exit Function
        End If
    Next i
    
    ' 2. 同じ名前の異なるディレクトリを確認
    Dim parentPath As String
    parentPath = Left(originalPath, InStrRev(originalPath, "\") - 1)
    
    Dim fileName As String
    fileName = Mid(originalPath, InStrRev(originalPath, "\") + 1)
    
    Dim altDirs() As String
    altDirs = Array("\backup", "\archive", "\old")
    
    For i = LBound(altDirs) To UBound(altDirs)
        altPath = parentPath & altDirs(i) & "\" & fileName
        
        If mFileOperations.FileExists(altPath) Then
            FindAlternativeFile = altPath
            Exit Function
        End If
    Next i
    
    FindAlternativeFile = ""
    Exit Function

ErrorHandler:
    FindAlternativeFile = ""
End Function

Private Function CheckBackupFile(ByVal originalPath As String) As String
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    backupPath = originalPath & ".bak"
    
    If mFileOperations.FileExists(backupPath) Then
        CheckBackupFile = backupPath
    Else
        CheckBackupFile = ""
    End If
    Exit Function

ErrorHandler:
    CheckBackupFile = ""
End Function

Private Sub LogError(ByRef errorDetail As ErrorInfo)
    On Error GoTo ErrorHandler
    
    ' 通常のログ出力
    mLogger.Log MODULE_NAME, "ファイルが見つかりません: " & errorDetail.Description, errorDetail.Code
    
    ' 重要なエラーは緊急ログにも記録
    If errorDetail.Severity >= ESError Then
        mEmergencyLogger.LogEmergency MODULE_NAME, errorDetail.ToString
    End If
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ファイルエラーのログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "LogError"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub

Private Sub NotifyUser(ByRef errorDetail As ErrorInfo, ByVal style As VbMsgBoxStyle)
    On Error GoTo ErrorHandler
    
    Dim message As String
    message = "ファイルが見つかりません: " & errorDetail.Description
    
    If errorDetail.HasAdditionalInfo("AlternativeFilePath") Then
        message = message & vbCrLf & vbCrLf & _
                 "代替ファイルが見つかりました: " & _
                 errorDetail.GetAdditionalInfo("AlternativeFilePath")
    ElseIf errorDetail.HasAdditionalInfo("BackupFilePath") Then
        message = message & vbCrLf & vbCrLf & _
                 "バックアップファイルが見つかりました: " & _
                 errorDetail.GetAdditionalInfo("BackupFilePath")
    End If
    
    mUserNotifier.Notify errorDetail, style, "ファイルエラー"
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ファイルエラーの通知中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "NotifyUser"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub