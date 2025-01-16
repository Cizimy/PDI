Attribute VB_Name = "modConfigManager"
Option Explicit

'*******************************************************************************
' モジュール: modConfigManager
' 目的：     アプリケーション設定の中央管理
' 作成日：   2025/01/17
'*******************************************************************************

' 設定ファイルのデフォルトパス
Private Const DEFAULT_CONFIG_PATH As String = "config/settings.json"

' 設定のキャッシュ
Private mConfigCache As Object
Private mConfigPath As String
Private mIsInitialized As Boolean

'*******************************************************************************
' 目的：    モジュールの初期化
' 引数：    configPath - 設定ファイルのパス（オプション）
' 戻り値：  なし
'*******************************************************************************
Public Sub Initialize(Optional ByVal configPath As String = "")
    If configPath = "" Then
        mConfigPath = DEFAULT_CONFIG_PATH
    Else
        mConfigPath = configPath
    End If
    
    ' 設定の読み込み
    LoadConfiguration
    mIsInitialized = True
End Sub

'*******************************************************************************
' 目的：    設定の読み込み
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Private Sub LoadConfiguration()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim jsonFile As Object
    Dim jsonText As String
    Dim scriptControl As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 設定ファイルが存在しない場合は、デフォルト設定を作成
    If Not fso.FileExists(mConfigPath) Then
        CreateDefaultConfig
    End If
    
    ' JSONファイルの読み込み
    Set jsonFile = fso.OpenTextFile(mConfigPath, 1)
    jsonText = jsonFile.ReadAll
    jsonFile.Close
    
    ' JSONのパース
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    Set mConfigCache = scriptControl.Eval("(" & jsonText & ")")
    
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 513, "modConfigManager", _
              "設定ファイルの読み込みに失敗しました: " & Err.Description
End Sub

'*******************************************************************************
' 目的：    デフォルト設定の作成
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Private Sub CreateDefaultConfig()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim jsonFile As Object
    Dim defaultConfig As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' デフォルト設定の定義
    defaultConfig = "{ " & _
        """logging"": { " & _
            """level"": ""INFO"", " & _
            """path"": ""logs/app.log"", " & _
            """maxSize"": 5242880, " & _
            """rotateCount"": 5 " & _
        "}, " & _
        """database"": { " & _
            """server"": """", " & _
            """database"": """", " & _
            """username"": """", " & _
            """password"": """" " & _
        "}, " & _
        """security"": { " & _
            """encryptionKey"": """", " & _
            """sessionTimeout"": 30 " & _
        "}, " & _
        """ui"": { " & _
            """theme"": ""default"", " & _
            """language"": ""ja"" " & _
        "} " & _
    "}"
    
    ' ディレクトリが存在しない場合は作成
    If Not fso.FolderExists(fso.GetParentFolderName(mConfigPath)) Then
        fso.CreateFolder fso.GetParentFolderName(mConfigPath)
    End If
    
    ' 設定ファイルの作成
    Set jsonFile = fso.CreateTextFile(mConfigPath, True)
    jsonFile.Write defaultConfig
    jsonFile.Close
    
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 514, "modConfigManager", _
              "デフォルト設定の作成に失敗しました: " & Err.Description
End Sub

'*******************************************************************************
' 目的：    設定値の取得
' 引数：    section - セクション名
'           key - キー名
'           defaultValue - デフォルト値（オプション）
' 戻り値：  設定値
'*******************************************************************************
Public Function GetValue(ByVal section As String, _
                        ByVal key As String, _
                        Optional ByVal defaultValue As Variant = Null) As Variant
    
    On Error GoTo ErrorHandler
    
    ' 初期化確認
    If Not mIsInitialized Then Initialize
    
    ' 設定値の取得
    If IsObject(mConfigCache(section)) Then
        If IsEmpty(mConfigCache(section)(key)) Then
            GetValue = defaultValue
        Else
            GetValue = mConfigCache(section)(key)
        End If
    Else
        GetValue = defaultValue
    End If
    
    Exit Function

ErrorHandler:
    GetValue = defaultValue
End Function

'*******************************************************************************
' 目的：    設定値の設定
' 引数：    section - セクション名
'           key - キー名
'           value - 設定値
' 戻り値：  なし
'*******************************************************************************
Public Sub SetValue(ByVal section As String, _
                   ByVal key As String, _
                   ByVal value As Variant)
                   
    On Error GoTo ErrorHandler
    
    ' 初期化確認
    If Not mIsInitialized Then Initialize
    
    ' 設定値の更新
    If IsObject(mConfigCache(section)) Then
        mConfigCache(section)(key) = value
    Else
        Err.Raise vbObjectError + 515, "modConfigManager", _
                  "無効なセクション名です: " & section
    End If
    
    ' 設定の保存
    SaveConfiguration
    
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 516, "modConfigManager", _
              "設定値の更新に失敗しました: " & Err.Description
End Sub

'*******************************************************************************
' 目的：    設定の保存
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Private Sub SaveConfiguration()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim jsonFile As Object
    Dim jsonText As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 設定をJSONに変換
    jsonText = ConvertToJson(mConfigCache)
    
    ' ファイルに保存
    Set jsonFile = fso.CreateTextFile(mConfigPath, True)
    jsonFile.Write jsonText
    jsonFile.Close
    
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 517, "modConfigManager", _
              "設定の保存に失敗しました: " & Err.Description
End Sub

'*******************************************************************************
' 目的：    オブジェクトをJSON形式に変換
' 引数：    obj - 変換対象のオブジェクト
' 戻り値：  JSON文字列
'*******************************************************************************
Private Function ConvertToJson(ByVal obj As Object) As String
    On Error Resume Next
    
    Dim key As Variant
    Dim item As Variant
    Dim result As String
    Dim isFirst As Boolean
    
    result = "{"
    isFirst = True
    
    ' オブジェクトのプロパティをループ
    For Each key In obj
        If Not isFirst Then
            result = result & ", "
        End If
        
        result = result & """" & key & """: "
        
        ' 値の型に応じた処理
        If IsObject(obj(key)) Then
            result = result & ConvertToJson(obj(key))
        ElseIf IsNull(obj(key)) Then
            result = result & "null"
        ElseIf IsNumeric(obj(key)) Then
            result = result & CStr(obj(key))
        Else
            result = result & """" & Replace(obj(key), """", "\""") & """"
        End If
        
        isFirst = False
    Next key
    
    result = result & "}"
    ConvertToJson = result
End Function

'*******************************************************************************
' 目的：    セクションの存在確認
' 引数：    section - セクション名
' 戻り値：  存在する場合はTrue
'*******************************************************************************
Public Function HasSection(ByVal section As String) As Boolean
    On Error Resume Next
    
    ' 初期化確認
    If Not mIsInitialized Then Initialize
    
    HasSection = IsObject(mConfigCache(section))
End Function

'*******************************************************************************
' 目的：    キーの存在確認
' 引数：    section - セクション名
'           key - キー名
' 戻り値：  存在する場合はTrue
'*******************************************************************************
Public Function HasKey(ByVal section As String, ByVal key As String) As Boolean
    On Error Resume Next
    
    ' 初期化確認
    If Not mIsInitialized Then Initialize
    
    If HasSection(section) Then
        HasKey = Not IsEmpty(mConfigCache(section)(key))
    Else
        HasKey = False
    End If
End Function
