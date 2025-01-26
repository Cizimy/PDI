VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' 定数定義
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = 2
Private Const HP_HASHSIZE As Long = 4

' メンバ変数
Private mCryptoProvider As LongPtr
Private mInitialized As Boolean
Private mAesKey As LongPtr
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mPerformanceMonitor = New clsPerformanceMonitor
    Set mLock = New clsLock
    InitializeCrypto
End Sub

Private Sub Class_Terminate()
    If mInitialized Then
        If mAesKey <> 0 Then
            CryptDestroyKey mAesKey
        End If
        CryptReleaseContext mCryptoProvider, 0
    End If
    Set mPerformanceMonitor = Nothing
    Set mLock = Nothing
End Sub

' ======================
' 内部メソッド
' ======================
Private Sub InitializeCrypto()
    If modWindowsAPI.CryptAcquireContext(mCryptoProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        mInitialized = True
    Else
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoProviderInitFailed
            .Description = "暗号化プロバイダーの初期化に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "InitializeCrypto"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
    End If
End Sub

Private Function DeriveAESKey(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' 既存のキーを破棄
    If mAesKey <> 0 Then
        CryptDestroyKey mAesKey
        mAesKey = 0
    End If
    
    ' キーのハッシュを作成
    Dim hHash As LongPtr
    If CryptCreateHash(mCryptoProvider, CALG_SHA_256, 0, 0, hHash) = 0 Then
        GoTo Cleanup
    End If
    
    ' キー文字列をハッシュ化
    If CryptHashData(hHash, ByVal key, Len(key), 0) = 0 Then
        CryptDestroyHash hHash
        GoTo Cleanup
    End If
    
    ' ハッシュからAESキーを生成
    If CryptDeriveKey(mCryptoProvider, CALG_AES_256, hHash, CRYPT_EXPORTABLE, mAesKey) = 0 Then
        CryptDestroyHash hHash
        GoTo Cleanup
    End If
    
    ' ハッシュオブジェクトの破棄
    CryptDestroyHash hHash
    
    ' CBCモードを設定
    Dim mode As Long
    mode = CRYPT_MODE_CBC
    If CryptSetKeyParam(mAesKey, KP_MODE, mode, 0) = 0 Then
        GoTo Cleanup
    End If
    
    DeriveAESKey = True
    
Cleanup:
    mLock.ReleaseLock
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrCryptoKeyDerivationFailed
        .Description = "AESキーの生成に失敗しました。"
        .Category = ECSecurity
        .Source = "clsCrypto"
        .ProcedureName = "DeriveAESKey"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
    DeriveAESKey = False
    Resume Cleanup
End Function

' ======================
' パブリックメソッド
' ======================
Public Function EncryptString(ByVal plainText As String, ByVal key As String) As String
    mPerformanceMonitor.StartMeasurement "String Encryption"
    
    mLock.AcquireLock
    
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "EncryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Encryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' キーの検証
    If Len(key) = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoKeyNotSpecified
            .Description = "空の暗号化キーが指定されました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "EncryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Encryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' AESキーの生成
    If Not DeriveAESKey(key) Then
        mPerformanceMonitor.EndMeasurement "String Encryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' データの準備
    Dim dataLen As Long
    Dim blockSize As Long
    blockSize = 16 ' AESブロックサイズ
    
    ' パディングを含むデータ長を計算
    dataLen = Len(plainText)
    If (dataLen Mod blockSize) <> 0 Then
        dataLen = dataLen + (blockSize - (dataLen Mod blockSize))
    End If
    
    ' バッファの準備
    Dim buffer() As Byte
    ReDim buffer(0 To dataLen - 1)
    CopyMemory buffer(0), ByVal StrPtr(plainText), Len(plainText) * 2
    
    ' 暗号化
    If CryptEncrypt(mAesKey, 0, True, 0, buffer(0), dataLen, dataLen) = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoEncryptionFailed
            .Description = "データの暗号化に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "EncryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Encryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    Dim result As String
    result = Base64Encode(buffer)
    
    mPerformanceMonitor.EndMeasurement "String Encryption"
    mLock.ReleaseLock
    EncryptString = result
End Function

Public Function DecryptString(ByVal cipherText As String, ByVal key As String) As String
    mPerformanceMonitor.StartMeasurement "String Decryption"
    
mLock.AcquireLock
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "DecryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Decryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' キーの検証
    If Len(key) = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoKeyNotSpecified
            .Description = "空の暗号化キーが指定されました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "DecryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Decryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' AESキーの生成
    If Not DeriveAESKey(key) Then
        mPerformanceMonitor.EndMeasurement "String Decryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' Base64デコード
    Dim buffer() As Byte
    buffer = Base64Decode(cipherText)
   
     ' データ長の取得
    Dim dataLen As Long
    dataLen = UBound(buffer) - LBound(buffer) + 1
    ' 復号化
    If CryptDecrypt(mAesKey, 0, True, 0, buffer(0), dataLen) = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoDecryptionFailed
            .Description = "データの復号化に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "DecryptString"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Decryption"
        mLock.ReleaseLock
        Exit Function
    End If
    
    Dim result As String
    result = Left$(StrConv(buffer, vbUnicode), dataLen)
    
    mPerformanceMonitor.EndMeasurement "String Decryption"
    mLock.ReleaseLock
    DecryptString = result
End Function

Public Function GenerateHash(ByVal inputString As String) As String
    mPerformanceMonitor.StartMeasurement "Hash Generation"
    mLock.AcquireLock
    
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
        mLock.ReleaseLock
        Exit Function
    End If
    
    Dim hHash As LongPtr
    Dim hashLen As Long
    Dim hashValue() As Byte
    Dim result As String
    
    ' ハッシュオブジェクトの作成
    If modWindowsAPI.CryptCreateHash(mCryptoProvider, CALG_SHA_256, 0, 0, hHash) = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoHashCreateFailed
            .Description = "ハッシュオブジェクトの作成に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' データのハッシュ化
    If modWindowsAPI.CryptHashData(hHash, ByVal inputString, Len(inputString), 0) = 0 Then
        modWindowsAPI.CryptDestroyHash hHash
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoHashDataFailed
            .Description = "データのハッシュ化に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' ハッシュサイズの取得
    If modWindowsAPI.CryptGetHashParam(hHash, HP_HASHSIZE, hashLen, 4, 0) = 0 Then
        modWindowsAPI.CryptDestroyHash hHash
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoHashDataFailed
            .Description = "ハッシュサイズの取得に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' ハッシュ値の取得
    ReDim hashValue(0 To hashLen - 1)
    If modWindowsAPI.CryptGetHashParam(hHash, HP_HASHVAL, hashValue(0), hashLen, 0) = 0 Then
        modWindowsAPI.CryptDestroyHash hHash
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoHashDataFailed
            .Description = "ハッシュ値の取得に失敗しました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
        mLock.ReleaseLock
        Exit Function
    End If
    
    ' ハッシュオブジェクトの破棄
    modWindowsAPI.CryptDestroyHash hHash
    
    ' バイト配列を16進数文字列に変換
    result = ""
    Dim i As Long
    For i = 0 To hashLen - 1
        result = result & Right$("0" & Hex$(hashValue(i)), 2)
    Next i
    
    mPerformanceMonitor.EndMeasurement "Hash Generation"
    mLock.ReleaseLock
    GenerateHash = result
    
End Function

' ======================
' ユーティリティメソッド
' ======================
Private Function Base64Encode(ByVal text As String) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = StringToBytes(text)
    
    Base64Encode = xmlNode.text
    
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function

Private Function Base64Decode(ByVal base64 As String) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64
    
    Base64Decode = BytesToString(xmlNode.nodeTypedValue)
    
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function

Private Function StringToBytes(ByVal text As String) As Byte()
    Dim bytes() As Byte
    bytes = text
    StringToBytes = bytes
End Function

Private Function BytesToString(ByRef bytes() As Byte) As String
    BytesToString = bytes
End Function

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    Private Function ValidateProvider() As Boolean
        ' 暗号化プロバイダーの状態を検証（テスト用）
        ValidateProvider = mInitialized And mCryptoProvider <> 0
    End Function
    
    Private Sub ResetProvider()
        ' 暗号化プロバイダーをリセット（テスト用）
        If mInitialized Then
            CryptReleaseContext mCryptoProvider, 0
            mInitialized = False
        End If
        InitializeCrypto
    End Sub
#End If