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
Private mPerformanceMonitor As clsPerformanceMonitor

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mPerformanceMonitor = New clsPerformanceMonitor
    InitializeCrypto
End Sub

Private Sub Class_Terminate()
    If mInitialized Then
        CryptReleaseContext mCryptoProvider, 0
    End If
    Set mPerformanceMonitor = Nothing
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
        End With
        modError.HandleError errInfo
    End If
End Sub

' ======================
' パブリックメソッド
' ======================
Public Function EncryptString(ByVal plainText As String, ByVal key As String) As String
    mPerformanceMonitor.StartMeasurement "String Encryption"
    
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "EncryptString"
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Encryption"
        Exit Function
    End If
    
    ' 単純なXOR暗号化（実際のプロダクションでは、より強力な暗号化アルゴリズムを使用すべき）
    Dim result As String
    Dim i As Long, j As Long
    Dim keyLen As Long
    
    keyLen = Len(key)
    ' 空のキーチェック
    If keyLen = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoKeyNotSpecified
            .Description = "空の暗号化キーが指定されました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "EncryptString"
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Encryption"
        Exit Function
    End If
    
    result = String$(Len(plainText), 0)
    For i = 1 To Len(plainText)
        j = ((i - 1) Mod keyLen) + 1
        Mid$(result, i, 1) = Chr$(Asc(Mid$(plainText, i, 1)) Xor Asc(Mid$(key, j, 1)))
    Next i
    
    ' Base64エンコード
    mPerformanceMonitor.EndMeasurement "String Encryption"
    EncryptString = Base64Encode(result)
End Function

Public Function DecryptString(ByVal cipherText As String, ByVal key As String) As String
    mPerformanceMonitor.StartMeasurement "String Decryption"
    
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "DecryptString"
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Decryption"
        Exit Function
    End If
    
    ' Base64デコード
    Dim decoded As String
    decoded = Base64Decode(cipherText)
    
    ' XOR復号化
    Dim result As String
    Dim i As Long, j As Long
    Dim keyLen As Long
    
    keyLen = Len(key)
    If keyLen = 0 Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoKeyNotSpecified
            .Description = "空の暗号化キーが指定されました。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "DecryptString"
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "String Decryption"
        Exit Function
    End If
    
    result = String$(Len(decoded), 0)
    For i = 1 To Len(decoded)
        j = ((i - 1) Mod keyLen) + 1
        Mid$(result, i, 1) = Chr$(Asc(Mid$(decoded, i, 1)) Xor Asc(Mid$(key, j, 1)))
    Next i
    
    mPerformanceMonitor.EndMeasurement "String Decryption"
    DecryptString = result
End Function

Public Function GenerateHash(ByVal inputString As String) As String
    mPerformanceMonitor.StartMeasurement "Hash Generation"
    
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "clsCrypto"
            .ProcedureName = "GenerateHash"
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
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
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
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
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
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
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
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
        End With
        modError.HandleError errInfo
        mPerformanceMonitor.EndMeasurement "Hash Generation"
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