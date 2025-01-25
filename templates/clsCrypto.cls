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

' Windows API宣言
Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As LongPtr, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal Algid As Long, _
    ByVal hKey As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef phHash As LongPtr) As Long

Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" ( _
    ByVal hHash As LongPtr) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal dwFlags As Long) As Long

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

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    InitializeCrypto
End Sub

Private Sub Class_Terminate()
    If mInitialized Then
        CryptReleaseContext mCryptoProvider, 0
    End If
End Sub

' ======================
' 内部メソッド
' ======================
Private Sub InitializeCrypto()
    If CryptAcquireContext(mCryptoProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        mInitialized = True
    Else
        Err.Raise vbObjectError + 1, "clsCrypto", "暗号化プロバイダーの初期化に失敗しました。"
    End If
End Sub

' ======================
' パブリックメソッド
' ======================
Public Function EncryptString(ByVal plainText As String, ByVal key As String) As String
    If Not mInitialized Then
        Err.Raise vbObjectError + 2, "clsCrypto", "暗号化プロバイダーが初期化されていません。"
    End If
    
    ' 単純なXOR暗号化（実際のプロダクションでは、より強力な暗号化アルゴリズムを使用すべき）
    Dim result As String
    Dim i As Long, j As Long
    Dim keyLen As Long
    
    keyLen = Len(key)
    If keyLen = 0 Then
        Err.Raise vbObjectError + 3, "clsCrypto", "暗号化キーが指定されていません。"
    End If
    
    result = String$(Len(plainText), 0)
    For i = 1 To Len(plainText)
        j = ((i - 1) Mod keyLen) + 1
        Mid$(result, i, 1) = Chr$(Asc(Mid$(plainText, i, 1)) Xor Asc(Mid$(key, j, 1)))
    Next i
    
    ' Base64エンコード
    EncryptString = Base64Encode(result)
End Function

Public Function DecryptString(ByVal cipherText As String, ByVal key As String) As String
    If Not mInitialized Then
        Err.Raise vbObjectError + 2, "clsCrypto", "暗号化プロバイダーが初期化されていません。"
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
        Err.Raise vbObjectError + 3, "clsCrypto", "暗号化キーが指定されていません。"
    End If
    
    result = String$(Len(decoded), 0)
    For i = 1 To Len(decoded)
        j = ((i - 1) Mod keyLen) + 1
        Mid$(result, i, 1) = Chr$(Asc(Mid$(decoded, i, 1)) Xor Asc(Mid$(key, j, 1)))
    Next i
    
    DecryptString = result
End Function

Public Function GenerateHash(ByVal inputString As String) As String
    If Not mInitialized Then
        Err.Raise vbObjectError + 2, "clsCrypto", "暗号化プロバイダーが初期化されていません。"
    End If
    
    Dim hHash As LongPtr
    Dim hashLen As Long
    Dim hashValue() As Byte
    Dim result As String
    
    ' ハッシュオブジェクトの作成
    If CryptCreateHash(mCryptoProvider, CALG_SHA_256, 0, 0, hHash) = 0 Then
        Err.Raise vbObjectError + 4, "clsCrypto", "ハッシュオブジェクトの作成に失敗しました。"
    End If
    
    ' データのハッシュ化
    If CryptHashData(hHash, ByVal inputString, Len(inputString), 0) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError + 5, "clsCrypto", "データのハッシュ化に失敗しました。"
    End If
    
    ' ハッシュサイズの取得
    If CryptGetHashParam(hHash, HP_HASHSIZE, hashLen, 4, 0) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError + 6, "clsCrypto", "ハッシュサイズの取得に失敗しました。"
    End If
    
    ' ハッシュ値の取得
    ReDim hashValue(0 To hashLen - 1)
    If CryptGetHashParam(hHash, HP_HASHVAL, hashValue(0), hashLen, 0) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError + 7, "clsCrypto", "ハッシュ値の取得に失敗しました。"
    End If
    
    ' ハッシュオブジェクトの破棄
    CryptDestroyHash hHash
    
    ' バイト配列を16進数文字列に変換
    result = ""
    Dim i As Long
    For i = 0 To hashLen - 1
        result = result & Right$("0" & Hex$(hashValue(i)), 2)
    Next i
    
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
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateProvider() As Boolean
        ValidateProvider = mInitialized And mCryptoProvider <> 0
    End Function
    
    Public Sub ResetProvider()
        If mInitialized Then
            CryptReleaseContext mCryptoProvider, 0
            mInitialized = False
        End If
        InitializeCrypto
    End Sub
#End If