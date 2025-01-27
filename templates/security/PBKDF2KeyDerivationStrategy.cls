VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PBKDF2KeyDerivationStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IKeyDerivationStrategy

' 定数定義
Private Const MS_ENHANCED_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const CALG_AES_256 As Long = &H6610

' メンバ変数
Private mCryptoProvider As LongPtr
Private mInitialized As Boolean

Private Sub Class_Initialize()
    InitializeCrypto
End Sub

Private Sub Class_Terminate()
    If mInitialized Then
        CryptReleaseContext mCryptoProvider, 0
    End If
End Sub

Private Sub InitializeCrypto()
    If modWindowsAPI.CryptAcquireContext(mCryptoProvider, vbNullString, MS_ENHANCED_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        mInitialized = True
    Else
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoProviderInitFailed
            .Description = "暗号化プロバイダーの初期化に失敗しました。"
            .Category = ECSecurity
            .Source = "PBKDF2KeyDerivationStrategy"
            .ProcedureName = "InitializeCrypto"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
    End If
End Sub

Private Function IKeyDerivationStrategy_DeriveKey(ByVal password As String, ByRef salt As Variant, ByVal iterations As Long) As Byte()
    If Not mInitialized Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrCryptoNotInitialized
            .Description = "暗号化プロバイダーが初期化されていません。"
            .Category = ECSecurity
            .Source = "PBKDF2KeyDerivationStrategy"
            .ProcedureName = "DeriveKey"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errInfo
        Exit Function
    End If
    
    ' ハッシュオブジェクトの作成
    Dim hHash As LongPtr
    If CryptCreateHash(mCryptoProvider, CALG_SHA_256, 0, 0, hHash) = 0 Then
        RaiseError ErrCryptoHashCreateFailed, "ハッシュオブジェクトの作成に失敗しました。"
        Exit Function
    End If
    
    ' パスワードとソルトを結合
    Dim combinedData() As Byte
    ReDim combinedData(0 To Len(password) + UBound(salt) + 1)
    CopyMemory combinedData(0), ByVal StrPtr(password), Len(password) * 2
    CopyMemory combinedData(Len(password)), salt(0), UBound(salt) + 1
    
    ' 反復処理でキーを生成
    Dim i As Long
    For i = 1 To iterations
        If CryptHashData(hHash, combinedData(0), UBound(combinedData) + 1, 0) = 0 Then
            CryptDestroyHash hHash
            RaiseError ErrCryptoHashDataFailed, "データのハッシュ化に失敗しました。"
            Exit Function
        End If
    Next i
    
    ' ハッシュ値の取得
    Dim hashLen As Long
    If CryptGetHashParam(hHash, HP_HASHSIZE, hashLen, 4, 0) = 0 Then
        CryptDestroyHash hHash
        RaiseError ErrCryptoHashDataFailed, "ハッシュサイズの取得に失敗しました。"
        Exit Function
    End If
    
    Dim derivedKey() As Byte
    ReDim derivedKey(0 To hashLen - 1)
    If CryptGetHashParam(hHash, HP_HASHVAL, derivedKey(0), hashLen, 0) = 0 Then
        CryptDestroyHash hHash
        RaiseError ErrCryptoHashDataFailed, "ハッシュ値の取得に失敗しました。"
        Exit Function
    End If
    
    ' ハッシュオブジェクトの破棄
    CryptDestroyHash hHash
    
    IKeyDerivationStrategy_DeriveKey = derivedKey
End Function

Private Sub RaiseError(ByVal errorCode As Long, ByVal description As String)
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = errorCode
        .Description = description
        .Category = ECSecurity
        .Source = "PBKDF2KeyDerivationStrategy"
        .ProcedureName = "DeriveKey"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub