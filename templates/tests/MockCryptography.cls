VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockCryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICryptography

Private mEncryptWasCalled As Boolean
Private mDecryptWasCalled As Boolean
Private mEncryptedValue As String
Private mDecryptedValue As String
Private mShouldThrowError As Boolean
Private mErrorNumber As Long
Private mErrorMessage As String

Private Sub Class_Initialize()
    ResetFlags
    mShouldThrowError = False
    mErrorNumber = 0
    mErrorMessage = ""
End Sub

'@Description("暗号化時の戻り値を設定")
Public Sub SetupEncryption(ByVal encryptedValue As String)
    mEncryptedValue = encryptedValue
End Sub

'@Description("復号化時の戻り値を設定")
Public Sub SetupDecryption(ByVal decryptedValue As String)
    mDecryptedValue = decryptedValue
End Sub

'@Description("エラーをシミュレートするように設定")
Public Sub SetupError(ByVal errorNumber As Long, ByVal errorMessage As String)
    mShouldThrowError = True
    mErrorNumber = errorNumber
    mErrorMessage = errorMessage
End Sub

'@Description("暗号化メソッドが呼び出されたかどうか")
Public Property Get EncryptWasCalled() As Boolean
    EncryptWasCalled = mEncryptWasCalled
End Property

'@Description("復号化メソッドが呼び出されたかどうか")
Public Property Get DecryptWasCalled() As Boolean
    DecryptWasCalled = mDecryptWasCalled
End Property

'@Description("フラグをリセット")
Public Sub ResetFlags()
    mEncryptWasCalled = False
    mDecryptWasCalled = False
    mEncryptedValue = ""
    mDecryptedValue = ""
End Sub

'ICryptography実装

Private Function ICryptography_Encrypt(ByVal value As String, ByVal key As String) As String
    If mShouldThrowError Then
        Err.Raise mErrorNumber, "MockCryptography", mErrorMessage
    End If
    
    mEncryptWasCalled = True
    If mEncryptedValue = "" Then
        ' デフォルトの動作：単純な文字列操作で暗号化をシミュレート
        ICryptography_Encrypt = "ENC:" & value
    Else
        ICryptography_Encrypt = mEncryptedValue
    End If
End Function

Private Function ICryptography_Decrypt(ByVal value As String, ByVal key As String) As String
    If mShouldThrowError Then
        Err.Raise mErrorNumber, "MockCryptography", mErrorMessage
    End If
    
    mDecryptWasCalled = True
    If mDecryptedValue = "" Then
        ' デフォルトの動作：単純な文字列操作で復号化をシミュレート
        If Left$(value, 4) = "ENC:" Then
            ICryptography_Decrypt = Mid$(value, 5)
        Else
            ICryptography_Decrypt = value
        End If
    Else
        ICryptography_Decrypt = mDecryptedValue
    End If
End Function

' 未使用のインターフェースメソッド
Private Function ICryptography_CreateHash() As Long
    ' テストでは使用しない
End Function

Private Sub ICryptography_DestroyHash(ByVal hHash As Long)
    ' テストでは使用しない
End Sub

Private Function ICryptography_DeriveKey(ByVal password As String, Optional ByVal salt As String) As String
    ' テストでは使用しない
End Function

Private Function ICryptography_GetCryptoContext() As Long
    ' テストでは使用しない
End Function

Private Sub ICryptography_ReleaseCryptoContext(ByVal hCryptProv As Long)
    ' テストでは使用しない
End Sub

Private Function ICryptography_HashData(ByVal hHash As Long, ByVal data() As Byte) As Boolean
    ' テストでは使用しない
End Function

Private Sub ICryptography_DestroyKey(ByVal hKey As Long)
    ' テストでは使用しない
End Sub