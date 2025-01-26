VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' 暗号化操作のインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、Windows CryptoAPIを使用した暗号化操作を提供します。
''' ハッシュ生成、鍵導出、データの暗号化/復号化などの機能を含みます。
''' </remarks>

''' <summary>
''' 暗号化コンテキストを取得します
''' </summary>
''' <param name="container">コンテナ名</param>
''' <param name="provider">プロバイダー名</param>
''' <param name="provType">プロバイダータイプ</param>
''' <param name="flags">フラグ</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptAcquireContext(ByVal container As String, _
                                  ByVal provider As String, _
                                  ByVal provType As Long, _
                                  ByVal flags As Long) As Boolean
End Function

''' <summary>
''' ハッシュオブジェクトを作成します
''' </summary>
''' <param name="algorithm">ハッシュアルゴリズム</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptCreateHash(ByVal algorithm As Long) As Boolean
End Function

''' <summary>
''' データのハッシュ値を計算します
''' </summary>
''' <param name="data">ハッシュ化するデータ</param>
''' <param name="dataLen">データ長</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptHashData(ByRef data As Any, _
                            ByVal dataLen As Long) As Boolean
End Function

''' <summary>
''' 鍵を導出します
''' </summary>
''' <param name="algorithm">暗号化アルゴリズム</param>
''' <param name="flags">フラグ</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptDeriveKey(ByVal algorithm As Long, _
                             ByVal flags As Long) As Boolean
End Function

''' <summary>
''' データを暗号化します
''' </summary>
''' <param name="data">暗号化するデータ</param>
''' <param name="dataLen">データ長</param>
''' <param name="bufLen">バッファ長</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptEncrypt(ByRef data As Any, _
                           ByRef dataLen As Long, _
                           ByVal bufLen As Long) As Boolean
End Function

''' <summary>
''' データを復号化します
''' </summary>
''' <param name="data">復号化するデータ</param>
''' <param name="dataLen">データ長</param>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptDecrypt(ByRef data As Any, _
                           ByRef dataLen As Long) As Boolean
End Function

''' <summary>
''' 鍵を破棄します
''' </summary>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptDestroyKey() As Boolean
End Function

''' <summary>
''' ハッシュオブジェクトを破棄します
''' </summary>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptDestroyHash() As Boolean
End Function

''' <summary>
''' 暗号化コンテキストを解放します
''' </summary>
''' <returns>成功時True、失敗時False</returns>
Public Function CryptReleaseContext() As Boolean
End Function