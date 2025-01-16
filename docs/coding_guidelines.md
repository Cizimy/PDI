# VBAコーディング規約

## 命名規則

### モジュール
- 標準モジュール: `mod〈機能名〉`
  例: `modDatabase`, `modFileHandler`
- クラスモジュール: `cls〈クラス名〉`
  例: `clsCustomer`, `clsOrder`
- フォーム: `frm〈画面名〉`
  例: `frmMain`, `frmSettings`

### 変数
```vb
' プライベート変数: 先頭小文字のキャメルケース
Dim firstName As String

' パブリック変数: 先頭大文字のパスカルケース
Public UserName As String

' 定数: すべて大文字（アンダースコア区切り）
Const MAX_RETRY_COUNT As Integer = 3
```

### プロシージャ
- サブプロシージャ: 動詞で開始
  ```vb
  Public Sub ProcessData()
  Public Sub UpdateCustomerInfo()
  ```
- 関数: 戻り値を示す名前
  ```vb
  Public Function GetCustomerName() As String
  Public Function IsValidDate(ByVal testDate As Date) As Boolean
  ```

## コメント規則

### プロシージャヘッダー
```vb
'*******************************************************************************
' 目的：    顧客データを処理し、結果をシートに出力する
' 引数：    customerID - 顧客ID
' 戻り値：  処理成功時 True, 失敗時 False
' 作成者：  作成者名
' 作成日：  YYYY/MM/DD
' 更新履歴：YYYY/MM/DD 更新内容
'*******************************************************************************
```

### インラインコメント
```vb
' 重要な処理の前には説明コメントを入れる
If condition Then
    ' エラー処理
    On Error GoTo ErrorHandler
End If
```

## エラー処理

```vb
Public Sub ExampleErrorHandling()
    On Error GoTo ErrorHandler
    
    ' メイン処理
    
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Exit Sub
End Sub
```

## コードの構造化

### モジュールの構成
1. Option宣言
2. 定数定義
3. 型定義
4. 変数定義
5. プロパティ
6. イベントハンドラ
7. パブリックメソッド
8. プライベートメソッド

### インデント
- 4スペースを使用
- タブ文字は使用しない

## パフォーマンス最適化

### ベストプラクティス
- ループ内でのシート操作を避ける
  ```vb
  Application.ScreenUpdating = False
  ' ループ処理
  Application.ScreenUpdating = True
  ```
- 配列を使用してシートデータを一括処理

### メモリ管理
- オブジェクト変数は明示的に解放
  ```vb
  Dim wb As Workbook
  Set wb = Nothing
  ```

## セキュリティ

- ユーザー入力は必ず検証
- 機密情報はワークシートに直接保存しない
- マクロのデジタル署名を使用

## バージョン管理

- モジュールヘッダーにバージョン情報を記載
- 重要な変更は必ずコメントに記録
- エクスポートしたコードファイルをGitで管理
