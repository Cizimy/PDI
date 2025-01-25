# VBAコーディング規約

## 命名規則

### モジュール
- 標準モジュール: `mod〈機能名〉`
  例: `modDatabase`, `modFileHandler`
- クラスモジュール: `cls〈クラス名〉`
  例: `clsCustomer`, `clsOrder`
- フォーム: `frm〈画面名〉`
  例: `frmMain`, `frmSettings`
- レポート: `rpt〈レポート名〉`
  例: `rptSalesReport`, `rptInventory`

### 変数
```vb
' プライベート変数: 先頭小文字のキャメルケース
Dim firstName As String

' パブリック変数: 先頭大文字のパスカルケース
Public UserName As String

' 定数: すべて大文字（アンダースコア区切り）
Const MAX_RETRY_COUNT As Integer = 3

' プロシージャ内変数は先頭小文字のキャメルケース
Dim localValue As Integer

' カスタムタイプ
Type typCustomerData
    ID As Long
    Name As String
End Type
```

### グローバル変数の使用制限
```vb
' 悪い例：グローバル変数の乱用
Public gCurrentUser As String
Public gIsProcessing As Boolean

' 良い例：プロパティを使用
Private mCurrentUser As String
Public Property Get CurrentUser() As String
    CurrentUser = mCurrentUser
End Property
Public Property Let CurrentUser(ByVal value As String)
    mCurrentUser = value
End Property

' 良い例：引数による値の受け渡し
Public Sub ProcessData(ByVal userName As String)
    ' 処理内容
End Sub

' やむを得ずグローバル変数を使用する場合のコメント例
''' <summary>
''' アプリケーション全体で共有する設定情報
''' グローバル変数として定義する理由：
''' - 複数フォーム間での設定共有が必要
''' - 頻繁なアクセスが発生するため、パフォーマンスを考慮
''' </summary>
Public gAppSettings As typAppSettings
```

### イベントハンドラ
```vb
' コントロール種別＋操作内容＋イベント種別
Private Sub btnSave_Click()
Private Sub txtName_Change()
```

### UIコントロール
```vb
' プレフィックス方式（3文字略称）
txtUserName   ' テキストボックス
lblStatus      ' ラベル
btnSubmit      ' ボタン
lstProducts    ' リストボックス
chkAgreement   ' チェックボックス
cmbCategory    ' コンボボックス
optPayment     ' オプションボタン
tblResults     ' テーブル
```

### 列挙型
```vb
Public Enum FileAccessMode
    famReadOnly = 1
    famReadWrite = 2
End Enum
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

## コード品質

### 必須宣言
```vb
Option Explicit  ' 全モジュールで必須
```

### マジックナンバーの回避
```vb
' 悪い例：マジックナンバーの使用
If status = 1 Then
    MsgBox "処理が完了しました"
ElseIf status = 2 Then
    MsgBox "エラーが発生しました"
End If

' 良い例：定数の使用
Private Const STATUS_SUCCESS As Long = 1
Private Const STATUS_ERROR As Long = 2

If status = STATUS_SUCCESS Then
    MsgBox "処理が完了しました"
ElseIf status = STATUS_ERROR Then
    MsgBox "エラーが発生しました"
End If

' 良い例：列挙型の使用
Public Enum ProcessStatus
    Success = 1
    Error = 2
    Pending = 3
End Enum

If status = ProcessStatus.Success Then
    MsgBox "処理が完了しました"
ElseIf status = ProcessStatus.Error Then
    MsgBox "エラーが発生しました"
End If
```

### GoTo文の使用制限
```vb
' 悪い例：制御フローにGoToを使用
If condition Then
    GoTo ProcessData
End If
' 処理
ProcessData:
    ' データ処理

' 良い例：構造化された制御フロー
If condition Then
    ProcessData
Else
    SkipProcessing
End If

' 例外：エラー処理での適切なGoTo使用
On Error GoTo ErrorHandler
    ' 処理内容
Exit Sub

ErrorHandler:
    ' エラー処理
    Resume Next
```

### 型安全な宣言
```vb
' 悪い例
Dim count, total  ' Variant型

' 良い例
Dim count As Long
Dim total As Double

' 複数変数宣言時の注意
Dim count As Long, total As Double ' 個別に型指定
```

### 配列操作ガイドライン
```vb
' 動的配列の明示的初期化
Dim data() As Variant
data = Array()

' 配列のサイズ変更は慎重に
ReDim Preserve data(1 To newSize)
```

## 外部参照とバインディング

### バインディング方針
```vb
' 基本は早期バインディングを使用
Dim excel As Excel.Application
Dim wb As Excel.Workbook

' 遅延バインディングが必要な場合は理由をコメントで明記
''' <remarks>
''' 配布環境でのExcelバージョン互換性のため遅延バインディングを使用
''' </remarks>
Dim excel As Object
```

### 参照設定のガイドライン
- プロジェクトの参照設定は明示的にバージョンを指定
- 必要最小限の参照のみを追加
- バージョン依存性の高い参照は遅延バインディングを検討

## フォーム設計ガイドライン

### ビジネスロジックの分離
```vb
' 悪い例：フォームに直接ビジネスロジックを記述
Private Sub btnSave_Click()
    ' フォーム上で直接データ処理
    Sheets("Data").Range("A1").Value = txtName.Value
End Sub

' 良い例：ビジネスロジックをクラスに分離
Private Sub btnSave_Click()
    Dim dataManager As New clsDataManager
    dataManager.SaveData GetFormData()
End Sub

Private Function GetFormData() As typFormData
    ' UIからデータを収集してクラスに渡す
End Function
```

### フォームの責務
- フォームはUIの表示と入力の受付のみを担当
- データの検証はビジネスロジッククラスで実施
- コントロールへの直接参照は最小限に抑える
- イベントハンドラはできるだけシンプルに保つ
- 複雑な処理は別クラスに委譲

## コメント規則

### XMLドキュメントコメント
```vb
''' <summary>
''' 顧客情報を更新する
''' </summary>
''' <param name="customerId">顧客ID（数値必須）</param>
''' <returns>更新成功時にTrueを返す</returns>
Public Function UpdateCustomer(ByVal customerId As Long) As Boolean
```

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

### 階層化エラーハンドリング
```vb
Public Sub MainProcedure()
    On Error GoTo GlobalHandler
    
    ' 詳細処理
    Call SubProcedure
    
    Exit Sub
    
GlobalHandler:
    Call LogError(Err.Number, Err.Description)
    Exit Sub
End Sub

Private Sub SubProcedure()
    On Error GoTo LocalHandler
    
    ' 処理内容
    
    Exit Sub
    
LocalHandler:
    ' 局所的な回復処理
    Resume Next
End Sub
```

### リソース管理
```vb
Public Sub ProcessFile()
    Dim fileNum As Integer
    Dim obj As Object
    
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open "data.txt" For Input As #fileNum
    Set obj = CreateObject("Some.Object")
    
    ' 処理内容
    
    Exit Sub

ErrorHandler:
    ' エラー処理
    
Cleanup:
    ' リソースの明示的解放
    If Not obj Is Nothing Then Set obj = Nothing
    Close #fileNum
End Sub
```

### リトライメカニズム
```vb
' リトライメカニズム
Const MAX_RETRIES = 3
Dim retryCount As Integer

On Error Resume Next
Do While retryCount < MAX_RETRIES
    ' 処理実行
    If Err.Number = 0 Then Exit Do
    retryCount = retryCount + 1
Loop
On Error GoTo 0
```

### エラーログ記録
```vb
Public Sub LogError(ByVal errNumber As Long, ByVal errDesc As String)
    ' エラー情報をファイル/DBに記録
    ' 発生時刻・プロシージャ名・モジュール名を含む
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

### 単一責任の原則（SRP）
```vb
' 悪い例：複数の責務を持つクラス
Public Class clsCustomerManager
    Public Sub SaveCustomer()
        ' 顧客データの保存処理
    End Sub
    
    Public Sub GenerateInvoice()
        ' 請求書生成処理
    End Sub
    
    Public Sub SendEmail()
        ' メール送信処理
    End Sub
End Class

' 良い例：責務を分割した設計
Public Class clsCustomerRepository
    Public Sub SaveCustomer()
        ' 顧客データの保存処理のみを担当
    End Sub
End Class

Public Class clsInvoiceGenerator
    Public Sub GenerateInvoice()
        ' 請求書生成処理のみを担当
    End Sub
End Class

Public Class clsEmailService
    Public Sub SendEmail()
        ' メール送信処理のみを担当
    End Sub
End Class

' 良い例：複雑なプロシージャの分割
' 悪い例：長大な処理を1つのプロシージャに記述
Public Sub ProcessOrder()
    ' 100行以上の処理...
End Sub

' 良い例：機能ごとに分割
Public Sub ProcessOrder()
    ValidateOrder
    CalculateTotal
    UpdateInventory
    SendConfirmation
End Sub

Private Sub ValidateOrder()
    ' 注文の検証処理
End Sub

Private Sub CalculateTotal()
    ' 合計金額の計算
End Sub

Private Sub UpdateInventory()
    ' 在庫の更新
End Sub

Private Sub SendConfirmation()
    ' 確認メールの送信
End Sub
```

### インデント
- 4スペースを使用
- タブ文字は使用しない

## パフォーマンス最適化

### 画面更新とイベント制御
```vb
' 画面更新の制御
Application.ScreenUpdating = False
' 処理実行
Application.ScreenUpdating = True

' イベントハンドリングの無効化
Application.EnableEvents = False
' 処理実行
Application.EnableEvents = True
```

### 計算設定制御
```vb
Application.Calculation = xlManual
' 重い処理実行
Application.Calculation = xlAutomatic
```

### オブジェクト操作ベストプラクティス
```vb
With Worksheets("Data")
    .Range(.Cells(1,1), .Cells(100,10)).Value = dataArray
End With

' Select/Activateの使用禁止
' 悪い例
Range("A1").Select
Selection.Value = 10

' 良い例
Range("A1").Value = 10
```

### メモリ効率の良い操作
```vb
' セル操作は配列経由で
Dim buffer As Variant
buffer = Range("A1:Z1000").Value
' 配列操作
Range("A1:Z1000").Value = buffer
```

### 文字列操作の最適化
```vb
' 悪い例：文字列の連続結合
Dim result As String
For i = 1 To 1000
    result = result & data(i) & ","
Next i

' 良い例：配列とJoinを使用
Dim items(1 To 1000) As String
For i = 1 To 1000
    items(i) = data(i)
Next i
result = Join(items, ",")

' 大量テキスト生成時は StringBuilder パターンを使用
' または配列に格納してから Join で結合
```

### メモリ管理
```vb
Dim wb As Workbook
Set wb = Nothing
```

## セキュリティ

### パスワード管理
```vb
' 禁止事項
Const PASSWORD = "secret123"  ' コード内直書き厳禁

' 推奨方法
Function GetPassword() As String
    ' 暗号化ストレージ/資格情報管理システムから取得
End Function
```

### コード保護
- プロジェクトロックは最終手段としてのみ使用
- デジタル署名必須化
- 重要な機密処理は.NETアセンブリ化を検討

### 入力検証
```vb
' 文字列長の検証
Const MAX_INPUT_LENGTH As Long = 255
If Len(strInput) > MAX_INPUT_LENGTH Then
    Err.Raise vbObjectError + 1000, , "入力が長すぎます"
End If

' 数値範囲の検証
If value < MIN_VALUE Or value > MAX_VALUE Then
    Err.Raise vbObjectError + 1001, , "値が範囲外です"
End If
```

### SQLインジェクション対策
```vb
' 悪い例
sql = "SELECT * FROM Users WHERE Name = '" & txtName & "'"

' 良い例
sql = "SELECT * FROM Users WHERE Name = ?"
cmd.Parameters.Append cmd.CreateParameter("@name", adVarChar, adParamInput, 255, txtName)
```

## テスト・検証

### 単体テスト基準
```vb
' テストプロシージャ命名規則
Sub Test_CalculateTax()
    ' テストコード
    Debug.Assert CalculateTax(100) = 10
End Sub
```

### 入力検証パターン
```vb
Function ValidateInput(ByVal inputValue As Variant) As Boolean
    ' 型チェック
    If Not IsNumeric(inputValue) Then Exit Function
    
    ' 範囲チェック
    If inputValue < 0 Or inputValue > 100 Then Exit Function
    
    ValidateInput = True
End Function
```

## バージョン管理

### コードエクスポート規則
```
/ProjectName
  ├── /src
  │   ├── /modules
  │   ├── /classes
  │   └── /forms
  ├── /docs
  └── /tests
```

### バージョンタグ形式
```
' メジャー.マイナー.パッチ＋アルファベット識別子
v2.1.5a  ' 開発版
v2.1.5rc ' リリース候補
v2.1.5   ' 正式版
```

## 高度な設計パターン

### 依存性注入
```vb
Public Sub ProcessData(ByVal dataRepository As IDataRepository)
    ' 具象クラスに依存しない実装
End Sub
```

### 非同期処理実装
```vb
Public Sub AsyncProcess()
    Dim task As New clsAsyncTask
    task.RunAsync AddressOf LongRunningProcess
End Sub
```

## リファクタリングガイドライン

### コードメトリクス基準
```
指標           許容値
プロシージャ行数  最大50行
循環的複雑度     最大10
パラメータ数     最大5
```

### 技術的負債管理
```vb
' TODO リファクタリング必要
' HACK 暫定対処コード
' WARNING パフォーマンス問題あり
```

### 変更管理表
モジュールヘッダーに追記する項目：
- 影響範囲
- 関連チケット番号
- 技術的負債フラグ
- テストステータス
