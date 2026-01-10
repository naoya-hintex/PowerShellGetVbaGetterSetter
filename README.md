# PowerShell｜VBA Property Get / Let / Set 自動生成スクリプト

## 📌 概要
本スクリプトは、ユーザーが入力した「プロパティ名」と「データ型」をもとに、
VBAの Property Get / Let / Set を自動生成する PowerShell スクリプトです。

データ型に応じて Set / Let を自動判定し、生成したコードはそのまま
クリップボードへ出力されます。

VBA開発における記述ミスの防止と、コーディング時間の短縮を目的として開発しました。

---

## 🚀 特徴 / 機能

### ・記述ミスを削減
VBEには、「プロパティ名」に基づき、「ゲッター」や「セッター」を自動生成する機能がなく、
開発者が手作業で作成する必要があります。よって、記述ミスが発生しやすいです。
「プロパティ名」、「データ型」さえ決まれば、開発作業を中断せず、スムーズにコーディングを進められます。

### ・作成したコメントは貼り付けるだけ
作成した「プロパティ」、「ゲッター」、「セッター」は自動的にクリップボードに設定されます。
クラスファイルに貼り付けるだけで、すぐに使用できます。

### ・ショートカットキーに設定して即実行
Powershellスクリプトはショートカットキーに割り当てることができます。
本スクリプトをショートカットキーに割り当てることで、
必要なタイミングですぐ「プロパティ」、「ゲッター」、「セッター」を作成できます。
日常的なVBA開発の効率化に最適です。

### ・シンプルでメンテナンスしやすい設計
可読性の高いコードとコメントが豊富なスクリプトでメンテナンスがしやすいです。
貼り付けるコードにコメントを入れることもできるので、開発ルールに合わせて調整できます。

---

## 🛠 使用技術
- **言語**：PowerShell 5.1
- **実行環境**：Windows 10 以降（「Set-Clipboard」を使用しているため、Windows限定）  

## 📂 使い方
### 1. ダウンロード・配置

「Download ZIP」より本リポジトリをダウンロードし、解凍後、任意のフォルダにスクリプトを配置してください。

例：C:\Scripts\Get-VbaGetterSetter.ps1

---

### 2. スクリプトの実行

PowerShell からスクリプトを実行します。

```powershell
.\Get-VbaGetterSetter.ps1

```

※ 実行ポリシーによりエラーが発生する場合は、以下を実行してください。

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

```

スクリプトの実行中、「プロパティ名」および「データ型」を入力するメッセージが表示されます。
入力後、Enterキーを押下してください。

プロパティ名の記入例
```powershell
プロパティ名を入力してください（例：Name）: test
```

データ型の記入例
```powershell
データ型を入力してください（例：String、Range、Object）: String
```

---

### 3. 出力結果の確認

実行が完了すると、クリップボードに、以下の文字列が作成されます。

出力イメージ

データ型が、オブジェクト型ではない（String、Long等）
```vba

'プロパティ
Private ptest As String

'セッター
Public Property Let test(Byval inputData As String)
    ptest = inputData    
End Property

'ゲッター
Public Property Get test() As String
    test = ptest
End Property

```
データ型が、オブジェクト型（Collection、Workbook等）
```vba


'プロパティ
Private psampleObject As Workbook

'セッター
Public Property Set sampleObject(Byval inputData As Workbook)
    Set psampleObject = inputData    
End Property

'ゲッター
Public Property Get sampleObject() As Workbook
    Set sampleObject = psampleObject
End Property

```

---

### 💡 想定ユースケース

- **クラスファイルを用いた、VBA開発を日常的に行っている方**

- **VBAの「ゲッター」、「セッター」を簡単に作成したいと考えている方**

- **開発作業を中断せず、スムーズにコーディングを進めたい方**

- **「ゲッター」、「セッター」の記述のばらつきを減らしたい方**

---

## 📝 ライセンス

This project is licensed under the MIT License.

You are free to use, modify, and distribute this script for personal or commercial purposes.

※ 本スクリプトは MIT License のもとで公開されています。
商用・非商用を問わず、自由に利用・改変・再配布が可能です。
