# Powershell VBAゲッターセッター自動生成スクリプト

## 📌 概要
ユーザが入力した「プロパティ名」、「データ型」をもとに「Getter」、「Setter」を自動生成します。
データ型に応じて Set / Let を自動判定し、作成したコードはそのままクリップボードへ出力します。
記述ミス防止と作業時間削減を目的とし、開発しました。

---

## 🚀 特徴 / 機能


---

## 🛠 使用技術
- **言語**：PowerShell 5.1
- **実行環境**：Windows 10 以降  

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
データ型を入力してください（例：String、Range、Object）: テストプロシージャです。
```

---

### 3. 出力結果の確認

実行が完了すると、クリップボードに、以下の文字列が作成されます。

出力イメージ



---

### 💡 想定ユースケース


---

## 📝 ライセンス

This project is licensed under the MIT License.

You are free to use, modify, and distribute this script for personal or commercial purposes.

※ 本スクリプトは MIT License のもとで公開されています。
商用・非商用を問わず、自由に利用・改変・再配布が可能です。
