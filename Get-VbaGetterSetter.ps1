# ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
# ファイル名    Get-VbaGetterSetter
# 概要         入力された「プロパティ名」と「データ型」をもとにVBA用の「プロパティ」、
#              「セッター」、「ゲッター」を作成し、クリップボードに設定します。
# 作成者        naoya-hintex
# 引数          なし
# 戻り値        なし
# ＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊

# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
# ユーザ入力
# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

#プロパティ名 
$propertyName = Read-Host "プロパティ名を入力してください（例：Name）"

# プロパティ名が未入力の場合、エラー終了
if ([string]::IsNullOrWhiteSpace($propertyName)) {
    Write-Host "プロパティ名が未入力です"
    return
}

# プロパティ名が英字始まりでない、英数字とアンダースコア以外の文字列が含まれている場合、エラー終了
if ($propertyName -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {
    Write-Host "プロパティ名は英字で始まり、英数字とアンダースコアのみ使用できます。"
    return
}

# プロパティ名が予約語の場合、エラー終了
$vbaReservedWords = @(
    "If","Then","Else","ElseIf","End",
    "For","Each","Next","Do","Loop","While","Wend",
    "Select","Case",
    "Dim","Private","Public","Static",
    "Set","Let","Get","New",
    "As","ByVal","ByRef",
    "Function","Sub","Property",
    "Option","Explicit",
    "True","False","Nothing"
)

if ($vbaReservedWords -contains $propertyName) {
    Write-Host "「$propertyName」はVBAの予約語のため使用できません。"
    return
}

# データ型
$dataType = Read-Host "データ型を入力してください（例：String、Range、Object）"

# データ型が未入力の場合、終了
if ([string]::IsNullOrWhiteSpace($dataType)) {
    Write-Host "データ型が未入力です"
    return
}

# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
# コード作成
# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

# ユーザ入力のデータ型が以下のオブジェクト型に一致する場合、Trueとする
$objectTypes = @("Range","Object","Collection","Variant","Workbook")
$isObject = $objectTypes -contains $dataType

# VBAの一般的なコーディング規約に基づき、プロパティ名の前に「p」をつける
$privateVar = "p" + $propertyName

# コードの作成
# isObjectがTrueなら「Set」のセッター、Falseなら「Let」のセッターを作る
$vbaCode = @"

'プロパティ
Private $privateVar As $dataType

'セッター
Public Property $(if($isObject){"Set"}else{"Let"}) $propertyName(Byval inputData As $dataType)
    $(if($isObject){"Set $privateVar = inputData"}else{"$privateVar = inputData"})    
End Property

'ゲッター
Public Property Get $propertyName() As $dataType
    $(if($isObject){"Set $propertyName = $privateVar"}else{"$propertyName = $privateVar"})
End Property

"@

# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
# 作成したコードをクリップボードに設定する
# ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

$vbaCode | Set-Clipboard
Write-Host "クリップボードに作成したコードを設定しました。"
