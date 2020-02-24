<#
.SYNOPSIS
.DESCRIPTION
  MoneyForwardのcsvが格納されているフォルダをDIRとして定義します
  指定した期間のうち、fugaに該当する全項目を表示し、またその合計値を出力します。
  処理概要は以下の通りです
    1. MoneyForwardのcsvを読み込む（全期間。精算済み含む。）
    1. 家賃は補助の分の5000を足す
    1. fugaの全期間合計値を出す
    1. 合計値に0.33を乗ずる
.PARAMETER part
  期間を指定して集計する
.PARAMETER csvout
  デスクトップにcsvを出力する
.PARAMETER hoge
  HOGE式関連項目のみ表示する
.EXAMPLE
  Summerize-Csv_ns.ps1
  Summerize-Csv_ns.ps1 -part
  Summerize-Csv_ns.ps1 -csvout
  Summerize-Csv_ns.ps1 -hoge
.NOTES
  Author : me
.LINK
#>

Param(
  [switch]$PART,    # 期間を指定して集計する
  [switch]$CSVOUT,  # デスクトップにcsvを出力する
  [switch]$HOGE     # HOGE関連項目のみ表示する
)

[string]$script:DIR = "C:\Users\Public\Documents"
[int]$script:HOUSE_RENT_AID = 5000 # 家賃補助費。収入なので正の数。
[int]$script:seisan = 0
if($HOGE){
  [scriptblock]$script:CONDITION = {$_.Memo -like "*HOGE精算*"}
}else{
  [scriptblock]$script:CONDITION = {$_.Memo -like "*fuga*"}
}

# fnShow_Costs用変数
[object]$script:FT_Property_1=@("Date",@{Name="Yen + HouseRentAid" ; Expression={([int]$_.Yen + $HOUSE_RENT_AID)} ; Alignment="right"},"Content","Tier1","Tier2","Memo")
[object]$script:FT_Property_2=@("Date",@{Name="Yen"                ; Expression={"{0:N0}" -f [int]$_.Yen}         ; Alignment="right"},"Content","Tier1","Tier2","Memo")
$script:HASH_HOUSERENT=@{
  msg          = [string]"家賃（補助費 $HOUSE_RENT_AID を差し引いた額） :"
  condition    = [scriptblock]{$_.Content -match "yachin"}
  args_ft_prop = $FT_Property_1
  args_ft_auto = $true
}
$script:HASH_SUIDO=@{
  msg          = [string]"光熱費 :"
  condition    = [scriptblock]{$_.Tier2 -match "光熱費"}
  args_ft_prop = $FT_Property_2
  args_ft_auto = $true
  # args_ft   = [object]@{ Property = $FT_Property_2 ; Autosize = $true}
}
$script:HASH_KONETSU=@{
  msg          = [string]"水道代 :"
  condition    = [scriptblock]{$_.Tier2 -match "水道代"}
  args_ft_prop = $FT_Property_2
  args_ft_auto = $true
}
$script:HASH_KOUGAKU=@{
  msg          = [string]"絶対値5万円以上のもの（家賃は除外） :"
  condition    = [scriptblock]{[Math]::abs($_.Yen) -ge 50000 -and $_.Content -notmatch "yachin" }
  args_ft_prop = $FT_Property_2
  args_ft_auto = $true
}
$script:HASH_DISNEY=@{
  msg          = [string]"Disney :"
  condition    = [scriptblock]{$_.Content -match ".*(disney|Disney|ディズニー).*" }
  args_ft_prop = $FT_Property_2
  args_ft_auto = $true
}
$script:HASH_ALREADY_PAID=@{
  msg          = [string]"manからのお支払い :"
  condition    = [scriptblock]{$_.Content -match "^振込.*kaku.*" -or $_.Memo -like "*fuga清算*" }
  args_ft_prop = $FT_Property_2
  args_ft_auto = $true
}

function fnTest_FolderPath(){
  # 指定ディレクトリが存在しない場合はカレントディレクトリを探索する
  [bool]$local:_ret = Test-Path -Path $DIR
  if ( -not $_ret ) {
    Write-Output $DIR
    Write-Output "上記ディレクトリが存在しないため、カレントディレクトリを探索します"
    Read-Host "キーを押下すると続行します"
    $DIR=Get-Location
  }
}

function fnGet_TargetCsvs(){
  # 収入・支出詳細*csvという名前のファイル一覧を変数に格納する
  $script:TARGET_CSVS = Get-ChildItem -Path $DIR | Where-Object { $_.Name -like "収入・支出詳細*csv"} # OutputがSystem.Stringになるのでdirの-nameパラメータは使用しない
  if ($TARGET_CSVS.length -eq 0){
    Write-Output "対象ファイルが存在しません。スクリプトを終了します。(100)"
    Read-Host "キーを押下すると続行します"
    exit 100
  }

  # newest/oldestなcsvも取得する
  $script:OLDEST_CSV = $TARGET_CSVS | Sort-Object -Property Name | Select-Object -First 1
  $script:NEWEST_CSV = $TARGET_CSVS | Sort-Object -Property Name | Select-Object -Last  1
}

function fnNew_ThisCsv(){
  # ヘッダを英語名に変更したthis.csvを新規に作成する
  # 通常配列@()の中に連想配列@{}を入れている。パイプで渡す想定なのでExpressionは式{}とし、"$_."をつけている
  [string]$script:THIS_CSV = Join-Path -path $DIR -ChildPath "this.csv"
  $local:_props = @(
    @{Name="isCalc"         ; Expression={$_."計算対象"}}
    @{Name="Date"           ; Expression={$_."日付"}}
    @{Name="Content"        ; Expression={$_."内容"}}
    @{Name="Yen"            ; Expression={$_."金額（円）"}}
    @{Name="BankingFacility"; Expression={$_."保有金融機関"}}
    @{Name="Tier1"          ; Expression={$_."大項目"}}
    @{Name="Tier2"          ; Expression={$_."中項目"}}
    @{Name="Memo"           ; Expression={$_."メモ"}}
    @{Name="Transfer"       ; Expression={$_."振替"}}
    @{Name="ID"             ; Expression={$_."ID"}}
  )
  $local:_args_ec=@{
    Path      = $THIS_CSV
    Delimiter = ","
    Encoding  = "Default"
    NoTypeInformation = $true
    Force             = $true
  }

  $TARGET_CSVS | ForEach-Object {
    Import-Csv -Path $_.Fullname -Encoding Default |
    Select-Object -Property $_props
  } | Export-Csv @_args_ec
}

function fnGet_SeisanTargetAll(){
  # Memo項を見て該当する項目だけImport-csvする
  $script:allTargetItems = Import-Csv $THIS_CSV -Encoding Default | Where-Object $CONDITION
}

function fnOutput_Csv([Object]$_csvObject){
  Write-Output "summary.csvをデスクトップに出力します"
  [string]$local:_summary=-join ([Environment]::GetFolderPath('Desktop'), "\summary.csv")
  $_csvObject | Export-Csv -Path $_summary -Encoding Default -Force
}

function fnDetermine_SpecifiedTargetPeriod(){
  $isloop = $true
  while ($isLoop) {
    Write-Output "指定された期間の項目のみを集計します"

    # 指定期間探索とエラーハンドリング
    [datetime]$local:_start = Read-Host "開始年月を入力して下さい (yyyyMM)"
    [datetime]$local:_end   = Read-Host "終了年月を入力して下さい (yyyyMM)"
    trap [ArgumentTransformationMetadataException] {
      'yyyyMM型 = 年月 を入力してください。処理を終了します'
      break
    }
    if ($_end -lt $_start) {
      [string]$local:_msg1="Warn : 開始年月と終了年月が期間逆転しています"
      [string]$local:_msg2="Warn : 再度入力してください"
      Write-Output $_msg1 $_msg2
      continue
    }
    if ( ($_start -lt $OLDEST_CSV_DATE) -Or ($_end -gt $NEWEST_CSV_DATE) ) {
      [string]$local:_msg1="Warn : 指定期間のcsvが存在しません"
      [string]$local:_msg2="Warn : 再度入力してください"
      Write-Output $_msg1 $_msg2
      continue
    }

    # 開始年月と終了年月を表示
    [datetime]$script:FROM=$_start # 月初
    [datetime]$script:TILL=$_end.AddMonths(1).AddDays(-1) # 月末
    [object]$local:_msg=@(
      "Info : 開始/終了期間は以下の通りです"
      "    $($FROM.toString("yyyy/MM"))"
      "    $($TILL.toString("yyyy/MM"))"
    )
    Write-Output $_msg
    Read-Host "Info : キーを押下すると処理を続行します"

    # ループを抜ける
    $isLoop=$false
  }
}

function fnDetermine_TargetPeriod(){
  # 一番古い日付のcsv、一番新しい日付のcsvのdateを取得する
  [datetime]$script:OLDEST_CSV_DATE=$OLDEST_CSV.Name.Split("_")[1]
  [datetime]$script:NEWEST_CSV_DATE=$NEWEST_CSV.Name.Split("_")[1]
  if($OLDEST_CSV_DATE.Day -ne 1){ exit 100 } # なんかフォーマットがおかしい場合はエラーにする
  if($NEWEST_CSV_DATE.Day -ne 1){ exit 100 } # ditto
  
  [datetime]$script:FROM=$OLDEST_CSV_DATE # 月初
  [datetime]$script:TILL=$NEWEST_CSV_DATE.AddMonths(1).AddDays(-1) # 月末
}

function fnGet_ThisItems(){
  # 対象期間のItemをのみを抜き出す
  $script:THIS_SEISAN_ITEMS = $allTargetItems | Where-Object {
    [datetime]$_.Date -ge $FROM -and
    [datetime]$_.Date -le $TILL
  }
}

function fnShow_Costs([hashtable]$_HASH){
  $local:_args=@{
    Property = $_HASH.args_ft_prop
    Autosize = $_HASH.args_ft_auto
  }
  Write-Output $_HASH.msg
  $THIS_SEISAN_ITEMS | Where-Object $_HASH.condition | Format-Table @_args
}

function fnCalc_Total() {
  # 合計値を計算する
  [scriptblock]$local:_condition1 = {$_.Content -like "yachin"}
  [scriptblock]$local:_condition2 = {$_.Memo    -like "*fuga清算*"}
  [scriptblock]$local:_condition3 = {$_.Content -like "*ふるさと納税*"}
  $THIS_SEISAN_ITEMS | ForEach-Object {
    [int]$script:seisan += [int]$_.Yen
    if($_condition1){ $seisan += $HOUSE_RENT_AID } # 家賃と判断される場合は都度補助費を足す
    if($_condition2){ $seisan -= $_.Yen} # 清算済み(値が正のもの=manからの振込)項目は除外する
    if($_condition3){ $seisan -= $_.Yen} # ふるさと納税は除外する
  }
}

function fnShow_Result(){
  # fnCalc_Totalで計算した$seisanの結果を出力する
  [int]$local:_num_orico=($THIS_SEISAN_ITEMS | Where-Object {$_.Content -match "yachin"}).length
  [int]$local:_seisan_Man = $seisan * 0.33 # manの負担
  [int]$local:_months= ( [int]$TILL.toString("yyyy") - [int]$FROM.toString("yyyy") ) * 12 + [int]$TILL.toString("MM") - [int]$FROM.toString("MM")
  $local:_items_already_paid = $THIS_SEISAN_ITEMS |
    Where-Object {$_.Memo -like "*fuga清算*" } |
    Measure-Object -Property Yen -Sum |
    Select-Object -ExpandProperty Sum
  [object]$local:_msgs=@( # Write-Outputで改行有出力にするためにstringではなくobjectにする
    "---"
    "集計期間 :"
    "  From   : $( $FROM.toString("yyyy/MM") )"
    "  Till   : $( $TILL.toString("yyyy/MM") )"
    "  Months : $_months "
    "  「yachin」の月数 : $_num_orico " # 月数と一致することを確認する
    "家賃補助 : $HOUSE_RENT_AID"
    "清算額合計 - 家賃補助*Months - 清算済み金額 - ふるさと納税 :"
    "  Amount : $('{0:N0}' -f $seisan)" # 数字をカンマ区切りで表示する
    "man負担(上記の33%) :"
    "  Amount : $('{0:N0}' -f $_seisan_Man)"
    "man清算済 :"
    "  Amount : $('{0:N0}' -f $_items_already_paid)"
    "今回の依頼分 :"
    "  Amount : $('{0:N0}' -f ($_seisan_Man + $_items_already_paid))"
    "---"
  )
  Write-Output $_msgs # 改行有出力にするために変数名をダブルクォートで囲まない
}

function fnMain(){
  fnTest_FolderPath     # ファイル存在確認
  fnGet_TargetCsvs      # 全csvの内容をオブジェクトに格納
  fnNew_ThisCsv         # オブジェクトを1つのcsv（this.csv）に出力
  fnGet_SeisanTargetAll # 条件（fuga/HOGE精算）に合致する項目のみオブジェクトに格納
  if($CSVOUT){
    fnOutput_Csv($allTargetItems) # summary.csvに出力
  }
  if($PART){
    fnDetermine_SpecifiedTargetPeriod  # 指定された期間を取得
  }else{
    fnDetermine_TargetPeriod           # 全csvのoldest/newestを取得
  }
  fnGet_ThisItems       # 指定期間に合致する項目を取得
  fnShow_Costs($HASH_HOUSERENT)
  fnShow_Costs($HASH_SUIDO)
  fnShow_Costs($HASH_KONETSU)
  fnShow_Costs($HASH_KOUGAKU)
  fnShow_Costs($HASH_DISNEY)
  fnShow_Costs($HASH_ALREADY_PAID)
  fnCalc_Total
  fnShow_Result
}

# Main
fnMain