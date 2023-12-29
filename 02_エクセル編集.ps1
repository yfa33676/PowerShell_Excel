$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.ScreenUpdating = $true

Import-Csv -Encoding UTF8 -Path .\2023年参加公演.txt |
Group-Object -Property タイトル, 日程, 会場, 開場, 開演  |
ForEach-Object {
  $タイトル = ($_.Name -split ",")[0]
  $日程 = ($_.Name -split ",")[1]
  $会場 = ($_.Name -split ",")[2]
  $開場 = ($_.Name -split ",")[3]
  $開演 = ($_.Name -split ",")[4]
  $出演 = $_.Group.出演 -join "、"

  $ファイル名 = $日程.Replace("/","").Replace(" ","") + "_" + $開演.Replace(":","").Replace(" ","") + "_" + $タイトル.Replace(" ","_") + ".xlsx"
  $ファイルパス = Join-Path -Path .\出力先 -ChildPath $ファイル名 | Resolve-Path

  $ブック = $excel.Workbooks.Open($ファイルパス)
  $シート = $ブック.Worksheets.Item("見出し")

  $シート.Range("A1").Value2 = $タイトル
  $シート.Range("A2").Value2 = "公演日：" + $日程
  $シート.Range("A3").Value2 = "会場：" + $会場
  $シート.Range("A4").Value2 = "開場：" + $開場 + "、開演：" + $開演
  $シート.Range("A5").Value2 = "出演：" + $出演

  $ブック.Close($true)
}

$excel.Quit()