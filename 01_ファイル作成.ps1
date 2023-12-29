Import-Csv -Encoding UTF8 -Path .\2023年参加公演.txt |
Group-Object -Property タイトル, 日程, 会場, 開場, 開演  |
ForEach-Object {
  $タイトル = ($_.Name -split ",")[0]
  $日程 = ($_.Name -split ",")[1]
  $会場 = ($_.Name -split ",")[2]
  $開場 = ($_.Name -split ",")[3]
  $開演 = ($_.Name -split ",")[4]

  $ファイル名 = $日程.Replace("/","").Replace(" ","") + "_" + $開演.Replace(":","").Replace(" ","") + "_" + $タイトル.Replace(" ","_") + ".xlsx"
  Copy-Item -Path .\【雛形】日程_タイトル.xlsx -Destination (Join-Path -Path .\出力先 -ChildPath $ファイル名)
}