# 在 Windows Microsoft Excel 中保存并关闭完整版，再运行本脚本。
# 仅在本文件所在目录更新同批次送审演示版与交流意见版，旧导出自动备份。
$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$fullPath = Join-Path $folder '租金市场法专业试用版_完整版_V0.2.xlsx'
$reviewPath = Join-Path $folder '租金市场法专业试用版_送审演示版_V0.2.xlsx'
$dialogPath = Join-Path $folder '租金市场法专业试用版_交流意见版_V0.2.xlsx'
if (-not (Test-Path -LiteralPath $fullPath)) { throw '找不到同目录的完整版，请将文件放在一起。' }
$batch = 'RP-V02-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
$coreSheets = @('测算结果','资产清单','计租与历史','市场案例库','三案例选用','三案例测算','修正明细','测算参数','租赁资料核定','特殊计租','设备附表','资料索引')
$existing = @($reviewPath,$dialogPath) | Where-Object { Test-Path -LiteralPath $_ }
if ($existing.Count -gt 0) {
 $backup = Join-Path (Join-Path $folder '历史导出') $batch
 New-Item -ItemType Directory -Path $backup -Force | Out-Null
 foreach ($path in $existing) { Copy-Item -LiteralPath $path -Destination $backup }
}
$excel = $null; $full = $null; $review = $null; $dialog = $null
try {
 $excel = New-Object -ComObject Excel.Application
 $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
 $full = $excel.Workbooks.Open($fullPath,0,$false)
 if ($full.ReadOnly) { throw '完整版被占用。请先保存并关闭完整版，再运行同步导出。' }
 $excel.CalculateFullRebuild()
 $result = $full.Worksheets.Item('测算结果')
 $overview = $full.Worksheets.Item('专业总览')
 $overview.Range('B5').Value2 = $batch
 $result.Range('A1').Value2 = '导出批次 '+$batch+'  全部为模拟数据'
 if ([string]$overview.Range('B10').Value2 -eq '待补计算资料') { throw '有范围内记录尚缺计算资料，请补齐后再导出。' }
 if ([double]$overview.Range('B16').Value2 -ne 0) { throw '三案例组合或修正存在异常，请按提示复核后再导出。' }
 foreach ($sheet in $full.Worksheets) {
  $sheet.PageSetup.Orientation=2
  $sheet.PageSetup.PaperSize=8
  $sheet.PageSetup.Zoom=$false
  $sheet.PageSetup.FitToPagesWide=1
  $sheet.PageSetup.FitToPagesTall=$false
  $sheet.PageSetup.LeftMargin=$excel.CentimetersToPoints(1)
  $sheet.PageSetup.RightMargin=$excel.CentimetersToPoints(1)
 }
 $printAreas = @{'专业总览'='A1:F34';'测算结果'='A1:N28';'资产分类汇总'='A1:E18';'两年金额对比'='A1:H30';'识别与分组'='A1:J27';'三案例选用'='A1:I19';'三案例测算'='A1:K43';'修正明细'='A1:N583';'市场案例库'='A1:R37';'资产清单'='A1:O27';'计租与历史'='A1:H28';'测算参数'='A1:D45';'租赁资料核定'='A1:P10';'特殊计租'='A1:K10';'设备附表'='A1:I9';'历史离散复核'='A1:I27';'大幅变动复核'='A1:J27';'内部复核参数'='A1:C14';'资料索引'='A1:F63';'异常练习'='A1:J24';'工时与反馈'='A1:F28'}
 foreach ($sheet in $full.Worksheets) { $sheet.PageSetup.PrintArea=$printAreas[$sheet.Name] }
 $result.PageSetup.FitToPagesTall=1
 $overview.Activate(); $overview.Range('A1').Select(); $excel.ActiveWindow.DisplayGridlines=$false; $excel.ActiveWindow.Zoom=85
 $full.Save()
 $full.SaveCopyAs($reviewPath)
 $review=$excel.Workbooks.Open($reviewPath,0,$false)
 $remove=@();foreach($sheet in $review.Worksheets) { if ($coreSheets -notcontains $sheet.Name) { $remove += $sheet.Name } }
 foreach($name in $remove) { $review.Worksheets.Item($name).Delete() }
 $review.Worksheets.Item('测算结果').Range('A2').Value2='租金市场法专业试用版送审演示结果'
 $excel.CalculateFullRebuild()
 if ($null -ne $review.LinkSources(1)) { throw '送审演示版出现表外链接，停止交付。' }
 $review.Worksheets.Item('测算结果').Activate();$review.Worksheets.Item('测算结果').Range('A1').Select()
 $review.Save()
 $dialog=$excel.Workbooks.Add()
 while ($dialog.Worksheets.Count -gt 1) { $dialog.Worksheets.Item($dialog.Worksheets.Count).Delete() }
 $sheet=$dialog.Worksheets.Item(1);$sheet.Name='交流意见'
 $sheet.Range('A1').Value2='导出批次 '+$batch+'  同批次结果快照'
 $sheet.Range('A2').Value2='租金市场法专业试用版交流意见表'
 $sheet.Range('A3').Value2='所有资产及金额均为模拟。请逐项填写确认意见；修改完整版后须重新同步导出。'
 $headers=@('序号','区域','管理单元','资产编码','资产名称','精确地址','实际用途','单价单位','计租数量','参考单价','本次单价','本年年租金元','差异率','范围说明','确认意见')
 for($c=1;$c -le 15;$c++) { $sheet.Cells.Item(5,$c).Value2=$headers[$c-1] }
 $data=$full.Worksheets.Item('资产清单')
 $matrix=New-Object 'object[,]' 20,15
 for($r=6;$r -le 25;$r++) {
  $ix=$r-6
  $matrix[$ix,0]=[double]($r-5)
  foreach($pair in @(@(1,2),@(2,3),@(3,1),@(4,4),@(5,5),@(6,8),@(7,10),@(8,11))) { $matrix[$ix,$pair[0]]=$data.Cells.Item($r,$pair[1]).Value2 }
  $matrix[$ix,9]=$result.Range('I'+$r).Value2
  $matrix[$ix,10]=$result.Range('G'+$r).Value2
  $matrix[$ix,11]=$result.Range('H'+$r).Value2
  $matrix[$ix,13]=$result.Range('L'+$r).Value2
 }
 for($ix=0;$ix -lt 20;$ix++) {
  for($cx=0;$cx -lt 15;$cx++) {
   $cellValue=$matrix[$ix,$cx]
   if($null -eq $cellValue){continue}
   if($cellValue -is [string]){ $cellText="'"+$cellValue }
   else { $cellText=[Convert]::ToString($cellValue,[Globalization.CultureInfo]::InvariantCulture) }
   $sheet.Cells.Item(($ix+6),($cx+1)).Formula=$cellText
  }
 }
 for($r=6;$r -le 25;$r++) { $sheet.Cells.Item($r,13).Formula='=IF(OR(COUNT(J'+$r+':K'+$r+')<>2,J'+$r+'<=0,N'+$r+'<>"纳入"),"",K'+$r+'/J'+$r+'-1)' }
 $sheet.Range('A1:O28').Font.Name='微软雅黑';$sheet.Range('A1:O28').Font.Size=10;$sheet.Range('A1:O28').VerticalAlignment=-4108
 $sheet.Range('A2:O2').Interior.Color=6108695;$sheet.Range('A2:O2').Font.Color=16777215;$sheet.Range('A2:O2').Font.Size=16;$sheet.Range('A2:O2').Font.Bold=$true;$sheet.Rows.Item(2).RowHeight=34
 $sheet.Range('A5:O5').Interior.Color=15849948;$sheet.Range('A5:O5').Font.Bold=$true;$sheet.Range('A5:O5').WrapText=$true;$sheet.Rows.Item(5).RowHeight=44
 $sheet.Range('A6:O25').Borders.LineStyle=1;$sheet.Range('A6:O25').Borders.Color=15198169;$sheet.Range('A6:O25').WrapText=$true;$sheet.Range('A6:O25').RowHeight=46
 $widths=@(8,15,18,16,23,22,17,18,14,17,18,20,17,14,44)
 for($c=1;$c -le 15;$c++) { $sheet.Columns.Item($c).ColumnWidth=$widths[$c-1] }
 $sheet.Range('I6:L25').NumberFormat='#,##0.00';$sheet.Range('M6:M25').NumberFormat='0.0%';$sheet.Range('O6:O25').Interior.Color=15137023;$sheet.Range('O6:O25').Font.Color=14110237
 $sheet.Range('A27').Value2='范围排除的0表示本次不计租，不表示资产市场价值为零。参考价仅作逐项比较，不代替本次测算。'
 $sheet.PageSetup.Orientation=2;$sheet.PageSetup.PaperSize=8;$sheet.PageSetup.Zoom=$false;$sheet.PageSetup.FitToPagesWide=1;$sheet.PageSetup.FitToPagesTall=1;$sheet.PageSetup.PrintArea='A1:O27'
 $sheet.Activate();$sheet.Range('A6').Select();$excel.ActiveWindow.FreezePanes=$true;$sheet.Range('A1').Select();$excel.ActiveWindow.DisplayGridlines=$false;$excel.ActiveWindow.Zoom=80
 $excel.CalculateFullRebuild()
 $dialog.SaveAs($dialogPath,51)
 for($r=6;$r -le 25;$r++) {
  foreach($c in @(7,8)) {
   $a=$result.Cells.Item($r,$c).Value2;$b=$review.Worksheets.Item('测算结果').Cells.Item($r,$c).Value2
   if ([Math]::Abs([double]$a-[double]$b) -gt 0.000001) { throw ('送审结果不同：第'+$r+'行') }
  }
  if ([Math]::Abs([double]$result.Cells.Item($r,8).Value2-[double]$sheet.Cells.Item($r,12).Value2) -gt 0.000001) { throw ('交流意见金额不同：第'+$r+'行') }
 }
 $dialog.Close($true);$dialog=$null;$review.Close($true);$review=$null;$full.Close($true);$full=$null
 Write-Output ('已同步导出，批次：'+$batch)
} finally {
 if($null -ne $dialog){$dialog.Close($false)};if($null -ne $review){$review.Close($false)};if($null -ne $full){$full.Close($false)}
 if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
# Remove Office author and local-folder metadata while preserving formulas and saved calculation values.
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
foreach($path in @($fullPath,$reviewPath,$dialogPath)) {
 $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try {
  foreach($entryName in @('docProps/core.xml','xl/workbook.xml')) {
   $entry=$zip.GetEntry($entryName);if($null -eq $entry){continue}
   $reader=New-Object IO.StreamReader($entry.Open());$content=$reader.ReadToEnd();$reader.Dispose()
   if($entryName -eq 'docProps/core.xml') {
    $content=[regex]::Replace($content,'<(dc:creator|cp:lastModifiedBy)(\s[^>]*)?>.*?</\1>','', [Text.RegularExpressions.RegexOptions]::Singleline)
   } else {
    $content=[regex]::Replace($content,'<mc:AlternateContent\b[^>]*>.*?</mc:AlternateContent>',{param($m) if($m.Value.Contains('absPath')){''}else{$m.Value}},[Text.RegularExpressions.RegexOptions]::Singleline)
   }
   $entry.Delete();$new=$zip.CreateEntry($entryName);$writer=New-Object IO.StreamWriter($new.Open(),(New-Object Text.UTF8Encoding($false)));$writer.Write($content);$writer.Dispose()
  }
 } finally { $zip.Dispose() }
}
