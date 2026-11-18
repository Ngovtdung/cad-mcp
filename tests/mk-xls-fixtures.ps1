# Tạo tests\fixtures-xls\book.xls và enc.xls (mật khẩu 1234) bằng Excel COM. Chạy 1 lần; tệp được giữ lại.
#   powershell -NoProfile -STA -File tests\mk-xls-fixtures.ps1
# Luôn mở 1 tiến trình Excel ẩn MỚI (New-Object), không bao giờ GetActiveObject -> không đụng Excel đang mở của user.
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot 'fixtures-xls'
New-Item -ItemType Directory -Force $dir | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AutomationSecurity = 3   # tắt macro
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  while ($wb.Worksheets.Count -gt 2) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $s = $wb.Worksheets.Item(1); $s.Name = 'Dữ liệu'
  $s.Range('A1').Value2 = 'Tên'; $s.Range('B1').Value2 = 'Ngày'; $s.Range('C1').Value2 = 'Tổng'
  $s.Range('A2').Value2 = 'Thép'; $s.Range('B2').Value2 = 45292; $s.Range('B2').NumberFormat = 'dd/mm/yyyy'
  $s.Range('C2').Formula = '=D2*2'; $s.Range('D2').Value2 = 42
  $s.Range('A3').Value2 = 'Gộp'; $null = $s.Range('A3:C3').Merge(); $s.Range('D3').Value2 = "x`ny"
  $h = $wb.Worksheets.Item(2); $h.Name = 'Ẩn'; $h.Range('A1').Value2 = 'bí mật'; $h.Visible = 0
  $wb.RemovePersonalInformation = $true   # xoá tác giả/người sửa cuối/WRITEACCESS khỏi tệp (không đổi cài đặt Office của người dùng)
  $wb.SaveAs((Join-Path $dir 'book.xls'), 56)
  $wb.SaveAs((Join-Path $dir 'enc.xls'), 56, '1234')
  $wb.Close($false)
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Get-ChildItem $dir | Select-Object Name, Length
