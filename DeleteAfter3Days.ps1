$todayDate = (get-date).AddDays(-3)
$test = Get-ChildItem -Path C:\users\rjefferson\Documents\Fam -File -Recurse | Where-Object { $_.LastWriteTime -lt $todayDate } | Remove-Item -Force -Recurse
$test
