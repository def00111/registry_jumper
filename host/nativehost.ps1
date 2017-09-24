try {
  $reader = New-Object System.IO.BinaryReader([System.Console]::OpenStandardInput())
  $len = $reader.ReadInt32()
  $buf = $reader.ReadBytes($len)
  $msg = [System.Text.Encoding]::UTF8.GetString($buf)
  
  $obj = $msg | ConvertFrom-Json
  $text = $obj.text
  
  $regJump = [System.IO.Path]::Combine($PSScriptRoot, "regjump.vbs")

  try {
    start wscript -ArgumentList "$regJump $($text)" -Wait -Verb "runas"
  }
  catch {}
}
finally {
  $reader.Dispose()
}
