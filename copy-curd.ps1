# remove-item -path f:\curd\ -include *.*   -force
start-transcript c:\AdminDir\debug.log

#$query = "select * from win32_pingstatus where address = 'd29tnv01'"
#$isping = get-wmiobject -query $query 
#if ($isping.statuscode -eq 0) {

   Write-Host Ping OK.

    foreach ($file in $(Get-childItem 'c:\AdminDir\Curd' -include CUR*.ARJ -Recurse ))
    {
      
      $hostout = $file.FullName + "..."
      write-host $hostout

      copy-item $file.FullName -destination c:\curd -recurse
      c:\AdminDir\2Val\dbf-01.ps1
      
    }
#}

stop-transcript