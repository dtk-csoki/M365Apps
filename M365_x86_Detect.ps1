#detection
    $product_name = "Microsoft 365 Apps"
    foreach ($product in $product_name) {
    $x32 = gci "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | foreach { gp $_.PSPath } | ? { $_ -match $product } | select Displayname, UninstallString
    $x64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue | foreach { gp $_.PSPath } | ? { $_ -match $product } | select DisplayName, UninstallString
        if ($x32){ 
            Write-Output "Office 365 x86 detected"
            Exit 0
        } elseif ($x64){ 
            #Write-Output  "Office 365 x64 detected"
            Exit -1
        }  else {
            Exit -1
        }     
    }
