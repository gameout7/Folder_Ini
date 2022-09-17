$folderStructure = Get-Content -path .\Folder_Structure.ini
[string]$ProjectPath = (get-location).path + "\" + $projectCountry + "_" + $($projectName.Replace(' ','_'))

'Folder Structer:'
foreach ($line in $folderStructure)
    {
#        if ($line -ne "" -and $line.startswith(";") -ne $true)
        if ($line -match '^\\')
            {
                $lineFormat = $line.Trim("")
#                $lineFormat = $lineFormat.Replace('','_')

                New-Item -path $($ProjectPath + $lineFormat) -ItemType Directory -Force | Out-Null
                $lineFormat
            } 
    }