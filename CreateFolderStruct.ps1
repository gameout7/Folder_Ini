$folderStructure = Get-Content -path .\Folder_Structure.ini
[string]$ProjectPath = (get-location).path + "\" + $projectCountry + "_" + $($projectName.Replace(' ','_'))

foreach ($line in $folderStructure)
    {
        if ($line -ne "" -and $line.startswith(";") -ne $true )
            {
                $lineFormat = $line.Trim("")
#                $lineFormat = $lineFormat.Replace('','_')

                New-Item -path $($ProjectPath + $lineFormat) -ItemType Directory -Force
            } 
    }