$folderStructure = Get-Content -path .\Folder_Structure.ini
foreach ($line in $folderStructure)
    {
        if ($line -ne "" -and $line.startswith(";") -ne $true )
            {
                New-Item -path $("Project" + $line.Trim("")) -ItemType Directory -Force
            } 
    }