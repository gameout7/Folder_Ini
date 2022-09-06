# $Read_Ini #raw text from ini file
# $Read_ini = Get-Content folder_ini.ini
# $Read_Ini

<#
function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $filePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = "Comment" + $CommentCount
            $ini[$section][$name] = $value
        }
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}
Get-IniContent ".\folder_ini.ini"
#>
$ini = Get-Content -Path .\folder_ini.ini 
#$ini
$section = "none" #section name in ini file
$projectName = "no projectname"
$ProjectNumber = "no projectnumber"
$Author = "no author"
$Manager = "no manager"
$Company = "no company"
$Site = @()
$Document = @()
$iniHash = @{}
$iniTemp = @()

#Ini File parser
foreach ($line in $ini)
    {
        if ($line -ne "" -and $line.startswith(";") -ne $true )
            {
# Checking section name
                if ($line.StartsWith("[General]") -eq $True) {$section = 'General'; continue}
                if ($line.StartsWith("[Site]") -eq $True) {$section = 'Site'; continue}
                if ($line.StartsWith("[Document]") -eq $True) {$section = 'Document'; continue}
# Section General                
                if ($section -eq 'General')
                {
                    $SpliteArray = $Line.Split("=")
                    if ($SpliteArray[0].trim('') -eq "ProjectName") {$projectName = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "ProjectNumber") {$ProjectNumber = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "Author") {$Author = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "Manager") {$Manager = $SpliteArray[1].trim('')}
                    if ($SpliteArray[0].trim('') -eq "Company") {$Company = $SpliteArray[1].trim('')}
                }
# Section Site
                if ($section -eq 'Site')
                {
                    $Site += $Line.Trim('')
                }
# Section Document
                if ($section -eq 'Document')
                {
                    $SpliteArray = $Line.Split(",")
                    $Document += @{DocumentName=$SpliteArray[0].trim(''); DocumentQt=$SpliteArray[1].trim(''); DocumentPath=$SpliteArray[2].trim('')}
                }                

            }

    }
    "Project Name is $projectName"
    "Project Number is $ProjectNumber"
    "Project Author is $Author"
    "Project Manager is $Manager"
    "Project Company is $Company"
$site
$Document

  #  $iniTemp