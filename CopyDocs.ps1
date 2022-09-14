foreach ($item in $ProjectItems)
    {
        If ($item.itemType -eq "visio")
        {
            $OfficeApplication = New-Object -ComObject Visio.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            $OfficeDocument = $OfficeApplication.Documents.open($templatename)

            $OfficeDocument.title = $item.ItemTitle
            $OfficeDocument.Description = $item.ItemNumber
            $OfficeDocument.Subject = $projectName
            $OfficeDocument.Company = $projectCompany
            $OfficeDocument.Category = $Item.ItemSite
            $OfficeDocument.Creator = $projectEngineer
            $OfficeDocument.Manager = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "excel")
        {
            $OfficeApplication = New-Object -ComObject excel.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            
            $OfficeDocument = $OfficeApplication.Workbooks.open($templatename)
            $OfficeDocument.BuiltInDocumentProperties("title") = $item.ItemTitle
            $OfficeDocument.BuiltInDocumentProperties("comments") = $item.ItemNumber
            $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $Item.ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "word")
        {
            $OfficeApplication = New-Object -ComObject word.application
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate

            $OfficeDocument = $OfficeApplication.Documents.open($templatename)
            $OfficeDocument.BuiltInDocumentProperties("title") = $item.ItemTitle
            $OfficeDocument.BuiltInDocumentProperties("comments") = $item.ItemNumber
            $OfficeDocument.BuiltInDocumentProperties("subject") = $projectName
            $OfficeDocument.BuiltInDocumentProperties("company") = $projectCompany
            $OfficeDocument.BuiltInDocumentProperties("category") = $Item.ItemSite
            $OfficeDocument.BuiltInDocumentProperties("author") = $projectEngineer
            $OfficeDocument.BuiltInDocumentProperties("manager") = $projectManager

            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            $OfficeDocument.saveas($path)
            $OfficeDocument.Close()
            $OfficeApplication.Quit()
        }
        If ($item.itemType -eq "none")
        {
            [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
            [string]$path =  $ProjectPath + "\" + $item.ItemPath
            Copy-Item -Path $templatename -Destination $path
        }


    }

    
<#
$OfficeApplication = New-Object -ComObject Visio.application

foreach ($item in $ProjectItems)
    {
        [string]$templatename =  (get-location).path + '\Templates\' + $item.itemtemplate
        $OfficeDocument = $OfficeApplication.Documents.open($templatename)

        $OfficeDocument.title = $item.ItemTitle
        $OfficeDocument.Description = $item.ItemNumber
        $OfficeDocument.Subject = $projectName
        $OfficeDocument.Company = $projectCompany
        $OfficeDocument.Category = $Item.ItemSite
        $OfficeDocument.Creator = $projectEngineer
        $OfficeDocument.Manager = $projectManager

        [string]$path =  $ProjectPath + "\" + $item.ItemPath
        $OfficeDocument.saveas($path)
        $OfficeDocument.Close()

    }

    $OfficeApplication.Quit()
#>