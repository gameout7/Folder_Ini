$Word = New-Object -ComObject Word.application
$Word.visible = $true
$WordDocument = $Word.Documents.add()
$Selection = $Word.Selection
$WordSection = $WordDocument.Sections.Item(1)

#Landscape orientation
$WordDocument.PageSetup.Orientation = 1

#Header

$Header = $WordSection.Headers.Item(1)
$Header.ParagraphFormat.Alignment = 0
$Header.Font.Size = 8
$Header.range.text = "header text"

#Footer
$Footer = $WordSection.footers.item(1)
$Footer.range.text = "footer text"

#Title
$Selection.ParagraphFormat.Alignment = 1
$Selection.Font.Bold = 1
$Selection.Font.Size = 18
$Selection.TypeText("Documentation List")
#Description
$Selection.TypeParagraph()
$Selection.TypeParagraph()
$Selection.ParagraphFormat.Alignment = 0
$Selection.Font.Bold = 0
$Selection.Font.Size = 11
$Selection.TypeText("Project Name: STC")
$Selection.TypeParagraph()
$Selection.TypeText("Project Code: 123")
$Selection.TypeParagraph()
$Selection.TypeText("Project Manager: Name")
$Selection.TypeParagraph()
$Selection.TypeText("Project Manager: Name")
