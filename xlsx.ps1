$global:rIdCount = 2

Function ExcelInit([string]$Path) {
    # loading WindowsBase.dll with the System.IO.Packaging namespace
    $Null = [Reflection.Assembly]::LoadWithPartialName("WindowsBase")

    $exPkg = $Null
    Try { # open the main package on disk or create if exists
        $exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Open)
    } 
    Catch {
        $exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Create)
    }

    return $exPkg
}

Function ExcelClose($exPkg) {
    if ($exPkg -ne $null) {
        $exPkg.Close()
    }
}

Function ExcelAdd-WorkSheet($exPkg, $WorkBookPart, $NewWorkSheetPartName) {
    # Increment rId
    $sheetId = $global:rIdCount
    $NewWorkBookRelId = "rId$sheetId"
    $global:rIdCount = $global:rIdCount + 1

	# create empty XML Document
	$New_Worksheet_xml = New-Object System.Xml.XmlDocument

    # Obtain a reference to the root node, and then add the XML declaration.
    $XmlDeclaration = $New_Worksheet_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
    $Null = $New_Worksheet_xml.InsertBefore($XmlDeclaration, $New_Worksheet_xml.DocumentElement)

    # Create and append the worksheet node to the document.
    $workSheetElement = $New_Worksheet_xml.CreateElement("worksheet")
	# add the Excel related office open xml namespaces to the XML document
    $Null = $workSheetElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $Null = $workSheetElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $Null = $New_Worksheet_xml.AppendChild($workSheetElement)

    # Create and append the sheetData node to the worksheet node.
    $Null = $New_Worksheet_xml.DocumentElement.AppendChild($New_Worksheet_xml.CreateElement("sheetData"))

	# create URI for worksheet package part
	$Uri_xl_worksheets_sheet_xml = New-Object System.Uri -ArgumentList ("/xl/worksheets/${NewWorkSheetPartName}.xml", [System.UriKind]::Relative)
	# try to open worksheet par, create if doesn't exist
   	try {
        $Part_xl_worksheets_sheet_xml = $exPkg.GetPart($Uri_xl_worksheets_sheet_xml)
    }
    catch {
    	$Part_xl_worksheets_sheet_xml = $exPkg.CreatePart($Uri_xl_worksheets_sheet_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    	# get writeable stream from part 
    	$dest = $part_xl_worksheets_sheet_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
    	# write $New_Worksheet_xml XML document to part stream
    	$New_Worksheet_xml.Save($dest)

        # create workbook to worksheet relationship
    	$Null = $WorkBookPart.CreateRelationship($Uri_xl_worksheets_sheet_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $NewWorkBookRelId)

    	# edit the xl\workbook.xml

    	# create empty XML Document
    	$WorkBookXmlDoc = New-Object System.Xml.XmlDocument
    	# load XML document from package part stream
    	$WorkBookXmlDoc.Load($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))

    	# create a new XML Node for the sheet 
    	$WorkBookXmlSheetNode = $WorkBookXmlDoc.CreateElement('sheet', $WorkBookXmlDoc.DocumentElement.NamespaceURI)
        $Null = $WorkBookXmlSheetNode.SetAttribute('name', $NewWorkSheetPartName)
        $Null = $WorkBookXmlSheetNode.SetAttribute('sheetId', $sheetId)
    	# try to create the ID Attribute with the r: Namespace (xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships") 
    	$NamespaceR = $WorkBookXmlDoc.DocumentElement.GetNamespaceOfPrefix("r")
    	If($NamespaceR) {
        	$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NamespaceR,$NewWorkBookRelId)
    	} Else {
    		$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NewWorkBookRelId)
    	}
    	
    	# add the new sheet node to XML document
    	$Null = $WorkBookXmlDoc.DocumentElement.Item("sheets").AppendChild($WorkBookXmlSheetNode)

    	# Save back the edited XML Document to package part stream
    	$WorkBookXmlDoc.Save($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Write))
    }

    return $Part_xl_worksheets_sheet_xml
}

Function ExcelNew-WorkBook($exPkg) {
    # create empty XML Document
    $xl_Workbook_xml = New-Object System.Xml.XmlDocument

    # Obtain a reference to the root node, and then add the XML declaration.
    $XmlDeclaration = $xl_Workbook_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
    $Null = $xl_Workbook_xml.InsertBefore($XmlDeclaration, $xl_Workbook_xml.DocumentElement)

    # Create and append the workbook node to the document.
    $workBookElement = $xl_Workbook_xml.CreateElement("workbook")
    # add the office open xml namespaces to the XML document
    $Null = $workBookElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $Null = $workBookElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $Null = $workBookElement.SetAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
    $Null = $workBookElement.SetAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    $Null = $workBookElement.SetAttribute("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    $Null = $workBookElement.SetAttribute("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" )
    $Null = $workBookElement.SetAttribute("xmlns:xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
    $Null = $workBookElement.SetAttribute("mc:Ignorable", "x14ac xr xr2 xr3" )
    $Null = $xl_Workbook_xml.AppendChild($workBookElement)

    # Create and append the sheets node to the workBook node.
    $Null = $xl_Workbook_xml.DocumentElement.AppendChild($xl_Workbook_xml.CreateElement("sheets"))

	# create URI for workbook.xml package part
	$Uri_xl_workbook_xml = New-Object System.Uri -ArgumentList ("/xl/workbook.xml", [System.UriKind]::Relative)
	# create workbook.xml part or open it if exists
	try {
        $Part_xl_workbook_xml = $exPkg.GetPart($Uri_xl_workbook_xml)
    }
    catch {
        $Part_xl_workbook_xml = $exPkg.CreatePart($Uri_xl_workbook_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    	# get writeable stream from workbook.xml part 
    	$dest = $part_xl_workbook_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
    	# write workbook.xml XML document to part stream
    	$xl_workbook_xml.Save($dest)

    	# create package general main relationships
    	$Null = $exPkg.CreateRelationship($Uri_xl_workbook_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1")
    }

    return $Part_xl_workbook_xml
}

Function ExcelAdd-WorkSheetRaw($workSheet, [string[]]$listParams) {
    # open worksheet xml
    $WorkSheetXmlDoc = New-Object System.Xml.XmlDocument
    # load XML document from package part stream
    $WorkSheetXmlDoc.Load($workSheet.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))

	$RowNode = $WorkSheetXmlDoc.CreateElement('row', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)

	ForEach($par in $listParams) {
		$CellNode = $WorkSheetXmlDoc.CreateElement('c', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
        $Null = $CellNode.SetAttribute('t', "inlineStr") 
		$Null = $RowNode.AppendChild($CellNode)

		$CellNodeIs = $WorkSheetXmlDoc.CreateElement('is', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
		$Null = $CellNode.AppendChild($CellNodeIs)

		$CellNodeIsT = $WorkSheetXmlDoc.CreateElement('t', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
		$CellNodeIsT.InnerText = [string]$par
		$Null = $CellNodeIs.AppendChild($CellNodeIsT)

		$Null = $WorkSheetXmlDoc.DocumentElement.Item("sheetData").AppendChild($RowNode)
	}

	$WorkSheetXmlDoc.Save($workSheet.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Write))
    $global:exPkg.Flush()
}
