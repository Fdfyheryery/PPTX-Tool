<#

.SYNOPSIS
WIP

.DESCRIPTION
WIP

.NOTES
WIP

#>

Add-Type -AssemblyName System.IO.Compression.FileSystem

#Emplacement des fichiers décompressés (images pour rapport HTML)
$appTempPath = $Env:temp + "\PPTX-Tool"

function GetEntryAsXML {
    param([System.IO.Compression.ZipArchiveEntry]$entry)

    $slide = $entry.Open()
    $reader = New-Object IO.StreamReader($slide)
    [xml]$entryXML = $reader.ReadToEnd()
    $reader.Close()
    $slide.Close()
    return $entryXML
}

#Contourne une limitation des classes PowerShell
function CallAnalyzeFromEntry {
    param([PPTXFile]$file, [System.IO.Compression.ZipArchiveEntry]$entry)

    $file.zipArchive = [System.IO.Compression.ZipArchive]::New($entry.Open())
    $file.AnalyzeFile()
    $file.zipArchive.Dispose()
}

#Contourne une limitation des classes PowerShell
function CallAnalyzeFromName {
    param([PPTXFile]$file, [string]$name)

    $file.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($name)
    $file.AnalyzeFile()
    $file.zipArchive.Dispose()
}

#Contourne une limitation des classes PowerShell
function ExtractImgToFile {
    param($entry, [string]$dPath)

    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dPath)
}

function GetImageFromXML {
    param([PPTXFile[]]$rIds, $pic)

    #rID
    $rId = $pic.blipfill.blip.embed

    #Ratio
    $cx = $pic.sppr.xfrm.ext.cx
    $cy = $pic.sppr.xfrm.ext.cy

    #Utilisation (Rognage) (10000 = 10.000%)
    $utilVertical = 100000 - ([int]$pic.blipfill.srcRect.t + [int]$pic.blipfill.srcRect.b)
    $utilHorizontal = 100000 - ([int]$pic.blipfill.srcRect.l + [int]$pic.blipfill.srcRect.r)


    if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
        $index = $rIds.name.indexof($rId)
        $rIds[$index].Total++

        if ($rIds[$index].cx -lt $cx) {
            $rIds[$index].cx = $cx
        }

        if ($rIds[$index].cy -lt $cy) {
            $rIds[$index].cy = $cy
        }

        if ($rIds[$index].UtilisationV -lt $utilVertical) {
            $rIds[$index].UtilisationV = $utilVertical
        }

        if ($rIds[$index].UtilisationH -lt $utilHorizontal) {
            $rIds[$index].UtilisationH = $utilHorizontal
        }
    }
    else {
        $newItem = [PPTXImage]::new($rId)
        $newItem.cx = $cx
        $newItem.cy = $cy
        $newItem.utilisationV = $utilVertical
        $newItem.utilisationH = $utilHorizontal
        $newItem.total = 1

        $rIds += $newItem
    }

    #Si l'image est un preview de vidéo, on ajoute la vidéo dans la liste
    if ($pic.nvpicpr.nvpr.videofile.link -ne $null) {
        $rId = $pic.nvpicpr.nvpr.videofile.link
        if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
            $index = $rIds.name.indexof($rId)
            $rIds[$index].Total++
        }
        else {
            $newItem = [PPTXVideo]::new($rId)
            $newItem.total = 1

            $rIds += $newItem
        }
    }
    return $rIDs
}

function GetRelsFromXML {
    param([PPTXFile]$file, [PPTXFile]$RIdItem, $xml, $slideNum)

    $image = $xmlNode.Target.split("/")[-1]

    #Images
    if ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
        
        if (($file.arrayImages.Count -gt 0) -and ($file.arrayImages.Name -contains $image)) {
            $indexImage = $file.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $slideNum

            if ($file.arrayImages[$indexImage].cx -lt $RIdItem.cx) {
                $file.arrayImages[$indexImage].cx = $RIdItem.cx
            }

            if ($file.arrayImages[$indexImage].cy -lt $RIdItem.cy) {
                $file.arrayImages[$indexImage].cy = $RIdItem.cy
            }

            if ($file.arrayImages[$indexImage].utilisationV -lt $RIdItem.utilisationV) {
                $file.arrayImages[$indexImage].utilisationV = $RIdItem.utilisationV
            }

            if ($file.arrayImages[$indexImage].utilisationH -lt $RIdItem.utilisationH) {
                $file.arrayImages[$indexImage].utilisationH = $RIdItem.utilisationH
            }
        }
        else {
            $newItem = [PPTXImage]::new($image)
            $newItem.cx = $RIdItem.cx
            $newItem.cy = $RIdItem.cy
            $newItem.utilisationV = $RIdItem.utilisationV
            $newItem.utilisationH = $RIdItem.utilisationH
            $newItem.slides = @($slideNum)
            $newItem.total = $RIdItem.total
            $newItem.decompressPath = $appTempPath + "\" + $file.name.Substring(0, $file.name.Length - 5)
            $file.arrayImages += $newItem
        }
    }

    #Videos
    elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video") {
        if (($file.arrayImages.Count -gt 0) -and ($file.arrayImages.Name -contains $image)) {
            $indexImage = $file.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $slideNum
        }
        else {
            $newItem = [PPTXVideo]::new($image)
            $newItem.slides = @($slideNum)
            $newItem.total = $RIdItem.total
            $file.arrayImages += $newItem
        }
    }

    #Word, Excel, PowerPoint
    elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
        if (($file.arrayImages.Count -gt 0) -and ($this.arrayImages.Name -contains $image)) {
            $indexImage = $this.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $slideNum
        }
        else {
            $itemType = $image.Substring($image.get_Length()-4)
            $newItemName = $file.name.Substring(0, $file.name.Length - 5) + "\" + $image

            if ($itemType -eq "pptx") {
                $newItem = [PPTXPowerPoint]::new($newItemName, $false)
            }
            elseif ($itemType -eq "xlsx") {
                $newItem = [PPTXExcel]::new($newItemName, $false)
            }
            elseif ($itemType -eq "docx") {
                $newItem = [PPTXWord]::new($newItemName, $false)
            }
            else {
                $newItem = [PPTXFile]::new()
                $newItem.name = $newItemName
            }
                                
            $newItem.slides = @($slideNum)
            $newItem.total = $RIdItem.total
            $file.arrayImages += $newItem
        }
    }
}

function CreateFileWarnings {
    param([PPTXFile]$pptxfile)

    foreach ($file in $pptxfile.arrayImages) {
        $filePath = ""
        $startPath = ""

        if ($pptxfile.GetType().Name -eq "PPTXPowerPoint") {
            $startPath = "ppt/"
        }

        elseif ($pptxfile.GetType().Name -eq "PPTXWord") {
            $startPath = "word/"
        }

        elseif ($pptxfile.GetType().Name -eq "PPTXExcel") {
            $startPath = "xl/"
        }

        if ($file.GetType().Name -eq "PPTXImage" -or $file.GetType().Name -eq "PPTXVideo") {
            $filePath = $startPath + "media/" + $file.Name
        }

        elseif ($file.GetType().Name -eq "PPTXPowerPoint" -or $file.GetType().Name -eq "PPTXExcel" -or $file.GetType().Name -eq "PPTXWord") {
            $filePath = $startPath + "embeddings/" + $file.Name.Split("\")[-1]
        }

        $entry = $pptxfile.zipArchive.GetEntry($filePath)
        $hasWarning = $file.CreateWarning($entry)
    }
}

function GenerateHTML {
    param([PPTXFile]$pptxfile, [bool]$isChild)

        $imgClass = "PPTXFile_img " + $pptxfile.GetType().Name + "_img"
        $style = ""
        $class = "order1"

        if ($isChild) {
            $imgClass = "PPTX_others " + $pptxfile.GetType().Name + "_img"
            $style = 'style="font-size:0.9em";'
            $class = "order4"
        }

        $html = ' <div class="PPTXFile ' + $class + '"><div class="line" ' + $style + '><div class="' + $imgClass + '">' + $pptxfile.GetType().Name[4] + '</div>'`
            + '<span class="name">' + $pptxfile.name + '</span></div>'

        if ($this.warning) {
            $html += '<div class="line line_child"><div class="PPTX_others ' + $pptxfile.GetType().Name + '_img">-</div>'`
            + '<span class="name nameLarge">Général</span><span class="slide"> </span><div class="colFlex">'

            foreach ($warning in $pptxfile.warning) {
                $html += '<span class="warning">' + $warning + '</span>'
            }

            $html += '</div></div>'
        }


        foreach ($file in $pptxfile.arrayImages) {
            if ($file.warning -or $file.GetType().Name -eq "PPTXPowerPoint" -or $file.GetType().Name -eq "PPTXExcel" -or $file.GetType().Name -eq "PPTXWord") {
                $html = $html + $file.GenerateHTML($true)
            }
        }

        $html = $html + '</div>'

        return $html
}

function GenerateHTMLReport {
    param([PPTXFile]$file)

    #Début du fichier html
    $html = 
@" 
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">

    <title>Résultats PPTX-Tool</title>
    <meta name="description" content="PPTX-Tool Results">
    <meta name="author" content="Fdfyheryery">
    
    <link rel="stylesheet" href="css/reset.css">
    <link rel="stylesheet" href="css/style.css">
	<style>
		<!-- Reset -->
		html, body, div, span, applet, object, iframe,
		h1, h2, h3, h4, h5, h6, p, blockquote, pre,
		a, abbr, acronym, address, big, cite, code,
		del, dfn, em, img, ins, kbd, q, s, samp,
		small, strike, strong, sub, sup, tt, var,
		b, u, i, center,
		dl, dt, dd, ol, ul, li,
		fieldset, form, label, legend,
		table, caption, tbody, tfoot, thead, tr, th, td,
		article, aside, canvas, details, embed, 
		figure, figcaption, footer, header, hgroup, 
		menu, nav, output, ruby, section, summary,
		time, mark, audio, video {
			margin: 0;
			padding: 0;
			border: 0;
			font-size: 100%;
			font: inherit;
			vertical-align: baseline;
		}
		article, aside, details, figcaption, figure, 
		footer, header, hgroup, menu, nav, section {
			display: block;
		}
		body {
			line-height: 1;
		}
		ol, ul {
			list-style: none;
		}
		blockquote, q {
			quotes: none;
		}
		blockquote:before, blockquote:after,
		q:before, q:after {
			content: '';
			content: none;
		}
		table {
			border-collapse: collapse;
			border-spacing: 0;
		}
	
		<!-- Style -->
		.content {
	        width: 97%;
	        margin: auto;
	        font-family: "Segoe UI";
        }

        .PPTXFile {
	        display: -moz-flex;
            display: -ms-flexbox;
	        display: flex;
	        -ms-flex-direction: column;
	        flex-direction: column;
	        margin: 10px;
	        padding: 5px 10px;
	        border: solid 1px #e2e2e2;
	        font-family: "Segoe UI";
	        font-weight: 400;
	        font-size: 16px;
        }

        .PPTXFile_img {
	        display: -moz-flex;
            display: -ms-flexbox;
	        display: flex;
	        height: 36px;
	        width: 40px;
	        font-weight: 300;
	        font-size: 22px;
	        color: #fefefe;
            align-items: center;
            justify-content: center;
	        margin-right: 6px;
	        padding-bottom: 4px;
        }

        .PPTX_others {
	        display: -moz-flex;
            display: -ms-flexbox;
	        display: flex;
	        height: 27px;
	        width: 30px;
	        background-color: #c2c2c2;
	        font-weight: 300;
	        font-size: 16px;
	        color: #fefefe;
            align-items: center;
            justify-content: center;
	        margin-right: 6px;
	        padding-bottom: 3px;
        }

        .PPTXImage {
	        max-width: 100px;
	        max-height: 75px;
	        margin: 10px 10px 0px 0px;
        }

        .PPTXWord_img {
	        background-color: #164FB3;
        }

        .PPTXExcel_img {
	        background-color: #06663C;
        }

        .PPTXPowerPoint_img {
	        background-color: #D64206;
        }

        .line {
	        display: -moz-flex;
            display: -ms-flexbox;
	        display: flex;
            -ms-flex-wrap: wrap;
            flex-wrap: wrap;
	        min-height: 40px;
            align-items: center;
            justify-content: left;
        }

        .line_child {
	        margin-left: 20px;
	        font-size: 13px;
        }

        .name {
	        min-width: 90px;
	        margin-left: 10px;
        }

        .nameLarge {
            margin-left: 74px;
        }

        .slide {
	        min-width: 70px;
	        margin-left: 20px;
	        font-size: 13px;
	        font-style: italic;
        }

        .colFlex {
	        display: -moz-flex;
            display: -ms-flexbox;
	        display: flex;
	        -ms-flex-direction: column;
	        flex-direction: column;
        }

        .warning {
	        font-size: 12px;
	        font-weight: 600;
	        color: #C8B906;
	        margin: 2px;
            max-width: 500px;
        }

        .order1 {
	        -ms-flex-order: -1;
	        order: -1;
        }

        .order2 {
	        -ms-flex-order: 2;
	        order: 2;
        }

        .order3 {
	        -ms-flex-order: 3;
	        order: 3;
        }

        .order4 {
	        -ms-flex-order: 4;
	        order: 3;
        }
	</style>
</head>

<body>
    <div class="content">
"@

    #Ajoute les items dynamiquement
    $html = $html + $file.GenerateHTML($false)
    $html = $html.Replace($appTempPath + "\", "")

    #Bloc de fin du fichier
    $html = $html + @"
	</div>
</body>
</html> 
"@
	
    return $html
}

Class PPTXFile
{
    [string]$name
    [int[]]$slides
    [int]$filesize
    [int]$total
    [string[]]$warning
}

Class PPTXImage : PPTXFile
{
    [int]$cx
    [int]$cy
    [int]$utilisationV
    [int]$utilisationH
    [string]$decompressPath

    PPTXImage ([string]$name)
    {
        $this.name = $name
    }

    [bool]CreateWarning($entry)
    {
        $hasWarning = $false

        $dPath = $this.decompressPath + "\" + $this.name
        $DirExists = Test-Path $dPath 
        if ($DirExists -eq $false) {
            New-Item -ItemType directory -Path $this.decompressPath
        }
        ExtractImgToFile $entry $dPath

        $objShell = New-Object -ComObject Shell.Application 
        $objFolder = $objShell.namespace($this.decompressPath) 
        $File = $objFolder.ParseName($this.name)

        #Calcul du ratio,  pas de metadata sur les emf et wmf
        $fileExt = $File.name.Split(".")[-1]
        if ($fileExt -ne "emf" -and $fileExt -ne "wmf") {
            $width = $objFolder.getDetailsOf($File, 162)
            $height = $objFolder.getDetailsOf($File, 164)

            $width = $width.replace(" pixels","").remove(0,1)
            $height = $height.replace(" pixels","").remove(0,1)

            $ratioX = ([double]$width * 9525) / $this.cx
            $ratioY = ([double]$height * 9525) / $this.cy

            if ($ratioX -ge 2 -and $ratioY -ge 2 -and [double]$width -gt 200) {
                $this.warning += "La taille de cette image est " + $ratioX.ToString("0.0") + " fois plus grande que son utilisation"
                $hasWarning = $true
            }

        }

        if ($this.utilisationH -le 90000) {
            $this.warning += "Seulement " + ($this.utilisationH / 1000).ToString("0") + "% de l'image est utilisée horizontalement"
        }

        if ($this.utilisationV -le 90000) {
            $this.warning += "Seulement " + ($this.utilisationV / 1000).ToString("0") + "% de l'image est utilisée verticalement"
        }

        $this.filesize = $entry.Length
        if ($this.filesize -gt 1MB) {
            $this.warning += "Cette image prend " + ($this.filesize / 1MB).ToString("0.00") + "MB"
            $hasWarning = $true
        }

        return $hasWarning;
    }

    [string]GenerateHTML([bool]$isChild)
    {
        $html = '<div class="line line_child"><div style="width:100px;"><img class="PPTXImage" src="' + $this.decompressPath `
            + "/" + $this.name +'" /></div><span class="name">' + $this.name + '</span><span class="slide">'
        
        for($i=0;$i -lt $this.slides.Length;$i++) {
            $html += $this.slides[$i]
            if ($i-lt $this.slides.Length - 1) {
                $html += ", "
            }
        }

        $html += '</span><div class="colFlex">'

        foreach ($warning in $this.warning) {
            $html += '<span class="warning">' + $warning + '</span>'
        }

        $html += '</div></div>'
        return $html
    }
}

Class PPTXVideo : PPTXFile
{
    #Cette information n'est probablement pas pertinente (devrait décompresser le fichier pour accéder aux metadatas)
    #[double]$length

    PPTXVideo ([string]$name)
    {
        $this.name = $name
    }

    [bool]CreateWarning($entry)
    {
        #Retourne toujours le poid du fichier vidéo comme avertissement
        $this.warning += "Cette vidéo prend " + ($this.filesize / 1MB).ToString("0.00") + "MB"
        return $true;
    }

    [string]GenerateHTML([bool]$isChild)
    {
        $html = '<div class="line line_child order2"><div class="PPTX_others">V</div>'`
            + '<span class="name nameLarge">' + $this.name + '</span><span class="slide">' + $this.slides + '</span><div class="colFlex">'

        foreach ($warning in $this.warning) {
            $html += '<span class="warning">' + $warning + '</span>'
        }

        $html += '</div></div>'
        return $html
    }
}

Class PPTXExcel : PPTXFile
{
    [PPTXFile[]]$arrayImages
    hidden $zipArchive
    [int]$conditionalFormat

    PPTXExcel ([string]$name, [bool]$analyseNow)
    {
        
        if ($analyseNow) {
            $this.name = $name.split("\")[-1]
            CallAnalyzeFromName $this $name
            $this.warning = $this.CreateWarning()
        }
        else {
            $this.name = $name
        }
    }

    hidden [void] AnalyzeFile() {
        #On retrouve les images sous xl/drawings/drawingX.xml
        #et les fichiers + règles de formattage conditionnel sous xl/sheets/sheetX.xml
        
        #On incrémente et vérifie si le drawing existe (commencent à 1)
        $i = 1
        $drawingExist = $true;
        while($drawingExist -eq $true) {
            $docPath = "xl/drawings/drawing" + $i + ".xml"
            $entry = $this.zipArchive.GetEntry($docPath)

            if ($entry) {
                $rIds = $null
                [PPTXFile[]]$rIds = @()

                $docContent = GetEntryAsXML $entry
            
                #Image
                $namespace = @{xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
                $pics = $docContent | Select-Xml -Namespace $namespace -XPath "//xdr:pic"
            
                foreach ($pic in $pics.Node) {
                    $rIds = GetImageFromXML $rIds $pic
                }

                #Référence dans le fichier xml.rels 
                $relsPath = "xl/drawings/_rels/drawing" + $i + ".xml.rels"
                $entry = $this.zipArchive.GetEntry($relsPath)

                if ($entry) {

                    [xml]$relsContent = GetEntryAsXML $entry

                    for($j=0;$j -lt $rIds.Length;$j++) {
                        if ($relsContent.relationships.Relationship.getType().Name -eq "XmlElement") {
                            $xmlNode = $relsContent.relationships.Relationship
                        }
                        else {
                            $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        }
                        
                        GetRelsFromXML $this $rIds[$j] $xmlNode $i
                    }  
                }
            }

            else {
                $drawingExist = $false
            }

            $i++
        }

        #On incrémente et vérifie si le sheet existe (commencent à 1)
        $i = 1
        $sheetExist = $true;
        while($sheetExist -eq $true) {
            $docPath = "xl/worksheets/sheet" + $i + ".xml"
            $entry = $this.zipArchive.GetEntry($docPath)

            if ($entry) {
                $rIds = $null
                [PPTXFile[]]$rIds = @()

                $docContent = GetEntryAsXML $entry
            
                #Image
                $namespace = @{mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"}
                $fallbacks = $docContent | Select-Xml -Namespace $namespace -XPath "//mc:Fallback"
            
                foreach ($fallback in $fallbacks.Node) {
                    #rID
                    $rId = $fallback.oleObject.id

                    if ($rId -ne $null) {
                        if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
                            $index = $rIds.name.indexof($rId)
                            $rIds[$index].Total++
                        }
                        else {
                            $itemType = $fallback.oleObject.progId

                            if ($itemType -eq "Word") {
                                $newItem = [PPTXWord]::new($rId, $false)
                            }
                            elseif ($itemType -eq "Exce") {
                                $newItem = [PPTXExcel]::new($rId, $false)
                            }
                            elseif ($itemType -eq "Présentation") {
                                $newItem = [PPTXPowerPoint]::new($rId, $false)
                            }
                            else {
                                $newItem = [PPTXFile]::new()
                                $newItem.name = $rId
                            }
                            
                            $newItem.total = 1
                            $rIds += $newItem
                        }
                    }
                }

                #Formattage conditionnel
                $this.conditionalFormat += $docContent.worksheet.conditionalFormatting.Count

                #Référence dans le fichier xml.rels 
                $relsPath = "xl/worksheets/_rels/sheet" + $i + ".xml.rels"
                $entry = $this.zipArchive.GetEntry($relsPath)

                if ($entry) {

                    [xml]$relsContent = GetEntryAsXML $entry

                    for($j=0;$j -lt $rIds.Length;$j++) {
                        if ($relsContent.relationships.Relationship.getType().Name -eq "XmlElement") {
                            $xmlNode = $relsContent.relationships.Relationship
                        }
                        else {
                            $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        }
                        GetRelsFromXML $this $rIds[$j] $xmlNode $i
                    }  
                }
            }

            else {
                $sheetExist = $false
            }

            $i++
        }      

        CreateFileWarnings $this
    }

    hidden [string[]]CreateWarning() {
        [string[]]$warningMsg = @()

        if ($this.conditionalFormat -gt 100) {
            $warningMsg += "Il y a " + $this.conditionalFormat + " règles de formattage conditionnel."
        }

        return $warningMsg
    }

    [bool]CreateWarning($entry)
    {
        CallAnalyzeFromEntry $this $entry

        $hasWarning = $false

        $this.filesize = $entry.Length
        if ($this.filesize -gt 1MB) {
            $this.warning += "Ce fichier Excel pèse " + ($this.filesize / 1MB).ToString("0.00") + "MB"
            return $true;
        }

        $this.warning += $this.CreateWarning()
        
        return $false;
    }

    [string]GenerateHTML([bool]$isChild)
    {
        return GenerateHTML $this $isChild
    }
}

Class PPTXPowerPoint : PPTXFile
{
    [PPTXFile[]]$arrayImages
    hidden $zipArchive

    PPTXPowerPoint ([string]$name, [bool]$analyseNow)
    {
        if ($analyseNow) {
            $this.name = $name.split("\")[-1]
            CallAnalyzeFromName $this $name
        }
        else {
            $this.name = $name
        }
    }

    hidden [void] AnalyzeFile() {
        #On incrémente et vérifie si la slide existe (pptx commencent à 1)
        $i = 1
        $slideExist = $true;
        while($slideExist -eq $true) {
        
            $slidePath = "ppt/slides/slide" + $i + ".xml"
            $entry = $this.zipArchive.GetEntry($slidePath)

            if ($entry) {
                $rIds = $null
                [PPTXFile[]]$rIds = @()

                $slideContent = GetEntryAsXML $entry

                #Image et Vidéo
                foreach ($pic in $slideContent.sld.csld.sptree.pic) {
                    $rIds = GetImageFromXML $rIds $pic
                }

                #Word, Excel, PowerPoint
                foreach ($graphic in $slideContent.sld.csld.sptree.graphicframe) {
                    #rID
                    $rId = $graphic.graphic.graphicdata.alternatecontent.fallback.oleobj.id

                    if ($rId -ne $null) {
                        if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
                            $index = $rIds.name.indexof($rId)
                            $rIds[$index].Total++
                        }
                        else {
                            $itemtype = $graphic.graphic.graphicData.AlternateContent.Fallback.oleObj.progId.Substring(0,4)

                            if ($itemType -eq "Word") {
                                $newItem = [PPTXWord]::new($rId, $false)
                            }
                            elseif ($itemType -eq "Exce") {
                                $newItem = [PPTXExcel]::new($rId, $false)
                            }
                            elseif ($itemType -eq "Powe") {
                                $newItem = [PPTXPowerPoint]::new($rId, $false)
                            }
                            else {
                                $newItem = [PPTXFile]::new()
                                $newItem.name = $rId
                            }
                            
                            $newItem.total = 1
                            $rIds += $newItem
                        }
                    }
                }

                #Référence dans le fichier xml.rels 
                $relsPath = "ppt/slides/_rels/slide" + $i + ".xml.rels"
                $entry = $this.zipArchive.GetEntry($relsPath)

                if ($entry) {

                    [xml]$relsContent = GetEntryAsXML $entry

                    for($j=0;$j -lt $rIds.Length;$j++) {
                        if ($relsContent.relationships.Relationship.getType().Name -eq "XmlElement") {
                            $xmlNode = $relsContent.relationships.Relationship
                        }
                        else {
                            $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        }  $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        GetRelsFromXML $this $rIds[$j] $xmlNode $i
                    }  
                }

                else {
                    #Normalement il y a toujours un fichier .rels d'associé à une diapositive
                    $errorMsg = "Erreur: Fichier " + $relsPath + " introuvable."
                    Write-Host $errorMsg
                }
            }

            else {
                $slideExist = $false
            }
            $i++
        }

        CreateFileWarnings $this
    }

    [bool]CreateWarning($entry)
    {
        CallAnalyzeFromEntry $this $entry

        #Pas d'avertissement sur les fichiers PowerPoint pour l'instant

        return $false;
    }

    [string]GenerateHTML([bool]$isChild)
    {
        return GenerateHTML $this $isChild
    }
}

Class PPTXWord : PPTXFile
{
    [PPTXFile[]]$arrayImages
    hidden $zipArchive

    PPTXWord ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name.split("\")[-1]
        if ($analyseNow) {
            $this.name = $name.split("\")[-1]
            CallAnalyzeFromName $this $name
        }
        else {
            $this.name = $name
        }
    }

    hidden [void] AnalyzeFile() {
        $docPath = "word/document.xml"
        $entry = $this.zipArchive.GetEntry($docPath)

        if ($entry) {
            $rIds = $null
            [PPTXFile[]]$rIds = @()

            $docContent = GetEntryAsXML $entry
            
            #Image
            $namespace = @{pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"}
            $pics = $docContent | Select-Xml -Namespace $namespace -XPath "//pic:pic"
            
            foreach ($pic in $pics.Node) {
                $rIds = GetImageFromXML $rIds $pic
            }

            #Word, Excel, PowerPoint
            $namespace = @{o = "urn:schemas-microsoft-com:office:office"}
            $OLEObjects = $docContent | Select-Xml -Namespace $namespace -XPath "//o:OLEObject"

            foreach ($OLEObject in $OLEObjects.Node) {
                #rID
                $rId = $OLEObject.id

                if ($rId -ne $null) {
                    if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
                        $index = $rIds.name.indexof($rId)
                        $rIds[$index].Total++
                    }
                    else {
                        $itemtype = $OLEObject.progId.Substring(0,4)

                        if ($itemType -eq "Word") {
                            $newItem = [PPTXWord]::new($rId, $false)
                        }
                        elseif ($itemType -eq "Exce") {
                            $newItem = [PPTXExcel]::new($rId, $false)
                        }
                        elseif ($itemType -eq "Powe") {
                            $newItem = [PPTXPowerPoint]::new($rId, $false)
                        }
                        else {
                            $newItem = [PPTXFile]::new()
                            $newItem.name = $rId
                        }
                            
                        $newItem.total = 1
                        $rIds += $newItem
                    }
                }
            }

            #Référence dans le fichier xml.rels 
            $relsPath = "word/_rels/document.xml.rels"
            $entry = $this.zipArchive.GetEntry($relsPath)

            if ($entry) {

                [xml]$relsContent = GetEntryAsXML $entry

                for($j=0;$j -lt $rIds.Length;$j++) {
                    if ($relsContent.relationships.Relationship.getType().Name -eq "XmlElement") {
                        $xmlNode = $relsContent.relationships.Relationship
                    }
                    else {
                        $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                    }
                    GetRelsFromXML $this $rIds[$j] $xmlNode 0
                }  
            }
        }

        CreateFileWarnings $this
    }

    [bool]CreateWarning($entry)
    {
        CallAnalyzeFromEntry $this $entry

        $this.filesize = $entry.Length
        if ($this.filesize -gt 3KB) {
            $this.warning += "Ce fichier Word pèse " + ($this.filesize / 1MB).ToString("0.00") + "MB"
            return $true;
        }
        return $false;
    }

    [string]GenerateHTML([bool]$isChild)
    {
        return GenerateHTML $this $isChild
    }
}

#Ouvre une fenêtre pour la sélection du fichier
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Word, Excel, PowerPoint (*.docx, *.xlxs, *pptx)|*.docx;*.pptx;*.xlsx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    #Crée un répertoire temporaire pour le rapport HTML (images)
    if(Test-Path $appTempPath) {
        Remove-Item $appTempPath -Force -Recurse
    }

    $newDir = New-Item -ItemType directory -Path $appTempPath

    #Lance l'analyse
    if ($openFileDialog.FileName.Substring($openFileDialog.FileName.Length - 4) -eq "pptx") {
        [PPTXPowerPoint]$analyzedFile = [PPTXPowerPoint]::new($openFileDialog.FileName, $true)
    }
    elseif ($openFileDialog.FileName.Substring($openFileDialog.FileName.Length - 4) -eq "docx") {
        [PPTXWord]$analyzedFile = [PPTXWord]::new($openFileDialog.FileName, $true)
    }
    elseif ($openFileDialog.FileName.Substring($openFileDialog.FileName.Length - 4) -eq "xlsx") {
        [PPTXExcel]$analyzedFile = [PPTXExcel]::new($openFileDialog.FileName, $true)
    }

    #Affichage temporaire
    #$analyzedFile.arrayImages | Where-Object {$_.warning}

    #Génération du rapport HTML
    $html = GenerateHTMLReport $analyzedFile
    $path = $appTempPath + "\results.html"
    $html | Out-File -filepath $path

    Invoke-Item $path
    start-sleep 5

    #Détruit les fichiers temporaires
    if(Test-Path $appTempPath) {
        Remove-Item $appTempPath -Force -Recurse
    }
    
}