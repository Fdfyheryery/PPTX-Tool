<#

.SYNOPSIS
WIP

.DESCRIPTION
WIP

.NOTES
WIP

#>

Add-Type -AssemblyName System.IO.Compression.FileSystem

function GetEntryAsXML {
    param([System.IO.Compression.ZipArchiveEntry]$entry)

    $slide = $entry.Open()
    $reader = New-Object IO.StreamReader($slide)
    [xml]$entryXML = $reader.ReadToEnd()
    $reader.Close()
    $slide.Close()
    return $entryXML
}

#Bypassing a class limitation
function CallAnalyzeFromEntry {
    param([PPTXFile]$file, [System.IO.Compression.ZipArchiveEntry]$entry)

    $file.zipArchive = [System.IO.Compression.ZipArchive]::New($entry.Open())
    $file.AnalyzeFile()
    $file.zipArchive.Dispose()
}

#Bypassing a class limitation
function CallAnalyzeFromName {
    param([PPTXFile]$file, [string]$name)

    $file.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($name)
    $file.AnalyzeFile()
    $file.zipArchive.Dispose()
}

function GetImageFromXML {
    param([PPTXFile[]]$rIds, $pic)

    #rID
    $rId = $pic.blipfill.blip.embed

    #Ratio
    if ($pic.sppr.xfrm.ext.cx -lt $pic.sppr.xfrm.ext.cy) {
        $ratio = $pic.sppr.xfrm.ext.cx / 914400
                    
    }
    else {
        $ratio = $pic.sppr.xfrm.ext.cy / 914400
    }

    #Utilisation (Rognage) (10000 = 10.000%)
    $utilVertical = 100000 - ([int]$pic.blipfill.srcRect.t + [int]$pic.blipfill.srcRect.b)
    $utilHorizontal = 100000 - ([int]$pic.blipfill.srcRect.l + [int]$pic.blipfill.srcRect.r)


    if (($rIds.Count -gt 0) -and ($rIds.Name -contains $rId)) {
        $index = $rIds.name.indexof($rId)
        $rIds[$index].Total++

        if ($rIds[$index].Ratio -gt $ratio) {
            $rIds[$index].Ratio = $ratio
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
        $newItem.ratio = $ratio
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
    param([PPTXFile]$file, [PPTXFile]$RIdItem, $xml)

    $image = $xmlNode.Target.split("/")[-1]

    #Images
    if ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
        
        if (($file.arrayImages.Count -gt 0) -and ($file.arrayImages.Name -contains $image)) {
            $indexImage = $file.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $i

            if ($file.arrayImages[$indexImage].Ratio -gt $RIdItem.ratio) {
                $file.arrayImages[$indexImage].Ratio = $RIdItem.Ratio
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
            $newItem.ratio = $RIdItem.ratio
            $newItem.utilisationV = $RIdItem.utilisationV
            $newItem.utilisationH = $RIdItem.utilisationH
            $newItem.slides = @($i)
            $newItem.total = $RIdItem.total
            $file.arrayImages += $newItem
        }
    }

    #Videos
    elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video") {
        if (($file.arrayImages.Count -gt 0) -and ($file.arrayImages.Name -contains $image)) {
            $indexImage = $file.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $i
        }
        else {
            $newItem = [PPTXVideo]::new($image)
            $newItem.slides = @($i)
            $newItem.total = $RIdItem.total
            $file.arrayImages += $newItem
        }
    }

    #Word, Excel, PowerPoint
    elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
        if (($file.arrayImages.Count -gt 0) -and ($this.arrayImages.Name -contains $image)) {
            $indexImage = $this.arrayImages.name.indexof($image)

            $file.arrayImages[$indexImage].Total = $file.arrayImages[$indexImage].Total + $RIdItem.Total
            $file.arrayImages[$indexImage].Slides += $i
        }
        else {
            $itemType = $image.Substring($image.get_Length()-4)

            if ($itemType -eq "pptx") {
                $newItem = [PPTXPowerPoint]::new($image, $false)
            }
            elseif ($itemType -eq "xlsx") {
                $newItem = [PPTXExcel]::new($image, $false)
            }
            elseif ($itemType -eq "docx") {
                $newItem = [PPTXWord]::new($image, $false)
            }
            else {
                $newItem = [PPTXFile]::new()
                $newItem.name = $image
            }
                                
            $newItem.slides = @($i)
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

        if ($file.GetType().Name -eq "PPTXImage" -or $file.GetType().Name -eq "PPTXVideo") {
            $filePath = $startPath + "media/" + $file.Name
        }

        elseif ($file.GetType().Name -eq "PPTXPowerPoint" -or $file.GetType().Name -eq "PPTXExcel" -or $file.GetType().Name -eq "PPTXWord") {
            $filePath = $startPath + "embeddings/" + $file.Name
        }

        $entry = $pptxfile.zipArchive.GetEntry($filePath)
        $hasWarning = $file.CreateWarning($entry)
    }
}

Class PPTXFile
{
    [string]$name
    [int[]]$slides
    [int]$filesize
    [int]$total
    [string]$warning
}

Class PPTXImage : PPTXFile
{
    [double]$ratio
    [int]$utilisationV
    [int]$utilisationH

    PPTXImage ([string]$name)
    {
        $this.name = $name
    }

    [bool]CreateWarning($entry)
    {
        $this.filesize = $entry.Length
        if ($this.filesize -gt 1KB) {
            $this.warning = "Cette image à un poid supérieur à 1KB"
            return $true;
        }
        return $false;
    }
}

Class PPTXVideo : PPTXFile
{
    [double]$length

    PPTXVideo ([string]$name)
    {
        $this.name = $name
    }

    [bool]CreateWarning($entry)
    {
        $this.filesize = $entry.Length
        if ($this.filesize -gt 10KB) {
            $this.warning = "Cette vidéo à un poid supérieur à 10KB"
            return $true;
        }
        return $false;
    }
}

Class PPTXExcel : PPTXFile
{
    [PPTXFile[]]$arrayFiles
    hidden $zipArchive

    PPTXExcel ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            #CallAnalyzeFromName $this $name
        }
    }

    [bool]CreateWarning($entry)
    {
        #CallAnalyzeFromEntry $this $entry

        $this.filesize = $entry.Length
        if ($this.filesize -gt 1KB) {
            $this.warning = "Ce fichier Excel à un poid supérieur à 1KB"
            return $true;
        }
        return $false;
    }
}

Class PPTXPowerPoint : PPTXFile
{
    [PPTXFile[]]$arrayImages
    hidden $zipArchive

    PPTXPowerPoint ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            CallAnalyzeFromName $this $name
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
                        $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        GetRelsFromXML $this $rIds[$j] $xmlNode
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

        $this.filesize = $entry.Length
        if ($this.filesize -gt 2KB) {
            $this.warning = "Ce fichier PowerPoint à un poid supérieur à 2KB"
            return $true;
        }
        return $false;
    }
}

Class PPTXWord : PPTXFile
{
    [PPTXFile[]]$arrayImages
    hidden $zipArchive

    PPTXWord ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            CallAnalyzeFromName $this $name
        }
    }

    hidden [void] AnalyzeFile() {
        $docPath = "word/document.xml"
        $entry = $this.zipArchive.GetEntry($docPath)

        if ($entry) {
            $rIds = $null
            [PPTXFile[]]$rIds = @()

            $docContent = GetEntryAsXML $entry
            
            $namespace = @{pic = "http://schemas.openxmlformats.org/drawingml/2006/picture"}
            $pics = $docContent | Select-Xml -Namespace $namespace -XPath "//pic:pic"
            
            foreach ($pic in $pics.Node) {
                $rIds = GetImageFromXML $rIds $pic
            }

            #Référence dans le fichier xml.rels 
            $relsPath = "word/_rels/document.xml.rels"
            $entry = $this.zipArchive.GetEntry($relsPath)

            if ($entry) {

                [xml]$relsContent = GetEntryAsXML $entry

                for($j=0;$j -lt $rIds.Length;$j++) {
                    $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                    GetRelsFromXML $this $rIds[$j] $xmlNode
                }  
            }
        }

        CreateFileWarnings $this
    }

    [bool]CreateWarning($entry)
    {
        #CallAnalyzeFromEntry $this $entry

        $this.filesize = $entry.Length
        if ($this.filesize -gt 3KB) {
            $this.warning = "Ce fichier Word à un poid supérieur à 3KB"
            return $true;
        }
        return $false;
    }
}

#Ouvre une fenêtre pour la sélection du fichier
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Word, Excel, PowerPoint (*.docx, *.xlxs, *pptx)|*.docx;*.pptx;*.xlsx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    if ($openFileDialog.FileName.Substring($openFileDialog.FileName.Length - 4) -eq "pptx") {
        [PPTXPowerPoint]$analyzedFile = [PPTXPowerPoint]::new($openFileDialog.FileName, $true)
    }
    elseif ($openFileDialog.FileName.Substring($openFileDialog.FileName.Length - 4) -eq "docx") {
        [PPTXWord]$analyzedFile = [PPTXWord]::new($openFileDialog.FileName, $true)
    }
    

    #Affichage temporaire
    $analyzedFile.arrayImages | Where-Object {$_.warning}
    
}