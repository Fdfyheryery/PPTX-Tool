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

    [bool]CreateWarning([System.IO.Compression.ZipArchiveEntry]$entry)
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

    [bool]CreateWarning([System.IO.Compression.ZipArchiveEntry]$entry)
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

    PPTXExcel ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            #$this.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($this.name)
            #$this.AnalyzeFile()
            #$this.zipArchive.Dispose()
        }
    }

    [bool]CreateWarning([System.IO.Compression.ZipArchiveEntry]$entry)
    {
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
    hidden [System.IO.Compression.ZipArchive]$zipArchive

    PPTXPowerPoint ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            $this.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($this.name)
            $this.AnalyzeFile()
            $this.zipArchive.Dispose()
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
                    #rID
                    $rId = $pic.blipfill.blip.embed

                    #Ratio (Image source : Image PPTX)
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

                    #Pour chaque rId: le nom du fichier associée, puis met à jour les informations ou créé l'entrée
                    for($j=0;$j -lt $rIds.Length;$j++) {
                        $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                        

                        #Images
                        if ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                            $image = $xmlNode.Target.Substring(9)
                            if (($this.arrayImages.Count -gt 0) -and ($this.arrayImages.Name -contains $image)) {
                                $indexImage = $this.arrayImages.name.indexof($image)

                                $this.arrayImages[$indexImage].Total = $this.arrayImages[$indexImage].Total + $rIds[$j].Total
                                $this.arrayImages[$indexImage].Slides += $i

                                if ($this.arrayImages[$indexImage].Ratio -gt $rIds[$j].ratio) {
                                    $this.arrayImages[$indexImage].Ratio = $rIds[$j].Ratio
                                }

                                if ($this.arrayImages[$indexImage].utilisationV -lt $rIds[$j].utilisationV) {
                                    $this.arrayImages[$indexImage].utilisationV = $rIds[$j].utilisationV
                                }

                                if ($this.arrayImages[$indexImage].utilisationH -lt $rIds[$j].utilisationH) {
                                    $this.arrayImages[$indexImage].utilisationH = $rIds[$j].utilisationH
                                }
                            }
                            else {
                                $newItem = [PPTXImage]::new($image)
                                $newItem.ratio = $rIds[$j].ratio
                                $newItem.utilisationV = $rIds[$j].utilisationV
                                $newItem.utilisationH = $rIds[$j].utilisationH
                                $newItem.slides = @($i)
                                $newItem.total = $rIds[$j].total
                                $this.arrayImages += $newItem
                            }
                        }

                        #Videos
                        elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video") {
                            $image = $xmlNode.Target.Substring(9)
                            if (($this.arrayImages.Count -gt 0) -and ($this.arrayImages.Name -contains $image)) {
                                $indexImage = $this.arrayImages.name.indexof($image)

                                $this.arrayImages[$indexImage].Total = $this.arrayImages[$indexImage].Total + $rIds[$j].Total
                                $this.arrayImages[$indexImage].Slides += $i
                            }
                            else {
                                $newItem = [PPTXVideo]::new($image)
                                $newItem.slides = @($i)
                                $newItem.total = $rIds[$j].total
                                $this.arrayImages += $newItem
                            }
                        }

                        #Word, Excel, PowerPoint
                        elseif ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package") {
                            $image = $xmlNode.Target.Substring(14)
                            if (($this.arrayImages.Count -gt 0) -and ($this.arrayImages.Name -contains $image)) {
                                $indexImage = $this.arrayImages.name.indexof($image)

                                $this.arrayImages[$indexImage].Total = $this.arrayImages[$indexImage].Total + $rIds[$j].Total
                                $this.arrayImages[$indexImage].Slides += $i
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
                                $newItem.total = $rIds[$j].total
                                $this.arrayImages += $newItem
                            }
                        }
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

        foreach ($file in $this.arrayImages) {
            $filePath = ""

            if ($file.GetType().Name -eq "PPTXImage" -or $file.GetType().Name -eq "PPTXVideo") {
                $filePath = "ppt/media/" + $file.Name
            }

            elseif ($file.GetType().Name -eq "PPTXPowerPoint" -or $file.GetType().Name -eq "PPTXExcel" -or $file.GetType().Name -eq "PPTXWord") {
                $filePath = "ppt/embeddings/" + $file.Name
            }

            $entry = $this.zipArchive.GetEntry($filePath)
            $hasWarning = $file.CreateWarning($entry)
        }
    }

    [bool]CreateWarning([System.IO.Compression.ZipArchiveEntry]$entry)
    {
        #TODO: Retirer les commentaires pour tester la récursivité
        #$this.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($entry)
        #$this.AnalyzeFile()
        #$this.zipArchive.Dispose()

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
    [PPTXFile[]]$arrayFiles

    PPTXWord ([string]$name, [bool]$analyseNow)
    {
        $this.name = $name
        if ($analyseNow) {
            #$this.zipArchive = [System.IO.Compression.ZipFile]::OpenRead($this.name)
            #$this.AnalyzeFile()
            #$this.zipArchive.Dispose()
        }
    }

    [bool]CreateWarning([System.IO.Compression.ZipArchiveEntry]$entry)
    {
        $this.filesize = $entry.Length
        if ($this.filesize -gt 3KB) {
            $this.warning = "Ce fichier Word à un poid supérieur à 3KB"
            return $true;
        }
        return $false;
    }
}

#Ouvre une fenêtre pour la sélection du fichier PowerPoint
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Powerpoint Presentations (*.pptx)|*.pptx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    [PPTXPowerPoint]$analyzedFile = [PPTXPowerPoint]::new($openFileDialog.FileName, $true)

    #Affichage temporaire
    $analyzedFile.arrayImages | Where-Object {$_.warning}
    
}