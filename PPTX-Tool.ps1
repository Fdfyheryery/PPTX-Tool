<#

.SYNOPSIS
WIP

.DESCRIPTION
WIP

.NOTES
WIP

#>

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

    [bool]CreateWarning()
    {
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

    [bool]CreateWarning()
    {
        if ($this.filesize -gt 10KB) {
            $this.warning = "Cette vidéo à un poid supérieur à 10KB"
            return $true;
        }
        return $false;
    }
}

function GetEntryAsXML {
    param([System.IO.Compression.ZipArchiveEntry]$entry)

    $slide = $entry.Open()
    $reader = New-Object IO.StreamReader($slide)
    [xml]$entryXML = $reader.ReadToEnd()
    $reader.Close()
    $slide.Close()
    return $entryXML
}

function FindUsedImages {
    param([string]$filename)

    [PPTXFile[]]$arrayImages = @()

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($filename)

    #On incrémente et vérifie si la slide existe (pptx commencent à 1)
    $i = 1
    $slideExist = $true;
    while($slideExist -eq $true) {
        
        $slidePath = "ppt/slides/slide" + $i + ".xml"
        $entry = $zipArchive.GetEntry($slidePath)

        if ($entry) {
            $rIds = $null
            [PPTXFile[]]$rIds = @()

            $slideContent = GetEntryAsXML $entry

            #Va chercher le ratio et rId pour chaque image de la slide
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

            #Va chercher la bonne référence dans le fichier xml.rels 
            $relsPath = "ppt/slides/_rels/slide" + $i + ".xml.rels"
            $entry = $zipArchive.GetEntry($relsPath)

            if ($entry) {

                [xml]$relsContent = GetEntryAsXML $entry

                #Va chercher, pour chaque rId, le nom de l'image associée, puis met à jour les informations (ou ajoute l'entrée si non-existant)
                for($j=0;$j -lt $rIds.Length;$j++) {
                    $xmlNode = $relsContent.relationships.Relationship.Where({$_.Id -eq $rIds[$j].name})
                    $image = $xmlNode.Target.Substring(9)

                    #Images
                    if ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                        if (($arrayImages.Count -gt 0) -and ($arrayImages.Name -contains $image)) {
                            $indexImage = $arrayImages.name.indexof($image)

                            $arrayImages[$indexImage].Total = $arrayImages[$indexImage].Total + $rIds[$j].Total
                            $arrayImages[$indexImage].Slides += $i

                            if ($arrayImages[$indexImage].Ratio -gt $rIds[$j].ratio) {
                                $arrayImages[$indexImage].Ratio = $rIds[$j].Ratio
                            }

                            if ($arrayImages[$indexImage].utilisationV -lt $rIds[$j].utilisationV) {
                                $arrayImages[$indexImage].utilisationV = $rIds[$j].utilisationV
                            }

                            if ($arrayImages[$indexImage].utilisationH -lt $rIds[$j].utilisationH) {
                                $arrayImages[$indexImage].utilisationH = $rIds[$j].utilisationH
                            }
                        }
                        else {
                            $newItem = [PPTXImage]::new($image)
                            $newItem.ratio = $rIds[$j].ratio
                            $newItem.utilisationV = $rIds[$j].utilisationV
                            $newItem.utilisationH = $rIds[$j].utilisationH
                            $newItem.slides = @($i)
                            $newItem.total = $rIds[$j].total
                            $arrayImages += $newItem
                        }
                    }

                    #Videos
                    if ($xmlNode.Type -eq "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video") {
                        if (($arrayImages.Count -gt 0) -and ($arrayImages.Name -contains $image)) {
                            $indexImage = $arrayImages.name.indexof($image)

                            $arrayImages[$indexImage].Total = $arrayImages[$indexImage].Total + $rIds[$j].Total
                            $arrayImages[$indexImage].Slides += $i
                        }
                        else {
                            $newItem = [PPTXVideo]::new($image)
                            $newItem.slides = @($i)
                            $newItem.total = $rIds[$j].total
                            $arrayImages += $newItem
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
    $zipArchive.Dispose()
    return $arrayImages
}

function GenerateWarnings {
    param([string]$filename, [PPTXFile[]]$fileArray)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($filename)

    foreach ($file in $fileArray) {
        $filePath = "ppt/media/" + $file.Name
        $entry = $zipArchive.GetEntry($filePath)
        $file.filesize = $entry.Length
        $hasWarning = $file.CreateWarning()
    }

    $zipArchive.Dispose()
    return $fileArray
}

#Ouvre une fenêtre pour la sélection du fichier PowerPoint
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Powerpoint Presentations (*.pptx)|*.pptx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    $images = FindUsedImages -filename $openFileDialog.FileName

    $images = GenerateWarnings -filename $openFileDialog.FileName -fileArray $images

    #Affichage temporaire
    $images | Where-Object {$_.warning}
}