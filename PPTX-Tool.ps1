<#

.SYNOPSIS
WIP

.DESCRIPTION
WIP

.NOTES
WIP

#>

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

    $arrayImages = @()

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
            $rIds = @()

            $slideContent = GetEntryAsText $entry

            #Va chercher le ratio et rId pour chaque image de la slide
            [xml]$slideContent = $slideContent

            foreach ($pic in $slideContent.sld.csld.sptree.pic) {
                $rId = $pic.blipfill.blip.embed
                if ($pic.sppr.xfrm.ext.cx -lt $pic.sppr.xfrm.ext.cy) {
                    $ratio = $pic.sppr.xfrm.ext.cx / 914400
                    
                }
                else {
                    $ratio = $pic.sppr.xfrm.ext.cy / 914400
                }

                if (($rIds.Count -gt 0) -and ($rIds.Values.Contains($rId))) {
                    $index = [math]::floor($rIds.Values.indexof($image)/$rIds[0].Count)
                    $rIds[$index].Total++

                    if ($rIds[$index].Ratio -gt $ratio) {
                        $rIds[$index].Ratio = $ratio
                    }
                }
                else {
                    $rIds += @{"rId" = $rId;"Total" = 1; "Ratio" = $ratio}
                }
            }

            #Va chercher la bonne référence dans le fichier xml.rels 
            $relsPath = "ppt/slides/_rels/slide" + $i + ".xml.rels"
            $entry = $zipArchive.GetEntry($relsPath)

            if ($entry) {

                [xml]$relsContent = GetEntryAsXML $entry

                #Va chercher, pour chaque rId, le nom de l'image associée, puis met à jour les informations (ou ajoute l'entrée si non-existant)
                for($j=0;$j -lt $rIds.Length;$j++) {
                    $image = $relsContent.Relationships.Relationship `
                    | Where-Object {($_.Id -eq $rIds[$j].rId) -and ($_.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")} `
                    | Foreach-Object {$_.Target.Substring(9)}
                    if (($arrayImages.Count -gt 0) -and ($arrayImages.Values.Contains($image))) {
                        $indexImage = [math]::floor($arrayImages.Values.indexof($image)/$rIds[0].Count)
                        $arrayImages[$indexImage].Total = $arrayImages[$indexImage].Total + $rIds[$j].Total
                        if ($arrayImages[$indexImage].Ratio -gt $rIds[$j].Ratio) {
                            $arrayImages[$indexImage].Ratio = $rIds[$j].Ratio
                        }
                    }
                    else {
                        $arrayImages += @{"Total"= $rIds[$j].Total; "Name" = $image; "Ratio" = $rIds[$j].Ratio}
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

function EvalImages {
    param([string]$filename, [hashtable[]]$hashImages)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($filename)

    foreach ($image in $hashImages) {
        $imgPath = "ppt/media/" + $image.Name
        $entry = $zipArchive.GetEntry($imgPath)

        # TODO: Générer les avertissements, exemple ci-dessous

        if (($entry.length / 1MB) -gt 1) {
            $image.FileSize = $entry.Length
            $image.FileType = "Image"
            $image.Message = "Cette image à un poid supérieur à 1MB"
        }
    }

    $zipArchive.Dispose()
    return $hashImages
}

#Ouvre une fenêtre pour la sélection du fichier PowerPoint
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Powerpoint Presentations (*.pptx)|*.pptx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    $images = FindUsedImages -filename $openFileDialog.FileName

    $images = EvalImages -filename $openFileDialog.FileName -hashImages $images

    #Affichage temporaire
    $images | Where-Object {$_.Message}
}