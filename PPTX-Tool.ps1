<#

.SYNOPSIS
WIP

.DESCRIPTION
WIP

.NOTES
WIP

#>

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
            $rIds = @(0,0)

            $slide = $entry.Open()
            $reader = New-Object IO.StreamReader($slide)
            $slideContent = $reader.ReadToEnd()

            #On incrémente et vérifie si le rId existe (rId utiles commencent à 2)
            $j = 2
            $rIdTotal = 1
            while($rIdTotal -gt 0) {
                
                $rId = "rId" + $j
                $rIdTotal = ([regex]::Matches($slideContent, $rId )).count
                
                if($rIdTotal -gt 0) {
                    $rIds += @{"Total" = $rIdTotal}
                }
                $j++
            }

            #Va chercher le ratio pour chaque rId (pour analyse future)
            #[xml]$slideContent = $reader.ReadToEnd()
            #TODO: À completer

            $reader.Close()
            $slide.Close()

            #Va chercher la bonne référence dans le fichier xml.rels 
            $relsPath = "ppt/slides/_rels/slide" + $i + ".xml.rels"
            $entry = $zipArchive.GetEntry($relsPath)

            if ($entry) {
                $rels = $entry.Open()
                $reader = New-Object IO.StreamReader($rels)
                [xml]$relsContent = $reader.ReadToEnd()

                #Va chercher, pour chaque rId, le nom de l'image associée, puis met à jour les informations (ou ajoute l'entrée si non-existant)
                for($j=2;$j -lt $rIds.Length;$j++) {
                    $rId = "rId" + $j
                    $image = $relsContent.Relationships.Relationship `
                    | Where-Object {($_.Id -eq $rId) -and ($_.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")} `
                    | Foreach-Object {$_.Target.Substring(9)}
                    if (($arrayImages.Count -gt 0) -and ($arrayImages.Values.Contains($image))) {
                        $indexImage = [math]::floor($arrayImages.Values.indexof($image)/2)
                        $arrayImages[$indexImage].Total = $rIds[$j].Total
                    }
                    else {
                        $arrayImages += @{"Total"= $rIds[$j].Total; "Name" = $image}
                    }
                }

                $reader.Close()
                $rels.Close()
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

#Choose File
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