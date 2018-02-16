<#

.SYNOPSIS
Text

.DESCRIPTION
Text

.NOTES
Text

#>

function FindUsedImages {
    param([string]$filename)

    $hashImages = @{}

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
                    $rIds += $rIdTotal
                }
                $j++
            }

            #[xml]$slideContent = $reader.ReadToEnd()

            $reader.Close()
            $slide.Close()

            #Va chercher la bonne référence dans le fichier xml.rels 
            $relsPath = "ppt/slides/_rels/slide" + $i + ".xml.rels"
            $entry = $zipArchive.GetEntry($relsPath)

            if ($entry) {
                $rels = $entry.Open()
                $reader = New-Object IO.StreamReader($rels)
                [xml]$relsContent = $reader.ReadToEnd()

                for($j=2;$j -lt $rIds.Length;$j++) {
                    $rId = "rId" + $j
                    $image = $relsContent.Relationships.Relationship `
                    | Where-Object {($_.Id -eq $rId) -and ($_.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")} `
                    | Foreach-Object {$_.Target.Substring(9)}
                    if ($hashImages.ContainsKey($image)) {
                        $hashImages.$image += $rIds[$j]
                    }
                    else {
                        $hashImages.$image = $rIds[$j]
                    }
                }

                $reader.Close()
                $rels.Close()
            }

            else {
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
    return $hashImages
}

function EvalImages {
    param([string]$filename, [hashtable]$hashImages)

    $warningTable = @()

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($filename)

    $hashImages.GetEnumerator() | sort -property Value | ForEach-Object {
        $imgPath = "ppt/media/" + $_.Key
        $entry = $zipArchive.GetEntry($imgPath)

        #($entry.length / 1MB).toString("0.00MB") - Uncompressed file size

        # TODO: Generate warnings, below is an example

        if (($entry.length / 1MB) -gt 1) {
            $warningTable += @{"FileType" = "Image";"FileSize"=$entry.length;"Message"="Cette image à un poid supérieur à 1MB"}
        }
    }

    $zipArchive.Dispose()
    return $warningTable
}

#Choose File
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object Windows.Forms.OpenFileDialog
$openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
$openFileDialog.filter = "Powerpoint Presentations (*.pptx)|*.pptx"
$result = $openFileDialog.ShowDialog()

if (($result -eq "OK") -and $openFileDialog.CheckFileExists) {

    $images = FindUsedImages -filename $openFileDialog.FileName
    $warnings = EvalImages -filename $openFileDialog.FileName -hashImages $images

    $warnings
}