# Define source and target paths
$phoneName = 'Apple iPhone'
$sourceFolders = @()
for ($year = 2006; $year -le 2024; $year++) {
    for ($month = 1; $month -le 12; $month++) {
        $yearMonth = "{0:D4}{1:D2}__" -f $year, $month
        $sourceFolder = "Internal Storage\$yearMonth"
        $sourceFolders += $sourceFolder
    }
}
$targetRootFolder = 'G:\iphone' # Root folder where folders with the same name will be created
$filter = '(.jpg)|(.mp4)|(.mov)|(.png)$'

function Get-ShellProxy {
    if(-not $global:ShellProxy) {
        $global:ShellProxy = New-Object -ComObject Shell.Application
    }
    $global:ShellProxy
}

function Get-Phone {
    param($phoneName)
    $shell = Get-ShellProxy
    $shellItem = $shell.NameSpace(17).Self
    $phone = $shellItem.GetFolder.Items() | Where-Object { $_.Name -eq $phoneName }
    return $phone
}

function Get-SubFolder {
    param($parent,[string]$path)
    $pathParts = @($path.Split([System.IO.Path]::DirectorySeparatorChar))
    $current = $parent

    foreach ($pathPart in $pathParts) {
        if ($pathPart) {
            $current = $current.GetFolder.Items() | Where-Object Name -eq $pathPart
        }
    }
    return $current
}

foreach ($sourceFolder in $sourceFolders) {
    $phoneFolderPath = $sourceFolder

    $phone = Get-Phone -phoneName $phoneName
    $folder = Get-SubFolder -parent $phone -path $phoneFolderPath

    $items = @( $folder.GetFolder.Items() | Where-Object { $_.Name -match $filter } )
    if ($items) {
        $destinationFolderName = Split-Path -Path $phoneFolderPath -Leaf
        $destinationFolderPath = Join-Path -Path $targetRootFolder -ChildPath $destinationFolderName

        if (-not (Test-Path $destinationFolderPath)) {
            $created = New-Item -ItemType Directory -Path $destinationFolderPath
        }

        $totalItems = $items.count
        if ($totalItems -gt 0) {
            Write-Verbose "Processing Path : $phoneName\$phoneFolderPath"
            Write-Verbose "Moving to : $destinationFolderPath"

            $shell = Get-ShellProxy
            $destinationFolder = $shell.Namespace($destinationFolderPath).Self
            $count = 0;

            foreach ($item in $items) {
                $fileName = $item.Name
            
                ++$count
                $percent = [int](($count * 100) / $totalItems)
                Write-Progress -Activity "Processing Files in $phoneName\$phoneFolderPath" `
                    -Status "Processing File ${count} / ${totalItems} (${percent}%)" `
                    -CurrentOperation $fileName `
                    -PercentComplete $percent
            
                # Check if the target file doesn't exist:
                $targetFilePath = Join-Path -Path $destinationFolderPath -ChildPath $fileName
                if (Test-Path -Path $targetFilePath) {
                    Write-Output "File already exists in the destination: $fileName"
                }
                else {
                    $destinationFolder.GetFolder.CopyHere($item)
                    if (Test-Path -Path $targetFilePath) {
                        Write-Output "File copied successfully to destination: $fileName"
                        # Optionally do something with the file after copying
                    }
                    else {
                        Write-Error "Failed to move file to destination:`n`t$targetFilePath"
                    }
                }
            }
            
        }
    }
}
Write-Output "Process Complete"
