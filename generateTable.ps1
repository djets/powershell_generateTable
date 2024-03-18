#K.V.Yankovich 2024
#А script for generating table in a docx document based on text file raw data

Add-Type -AssemblyName System.Windows.Forms
function Select-File ([string]$title, [string]$filter)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $title
    $OpenFileDialog.InitialDirectory = ".\"
    $OpenFileDialog.filter = $filter
    $openFileDialog.ShowHelp = $true
If ($OpenFileDialog.ShowDialog() -eq "Cancel")
{
    [System.Windows.Forms.MessageBox]::Show("No File Selected. Please select a file !", "Error", 0, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    Return null
}
    $Global:SelectedFile = $OpenFileDialog.FileName
    Return $SelectedFile #add this return
}

function Get-HashtableEmployees {
    param (
        [Parameter (Mandatory = $true)]
            $arrString,
        [Parameter (Mandatory = $true)]
            $pattern
    )
    $hashtableEmployees = @{}
    $rawFields = $pattern | Select-String "(?<=<)\w+(?<!>)" -AllMatches
    [string[]]$fields = $rawFields.Matches.Value
    $arrString | Select-String -Pattern $pattern | ForEach-Object {
        $hashtableFields = @{}
        for ($i = 0; $i -le $fields.Length; $i++) {
            $n = $i + 1
            [string]$stringValue = $($_.matches.groups[$n])
            $stringValue = $stringValue.TrimStart("")
            $stringValue = $stringValue.TrimEnd("")
            $stringValue = $stringValue -replace '\s{2,}', ' '
            if ($stringValue) {
                $hashtableFields.add($fields[$i], $stringValue)
            }
        }
        $hashtableEmployees.add([guid]::NewGuid().ToString(), $hashtableFields)
    }
    return $hashtableEmployees 
}

$todocConf = Select-File -title "Выбирите файл конфигурации" -filter "Configuration file (*.conf) | *.conf"
#Adding variables from the configuration file
foreach ($i in $(Get-Content -Path $todocConf)){
    Set-Variable -Name $i.split("=")[0] -Value $i.split("=",2)[1]
}
[System.Collections.Generic.List[string]]$filterList = $filter_list.Split(",")
$dataFile = Select-File -title "Выбирите файл с исходными данными" -filter "Txt file (*.txt) | *.txt" 
$templateFile = Select-File -title "Выбирите файл с шаблоном" -filter "Docx file (*.docx) | *.docx" 
[string[]]$arrStr = Get-Content $dataFile
$runPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition


if($arrStr.Count -ge 0) {
    $hashsetEmployees = Get-HashtableEmployees $arrStr $pattern_search
    
    Write-Host "Выбраных записей -" $hashSetEmployees.Count

    #filter by object fields
    if($filter_enable -eq "true") { 
        $removeList = [System.Collections.Generic.List[string]]::new()
        if ($filter_invers -eq "true") {
            foreach ($itemKey in $hashsetEmployees.Keys) {
                if ($filterList -contains $hashsetEmployees.$itemKey.$filter_field) {
                    $removeList.Add($itemKey)
                }
            }
        } elseif ($filter_invers -eq "false") {
            foreach ($itemKey in $hashsetEmployees.Keys) {
                if ($filterList -notcontains $hashsetEmployees.$itemKey.$filter_field) {
                    $removeList.Add($itemKey)
                }
            }
        }       
        if ($removeList.Count -gt 0) {
            $removeList | ForEach-Object { $hashsetEmployees.Remove($_) }
        }
    }

    #sorting by a specific field
    $sortedSetEmployees = [System.Collections.SortedList]::new()
    foreach ($key in $hashSetEmployees.Keys) {
        try {
            $sorting_filed
            $sortedSetEmployees.add($hashSetEmployees.$key.$sorting_field, $key)
        }
        catch {
            [string]$str = $hashSetEmployees.$key.$sorting_filed
            Write-Warning  "$str - дублирующая запись"
            # $PSItem.Exception
            }
    }
    Write-Host "Отфильтрованных, отсортированных записей -" $sortedSetEmployees.Count

    $word = New-Object –comobject Word.Application
    $word.visible = $false
    $document = $word.documents.open($templateFile, $false, $false)
    $table = $document.Tables.item([int]$table_number)

    $count = 1
    [Int32]$xRow = 2
    foreach ($employeeKey in $sortedSetEmployees.Keys) {
        $hashKeyFio = $sortedSetEmployees.$employeeKey
        $hashtableRecord = $hashsetEmployees.$hashKeyFio

        $employeeFio = $hashtableRecord.fio
        if ($employees_initials -eq "true") {
            $arrFio = $employeeFio.Split(" ")
            $lastName = $arrFio[0]
            $firstName = ($arrFio[1]).ToCharArray()
            $midleName = ($arrFio[2]).ToCharArray()
            $employeeFio = $lastName + " " + $firstName[0] + "." + $midleName[0] + "."
        }
        if ($employees_position -eq "true") {
            Write-Host $employeeKey - $hashtableRecord.position
            $table.Cell($xRow, 1).Range.Text = [string]($xRow - 1).ToString()
            $table.Cell($xRow, 2).Range.Text = $employeeFio
            $table.Cell($xRow, 3).Range.Text = $hashtableRecord.position
            $xRow++
            $newRow = $table.Rows.Add()
            
        } else {
            Write-Host $employeeKey
            $table.Cell($xRow, 1).Range.Text = [string]($xRow - 1).ToString()
            $table.Cell($xRow, 2).Range.Text = $employeeFio
            $xRow++
            $newRow = $table.Rows.Add()
        }
    }    
    
    $count
    $date = Get-Date -Format("yyyyMMdd")
    $saltNameFile = Split-Path $templateFile -leaf
    [string]$fileName = $date + "-" + $saltNameFile
    # Save as and close the document
    # Save as and close the document (directory of the PowerShell script)
    $document.SaveAs([ref]"$runPath\$fileName")
    $document.Close(-1)
    $word.Quit()
    
    # Освобождение ресурсов
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Remove-Variable word
}




    
