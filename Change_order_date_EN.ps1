<#  
    Writes $str_to_write to any combination of console and log files using 2^n logic.
    i.e. $type = 3 (1+2) will write into Console and Normal Log,
    $type = 6 (2+4) will write to normal and error log, but not to console.
#>
Function Write-Log_Host($str_to_write, $type) {
    $dtstring ="[" +  (Get-Date -Format "dd.MM.yyyy HH:mm:ss") + "] " + $str_to_write
    $loggingTypes = @([pscustomobject]@{ID=1;Type="CONSOLE";Dest = "HOST"},
                      [pscustomobject]@{ID=2;Type="NORMAL LOG";Dest = $sourcefolder+"\log.txt"},
                      [pscustomobject]@{ID=4;Type="ERROR LOG"; Dest = $sourcefolder+"\failed.txt"})
    $typeRem = $type
    $i = $loggingTypes.Count -1
    while ($typeRem -gt 0) {
        while (($typeRem - [int32]$loggingTypes[$i].ID) -lt 0){$i--}
        if ($typeRem -eq 1) {Write-Host $dtstring}
        else {Add-Content $loggingTypes[$i].Dest $dtstring}
        $typeRem -= [int32]$loggingTypes[$i].ID
    }
}
<#
    Returns folder, selected by user or quits the app if cancelled.
#>
Function Get-Folder(){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Specify path to documents"
    $foldername.rootfolder = "MyComputer"
    if ($foldername.ShowDialog() -eq "OK"){return $foldername.SelectedPath}
    else{Read-Host -Prompt 'Folder selection cancelled. Press "Enter" to exit the program';Exit}
}
$sourcefolder = Get-Folder
$failedFilesCount = 0
<#
    Case 1: Script will delete the print date from the order in top-right corner;
    Case 2: Script will replace the print date with the order issue date, taken from order text.
#>
$promptstring = 'Erase the print date from the document (1) or replace with the order date (2)?'
do {$replMode = Read-Host -Prompt $promptstring} while (!(($replMode -eq 1) -or ($replMode -eq 2)))
<#    
    Ask user for a print date.
#>
$promptstring = 'Specify the print date, formatted as DD.MM.YYYY:'
do {$orderexportdate = Read-Host -Prompt $promptstring} while (!($orderexportdate -match '\d\d\.\d\d\.\d\d\d\d'))
$orderexportdate += ' года'
<#
    Get all documents in folder (non-recursively), count them and log the total number.
#>
$list = Get-ChildItem -Path $sourcefolder -Filter *.doc*
$totalDocs = $list.Count
Write-Log_Host ("Documents found: " + $totalDocs) 3
$counter = 0
<#
    Open invisible Word app. 
    Later the script will open each document in this app instance.
#>
$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $false
foreach($order in $list){
    # Open the document
    $objDoc = $objWord.Documents.Open($order.FullName,$true)
    $objSelection = $objWord.Selection
    <#
        If user selected the option to remove the date (1), it will be replaced with empty string.
        Otherwise: Find the order date from text using Regex
                   If order date not found - log the error and move on.
    #>
    if ($replMode -eq 1) {$dateToReplace = ""} elseif ($replMode -eq 2)
    {
        if ($objDoc.Content.Text -match "Заказ № \d{8} от (\d\d\.\d\d\.\d\d\d\d) г\.") {$dateToReplace = $matches[1]} else 
        {
            Write-Log_Host ("Order date not found in file " + $order.Name) 7
            $failedFilesCount++
            $objDoc.Close()
            $counter++
            Continue
        }
    }
    <#
    Find.Execute($ReplaceWhat,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,
        $MatchAllWordForms,$Forward,$Wrap,$Format,$replaceWith,$wdReplaceOne)) 
    wdReplaceAll	2	Replace all occurrences.
    wdReplaceNone	0	Replace no occurrences.
    wdReplaceOne	1	Replace the first occurrence encountered.
    #>

    <#
        Replace the print date (specified by user) with empty string or odred date.
        If the print date not found - log the error and move on.
    #>
    if ($objSelection.Find.Execute($orderexportdate,$False,$True,$False,$False,$False,$True,1,$False,$dateToReplace,1))
    {
        if ($replMode -eq 1) {$logstring = "Date "+$orderexportdate+" in file "+$order.Name+" successfully deleted"}
        else {$logstring = "Date "+$orderexportdate+" in file "+$order.Name+" replaced with date "+$dateToReplace}
        Write-Log_Host $logstring 3
    } else 
    {
        Write-Log_Host ("Couldn't find print date "+$orderexportdate+" in file "+$order.Name) 7
        $failedFilesCount++
    }
    # Save and close the document.
    $objDoc.Save()
    $objDoc.Close()
    # Calculate values for progress bar.
    $counter++
    $percent = [math]::floor([int]$counter*100/$totalDocs)
    $statusmsg = 'Documents processed - ' + $counter + ' out of ' + $totalDocs
    Write-Progress -Id 1 -Activity 'Date replace' -Status $statusmsg -PercentComplete $percent
}
# Close the Word App instance
$objWord.Quit()
# Hide progress bar
Write-Progress -Id 1 -Activity "Date replace" -Completed
# Log and show the final result - how many files were changed and how many errors there were.
$logstring = "Processed documents - "+$counter+" out of "+$totalDocs+". Information saved in file 'log.txt'"
if ($failedFilesCount -gt 0) {$logstring +="`r`n"+"Processing errors - "+$failedFilesCount+". Information about errors saved in file 'failed.txt'"}
Write-Log_Host $logstring 3
Read-Host -Prompt 'Press "Enter" to exit the program'



