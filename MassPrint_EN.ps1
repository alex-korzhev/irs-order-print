<#--------- E X P L A N A T I O N ----
    Since -Verb print does not specify the printer to use, and "printto" throws an error
    in Word documents, the script will switch the default printer every SD times, which is 5 by default.
    Script uses Np numbers and counts printed documents for each printer 
    to specify the next document. 
    Since printing usually takes more time than putting a file in a print queue it will result in a 
    simultaneous printing. After collecting the documents they will be ordered by filename.
    
    Suppose we printing 53 documents on 3 printers.
    Np numbers will be 18 18 17. Printing by 5 and switching the printer will result in
    printing order like this:
    p[1] prints docs 1,2,3,4,5.
    p[2] prints docs 19,20,21,22,23
    p[3] prints docs 37,38,39,40,41
    p[1] prints docs 6,7,8,9,10
    and so on.
    When we collect the documents from printers 1, 2 and 3 they will be ordered by name.

    Algorithm:

    1) Get all documents (Nmax) in selected folder and put them into a list
    2) Ask user which printers to use -> p
    3) Ask user how many documents to print (1 - Nmax) -> N
    4) Divide N document between p printers (Np numbers) using formula:
        Np(i) = Ceiling(N/p), if i < %(N/p), Np(i) = Floor(N/p), if i >= %(N/p)
    5) Block documents in csv file by setting print status to 2 (in progress).
    In a loop:
    6) Set default printer 
    7) Get the next document for printing using NP numbers and counters for each printer.
    8) Save progress into csv file each time the printer loop goes full circle.
    After loop:
    9) Save progress and ask again how many documents to print.
#>
<#--------- G L O B A L S ----------#>
<#
    Get the folder from user, set the name of csv-file, set some global variables.
#>
$sourcefolder = Get-Folder
$gl_csv = $sourcefolder+"\FILES_PRINTED.CSV"
$excludelist = @("FILES_PRINTED.CSV","log.txt","failed.txt")
$failedFilesCount = 0
$printcount = -1
$defaultdelay = 200
$CR_LF = "`r`n"

<#----------PRINT_STATUS---------------#>

# -1 = ERROR
#  0 = NOT PRINTED
#  1 = PRINTED
#  2 = IN PROGRESS

<#--------- F U N C T I O N S ---------#>

<#
    Saves csv-file over the series of $triesMax tries,
    waiting $delay milliseconds between each try.
    Returns $true, if csv is saved, $false if out of tries.
#>
Function TryWrite-CSV ($documents, $delay, $triesMax) {
    $tries = 0
    $isWritten = $false
    While (!$isWritten) {
        if ($tries -eq $triesMax) {return $false}
        try {$documents | Export-Csv -Delimiter ";" -Path $gl_csv -NoTypeInformation;$isWritten = $true}
        catch [system.exception] {Start-Sleep -Milliseconds $delay}
        finally{$tries++}
    }
    return $true
}  

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
    Almost evenly distributes $N documents between $p printers.
    First printers get roundup($N/p), last get rounddown ($N/p).
    Also works if $N is completely divisible by $p.
#>
Function Get-NP_numbers($N, $p) {
    $np = New-Object System.Collections.Generic.List[System.Object]
    for ($i = 0; $i -lt $p; $i++){
        if ($i -lt $N%$p){$a = [Math]::Ceiling($N/$p)} else {$a = [Math]::Floor($N/$p)}
        $np.Add($a)
    }
    return $np
}

<#
    Shows the list of connected printers.
    Asks user to input corresponding numbers divided by whitespace.
    Returns list of selected printers.
#>
Function Select-Printers(){
    $printers = Get-CimInstance Win32_Printer | Where-Object {$_.PrinterStatus -eq 3 -or $_.PrinterStatus -eq 1}
    Write-Host "Connected printers:"
    $i = 1
    foreach ($printer in $printers){
        $printerstring = '[' + $i + '] - ' + $printer.Name
        Write-Host $printerstring
        $i++
    }
    $selected_printers = Read-Host -Prompt "Select printers by typing their id divided by whitespace"
    $sp = $selected_printers.Split(" ")
    $resulting_printers = New-Object System.Collections.Generic.List[System.Object]
    foreach ($p in $sp){$resulting_printers.Add($printers[$p-1])}
    return $resulting_printers 
}

<#
    Reduces the number of selected printers, when $p < $N
    It is a rare case, but will lead to an error,
    when distributing documents between printers.
    Cannot be merged with Select-Printers, because
    printer selection comes before than documents selection
    and also out of the main loop.
#>
Function Filter-Printers($selectedPrinters, $docsToPrint) {
    $resulting_printers = New-Object System.Collections.Generic.List[System.Object]
    $i = 0
    foreach ($p in $selectedPrinters){
        $i++
        $resulting_printers.Add($selectedPrinters[$i-1])
        if ($i -eq $docsToPrint){break}
    }
    return $resulting_printers
}

<#
    Returns folder selected by user or quits the app if cancelled.
#>
Function Get-Folder(){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Specify path to documents"
    $foldername.rootfolder = "MyComputer"
    if ($foldername.ShowDialog() -eq "OK"){return $foldername.SelectedPath}
    else{Read-Host -Prompt 'Folder selection cancelled. Press "Enter" to exit the program';Exit}
}

<#
    Compares the existing csv file to current documents in selected folder.
    Adds new documents and saves the list.
    Does not delete missing documents - if the document is not present
    at the moment of printing - will log an error.
#>
Function Compare-DocumentsAndCsv (){
    $currentFiles = Get-ChildItem -Path $sourcefolder -Exclude $excludelist | where { ! $_.PSIsContainer } | select -ExpandProperty Name 
    $prevFiles = Import-Csv -Path $gl_csv -Delimiter ";"
    $prevFilesNames = $prevFiles | select -ExpandProperty Doc
    $newFiles =  Compare-Object -ReferenceObject $prevFilesNames -DifferenceObject $currentFiles -PassThru
    if ($newFiles) {
        $lastCSVID = ($prevFiles | select -ExpandProperty ID | measure -Maximum).Maximum
        $docList = New-Object System.Collections.Generic.List[System.Object]
        foreach ($file in $newFiles) {
            $lastCSVID++
            $docList.add([pscustomobject]@{ID = $lastCSVID; Doc = $file; PrintStatus = 0})
        }
        $prevFiles += $docList 
        if (!(TryWrite-CSV $prevFiles 500 8)) {
            Write-Log_Host "CSV file is missing or currently inaccessible" 7
            Read-Host -Prompt 'Press "Enter" to exit the program'
            Exit       
        }
    }
    return $prevFiles
}

<#
    Creates new list of documents if csv-file not exists,
    compares list to documents otherwise.
#>
Function Load-DocumentsFromCsv(){
    if (!(test-path $gl_csv)){
        $list = Get-ChildItem -Path $sourcefolder -Exclude $excludelist | where { ! $_.PSIsContainer }
        $documents = New-Object System.Collections.Generic.List[System.Object]
        for ($i = 0; $i -lt $list.Count; $i++) {$documents.add([pscustomobject]@{ID = ($i+1); Doc = $list[$i].Name;PrintStatus = 0})}
        $documents | Export-Csv -Delimiter ";" -Path $gl_csv -NoTypeInformation
    } else {$documents = (Compare-DocumentsAndCsv | Where-Object {$_.PrintStatus -eq 0 -or $_.PrintStatus -eq -1})}
    return $documents
}

<#
    Gets integer from user, representing the amount of documents to print.
#>
Function Ask-DocsToPrint ($documents){
    $isInvalid = $true
    do {
        $promptstring = "How many documents to print?" + "(1-"+[string]$documents.Count+")"
        [int]$docs_to_print = Read-Host -Prompt $promptstring
        if ($docs_to_print -is [int]) {
            if (($docs_to_print -gt 0) -and ($docs_to_print -le $documents.Count)){$isInvalid = $false}
        }
    }
    while ($isInvalid)
    return $docs_to_print
}

<#
    Added for cycling through the list of whatever needed.
#>
Function Select-NextItem ($current, $max) {
    if ($current -eq ($max-1)) {return 0} else {return ($current + 1)}
}

<#
    Opens csv file, changes $PrintStatus of printed files,
    Tries to save csv file - if saving is necessary - will throw an error and quit,
    otherwise just returns $false.
    Returns $true if csv was saved.
#>
Function Save-DocumentsToCsv($newDocuments, $docsToPrintCount, $isErrorAcceptable, $triesMax, $delay){
    
    $oldDocuments = Import-Csv -Path $gl_csv -Delimiter ";"
    for ($i = 0; $i -lt $docsToPrintCount; $i++) {$oldDocuments[$($oldDocuments.ID).indexof([string]$newDocuments[$i].ID)].PrintStatus = $newDocuments[$i].PrintStatus}
    if (!(TryWrite-CSV $oldDocuments $triesMax $delay)) {
        if ($isErrorAcceptable) {return $false}
        else {
            Write-Log_Host "CSV file is missing or currently inaccessible" 7
            Read-Host -Prompt 'Press "Enter" to exit the program'
            Exit
        }
    }
    return $true
}

<# ---------B O D Y ---------#>
# Select printers ($p later)
$selectedPrintersBF = Select-Printers $docsToPrintCount
do  {

    # Get a list of documents
    $docs = Load-DocumentsFromCsv 
    if ($docs.Count -eq 0) {
        Write-Log_Host ("All documents from folder " + $sourcefolder +" are already printed!") 7
        Read-Host -Prompt 'Press "Enter" to exit the program'
        Exit
    }
    Write-Log_Host ("Unprinted documents total - "+$docs.Count) 3
    # Ask user how much to print ($N later)
    $docsToPrintCount = Ask-DocsToPrint($docs) 
    <# 
        Filter printers (drop unnecessary printers if $N < $p
        Filter original selection for each loop,
        because on the next iteration $N may be greater than $p,
        maybe 1st iteration will be used as test by user.
    $selectedPrinters = Filter-Printers $selectedPrintersBF $docsToPrintCount
    <#
        Added to prevent an error, when the entire list of 1 printer
        is considered a Win32_Printer, not an item in a list.
    #>
    if ($selectedPrinters -isnot [Array]) {$printersCount = 1}
    else {$printersCount = $selectedPrinters.Count}
    <# 
        Swith the printer in use after $printerSwitchDocumentDelay documents printed.
        Фdded so as not to change the default printer after each document.
    #>
    if ([Math]::Ceiling($docsToPrintCount/$printersCount) -gt 5) {$printerSwitchDocumentDelay = 5} 
    else {$printerSwitchDocumentDelay = [Math]::Ceiling($docsToPrintCount/$printersCount)}
    <# Distribute documents between printers #>
    $npNumbers = Get-NP_Numbers $docsToPrintCount $printersCount
    $printedByPrinter = new-object int[] $printersCount
    $printerSwitcher = 0
    $SDcounter = 0
    <#
        Add a blocking mechanism in order to spread the task between workstations.
    #>
    for ($i = 0; $i -lt $docsToPrintCount; $i++) {$docs[$i].PrintStatus = 2}
    (Save-DocumentsToCsv $docs $docsToPrintCount $false 1000 10) | Out-Null
    # Log task parameters for later troubleshooting
    $logstring =  ("Task started with parameters:" + $CR_LF)
    $logstring += ("Unprinted documents total = " + $docs.Count + ", selected = " + $docsToPrintCount) + $CR_LF
    $logstring += ("Printers selected = " + $printersCount + $CR_LF)
    $logstring += ("SDnumber = " + $printerSwitchDocumentDelay + $CR_LF + "N =")
    foreach ($n in $npNumbers) {$logstring += (" " + $n)}
    Write-Log_Host $logstring 2
    # Set first selectd printer as default
    Invoke-CimMethod -InputObject $selectedPrinters[0] -MethodName SetDefaultPrinter | Out-Null
    for ($i = 0; $i -lt $docsToPrintCount; $i++){
        <#
            Loop for the last portion of documents, when each printer
            prints less, than $printerSwitchDocumentDelay.
        #>
        $ERR_NP_NUMBER = 0
        While ($printedByPrinter[$printerSwitcher] -eq $npNumbers[$printerSwitcher]){
            if ($ERR_NP_NUMBER -eq $printersCount) {
                    Write-Log_Host "Error in printer count or Np numbers" 6
                    Save-DocumentsToCsv $docs $docsToPrintCount $false 1000 10
                    Exit
                }
            $printerSwitcher = Select-NextItem $printerSwitcher $printersCount
            Invoke-CimMethod -InputObject $selectedPrinters[$printerSwitcher] -MethodName SetDefaultPrinter | Out-Null
            $SDcounter = 0
            $ERR_NP_NUMBER++
        }
        # Switch the printer in use after $printerSwitchDocumentDelay documents printed
        if ($SDcounter -eq $printerSwitchDocumentDelay){
            $SDcounter = 0
            $printerSwitcher = Select-NextItem $printerSwitcher $printersCount
            if ($printerSwitcher -eq 0) {
                if (Save-DocumentsToCsv $docs $docsToPrintCount $true 500 5) {Write-Log_Host "Intermediate save successfull" 2}
                else {Write-Log_Host "Error: Intermediate save unsuccessfull" 6}
            }
            Invoke-CimMethod -InputObject $selectedPrinters[$printerSwitcher] -MethodName SetDefaultPrinter | Out-Null
        }
        # Get the index of document to print - see EXPLANATION
        $printerShift = 0
        if ($printerSwitcher -ne 0) {for ($j = 1; $j -le $printerSwitcher; $j++){$printerShift += $npNumbers[$j-1];}}
        $currentDocument = $printerShift + $printedByPrinter[$printerSwitcher]

        <#
            Try to print a document from list.
            Log an error if one occures, set PrintSTatus to -1 and move on.
        #>
        try {
            Write-Log_Host ("Printing document - " + $docs[$currentDocument].Doc + "  ||  using " + $selectedPrinters[$printerSwitcher].Name) 2
            Start-Process -FilePath ($sourcefolder + "\" + $docs[$currentDocument].Doc) -Verb Print -Wait
            #Start-Sleep -Milliseconds $defaultdelay
            $docs[$currentDocument].PrintStatus = 1
        }
        catch [system.exception] {
            $docs[$currentDocument].PrintStatus = -1
            Write-Log_Host ("Error: file " + $docs[$currentDocument].Doc + " - missing or currently inaccessible.") 6
            $failedFilesCount++
        }
        finally{$printedByPrinter[$printerSwitcher]++;$SDcounter++}

        # Calculate progress bar status and values for each printer
        $statusmsg = 'Documents processed - ' + ($i+1) + ' из ' + $docsToPrintCount + '       '
        for ($wp = 0; $wp -lt $printersCount; $wp++){
            $statusmsg+=' || Printer № ' + ($wp + 1) + ' - ' + ($printedByPrinter[$wp]) + '/' + $npNumbers[$wp]
        }
        $statusmsg+=' ||'
        $percent = [math]::floor([int]($i+1)*100/$docsToPrintCount)
        Write-Progress -id 1 -Activity 'Document printing' -Status $statusmsg -PercentComplete $percent
    }

    # Hide progress bar
    Write-Progress -id 1 -Activity 'Document printing' -Completed

    <#
        Save progress to csv file. 
        Big delay and tries numbers because CSV-file errors not accepted in final saving.
    #>
    if (Save-DocumentsToCsv $docs $docsToPrintCount $false 1500 20) {Write-Log_Host "Final progress successfully saved to csv." 2}

    # Show and log final report.
    Write-Log_Host ("Printing completed. Totally printed - " + ($docsToPrintCount - $failedFilesCount) + " documents") 3
    # Show info about error log if there were any errors.
    if ($failedFilesCount -gt 0) {
        Write-Log_Host ("File processing errors - " + $failedFilesCount + " Information saved in file failed.txt") 3;
        Write-Log_Host ("----------------") 4;
    }
    Write-Log_Host ("----------------") 2
    # Loops the entire program back to selecting number of documents.
    $answ = Read-Host -Prompt 'Type any key to re-run the program or press "Enter" to exit the program' 
} While ($answ -ne "")