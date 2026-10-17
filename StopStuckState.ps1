# 1. Define the path to the Excel file you are trying to open
$FilePath = "https://reedelsevier.sharepoint.com/sites/OG-P2PHdesk/Lists/RequestList/Attachments/530/BUYCRASH - 2.xlsx?web=1" 

# 2. Create the Excel COM Object
# Ensure the Excel application itself is not running in the background from a prior failed attempt.
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false # Keep Excel hidden

# 3. Suppress alerts to prevent the "File In Use" dialog from popping up
$Excel.DisplayAlerts = $false

# 4. Attempt to open the Workbook
# The Open method will try to open the file read-only if it's in use, 
# but it won't prompt the user because DisplayAlerts is $false.
try {
    $Workbook = $Excel.Workbooks.Open($FilePath)
    
    # *** Your core script logic goes here ***
    Write-Host "File opened successfully (possibly read-only)."

} catch {
    Write-Error "Failed to open or process the Excel file: $($_.Exception.Message)"
} finally {
    # 5. Clean up the COM objects
    if ($Workbook) { $Workbook.Close($false) } # $false means don't save changes
    if ($Excel) {
        $Excel.Quit() 
        # Crucial step: Release the COM object from memory
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        Remove-Variable Excel
    }
}