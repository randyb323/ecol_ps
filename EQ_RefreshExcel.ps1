
function RefreshExcel
{

param([string[]]$files , [string]$path)
 
    $excelObj = New-Object -ComObject Excel.Application
    write-host "Created Excel object"

    foreach($f in $files)
    {
        ##$filePath = "\\corpsv01\ecollaboration\Tableau\RootSiteSizeData.xlsx"
        $filePath = $path + $f 
        write-host "Current filePath = " $filePath 
       
        $excelObj.Visible = $False
        $excelObj.DisplayAlerts = $False

        #Open the workbook
        $workBook = $excelObj.Workbooks.Open($filePath)

        write-host "Opened workbook: "  $workBook.Name 

        $workSheet = $workBook.Worksheets.Item(1)
        $workSheet.Select()

        #Refresh all data in this workbook
        write-host "About to refresh EQ Excel files ...." 
        try 
        {
            $workBook.RefreshAll()
        }
        catch [System.Exception] 
        {
            $exception = $_.Exception

            while ($null -ne $exception.InnerException)
            {
                $exception = $exception.InnerException
            }

            # Display the properties of the original exception
            $exception | Format-List * -Force
        }
           
        $workBook.Save()        
        $workbook.Close()

        write-host "Refresh complete"
    }
    
    #Uncomment this line if you want Excel to close on its own
    $excelObj.Quit()
}


$myFiles = @("ActionItems.xlsx", "HighRiskActions.xlsx")

$myPath = "\\corpsv01\ecollaboration\Tableau\EQ\"

RefreshExcel $myFiles $myPath