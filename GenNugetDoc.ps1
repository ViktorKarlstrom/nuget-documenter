function GenNugetDoc($BaseDirectory) {


    $userInput = Read-Host -Prompt "Press Enter for latest stable, enter '9' for prerelease"
    $getPrerelease = "false"
    
    if($userInput -eq '9'){
        $getPrerelease = "true"
    }
    Write-Host "This might take a minute, go get a coffee..."

    # Create an Excel object
    $ExcelObj = New-Object -comobject Excel.Application
    $ExcelObj.Visible = $true

    # Add workbook
    $ExcelWorkBook = $ExcelObj.Workbooks.Add()
    $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)

    # Rename worksheet
    $ExcelWorkSheet.Name = "$($BaseDirectory -replace '[^a-zA-Z0-9]', '') NuGet"

    # Fill in table head
    $ExcelWorkSheet.Cells.Item(1,1) = 'Package'
    $ExcelWorkSheet.Cells.Item(1,2) = 'Original version'
    $ExcelWorkSheet.Cells.Item(1,3) = 'Latest available version'
    $ExcelWorkSheet.Cells.Item(1,4) = 'Solution version'
    $ExcelWorkSheet.Cells.Item(1,5) = 'Is major upgrade'
    $ExcelWorkSheet.Cells.Item(1,6) = 'Is updated to latest'
    $ExcelWorkSheet.Cells.Item(1,7) = 'Status'
    $ExcelWorkSheet.Cells.Item(1,8) = 'Project'

    # Format columns to text
    $ExcelWorkSheet.Columns.Item(2).NumberFormat = "@"
    $ExcelWorkSheet.Columns.Item(4).NumberFormat = "@"

    # Make the table head bold, set the font size and the column width
    $ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
    $ExcelWorkSheet.Rows.Item(1).Font.size = 12

    $ExcelWorkSheet.Columns.Format

    # Recursively get all the packages.config. Exclude config in folder "PackageTmp".
    $PACKAGECONFIGS = Get-ChildItem -Recurse -Force $BaseDirectory -ErrorAction SilentlyContinue |
        Where-Object { (( $_.Name -eq "packages.config")) -and $(Split-Path (Split-Path $_.FullName -Parent) -Leaf) -notlike 'PackageTmp'} 

    $counter = 2

    ForEach($PACKAGECONFIG in $PACKAGECONFIGS)
    {

        $path = $PACKAGECONFIG.FullName
        [xml]$packages = Get-Content $path

        ForEach($package in $packages.packages.package)
        {
            $response = Invoke-RestMethod "https://azuresearch-usnc.nuget.org/query?q=$($package.id)&prerelease=$($getPrerelease)" -contenttype 'application/json'
            $project = Split-Path (Split-Path $path -Parent) -Leaf

            $ExcelWorkSheet.Columns.Item(1).Rows.Item($counter) = $package.id
            $ExcelWorkSheet.Columns.Item(2).Rows.Item($counter) = $package.version
            $ExcelWorkSheet.Columns.Item(4).Rows.Item($counter) = $package.version
            
            if($response.totalHits -gt 0) {
                $ExcelWorkSheet.Columns.Item(3).Rows.Item($counter) = $response.data[0].version
                $ExcelWorkSheet.Cells.Item($counter,5).FormulaLocal = "=(IF(VALUE(LEFT(C$counter;(FIND(`".`";C$counter;1)-1))) > VALUE(LEFT(B$counter;(FIND(`".`";B$counter;1)-1))); `"Yes`"; `"No`"))"
                $ExcelWorkSheet.Cells.Item($counter,6).FormulaLocal = "=IF(G$counter=`"Uninstalled`";`"N/A`";IF(C$counter=D$counter; `"Yes`"; `"No`"))"

            } else {
                $ExcelWorkSheet.Columns.Item(3).Rows.Item($counter) = "Not found"
            }

            $ExcelWorkSheet.Columns.Item(8).Rows.Item($counter) = $project

            $counter++
        }
    }

    # Fit columns to content
    $usedRange = $ExcelWorkSheet.UsedRange	
    $usedRange.EntireColumn.AutoFit() | Out-Null
    
    # Save the report and close Excel:
    $timestamp = (Get-Date -UFormat "%Y-%m-%d_%H-%M-%S").tostring()
    $fileName = "NuGet-$($timestamp).xlsx";
    $location = "C:"
    $savedir = "$($location)\$($fileName)";
    $ExcelWorkBook.SaveAs($savedir)
    # $ExcelWorkBook.close($true)

    Write-Host $($counter-2) "rows written"
    Write-Host "$fileName saved at $location"
    Invoke-Item C:/
}
