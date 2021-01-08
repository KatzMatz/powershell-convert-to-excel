# Define Class

class ConverterCsvToXlsx {

    [void] convertCsvToXlsx([String] $csvPath, [String] $xlsxPath) {
        if (!(Tets-Path $csvPath)) {
            return
        }

        
        try {
            # Import the csv file
            $data = Import-Csv $csvPath -Encoding UTF8
            
            # Open excel
            $excel = New-Object -ComObject Excel.Application # Component Object model
            
            # Generate or Open a book
            if (Test-Path $xlsxPath) {
                # Open
                $book = $excel.Workbooks.Open($xlsxPath)
                
                $preSheet = $book.Sheets($book.Sheets.Count)
                $book.Worksheets.Add([System.Reflection.Missing]::Value, $preSheet)
                $sheet = $book.Sheets($book.Sheets.Count)

                $preSheet = $null
            } else {
                # Generate
                $book = $excel.Workbooks.Add()
                $sheet = $book.Sheets(1)
            }

            # Set sheet name
            $csvName = Split-Path $csvPath -Leaf
            $sheetName = $csvName.Replace(".csv", "")
            $sheet.Name = $sheetName

            # The cell position variables
            $numRow = 1
            $numColumn = 1

            if ($data.Count -ne 0) {
                # get header properties
                $props = $data[0].psoobject.properties.name
                
                # set cell type as string and line style
                $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($data.Count, $props.Length)).NumberFormat = '@'
                $sheet.Range($sheet.Cells(1, 1), $sheet.Celss($data.Count, $props.Length)).Borders.LineSyle = 1 # coutinuous
                $sheet.Range($sheet.Cells(1, 1), $sheet.Celss($data.Count, $props.Length)).Borders.Weight = 2

                # Output header row
                foreach ($propItem in $props) {
                    # set value
                    $sheet.Cells.Item($numRow, $numColumn) = [String]$propItem
                    #set line style
                    
                    # change position to right
                    $numColumn++
                }

                # change position to bottom
                $numRow++

                # Output contents
                foreach ($item in $data) {
                    $numColumn = 1

                    foreach ($propItem in $props) {
                        $sheet.Cells.Item($numRow,$numColumn) = [String]$propItem
                        #set line style
                        # $sheet.Cells.Item($numRow, $numColumn).Borders.LineStyle = 1 # coutinuous
                        # $sheet.Cells.Item($numRow, $numColumn).Borders.Weight = 2
                        # change position to right
                        $numColumn++
                    }

                    $numRow++
                }
            }

            if (Test-Path $xlsxPath) {
                $book.Save()
            } else {
                $book.SaveAs($xlsxPath)
            }
            
            $excel.Quit()
        } finally {
            # GC
            $sheet = $null
            $book = $null
            $excel = $null
            [GC]::Collect()
        }

    }

    [String[]] getCsvPathList([String] $dirPath) {
        [String[]] $csvPathList = @()

        $childItems = Get-ChildItem $dirPath

        foreach ($item in $childItems) {
            if ($item.Name.Contains(".csv")) {
                $csvPathList += $item.Name
            }
        }

        return $csvPathList
    }

}

