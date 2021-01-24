
function Convert-PSCustomObject-To-Excel {

    <#
    .SYNOPSIS
        The function which converts PSCustomObjects to Excel file

    .DESCRIPTION
        This function converts PSCustomObject to Excel file.

    .EXAMPLE
        Convert-PSCustomObject-To-Excel -InputObject $customObject -OutputExcelPath $OutputExcelPath

    .PARAMETER InputObject
        PSCustomObject variable
        This function can receive the list of PSCustomObject.
        
    .PARAMETER OutputExcelPath
        Output file path
    
    #>
    param (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        $InputObject,
        
        [Parameter(Mandatory=$true)]
        [String]
        $OutputExcelPath,
        
        [Switch]
        $Border
    )
   
    Begin {
        $objectList = @()

        if ($InputObject.Count) {
            $objectList = $InputObject
        }        
    }
    
    Process {
        if (!$InputObject.Count) {
            $objectList += $InputObject
        }   
    }

    End {
        $excel = $null
        $book = $null
        $sheet = $null
        if ($objectList.Length -eq 0) {
            Write-Host "object has no length"
        } else {

            try {
                Write-Host "Open Excel ..."
                $excel = New-Object -ComObject Excel.Application
                $excel.DisplayAlerts = $false
             
                if (!(Test-Path $OutputExcelPath)) {
                                
                    Write-Host "Add book ..."
                    $book = $excel.Workbooks.Add()
                    Write-Host "Get Sheet..."
                    $sheet = $excel.WorkSheets.Item(1)
                                    
                } else {
                    
                    Write-Host "Open book..."
                    $book = $excel.Workbooks.Open($OutputExcelPath)
                    $presheet = $book.Sheets($book.Sheets.Count)
                    $book.Worksheets.Add([System.Reflection.Missing]::Value, $preSheet) | Out-Null
                    Write-Host "Add sheet..."
                    $sheet = $book.Sheets($book.Sheets.Count)
                    
                    $preSheet = $null
                }
                
                # Positions
                [int] $rowPosition = 1
                [int] $columnPosition = 1
                $properties = $InputObject[0].PSObject.Properties.Name

                # set cell style
                $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($objectList.Length+1, $properties.Length)).NumberFormat = '@'
                if ($Border -eq $true) {
                    $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($objectList.Length+1, $properties.Length)).Borders.LineStyle = 1 # coutinuous
                    $sheet.Range($sheet.Cells(1, 1), $sheet.Cells($objectList.Length+1, $properties.Length)).Borders.Weight = 2
                }
               
                # Output properties
                foreach ($prop in $properties) {
                    $sheet.Cells.Item($rowPosition, $columnPosition) = [String]$prop
                    $columnPosition++
                }
                
                $rowPosition++
                $columnPosition = 1
                foreach ($item in $objectList) {

                    foreach ($prop in $properties) {
                        $sheet.Cells.Item($rowPosition, $columnPosition) = [String]$item.$prop
                        $columnPosition++
                    }
                    $rowPosition++
                    $columnPosition = 1
                }
                
    
                if (Test-Path $OutputExcelPath) {
                    $book.Save()
                } else {
                    $book.SaveAs($OutputExcelPath)
                }

                Write-Host "Saved..."
            } catch {
                Write-Host $_.Exception
                Write-Host "Exception in end"
            } finally {
                if ($null -ne $excel) {
                    Write-Host "Close Excel"
                    $excel.Quit()
                    $sheet = $null
                    $book = $null
                    $excel = $null
                }
            }
        }

    }  

}