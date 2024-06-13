
    param (
        [Parameter(Mandatory=$true)]
        [string]$csvPath,    # Chemin vers le fichier CSV
	[string]$excelPath   # Chemin vers le fichier xlsx 
    )
    

    # Créer une instance d'Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true  # Rendre Excel visible

    try {

	try {
		Remove-Item -Path $excelPath
	}
    	catch {
        	Write-Host "Erreur : $_"
    	}
    	finally {
	}

        # Ouvrir le fichier CSV
        $workbook = $excel.Workbooks.Open($csvPath)
        $worksheet = $workbook.Sheets.Item(1)
        $usedRange = $worksheet.UsedRange  # Obtenez la plage utilisée de la feuille

        # Appliquer TextToColumns pour définir le délimiteur correct
        $usedRange.TextToColumns($usedRange, [Microsoft.Office.Interop.Excel.XlTextParsingType]::xlDelimited, [Microsoft.Office.Interop.Excel.XlTextQualifier]::xlTextQualifierDoubleQuote, $false, $false, $false, $true, $false, $false, $null)

        # Enregistrer le fichier nettoyé en tant que nouveau fichier Excel
        $workbook.SaveAs($excelPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
        # $workbook.Close($true)
    }
    catch {
        Write-Host "Erreur : $_"
    }
    finally {
        # Fermer Excel
        # $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    Write-Host "Fichier CSV ouvert, nettoyé et enregistré en tant que fichier Excel."



