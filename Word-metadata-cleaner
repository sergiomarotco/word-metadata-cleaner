#Скрипт ищет все подпапки в указанной папке с глубиной 1, и для каждой папки отдельно
#ищет в цикле все файлы, это сделано для того, что если у вас в папке будут десятки или
#сотни тысяч файлов, powershell просто вылетит, так и не найдя файлы. А при таком подходе
#не будет переполнения размера массива и вылетов.
#Скрипт не обрабатывает ошибки открытия файлов доступных только на чтение, запароленных,
#т.е. всех тех при открытии которых выскакивает какое-либо окошко
$StartPath = "c:\"
$pathAll = Get-ChildItem -Name -Path $StartPath -Exclude *.*
$Errors = 0
$iterator = 1
$countAll = $StartPath.length
foreach($path in $pathAll) 
{    
    $path=$StartPath+$path    
    Write-Host "Составляю список файлов *.doc, *.docx в папке $path"
    Add-Type -AssemblyName Microsoft.Office.Interop.Word
    $WdRemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type]
    $wordFiles       = Get-ChildItem -Path $path -include *.doc, *.docx -recurse
    $objword         = New-Object -ComObject word.application
    $i = 1
    $wordFileslength = $wordFiles.length
    $count = $wordFiles.length
    "Найдено Word файлов: $wordFileslength"
    $objword.visible = $false
    foreach($obj in $wordFiles) 
    { 
        $progressBar = [int]($i*100/$count)
        try{
            $documents = $objword.Documents.Open($obj.fullname)
            “$i / $count ($progressBar%) Удаляю метаданные из $obj”
            $i = $i + 1
            $documents.RemoveDocumentInformation($WdRemoveDocType::wdRDIAll)
            $documents.Save()
            $objword.documents.close()            
        }
        catch{
            Write-Host "Ошибка очистки метаданных в $obj"
            $Errors = $Errors + 1
        }
    } 
    $objword.Quit()
    $progressBarAll = [int]($iterator*100/$countAll)
    Write-Host "Удаление метаданных в $path завершено в папках: $iterator / $countAll ($progressBarAll%)"
    $iterator = $iterator +1
}
Write-Host "Удаление метаданных завершено! Ошибок: $Errors"
