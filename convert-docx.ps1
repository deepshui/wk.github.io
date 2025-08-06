try {
    $word = New-Object -ComObject Word.Application
    $docPath = Join-Path -Path $PSScriptRoot -ChildPath '文字文稿1.docx'
    $doc = $word.Documents.Open($docPath)
    $content = $doc.Content.Text
    $outputPath = Join-Path -Path $PSScriptRoot -ChildPath 'content\posts\doc-content.txt'
    $content | Out-File -FilePath $outputPath -Encoding utf8
    $doc.Close()
    $word.Quit()
    Write-Output 'Document converted successfully'
} catch {
    Write-Error "Error processing document: $_"
}