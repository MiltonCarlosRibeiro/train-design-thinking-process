# Revincula os vídeos das apresentações à pasta "videos" do computador atual.
#
# O PowerPoint só toca vídeo vinculado por caminho completo (C:\...\videos\M1.mp4). Ao copiar a pasta do
# treinamento para outro computador, esse caminho muda. Este script reescreve, dentro de cada .pptx da pasta
# do treinamento, os vínculos que apontam para ...\videos\<arquivo> para <esta pasta>\videos\<arquivo>.
# Não precisa de PowerPoint nem de Python. Feche as apresentações antes de rodar.
param([string]$Pasta = (Split-Path -Parent $PSScriptRoot))

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Pasta = (Resolve-Path -LiteralPath $Pasta).Path.TrimEnd("\")
$videos = Join-Path $Pasta "videos"
if (-not (Test-Path -LiteralPath $videos)) {
    Write-Host "ERRO: pasta de vídeos não encontrada: $videos" -ForegroundColor Red
    exit 1
}
# Formato usado pelo PowerPoint: file:///C:\pasta%20com%20espaco\videos\arquivo.mp4
$base = "file:///" + ($videos -replace " ", "%20") + "\"
$padrao = 'Target="file:///[^"]*?\\videos\\([^"]+)"'
$falhas = 0

$apresentacoes = Get-ChildItem -LiteralPath $Pasta -Filter "*.pptx" | Where-Object { $_.Name -notlike '~$*' }
foreach ($arq in $apresentacoes) {
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($arq.FullName, "Update")
    } catch {
        Write-Host "  $($arq.Name): não foi possível abrir — feche o arquivo no PowerPoint e rode de novo." -ForegroundColor Red
        $falhas++
        continue
    }
    $trocas = 0
    $usados = New-Object System.Collections.Generic.HashSet[string]
    try {
        $entradas = @($zip.Entries | Where-Object { $_.FullName -like "ppt/slides/_rels/*.rels" })
        foreach ($e in $entradas) {
            $leitor = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
            $xml = $leitor.ReadToEnd()
            $leitor.Close()
            $ms = [regex]::Matches($xml, $padrao)
            if ($ms.Count -eq 0) { continue }
            foreach ($m in $ms) { [void]$usados.Add($m.Groups[1].Value) }
            $novo = [regex]::Replace($xml, $padrao, { param($m) 'Target="' + $base + $m.Groups[1].Value + '"' })
            if ($novo -ne $xml) {
                $nome = $e.FullName
                $e.Delete()
                $nova = $zip.CreateEntry($nome, [System.IO.Compression.CompressionLevel]::Optimal)
                $escritor = New-Object System.IO.StreamWriter($nova.Open(), (New-Object System.Text.UTF8Encoding($false)))
                $escritor.Write($novo)
                $escritor.Close()
                $trocas += $ms.Count
            }
        }
    } finally {
        $zip.Dispose()
    }
    if ($usados.Count -eq 0) { continue }
    if ($trocas) { Write-Host "`n$($arq.Name): $trocas vínculo(s) atualizado(s)" }
    else { Write-Host "`n$($arq.Name): os vínculos já apontavam para esta pasta" }
    foreach ($v in $usados) {
        $caminho = Join-Path $videos ([uri]::UnescapeDataString($v))
        if (Test-Path -LiteralPath $caminho) {
            Write-Host "  ok       videos\$([uri]::UnescapeDataString($v))" -ForegroundColor Green
        } else {
            Write-Host "  FALTANDO videos\$([uri]::UnescapeDataString($v))" -ForegroundColor Yellow
            $falhas++
        }
    }
}
Write-Host ""
if ($falhas) { Write-Host "Concluído com $falhas aviso(s)." -ForegroundColor Yellow; exit 1 }
Write-Host "Pronto! Os vídeos agora apontam para: $videos" -ForegroundColor Green
