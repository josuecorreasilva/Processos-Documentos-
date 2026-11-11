$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$dest = "C:\Users\Julia\Documents\painel-compras-empenhos\js\seed.js"
$pt = [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR")
$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Parse-Brl([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  $s = $text -replace "R\$", "" -replace "\s", ""
  if ($s -match "^-?\d{1,3}(\.\d{3})+,\d+$" -or $s -match "^-?\d+,\d+$") {
    $s = $s -replace "\.", "" -replace ",", "."
  }
  $n = 0.0
  $ok = [double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $inv, [ref]$n)
  if ($ok) { return [math]::Round($n, 4) }
  return $null
}

function Read-Cell($ws, $r, $c) {
  $cell = $ws.Cells.Item($r, $c)
  $t = ([string]$cell.Text)
  if ($null -eq $t) { $t = "" }
  $t = $t.Trim()
  $num = $null
  $date = $null
  if ($t -match '^#+$' -or $t -match 'R\$') {
    $v = $cell.Value2
    if ($null -ne $v -and $v -isnot [string]) {
      $num = [math]::Round([double]$v, 4)
    } else {
      $num = Parse-Brl $t
    }
  }
  if ($t -match '^(\d{1,2})/(\d{1,2})/(\d{2,4})$') {
    $dd = [int]$Matches[1]; $mm = [int]$Matches[2]; $yy = [int]$Matches[3]
    if ($yy -lt 100) { $yy += 2000 }
    if ($mm -ge 1 -and $mm -le 12 -and $dd -ge 1 -and $dd -le 31) {
      $date = "{0:D4}-{1:D2}-{2:D2}" -f $yy, $mm, $dd
    }
  }
  return @{ t = $t; n = $num; d = $date }
}

function Read-Num($ws, $r, $c) {
  $cell = $ws.Cells.Item($r, $c)
  $v = $cell.Value2
  if ($null -eq $v -or "$v" -eq "") {
    return Parse-Brl ([string]$cell.Text)
  }
  if ($v -is [string]) { return Parse-Brl $v }
  return [math]::Round([double]$v, 4)
}

function Get-EmendaKey([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $line = ($s -split "[\r\n]")[0].Trim()
  if ($line -match '(\d{1,4})\s*/\s*(\d{4})') {
    return ("{0:D3}/{1}" -f [int]$Matches[1], $Matches[2])
  }
  return $line
}

function Get-SisKey([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $line = ($s -split "[\r\n]")[0].Trim()
  if ($line -match '(\d{2,6})\s*/\s*(\d{4})') {
    return "$($Matches[1])/$($Matches[2])"
  }
  return $line
}

function Fold-Key([string]$s) {
  if (-not $s) { return "" }
  $norm = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $norm.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne "NonSpacingMark") {
      [void]$sb.Append($ch)
    }
  }
  return ($sb.ToString().ToUpperInvariant() -replace '\s+', ' ').Trim()
}

function Norm-Setor([string]$s) {
  $t = ($s -replace '\s+', ' ').Trim()
  if (-not $t) { return "" }
  $k = Fold-Key $t
  $map = @{
    "COMPRAS" = "Compras"
    "DEPARTAMENTO DE COMPRAS" = "Compras"
    "COMPRAS RP" = "Compras"
    "DAF" = "DAF"
    "DIRETORIA ADMINISTRATIVA E FINANCEIRA" = "DAF"
    "ENCAMINHADO PARA O DAF" = "DAF"
    "SETOR DE LICITACOES" = "Licitacoes"
    "EMENDAS" = "Emendas"
    "DIRETORIA DE ATENCAO PRIMARIA" = "Atencao Primaria"
    "SETOR DE EMPENHOS" = "Empenhos"
    "APOIO EMPENHOS" = "Empenhos"
    "CCGOF" = "CCGOF"
    "SETOR DE COPES" = "COPES"
    "DPTO PLANEJAMENTO" = "Planejamento"
    "DEPARTAMENTO DE PLANEJAMENTO" = "Planejamento"
  }
  if ($map.ContainsKey($k)) {
    $nome = $map[$k]
    if ($nome -eq "Licitacoes") { return "Licitações" }
    if ($nome -eq "Atencao Primaria") { return "Atenção Primária" }
    return $nome
  }
  return $t
}

$seq = 0
function Next-Id([string]$prefix) {
  $script:seq++
  return "$prefix-$($script:seq)"
}

function Read-Block($ws, $start, $maxRow, $cols, $streakLimit) {
  $rows = New-Object System.Collections.Generic.List[object]
  $empty = 0
  for ($r = $start; $r -le $maxRow; $r++) {
    $cells = @()
    $any = $false
    for ($c = 1; $c -le $cols; $c++) {
      $cell = Read-Cell $ws $r $c
      $cells += $cell
      if ($cell.t) { $any = $true }
    }
    if (-not $any) {
      $empty++
      if ($empty -ge $streakLimit) { break }
      continue
    }
    $empty = 0
    $rows.Add(@{ r = $r; cells = $cells })
  }
  return $rows
}

function At($cells, $i) { return $cells[$i] }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false

$empenhos = New-Object System.Collections.Generic.List[object]
$itens = New-Object System.Collections.Generic.List[object]
$liberacoes = New-Object System.Collections.Generic.List[object]
$emendasMap = @{}
$processosMap = @{}
$atas = New-Object System.Collections.Generic.List[object]
$lotes = New-Object System.Collections.Generic.List[object]
$consolidados = New-Object System.Collections.Generic.List[object]

function Add-Emenda([string]$key, [hashtable]$patch) {
  if (-not $key) { return }
  if (-not $script:emendasMap.ContainsKey($key)) {
    $script:emendasMap[$key] = @{
      id = (Next-Id "emenda")
      numero = $key
      ano = $(if ($key -match '/(\d{4})$') { $Matches[1] } else { "" })
      vereador = ""
      partido = ""
      secretaria = ""
      valor = $null
      ementa = ""
      local = ""
      objeto = ""
      justificativa = ""
      observacao = ""
      sis = ""
      memorando = ""
      ubs = ""
      materiaisDefinidos = ""
      smsTerceiro = ""
      orgao = ""
      destino = ""
      natureza = ""
      obra = ""
      definicao = ""
      statusExecucao = ""
    }
  }
  $row = $script:emendasMap[$key]
  foreach ($p in $patch.Keys) {
    $val = $patch[$p]
    if ($null -eq $val) { continue }
    if ($val -is [string] -and $val -eq "" -and $row[$p]) { continue }
    if ($p -eq "valor" -and $null -eq $val) { continue }
    if ($p -eq "valor" -and $null -ne $row[$p] -and $null -ne $val) {
      if ([double]$row[$p] -eq 0 -and [double]$val -ne 0) { $row[$p] = $val }
      continue
    }
    if (-not $row.ContainsKey($p) -or -not $row[$p]) { $row[$p] = $val }
  }
}

function Touch-Processo($rec, [bool]$prefer) {
  $key = $rec.sis
  if (-not $key) { $key = ("sem-sis:" + $rec.objeto + "|" + $rec.data) }
  if (-not $script:processosMap.ContainsKey($key)) {
    $rec.id = Next-Id "proc"
    $script:processosMap[$key] = $rec
    return
  }
  $old = $script:processosMap[$key]
  foreach ($p in @("objeto", "data", "dataTexto", "setor", "descricao", "observacao", "ultimaVerificacao", "ultimaTexto", "dataTramitacao", "ano")) {
    $novo = $rec[$p]
    if ($prefer -and $novo) { $old[$p] = $novo; continue }
    if ((-not $old[$p]) -and $novo) { $old[$p] = $novo }
  }
}

try {
  Write-Output "Lendo controle de empenhos..."
  $wb = $excel.Workbooks.Open("C:\Users\Julia\Downloads\Controle de Empenhos - Atenção Primária.xlsx", 0, $true)
  $ws = $wb.Worksheets.Item("Página1")
  foreach ($row in (Read-Block $ws 10 80 9 8)) {
    $c = $row.cells
    $numero = (At $c 1).t
    if (-not $numero) { continue }
    if ($numero -match '^(TOTAL|PREFEITURA)') { continue }
    $valor = Read-Num $ws $row.r 7
    $documento = ""
    try {
      $linkCell = $ws.Cells.Item($row.r, 2)
      if ($linkCell.Hyperlinks.Count -gt 0) { $documento = [string]$linkCell.Hyperlinks.Item(1).Address }
    } catch {}
    $empenhos.Add(@{
      id = (Next-Id "emp")
      numero = $numero
      documento = $documento
      data = (At $c 0).d
      dataTexto = (At $c 0).t
      fornecedor = (At $c 2).t
      cnpj = (At $c 3).t
      objeto = (At $c 4).t
      unidade = (At $c 5).t
      valor = $valor
      status = (At $c 7).t
      observacoes = (At $c 8).t
      origem = "controle"
      ano = $(if ($numero -match '/(20\d{2})\s*$') { $Matches[1] } else { "" })
    })
  }

  $ws = $wb.Worksheets.Item("Atas de Registro de Preços")
  $ata = @{
    id = (Next-Id "ata")
    numero = "199/2026"
    fornecedor = "Gráfica Alta Definição Ltda."
    processo = "PE nº 058/2026"
    vigencia = "12 meses"
    orgao = "SMS Pelotas"
    valorTotal = $null
    nota = ""
    itens = (New-Object System.Collections.Generic.List[object])
  }
  $maxA = [Math]::Min($ws.UsedRange.Rows.Count, 40)
  for ($r = 1; $r -le $maxA; $r++) {
    $a = (Read-Cell $ws $r 1).t
    if ($a -match 'ATA DE REGISTRO DE PRE[CÇ]OS') { $ata.nota = $a }
    if ($a -match 'Vig[eê]ncia:\s*([^•]+)') { $ata.vigencia = $Matches[1].Trim() }
    if ($a -match 'N[º°]\s*(\d+\s*/\s*\d{4})') { $ata.numero = ($Matches[1] -replace '\s', '') }
    if ($a -match 'FORNECEDOR:\s*([^•]+)') { $ata.fornecedor = $Matches[1].Trim(' ', '.') }
    if ($a -match 'PE\s*n[º°.]?\s*[\d/]+') { $ata.processo = $Matches[0] }
  }
  foreach ($row in (Read-Block $ws 10 30 8 6)) {
    $c = $row.cells
    $item = (At $c 0).t
    if ($item -notmatch '^\d+$') { continue }
    $qtd = Read-Num $ws $row.r 3
    $vu = Read-Num $ws $row.r 5
    $vt = Read-Num $ws $row.r 6
    $ata.itens.Add(@{
      item = $item
      descricao = (At $c 1).t
      quantidade = $qtd
      unidade = (At $c 3).t
      valorUnitario = $vu
      valorTotal = $vt
      destinacao = (At $c 6).t
      observacoes = (At $c 7).t
    })
  }
  $somaAta = 0.0
  foreach ($it in $ata.itens) { if ($null -ne $it.valorTotal) { $somaAta += [double]$it.valorTotal } }
  $ata.valorTotal = [math]::Round($somaAta, 2)
  $ata.itens = @($ata.itens.ToArray())
  $atas.Add($ata)
  $wb.Close($false)

  Write-Output "Lendo relacao de empenhos..."
  $wb = $excel.Workbooks.Open("C:\Users\Julia\Downloads\Relação Empenhos - Compras e Emendas.xlsx", 0, $true)
  $ws = $wb.Worksheets.Item("Relação Empenhos - Emendas")
  $docsRelacao = @{}
  foreach ($row in (Read-Block $ws 3 400 17 20)) {
    $c = $row.cells
    $numero = (At $c 0).t
    if (-not $numero -or $numero -match 'Empenho') { continue }
    $qtd = Read-Num $ws $row.r 10
    $vu = Read-Num $ws $row.r 12
    $vt = Read-Num $ws $row.r 13
    $itens.Add(@{
      id = (Next-Id "item")
      numeroEmpenho = $numero
      data = (At $c 1).d
      dataTexto = (At $c 1).t
      sis = (Get-SisKey (At $c 2).t)
      tipo = (At $c 3).t
      emenda = (Get-EmendaKey (At $c 4).t)
      autor = (At $c 5).t
      solicitante = (At $c 6).t
      fornecedor = (At $c 7).t
      descricao = (At $c 8).t
      quantidade = $qtd
      unidade = (At $c 10).t
      valorUnitario = $vu
      valorTotal = $vt
      dataContato = (At $c 13).d
      dataContatoTexto = (At $c 13).t
      previsaoEntrega = (At $c 14).d
      previsaoTexto = (At $c 14).t
      observacoes = (At $c 15).t
      destino = (At $c 16).t
      origem = "relacao"
    })
    if (-not $docsRelacao.ContainsKey($numero)) {
      $link = ""
      try {
        $linkCell = $ws.Cells.Item($row.r, 1)
        if ($linkCell.Hyperlinks.Count -gt 0) { $link = [string]$linkCell.Hyperlinks.Item(1).Address }
      } catch {}
      $docsRelacao[$numero] = $link
    }
  }
  $vistosEmp = @{}
  foreach ($e in $empenhos) { $vistosEmp[$e.numero.ToUpper().Replace(" ", "")] = $true }
  $porNum = @{}
  foreach ($it in $itens) {
    if ($it.origem -ne "relacao" -or -not $it.numeroEmpenho) { continue }
    if (-not $porNum.ContainsKey($it.numeroEmpenho)) { $porNum[$it.numeroEmpenho] = New-Object System.Collections.Generic.List[object] }
    $porNum[$it.numeroEmpenho].Add($it)
  }
  foreach ($k in $porNum.Keys) {
    $norm = $k.ToUpper().Replace(" ", "")
    if ($vistosEmp.ContainsKey($norm)) { continue }
    $grupo = $porNum[$k]
    $primeiro = $grupo[0]
    $somaRel = 0.0
    $sises = New-Object System.Collections.Generic.List[string]
    $ems = New-Object System.Collections.Generic.List[string]
    $desc = ""
    foreach ($it in $grupo) {
      if ($null -ne $it.valorTotal) { $somaRel += [double]$it.valorTotal }
      if ($it.sis -and -not $sises.Contains($it.sis)) { $sises.Add([string]$it.sis) }
      if ($it.emenda -match '^\d{3}/\d{4}$' -and -not $ems.Contains($it.emenda)) { $ems.Add([string]$it.emenda) }
      if (-not $desc -and $it.descricao) { $desc = [string]$it.descricao }
    }
    $anoRel = ""
    if ($k -match '/(20\d{2})\s*$') { $anoRel = $Matches[1] }
    $obsRel = @()
    foreach ($s in $sises) { $obsRel += "SIS $s" }
    foreach ($em in $ems) { $obsRel += "Emenda $em" }
    $slug = ($norm -replace '[^A-Z0-9]+', '-').Trim('-')
    $empenhos.Add(@{
      id = "emp-rel-$slug"
      numero = $k
      documento = $(if ($docsRelacao.ContainsKey($k)) { $docsRelacao[$k] } else { "" })
      data = $primeiro.data
      dataTexto = $primeiro.dataTexto
      fornecedor = $primeiro.fornecedor
      cnpj = ""
      objeto = $desc
      unidade = "Secretaria Municipal de Saúde / Atenção Primária em Saúde"
      valor = [math]::Round($somaRel, 2)
      status = ""
      observacoes = ($obsRel -join "; ")
      origem = "relacao"
      ano = $anoRel
      sis = $(if ($sises.Count) { $sises[0] } else { "" })
    })
  }

  $ws = $wb.Worksheets.Item("Relação de Emendas")
  foreach ($row in (Read-Block $ws 2 250 12 15)) {
    $c = $row.cells
    $key = Get-EmendaKey (At $c 0).t
    if (-not $key -or $key -match 'PARTIDO|EMENDA') { continue }
    $valor = Read-Num $ws $row.r 4
    Add-Emenda $key @{
      vereador = (At $c 2).t
      partido = (At $c 1).t
      valor = $valor
      smsTerceiro = (At $c 4).t
      orgao = (At $c 5).t
      destino = (At $c 6).t
      natureza = (At $c 7).t
      obra = (At $c 8).t
      definicao = (At $c 9).t
      statusExecucao = (At $c 10).t
      memorando = (At $c 11).t
    }
  }
  $wb.Close($false)

  Write-Output "Lendo emendas impositivas..."
  $wb = $excel.Workbooks.Open("C:\Users\Julia\Downloads\Planilha Emendas Impositivas 2025_2026 DAP (1).xlsx", 0, $true)
  $ws = $wb.Worksheets.Item("Planilha 2025 Original")
  foreach ($row in (Read-Block $ws 3 120 10 15)) {
    $c = $row.cells
    $key = Get-EmendaKey (At $c 0).t
    if (-not $key) { continue }
    Add-Emenda $key @{
      vereador = (At $c 1).t
      secretaria = (At $c 2).t
      valor = (Read-Num $ws $row.r 4)
      ementa = (At $c 4).t
      local = (At $c 5).t
      objeto = (At $c 6).t
      justificativa = (At $c 7).t
      observacao = (At $c 8).t
      sis = (Get-SisKey (At $c 9).t)
    }
  }
  $ws = $wb.Worksheets.Item("Planilha Resumida")
  foreach ($row in (Read-Block $ws 3 80 8 12)) {
    $c = $row.cells
    $key = Get-EmendaKey (At $c 0).t
    if (-not $key) { continue }
    Add-Emenda $key @{
      vereador = (At $c 1).t
      secretaria = (At $c 2).t
      valor = (Read-Num $ws $row.r 4)
      ementa = (At $c 4).t
      ubs = (At $c 5).t
      objeto = (At $c 6).t
      observacao = (At $c 7).t
    }
  }
  $ws = $wb.Worksheets.Item("Materiais, equip. e insumos a s")
  foreach ($row in (Read-Block $ws 3 40 9 10)) {
    $c = $row.cells
    $key = Get-EmendaKey (At $c 0).t
    if (-not $key) { continue }
    Add-Emenda $key @{
      vereador = (At $c 1).t
      secretaria = (At $c 2).t
      valor = (Read-Num $ws $row.r 4)
      ementa = (At $c 4).t
      ubs = (At $c 5).t
      objeto = (At $c 6).t
      observacao = (At $c 7).t
      materiaisDefinidos = (At $c 8).t
    }
  }

  $ws = $wb.Worksheets.Item("Divisão em lotes de mat. já def")
  $loteNome = ""
  foreach ($row in (Read-Block $ws 1 80 4 12)) {
    $c = $row.cells
    $a = (At $c 0).t
    if (-not $a) { continue }
    if ($a -match '^LOTE\b') { $loteNome = ($a -split "[\r\n]")[0].Trim(); continue }
    if ($a -match '^(Item|Descri)') { continue }
    $sis = ""
    if ($a -match 'SIS\s*:?\s*(\d{3,6})') { $sis = $Matches[1] }
    $nome = ($a -split '(?i)SIS')[0].Trim()
    if (-not $nome) { continue }
    $lotes.Add(@{
      id = (Next-Id "lote")
      lote = $loteNome
      item = $nome
      sis = $sis
      descricao = (At $c 1).t
      destinos = (At $c 2).t
      emendas = (At $c 3).t
    })
  }

  $ws = $wb.Worksheets.Item("Consolidação de Materiais")
  foreach ($row in (Read-Block $ws 2 40 6 8)) {
    $c = $row.cells
    $cat = (At $c 0).t
    $item = (At $c 1).t
    if (-not $item -or $item -match '^Item$') { continue }
    $consolidados.Add(@{
      id = (Next-Id "cons")
      categoria = $cat
      item = $item
      qtdDefinidos = (Read-Num $ws $row.r 3)
      qtdOriginais = (Read-Num $ws $row.r 4)
      total = (Read-Num $ws $row.r 5)
      destinos = (At $c 5).t
    })
  }
  $wb.Close($false)

  Write-Output "Lendo acompanhamento de compras..."
  $wb = $excel.Workbooks.Open("C:\Users\Julia\Downloads\Acompanhamento Compras.xlsx", 0, $true)

  function Add-ProcFrom($ws, $start, $max, $mapCols, [bool]$prefer) {
    foreach ($row in (Read-Block $ws $start $max 7 12)) {
      $c = $row.cells
      $objeto = (At $c $mapCols.objeto).t
      $sisRaw = (At $c $mapCols.sis).t
      if ($objeto -match '^(OBJETO|DATA)$' -or $sisRaw -match '^(SIS|OBJETO)$') { continue }
      if (-not $objeto -and -not $sisRaw) { continue }
      $dataCell = At $c $mapCols.data
      $ultCell = At $c $mapCols.ultima
      $setorRaw = (At $c $mapCols.setor).t
      $sis = Get-SisKey $sisRaw
      $ano = ""
      if ($sis -match '/(\d{4})$') { $ano = $Matches[1] }
      elseif ($dataCell.d -match '^(\d{4})') { $ano = $Matches[1] }
      $rec = @{
        sis = $sis
        objeto = $objeto
        data = $dataCell.d
        dataTexto = $dataCell.t
        setor = (Norm-Setor $setorRaw)
        setorOriginal = $setorRaw
        descricao = ""
        observacao = ""
        ultimaVerificacao = $ultCell.d
        ultimaTexto = $ultCell.t
        dataTramitacao = ""
        ano = $ano
        historico = ""
      }
      if ($mapCols.ContainsKey("descricao")) { $rec.descricao = (At $c $mapCols.descricao).t }
      if ($mapCols.ContainsKey("observacao")) { $rec.observacao = (At $c $mapCols.observacao).t }
      if ($mapCols.ContainsKey("dataTramitacao")) {
        $dt = At $c $mapCols.dataTramitacao
        $rec.dataTramitacao = $(if ($dt.d) { $dt.d } else { $dt.t })
      }
      Touch-Processo $rec $prefer
    }
  }

  Add-ProcFrom $wb.Worksheets.Item("SOLICITAÇÕES") 2 200 @{
    data = 0; sis = 1; objeto = 2; setor = 3; observacao = 4; ultima = 5
  } $false
  Add-ProcFrom $wb.Worksheets.Item("2024") 2 80 @{
    objeto = 0; data = 1; sis = 2; setor = 3; observacao = 4; ultima = 5
  } $false
  Add-ProcFrom $wb.Worksheets.Item("2025") 2 120 @{
    objeto = 0; data = 1; sis = 2; setor = 3; observacao = 4; ultima = 5
  } $false
  Add-ProcFrom $wb.Worksheets.Item("2026") 2 160 @{
    objeto = 0; data = 1; sis = 2; setor = 3; dataTramitacao = 4; descricao = 5; ultima = 6
  } $true

  $ws = $wb.Worksheets.Item("EMPENHOS")
  $sisAtual = ""
  $dataAtual = ""
  foreach ($row in (Read-Block $ws 2 80 5 8)) {
    $c = $row.cells
    if ((At $c 0).t -and (At $c 0).t -ne "SIS") { $sisAtual = Get-SisKey (At $c 0).t }
    if ((At $c 3).d) { $dataAtual = (At $c 3).d }
    elseif ((At $c 3).t -and (At $c 3).t -ne "DATA DE LIBERAÇÃO") { $dataAtual = (At $c 3).t }
    $num = (At $c 1).t
    if (-not $num -or $num -match 'EMPENHO') { continue }
    $liberacoes.Add(@{
      id = (Next-Id "lib")
      sis = $sisAtual
      numero = ($num -replace '\s', '')
      valor = (Read-Num $ws $row.r 3)
      data = $(if ((At $c 3).d) { (At $c 3).d } else { $dataAtual })
      dataTexto = (At $c 3).t
      observacao = (At $c 4).t
      origem = "acompanhamento"
    })
  }
  $wb.Close($false)

  $soma = 0.0
  foreach ($e in $empenhos) { if ($null -ne $e.valor) { $soma += [double]$e.valor } }
  $st = @{}
  foreach ($e in $empenhos) {
    $k = $e.status
    if (-not $st.ContainsKey($k)) { $st[$k] = 0 }
    $st[$k]++
  }
  Write-Output ("EMPENHOS " + $empenhos.Count + " SOMA " + $soma.ToString("N2", $pt))
  Write-Output ("STATUS " + (($st.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join " | "))
  Write-Output ("ITENS " + $itens.Count)
  Write-Output ("EMENDAS " + $emendasMap.Count)
  Write-Output ("PROCESSOS " + $processosMap.Count)
  Write-Output ("LIBERACOES " + $liberacoes.Count)
  Write-Output ("ATAS " + $atas.Count + " ITENS_ATA " + $atas[0].itens.Count)
  Write-Output ("LOTES " + $lotes.Count)
  Write-Output ("CONSOLID " + $consolidados.Count)

  $root = @{
    geradoEm = (Get-Date).ToString("yyyy-MM-dd")
    empenhos = @($empenhos.ToArray())
    itens = @($itens.ToArray())
    liberacoes = @($liberacoes.ToArray())
    emendas = @($emendasMap.Values)
    processos = @($processosMap.Values)
    atas = @($atas.ToArray())
    lotes = @($lotes.ToArray())
    consolidados = @($consolidados.ToArray())
  }

  Add-Type -AssemblyName System.Web.Extensions
  $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = 67108864
  $json = $ser.Serialize($root)
  $js = "window.SEED = $json;`n"
  $utf8 = New-Object System.Text.UTF8Encoding $false
  [System.IO.File]::WriteAllText($dest, $js, $utf8)
  Write-Output ("SEED_BYTES " + (Get-Item -LiteralPath $dest).Length)
}
finally {
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
