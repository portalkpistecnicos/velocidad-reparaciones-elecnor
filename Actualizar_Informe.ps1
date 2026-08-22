#Requires -Version 5.1
# Actualiza el informe P31 - Velocidad de Reparaciones (pagina GitHub) a partir del CSV mas reciente en BBDD.

$ErrorActionPreference = 'Stop'

# ---------------- Configuracion ----------------
$BbddFolder = "G:\Mi unidad\Respaldo\Documents\KPI's\Velocidad Reparaciones\BBDD"
$RepoPath   = "C:\Users\xarancibia\Documents\GitHub\velocidad-reparaciones-elecnor"
$HtmlOutput = Join-Path $RepoPath "index.html"
$LogoPath   = "G:\Mi unidad\Respaldo\Documents\Elecnor\logo elecnor.png"
$Meta       = 85.0
$utf8NoBom  = New-Object System.Text.UTF8Encoding($false)

$agenciasOrden  = @('VINA DEL MAR','VALPARAISO','SAN ANTONIO')
$nombresAgencia = @{ 'VINA DEL MAR'='Viña del Mar'; 'VALPARAISO'='Valparaíso'; 'SAN ANTONIO'='San Antonio' }
$mesesEs = @{1='Enero';2='Febrero';3='Marzo';4='Abril';5='Mayo';6='Junio';7='Julio';8='Agosto';9='Septiembre';10='Octubre';11='Noviembre';12='Diciembre'}

Write-Host "=== Actualizacion informe P31 - Velocidad de Reparaciones ===" -ForegroundColor Cyan

# ---------------- Helpers de formato (es-CL: miles con punto, decimales con coma) ----------------
function Fmt1([double]$n) {
    $s = $n.ToString("N1", [System.Globalization.CultureInfo]::InvariantCulture)
    return $s.Replace(',', '§').Replace('.', ',').Replace('§', '.')
}
function FmtInt([long]$n) {
    $s = $n.ToString("N0", [System.Globalization.CultureInfo]::InvariantCulture)
    return $s.Replace(',', '.')
}
function FmtCss([double]$n) {
    return $n.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
}
function StatusVar([double]$pct) {
    if ($pct -ge $Meta) { return 'good' }
    elseif ($pct -ge 65) { return 'warn' }
    else { return 'bad' }
}

# ---------------- 1. Ubicar CSV mas reciente ----------------
$csv = Get-ChildItem -Path $BbddFolder -Filter "*.csv" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $csv) { throw "No se encontro ningun archivo .csv en $BbddFolder" }
Write-Host "Fuente: $($csv.Name)  (modificado $($csv.LastWriteTime))"

$raw = Import-Csv -Path $csv.FullName -Delimiter ';' -Encoding UTF8
Write-Host "Registros leidos: $($raw.Count)"

# ---------------- 2. Aplicar metodologia oficial P31 ----------------
$data = @($raw | Where-Object {
    $_.tecno_acceso -eq 'FO' -and
    $_.ambito -eq 'ATENCION CLIENTES' -and
    $_.es_falla_masiva -eq '0'
})
$excluidos = $raw.Count - $data.Count

foreach ($r in $data) {
    $dias = 0
    [void][int]::TryParse($r.q_dias_efectivos, [ref]$dias)
    $r | Add-Member -NotePropertyName CumpleOficial -NotePropertyValue ($dias -lt 2) -Force
}

$totalGlobal  = $data.Count
$cumpleGlobal = @($data | Where-Object { $_.CumpleOficial }).Count
$pctGlobal    = if ($totalGlobal -gt 0) { [math]::Round(($cumpleGlobal / $totalGlobal) * 100, 1) } else { 0 }
$brechaGlobal = [math]::Round($Meta - $pctGlobal, 1)
$faltanGlobal = [math]::Max(0, [math]::Ceiling($totalGlobal * ($Meta/100)) - $cumpleGlobal)

# ---------------- 3. Por agencia y por bucket ----------------
$agenciasInfo = @()
foreach ($ag in $agenciasOrden) {
    $grp = @($data | Where-Object { $_.agencia.Trim().ToUpper() -eq $ag })
    if ($grp.Count -eq 0) { continue }
    $c = @($grp | Where-Object { $_.CumpleOficial }).Count
    $t = $grp.Count
    $pct = [math]::Round(($c/$t)*100,1)
    $brecha = [math]::Round($Meta - $pct,1)
    $falta = [math]::Max(0,[math]::Ceiling($t*($Meta/100)) - $c)

    $buckets = @($grp | Group-Object { $_.cod_bucket.Trim() } | Sort-Object Count -Descending | ForEach-Object {
        $bc = @($_.Group | Where-Object { $_.CumpleOficial }).Count
        $bt = $_.Count
        $bpct = [math]::Round(($bc/$bt)*100,1)
        $bfalta = [math]::Max(0,[math]::Ceiling($bt*($Meta/100)) - $bc)
        [PSCustomObject]@{ Nombre=$_.Name; Total=$bt; Cumple=$bc; Pct=$bpct; Falta=$bfalta }
    })

    $agenciasInfo += [PSCustomObject]@{
        Key=$ag; Nombre=$nombresAgencia[$ag]; Total=$t; Cumple=$c; Pct=$pct; Brecha=$brecha; Falta=$falta; Buckets=$buckets
    }
}
$agenciasInfo = @($agenciasInfo | Sort-Object Total -Descending)
$peorAgencia  = $agenciasInfo | Sort-Object Pct | Select-Object -First 1
$mejorAgencia = $agenciasInfo | Sort-Object Pct -Descending | Select-Object -First 1

# ---------------- 4. Causas de incumplimiento ----------------
$incumplidos = @($data | Where-Object { -not $_.CumpleOficial })
$nInc = $incumplidos.Count
$causasRaw = @($incumplidos | Group-Object { if ([string]::IsNullOrWhiteSpace($_.area_resp_quiebre_pcita)) { 'Sin especificar' } else { $_.area_resp_quiebre_pcita } } | Sort-Object Count -Descending)
$causasTop = @($causasRaw | Select-Object -First 5 | ForEach-Object {
    [PSCustomObject]@{ Nombre=$_.Name; Count=$_.Count; Pct=[math]::Round(($_.Count/[math]::Max(1,$nInc))*100,1) }
})
$causasResto = @($causasRaw | Select-Object -Skip 5)
if ($causasResto.Count -gt 0) {
    $restoCount = ($causasResto | Measure-Object -Property Count -Sum).Sum
    $causasTop += [PSCustomObject]@{ Nombre='Otras causas menores'; Count=$restoCount; Pct=[math]::Round(($restoCount/[math]::Max(1,$nInc))*100,1) }
}

# ---------------- 5. Tendencia diaria ----------------
$diario = @($data | Group-Object { [int]$_.dia_cierre } | Sort-Object { [int]$_.Name } | ForEach-Object {
    $c = @($_.Group | Where-Object { $_.CumpleOficial }).Count
    $t = $_.Count
    [PSCustomObject]@{ Dia=[int]$_.Name; Total=$t; Cumple=$c; Pct=[math]::Round(($c/$t)*100,1) }
})
$diarioFiltrado = @($diario | Where-Object { $_.Total -ge 20 })
$diarioOmitido  = @($diario | Where-Object { $_.Total -lt 20 })

# ---------------- 6. Periodo / alcance ----------------
$periodo = "Periodo no identificado"
if ($csv.Name -match '(\d{4})-(\d{2})') {
    $anio = [int]$Matches[1]; $mes = [int]$Matches[2]
    if ($mesesEs.ContainsKey($mes)) { $periodo = "$($mesesEs[$mes]) $anio" }
}
$alcance = ($agenciasInfo | ForEach-Object { $_.Nombre }) -join ' &middot; '
$metaTxt = Fmt1 $Meta
$fechaGeneracion = Get-Date -Format "dd-MM-yyyy HH:mm"

# ---------------- 7. Logo en base64 ----------------
$logoB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($LogoPath))

# ================= Construccion de fragmentos HTML =================

# --- KPI cards ---
$kpi1Class = if ($pctGlobal -ge $Meta) { 'kpi good' } else { 'kpi bad' }
if ($brechaGlobal -gt 0) { $brechaTxt = "&minus;$(Fmt1 $brechaGlobal) pts"; $kpi3Class = 'kpi bad'; $brechaSub = "&asymp; $(FmtInt $faltanGlobal) casos adicionales" }
elseif ($brechaGlobal -lt 0) { $brechaTxt = "+$(Fmt1 ([math]::Abs($brechaGlobal))) pts"; $kpi3Class = 'kpi good'; $brechaSub = "meta superada" }
else { $brechaTxt = "0,0 pts"; $kpi3Class = 'kpi good'; $brechaSub = "meta alcanzada" }

$kpiCards = @"
    <div class="$kpi1Class">
      <p class="kpi-label">Cumplimiento global</p>
      <p class="kpi-value">$(Fmt1 $pctGlobal)%</p>
      <p class="kpi-sub">$(FmtInt $cumpleGlobal) de $(FmtInt $totalGlobal) reparaciones</p>
    </div>
    <div class="kpi neutral">
      <p class="kpi-label">Meta corporativa</p>
      <p class="kpi-value">$metaTxt%</p>
      <p class="kpi-sub">única meta, $($agenciasInfo.Count) agencias</p>
    </div>
    <div class="$kpi3Class">
      <p class="kpi-label">Brecha</p>
      <p class="kpi-value">$brechaTxt</p>
      <p class="kpi-sub">$brechaSub</p>
    </div>
    <div class="kpi neutral">
      <p class="kpi-label">Universo analizado</p>
      <p class="kpi-value">$(FmtInt $totalGlobal)</p>
      <p class="kpi-sub">de $(FmtInt $raw.Count) registros del período</p>
    </div>
"@

# --- Nota exclusion metodologia ---
if ($excluidos -gt 0) {
    $notaExclusion = "Sobre los $(FmtInt $raw.Count) registros brutos del archivo se excluyeron $excluidos casos fuera de alcance (ámbito distinto de atención a clientes, falla masiva o acceso no FO), dejando un universo válido de $(FmtInt $totalGlobal) reparaciones."
} else {
    $notaExclusion = "Los $(FmtInt $raw.Count) registros del archivo cumplen el alcance del reporte (FO, atención a clientes, sin fallas masivas)."
}

# --- Seccion 02: agencias ---
$agenciasBajoMeta = @($agenciasInfo | Where-Object { $_.Pct -lt $Meta })
if ($agenciasBajoMeta.Count -eq $agenciasInfo.Count -and $agenciasInfo.Count -gt 0) {
    $ledeAgencias = "Ninguna de las $($agenciasInfo.Count) agencias alcanza el $metaTxt% exigido."
} elseif ($agenciasBajoMeta.Count -eq 0) {
    $ledeAgencias = "Las $($agenciasInfo.Count) agencias alcanzan la meta de $metaTxt%."
} else {
    $ledeAgencias = "$($agenciasBajoMeta.Count) de $($agenciasInfo.Count) agencias no alcanza el $metaTxt% exigido."
}
if ($peorAgencia) {
    $shareVol = [math]::Round(($peorAgencia.Total/$totalGlobal)*100,0)
    $ledeAgencias += " $($peorAgencia.Nombre) concentra el $shareVol% del volumen total y el desempeño más bajo ($(Fmt1 $peorAgencia.Pct)%)."
}

$barrasAgencia = ""
$filasAgencia = ""
foreach ($info in $agenciasInfo) {
    $sv = StatusVar $info.Pct
    $barrasAgencia += @"
      <div class="bar-row">
        <div class="bar-label">$($info.Nombre)<small>$(FmtInt $info.Total) reparaciones</small></div>
        <div class="bar-track">
          <div class="bar-fill" style="width:$(FmtCss $info.Pct)%; background:var(--$sv);"></div>
          <div class="bar-goal" style="left:$(FmtCss $Meta)%;"></div>
        </div>
        <div class="bar-val" style="color:var(--$sv);">$(Fmt1 $info.Pct)%</div>
      </div>

"@
    $brechaCelda = if ($info.Brecha -gt 0) { "&minus;$(Fmt1 $info.Brecha)" } elseif ($info.Brecha -lt 0) { "+$(Fmt1 ([math]::Abs($info.Brecha)))" } else { "0,0" }
    $filasAgencia += "          <tr><td class=`"name`">$($info.Nombre)</td><td class=`"num`">$(FmtInt $info.Total)</td><td class=`"num`">$(FmtInt $info.Cumple)</td><td class=`"num`"><span class=`"pill $sv`"><span class=`"dot`"></span>$(Fmt1 $info.Pct)%</span></td><td class=`"num`">$brechaCelda</td><td class=`"num`">$(FmtInt $info.Falta)</td></tr>`n"
}
$brechaGlobalCelda = if ($brechaGlobal -gt 0) { "&minus;$(Fmt1 $brechaGlobal)" } elseif ($brechaGlobal -lt 0) { "+$(Fmt1 ([math]::Abs($brechaGlobal)))" } else { "0,0" }
$faltanGlobalTxt = if ($faltanGlobal -gt 0) { "$(FmtInt $faltanGlobal)" } else { "0" }

# --- Seccion 03: causas ---
$causasRows = ""
foreach ($c in $causasTop) {
    $causasRows += @"
        <div class="cause-row">
          <span>$($c.Nombre)</span>
          <div class="cause-track"><div class="cause-fill" style="width:$(FmtCss $c.Pct)%;"></div></div>
          <span class="cause-pct">$(Fmt1 $c.Pct)%</span>
        </div>

"@
}
$textoCausas = ""
if ($causasTop.Count -gt 0) {
    $top1 = $causasTop[0]
    $textoCausas = "El $(Fmt1 $top1.Pct)% de los incumplimientos corresponde a &ldquo;$($top1.Nombre)&rdquo;"
    if ($causasTop.Count -gt 1) {
        $top2 = $causasTop[1]
        $textoCausas += ", seguido por &ldquo;$($top2.Nombre)&rdquo; con $(Fmt1 $top2.Pct)%."
    } else {
        $textoCausas += "."
    }
}

# --- Seccion 04: buckets ---
$allBuckets = @($agenciasInfo | ForEach-Object { $_.Buckets })
$bucketsBajoMeta = @($allBuckets | Where-Object { $_.Pct -lt $Meta })
if ($allBuckets.Count -gt 0 -and $bucketsBajoMeta.Count -eq $allBuckets.Count) {
    $ledeBuckets = "Ningún bucket alcanza individualmente el $metaTxt%."
} elseif ($bucketsBajoMeta.Count -eq 0) {
    $ledeBuckets = "Todos los buckets alcanzan la meta de $metaTxt%."
} else {
    $ledeBuckets = "$($bucketsBajoMeta.Count) de $($allBuckets.Count) buckets no alcanza el $metaTxt%."
}
if ($peorAgencia) {
    $ledeBuckets += " En $($peorAgencia.Nombre) el rezago es transversal a los $($peorAgencia.Buckets.Count) buckets, señal de una causa sistémica de agencia más que de un equipo puntual."
}

$bucketsGrid = ""
foreach ($info in $agenciasInfo) {
    $filasBucket = ""
    foreach ($b in $info.Buckets) {
        $sv = StatusVar $b.Pct
        $filasBucket += "            <tr><td class=`"name`">$($b.Nombre)</td><td class=`"num`">$(FmtInt $b.Total)</td><td class=`"num`"><span class=`"pill $sv`"><span class=`"dot`"></span>$(Fmt1 $b.Pct)%</span></td><td class=`"num`">$(FmtInt $b.Falta)</td></tr>`n"
    }
    $bucketsGrid += @"
      <div>
        <h3>$($info.Nombre)</h3>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Bucket</th><th class="num">Total</th><th class="num">%</th><th class="num">Falta</th></tr></thead>
          <tbody>
$filasBucket          </tbody>
        </table></div>
      </div>

"@
}

# --- Seccion 05: tendencia diaria (SVG) ---
$ledeTendencia = ""
if ($diarioFiltrado.Count -gt 0) {
    $minPct = ($diarioFiltrado | Measure-Object -Property Pct -Minimum).Minimum
    $maxPct = ($diarioFiltrado | Measure-Object -Property Pct -Maximum).Maximum
    $ledeTendencia = "El cumplimiento diario osciló entre $(Fmt1 $minPct)% y $(Fmt1 $maxPct)% a lo largo del período" + $(if ($maxPct -lt $Meta) { ", sin alcanzar la meta de $metaTxt% en ningún día con volumen representativo." } else { "." })
}

$goalY = 228 - ($Meta * 2.08)
$polyPoints = ""
$circles = ""
$dayLabels = ""
if ($diarioFiltrado.Count -gt 0) {
    $minDay = ($diarioFiltrado | Measure-Object -Property Dia -Minimum).Minimum
    $maxDay = ($diarioFiltrado | Measure-Object -Property Dia -Maximum).Maximum
    $rango = [math]::Max(1, $maxDay - $minDay)
    $pts = @()
    foreach ($d in $diarioFiltrado) {
        $x = [math]::Round(60 + (($d.Dia - $minDay) / $rango) * 768, 1)
        $y = [math]::Round(228 - ($d.Pct * 2.08), 1)
        $pts += "$x,$y"
        $circles += "          <circle cx=`"$x`" cy=`"$y`" r=`"3.2`"/>`n"
        $dayLabels += "          <text x=`"$x`" y=`"248`" text-anchor=`"middle`">$($d.Dia)</text>`n"
    }
    $polyPoints = $pts -join ' '
}

$notaTendencia = ""
if ($diarioOmitido.Count -gt 0) {
    $diasOmitidosTxt = ($diarioOmitido | ForEach-Object { $_.Dia }) -join ', '
    $notaTendencia = "Se omiten del gráfico los días $diasOmitidosTxt por bajo volumen (menos de 20 casos), que distorsionan el porcentaje diario."
}

# --- Seccion 06: conclusiones ---
$recoItems = ""
$recoItems += "      <li><p><strong>Brecha de $(Fmt1 ([math]::Abs($brechaGlobal))) puntos</strong> respecto a la meta de $metaTxt%: " + $(if ($faltanGlobal -gt 0) { "se requieren cerca de $(FmtInt $faltanGlobal) reparaciones adicionales cumpliendo el estándar de &le; 1 día efectivo para alcanzarla con el volumen actual." } else { "la meta ya se encuentra alcanzada a nivel consolidado." }) + "</p></li>`n"
if ($peorAgencia) {
    $shareVol = [math]::Round(($peorAgencia.Total/$totalGlobal)*100,0)
    $recoItems += "      <li><p><strong>$($peorAgencia.Nombre) es la prioridad de intervención:</strong> concentra el $shareVol% del volumen y el desempeño más bajo ($(Fmt1 $peorAgencia.Pct)%), con $(FmtInt $peorAgencia.Falta) casos adicionales necesarios para alcanzar la meta.</p></li>`n"
}
if ($causasTop.Count -gt 0) {
    $top1 = $causasTop[0]
    $recoItems += "      <li><p><strong>$(Fmt1 $top1.Pct)% del incumplimiento</strong> está asociado a &ldquo;$($top1.Nombre)&rdquo;: es el primer foco de mejora operativa.</p></li>`n"
}
if ($mejorAgencia -and $peorAgencia -and $mejorAgencia.Key -ne $peorAgencia.Key) {
    $recoItems += "      <li><p><strong>$($mejorAgencia.Nombre), con la brecha menor</strong> ($(FmtInt $mejorAgencia.Falta) casos para la meta), es el objetivo más alcanzable en el corto plazo y puede servir de referencia de buenas prácticas para las otras agencias.</p></li>`n"
}

# ================= Plantilla HTML =================
$template = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Velocidad de Reparaciones — Elecnor</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<style>
  :root{
    --ink:#1b2430;
    --ink-soft:#51606d;
    --ink-faint:#8493a0;
    --paper:#f4f6f5;
    --surface:#ffffff;
    --surface-2:#eef1f0;
    --line:#d7dedc;
    --accent:#004187;
    --accent-ink:#003267;
    --accent-soft:#e5edf5;
    --brand-orange:#f47c02;
    --brand-orange-soft:#fdecd9;
    --good:#2f7d4f;
    --good-soft:#e6f3ec;
    --warn:#a3701c;
    --warn-soft:#f8f0dd;
    --bad:#af3a2d;
    --bad-soft:#f8e8e6;
    --shadow: 0 1px 2px rgba(27,36,48,.06), 0 6px 20px -8px rgba(27,36,48,.15);
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      --ink:#e7ecef;
      --ink-soft:#a9b6c0;
      --ink-faint:#71818d;
      --paper:#111820;
      --surface:#182029;
      --surface-2:#1f2933;
      --line:#2c3742;
      --accent:#6fa8dc;
      --accent-ink:#a9cdeb;
      --accent-soft:#17283a;
      --brand-orange:#ff9c40;
      --brand-orange-soft:#332616;
      --good:#5fbb87;
      --good-soft:#152a1e;
      --warn:#dcab5c;
      --warn-soft:#2c2416;
      --bad:#e28d80;
      --bad-soft:#2f1c19;
      --shadow: 0 1px 2px rgba(0,0,0,.3), 0 8px 24px -10px rgba(0,0,0,.5);
    }
  }
  :root[data-theme="dark"]{
    --ink:#e7ecef;
    --ink-soft:#a9b6c0;
    --ink-faint:#71818d;
    --paper:#111820;
    --surface:#182029;
    --surface-2:#1f2933;
    --line:#2c3742;
    --accent:#6fa8dc;
    --accent-ink:#a9cdeb;
    --accent-soft:#17283a;
    --brand-orange:#ff9c40;
    --brand-orange-soft:#332616;
    --good:#5fbb87;
    --good-soft:#152a1e;
    --warn:#dcab5c;
    --warn-soft:#2c2416;
    --bad:#e28d80;
    --bad-soft:#2f1c19;
    --shadow: 0 1px 2px rgba(0,0,0,.3), 0 8px 24px -10px rgba(0,0,0,.5);
  }

  *{box-sizing:border-box;}
  body{
    margin:0;
    background:var(--paper);
    color:var(--ink);
    font-family:"IBM Plex Sans", "Segoe UI", sans-serif;
    font-size:15.5px;
    line-height:1.6;
    -webkit-font-smoothing:antialiased;
  }
  .sheet{
    max-width:920px;
    margin:0 auto;
    padding:56px 48px 80px;
  }
  .nums{ font-family:"IBM Plex Mono", ui-monospace, monospace; font-variant-numeric:tabular-nums; }

  .masthead{
    display:flex;
    justify-content:space-between;
    align-items:flex-end;
    gap:24px;
    padding-bottom:20px;
    border-bottom:3px solid var(--accent);
    margin-bottom:34px;
  }
  .masthead-brand{ display:flex; align-items:center; gap:22px; }
  .brand-logo{ height:44px; width:auto; display:block; flex:none; }
  .masthead-divider{ width:1px; align-self:stretch; min-height:46px; background:var(--line); }
  .brand-accent{ width:46px; height:4px; background:var(--brand-orange); border-radius:2px; margin-top:11px; }
  .eyebrow{
    font-family:"IBM Plex Mono", monospace;
    font-size:11.5px;
    letter-spacing:.14em;
    text-transform:uppercase;
    color:var(--accent-ink);
    font-weight:600;
    margin:0 0 10px;
  }
  h1.title{
    font-family:"Archivo", sans-serif;
    font-weight:800;
    font-size:clamp(28px,4vw,38px);
    line-height:1.08;
    letter-spacing:-.01em;
    margin:0;
    color:var(--ink);
    text-wrap:balance;
  }
  .masthead-meta{
    text-align:right;
    font-size:13px;
    color:var(--ink-soft);
    line-height:1.7;
    white-space:nowrap;
  }
  .masthead-meta strong{ color:var(--ink); font-weight:600; }

  .kpis{
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:14px;
    margin:0 0 40px;
  }
  .kpi{
    background:var(--surface);
    border:1px solid var(--line);
    border-radius:3px;
    padding:16px 18px;
    box-shadow:var(--shadow);
    border-left:4px solid var(--accent);
  }
  .kpi.bad{ border-left-color:var(--bad); }
  .kpi.good{ border-left-color:var(--good); }
  .kpi.neutral{ border-left-color:var(--ink-faint); }
  .kpi-label{
    font-size:10.5px;
    letter-spacing:.1em;
    text-transform:uppercase;
    color:var(--ink-faint);
    font-weight:600;
    margin:0 0 8px;
  }
  .kpi-value{
    font-family:"IBM Plex Mono",monospace;
    font-size:27px;
    font-weight:600;
    color:var(--ink);
    margin:0 0 4px;
  }
  .kpi.bad .kpi-value{ color:var(--bad); }
  .kpi.good .kpi-value{ color:var(--good); }
  .kpi-sub{ font-size:12.5px; color:var(--ink-soft); }

  section{ margin-bottom:40px; }
  .sec-head{
    display:flex;
    align-items:baseline;
    gap:10px;
    margin-bottom:14px;
    padding-bottom:8px;
    border-bottom:1px solid var(--line);
  }
  .sec-num{
    font-family:"IBM Plex Mono",monospace;
    font-size:12px;
    color:var(--accent);
    font-weight:600;
  }
  h2{
    font-family:"Archivo", sans-serif;
    font-weight:700;
    font-size:19px;
    letter-spacing:-.005em;
    margin:0;
    color:var(--ink);
  }
  h3{
    font-family:"Archivo", sans-serif;
    font-weight:700;
    font-size:14.5px;
    margin:22px 0 10px;
    color:var(--ink);
  }
  p{ margin:0 0 12px; max-width:72ch; color:var(--ink); }
  p.lede{ color:var(--ink-soft); }
  .note{
    font-size:12.5px;
    color:var(--ink-faint);
    font-style:italic;
    max-width:72ch;
  }

  .callout{
    background:var(--accent-soft);
    border:1px solid var(--line);
    border-left:3px solid var(--accent);
    border-radius:3px;
    padding:16px 20px;
    margin-bottom:8px;
  }
  .callout ul{ margin:0; padding-left:20px; }
  .callout li{ margin-bottom:7px; font-size:14px; }
  .callout li:last-child{ margin-bottom:0; }
  .callout code{
    font-family:"IBM Plex Mono",monospace;
    background:var(--surface);
    border:1px solid var(--line);
    padding:1px 5px;
    border-radius:2px;
    font-size:12.5px;
  }

  .tbl-wrap{ overflow-x:auto; border:1px solid var(--line); border-radius:3px; box-shadow:var(--shadow); }
  table{ width:100%; border-collapse:collapse; font-size:13.5px; background:var(--surface); }
  thead th{
    background:var(--ink);
    color:var(--paper);
    text-align:left;
    font-weight:600;
    font-size:11px;
    letter-spacing:.05em;
    text-transform:uppercase;
    padding:10px 14px;
    white-space:nowrap;
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]) thead th{ background:var(--surface-2); color:var(--ink); }
  }
  :root[data-theme="dark"] thead th{ background:var(--surface-2); color:var(--ink); }
  tbody td{
    padding:9px 14px;
    border-top:1px solid var(--line);
    color:var(--ink);
  }
  tbody tr:nth-child(even){ background:var(--surface-2); }
  tbody tr.total td{ font-weight:700; border-top:2px solid var(--ink); background:var(--surface); }
  td.num, th.num{ text-align:right; font-family:"IBM Plex Mono",monospace; font-variant-numeric:tabular-nums; }
  td.name{ font-weight:600; }

  .pill{
    display:inline-flex; align-items:center; gap:6px;
    font-family:"IBM Plex Mono",monospace;
    font-weight:600; font-size:13px;
    padding:2px 0;
  }
  .dot{ width:7px; height:7px; border-radius:50%; flex:none; }
  .pill.good{ color:var(--good); } .pill.good .dot{ background:var(--good); }
  .pill.warn{ color:var(--warn); } .pill.warn .dot{ background:var(--warn); }
  .pill.bad{ color:var(--bad); } .pill.bad .dot{ background:var(--bad); }

  .chart-card{
    background:var(--surface); border:1px solid var(--line); border-radius:3px;
    padding:22px 24px 16px; box-shadow:var(--shadow); margin-bottom:18px;
  }
  .bar-row{ display:grid; grid-template-columns:150px 1fr 64px; align-items:center; gap:14px; margin-bottom:16px; }
  .bar-row:last-child{ margin-bottom:0; }
  .bar-label{ font-size:13px; font-weight:600; color:var(--ink); }
  .bar-label small{ display:block; font-weight:400; color:var(--ink-faint); font-size:11px; font-family:"IBM Plex Mono",monospace; }
  .bar-track{ position:relative; height:20px; background:var(--surface-2); border-radius:2px; overflow:visible; }
  .bar-fill{ position:absolute; left:0; top:0; height:100%; border-radius:2px; background:var(--bad); }
  .bar-goal{ position:absolute; top:-5px; bottom:-5px; width:2px; background:var(--brand-orange); }
  .bar-goal::after{
    content:"meta __META_TXT__%"; position:absolute; top:-17px; left:50%; transform:translateX(-50%);
    font-size:9.5px; font-family:"IBM Plex Mono",monospace; color:var(--ink-faint); white-space:nowrap;
  }
  .bar-val{ font-family:"IBM Plex Mono",monospace; font-weight:600; text-align:right; color:var(--bad); font-size:14px; }

  .causes{ display:flex; flex-direction:column; gap:10px; }
  .cause-row{ display:grid; grid-template-columns:1fr 220px 50px; align-items:center; gap:12px; font-size:13.5px; }
  .cause-track{ height:14px; background:var(--surface-2); border-radius:2px; overflow:hidden; border:1px solid var(--line); }
  .cause-fill{ height:100%; background:var(--accent); }
  .cause-pct{ font-family:"IBM Plex Mono",monospace; text-align:right; color:var(--ink-soft); }

  .trend-card{ background:var(--surface); border:1px solid var(--line); border-radius:3px; padding:20px 24px; box-shadow:var(--shadow); }
  .trend-svg{ width:100%; height:auto; display:block; }

  .crew-grid{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:16px; }
  @media (max-width:820px){ .crew-grid{ grid-template-columns:1fr; } .kpis{ grid-template-columns:1fr 1fr; } }

  ol.reco{ margin:0; padding-left:0; list-style:none; counter-reset:reco; }
  ol.reco li{
    counter-increment:reco;
    display:grid; grid-template-columns:30px 1fr; gap:12px;
    padding:12px 0; border-top:1px solid var(--line);
  }
  ol.reco li:first-child{ border-top:none; }
  ol.reco li::before{
    content:counter(reco);
    font-family:"IBM Plex Mono",monospace; font-weight:600; color:var(--accent);
    font-size:14px;
  }
  ol.reco li p{ margin:0; }
  ol.reco li strong{ color:var(--ink); }

  footer{
    margin-top:52px; padding-top:18px; border-top:1px solid var(--line);
    font-size:11.5px; color:var(--ink-faint); display:flex; justify-content:space-between; gap:20px; flex-wrap:wrap;
  }

  @media print{
    body{ background:#fff; }
    .sheet{ padding:0; max-width:none; }
    section{ break-inside:avoid; }
    .tbl-wrap, .chart-card, .trend-card{ box-shadow:none; }
  }
</style>
</head>
<body>
<div class="sheet">

  <div class="masthead">
    <div class="masthead-brand">
      <img class="brand-logo" src="data:image/png;base64,__LOGO_B64__" alt="Elecnor Chile">
      <div class="masthead-divider"></div>
      <div>
        <p class="eyebrow">Informe de gestión operativa &middot; P31 &middot; Actualización automática</p>
        <h1 class="title">Velocidad de Reparaciones</h1>
        <div class="brand-accent"></div>
      </div>
    </div>
    <div class="masthead-meta">
      <div><strong>Período</strong> __PERIODO__</div>
      <div><strong>Alcance</strong> __ALCANCE__</div>
      <div><strong>Meta</strong> __META_TXT__% de cumplimiento</div>
    </div>
  </div>

  <div class="kpis">
__KPI_CARDS__
  </div>

  <section>
    <div class="sec-head"><span class="sec-num">01</span><h2>Metodología y consideraciones</h2></div>
    <p class="lede">El indicador se recalculó a partir del archivo base <code style="font-family:'IBM Plex Mono',monospace;font-size:13px;">__NOMBRE_ARCHIVO__</code> aplicando estrictamente los criterios oficiales del reporte P31.</p>
    <div class="callout">
      <ul>
        <li><strong>Fórmula:</strong> cierres con <code>q_dias_efectivos &lt; 2</code> (es decir, &le; 1 día efectivo) &divide; total de reparaciones del mes.</li>
        <li><strong>Acceso:</strong> solo <code>tecno_acceso = FO</code> (fibra óptica).</li>
        <li><strong>Ámbito:</strong> solo <code>ambito = ATENCION CLIENTES</code>.</li>
        <li><strong>Fallas masivas:</strong> se excluyen los registros con <code>es_falla_masiva = 1</code>.</li>
      </ul>
    </div>
    <p class="note">__NOTA_EXCLUSION__</p>
  </section>

  <section>
    <div class="sec-head"><span class="sec-num">02</span><h2>Resultado por agencia vs. meta</h2></div>
    <p class="lede">__LEDE_AGENCIAS__</p>

    <div class="chart-card">
__BARRAS_AGENCIA__    </div>

    <div class="tbl-wrap">
      <table>
        <thead><tr>
          <th>Agencia</th><th class="num">Total</th><th class="num">Cumple</th><th class="num">% cumplimiento</th><th class="num">Brecha (pts)</th><th class="num">Casos para meta</th>
        </tr></thead>
        <tbody>
__FILAS_AGENCIA__          <tr class="total"><td class="name">Total</td><td class="num">__TOTAL_GLOBAL__</td><td class="num">__CUMPLE_GLOBAL__</td><td class="num">__PCT_GLOBAL__%</td><td class="num">__BRECHA_GLOBAL__</td><td class="num">__FALTAN_GLOBAL__</td></tr>
        </tbody>
      </table>
    </div>
  </section>

  <section>
    <div class="sec-head"><span class="sec-num">03</span><h2>Causas del incumplimiento</h2></div>
    <p class="lede">Sobre los __N_INCUMPLE__ casos que no cumplen el estándar de &le; 1 día efectivo, así se distribuye la causa raíz según el área responsable del quiebre:</p>

    <div class="chart-card">
      <div class="causes">
__CAUSAS_ROWS__      </div>
    </div>
    <p>__TEXTO_CAUSAS__</p>
  </section>

  <section>
    <div class="sec-head"><span class="sec-num">04</span><h2>Detalle por bucket</h2></div>
    <p class="lede">__LEDE_BUCKETS__</p>

    <div class="crew-grid">
__BUCKETS_GRID__    </div>
  </section>

  <section>
    <div class="sec-head"><span class="sec-num">05</span><h2>Tendencia diaria</h2></div>
    <p class="lede">__LEDE_TENDENCIA__</p>

    <div class="trend-card">
      <svg class="trend-svg" viewBox="0 0 860 260" preserveAspectRatio="xMidYMid meet" role="img" aria-label="Cumplimiento diario frente a la meta">
        <g font-family="IBM Plex Mono, monospace" font-size="10">
          <line x1="40" y1="20" x2="840" y2="20" stroke="var(--line)" stroke-width="1"/>
          <text x="30" y="24" text-anchor="end" fill="var(--ink-faint)">100</text>
          <line x1="40" y1="72" x2="840" y2="72" stroke="var(--line)" stroke-width="1"/>
          <text x="30" y="76" text-anchor="end" fill="var(--ink-faint)">75</text>
          <line x1="40" y1="124" x2="840" y2="124" stroke="var(--line)" stroke-width="1"/>
          <text x="30" y="128" text-anchor="end" fill="var(--ink-faint)">50</text>
          <line x1="40" y1="176" x2="840" y2="176" stroke="var(--line)" stroke-width="1"/>
          <text x="30" y="180" text-anchor="end" fill="var(--ink-faint)">25</text>
          <line x1="40" y1="228" x2="840" y2="228" stroke="var(--line)" stroke-width="1"/>
          <text x="30" y="232" text-anchor="end" fill="var(--ink-faint)">0</text>
        </g>
        <line x1="40" y1="__GOAL_Y__" x2="840" y2="__GOAL_Y__" stroke="var(--brand-orange)" stroke-width="1.5" stroke-dasharray="4 4"/>
        <text x="844" y="__GOAL_Y_LABEL__" font-family="IBM Plex Mono, monospace" font-size="10" fill="var(--brand-orange)">meta __META_TXT__%</text>

        <polyline fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round" points="__POLY_POINTS__"/>
        <g fill="var(--accent)">
__CIRCLES__        </g>
        <g font-family="IBM Plex Mono, monospace" font-size="10" fill="var(--ink-faint)">
__DAY_LABELS__        </g>
      </svg>
    </div>
    <p class="note">__NOTA_TENDENCIA__</p>
  </section>

  <section>
    <div class="sec-head"><span class="sec-num">06</span><h2>Conclusiones y recomendaciones</h2></div>
    <ol class="reco">
__RECO_ITEMS__    </ol>
  </section>

  <footer>
    <span>Fuente: __NOMBRE_ARCHIVO__ &middot; Elecnor &middot; universo __TOTAL_GLOBAL__ de __TOTAL_RAW__ registros</span>
    <span>Informe generado el __FECHA_GENERACION__</span>
  </footer>

</div>
</body>
</html>
'@

$html = $template
$html = $html.Replace('__LOGO_B64__', $logoB64)
$html = $html.Replace('__PERIODO__', $periodo)
$html = $html.Replace('__ALCANCE__', $alcance)
$html = $html.Replace('__META_TXT__', $metaTxt)
$html = $html.Replace('__KPI_CARDS__', $kpiCards)
$html = $html.Replace('__NOMBRE_ARCHIVO__', $csv.Name)
$html = $html.Replace('__NOTA_EXCLUSION__', $notaExclusion)
$html = $html.Replace('__LEDE_AGENCIAS__', $ledeAgencias)
$html = $html.Replace('__BARRAS_AGENCIA__', $barrasAgencia)
$html = $html.Replace('__FILAS_AGENCIA__', $filasAgencia)
$html = $html.Replace('__TOTAL_GLOBAL__', (FmtInt $totalGlobal))
$html = $html.Replace('__CUMPLE_GLOBAL__', (FmtInt $cumpleGlobal))
$html = $html.Replace('__PCT_GLOBAL__', (Fmt1 $pctGlobal))
$html = $html.Replace('__BRECHA_GLOBAL__', $brechaGlobalCelda)
$html = $html.Replace('__FALTAN_GLOBAL__', $faltanGlobalTxt)
$html = $html.Replace('__N_INCUMPLE__', (FmtInt $nInc))
$html = $html.Replace('__CAUSAS_ROWS__', $causasRows)
$html = $html.Replace('__TEXTO_CAUSAS__', $textoCausas)
$html = $html.Replace('__LEDE_BUCKETS__', $ledeBuckets)
$html = $html.Replace('__BUCKETS_GRID__', $bucketsGrid)
$html = $html.Replace('__LEDE_TENDENCIA__', $ledeTendencia)
$html = $html.Replace('__GOAL_Y__', (Fmt1 $goalY).Replace(',','.'))
$html = $html.Replace('__GOAL_Y_LABEL__', (Fmt1 ($goalY+4)).Replace(',','.'))
$html = $html.Replace('__POLY_POINTS__', $polyPoints)
$html = $html.Replace('__CIRCLES__', $circles)
$html = $html.Replace('__DAY_LABELS__', $dayLabels)
$html = $html.Replace('__NOTA_TENDENCIA__', $notaTendencia)
$html = $html.Replace('__RECO_ITEMS__', $recoItems)
$html = $html.Replace('__TOTAL_RAW__', (FmtInt $raw.Count))
$html = $html.Replace('__FECHA_GENERACION__', $fechaGeneracion)

[System.IO.File]::WriteAllText($HtmlOutput, $html, $utf8NoBom)
Write-Host "HTML actualizado: $HtmlOutput" -ForegroundColor Green
Write-Host "Cumplimiento global: $(Fmt1 $pctGlobal)% (meta $metaTxt%)"

# ---------------- 8. Commit y push ----------------
Push-Location $RepoPath
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try {
    git add index.html 2>&1 | Out-Null
    $changes = git status --porcelain
    if ([string]::IsNullOrWhiteSpace($changes)) {
        Write-Host "Sin cambios que publicar en GitHub." -ForegroundColor Yellow
    } else {
        $hoy = Get-Date -Format "dd-MM-yyyy"
        git commit -m "Actualizacion automatica $hoy - fuente $($csv.Name)" 2>&1 | Out-String | Write-Host
        git push origin main 2>&1 | Out-String | Write-Host
        Write-Host "Publicado: https://xarancibias.github.io/velocidad-reparaciones-elecnor/" -ForegroundColor Green
    }
} finally {
    $ErrorActionPreference = $prevEAP
    Pop-Location
}

Write-Host "=== Listo ===" -ForegroundColor Cyan
