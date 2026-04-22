$DossierRapports = "$env:USERPROFILE\Documents\SantePC"
if (-not (Test-Path $DossierRapports)) {
    New-Item -ItemType Directory -Path $DossierRapports -Force | Out-Null
}

$Date = Get-Date
$Horodatage = $Date.ToString("yyyy-MM-dd_HH-mm")
$FichierRapport = Join-Path $DossierRapports "rapport-sante_$Horodatage.html"
$DateAffichee = $Date.ToString("dddd dd MMMM yyyy - HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)

Write-Host "Collecte des informations en cours..." -ForegroundColor Cyan

Write-Host "  - Disques" -ForegroundColor Gray
$Disques = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    $totalGo = [math]::Round($_.Size / 1GB, 1)
    $libreGo = [math]::Round($_.FreeSpace / 1GB, 1)
    $utilGo  = [math]::Round($totalGo - $libreGo, 1)
    $pct     = if ($totalGo -gt 0) { [math]::Round(($utilGo / $totalGo) * 100, 1) } else { 0 }
    [PSCustomObject]@{
        Lettre   = $_.DeviceID
        Nom      = if ($_.VolumeName) { $_.VolumeName } else { "(sans nom)" }
        TotalGo  = $totalGo
        LibreGo  = $libreGo
        UtilGo   = $utilGo
        Pct      = $pct
    }
}

Write-Host "  - Memoire RAM" -ForegroundColor Gray
$Os = Get-CimInstance Win32_OperatingSystem
$RamTotalGo    = [math]::Round($Os.TotalVisibleMemorySize / 1MB, 1)
$RamLibreGo    = [math]::Round($Os.FreePhysicalMemory / 1MB, 1)
$RamUtilGo     = [math]::Round($RamTotalGo - $RamLibreGo, 1)
$RamPct        = if ($RamTotalGo -gt 0) { [math]::Round(($RamUtilGo / $RamTotalGo) * 100, 1) } else { 0 }
$ModulesRam    = Get-CimInstance Win32_PhysicalMemory | ForEach-Object {
    [PSCustomObject]@{
        Emplacement = $_.DeviceLocator
        Capacite    = [math]::Round($_.Capacity / 1GB, 0)
        Vitesse     = $_.Speed
        Fabricant   = $_.Manufacturer
    }
}

Write-Host "  - Batterie" -ForegroundColor Gray
$Batterie = $null
$InfosBat = $null
try {
    $InfosBat = Get-CimInstance Win32_Battery -ErrorAction Stop
} catch {}

if ($InfosBat) {
    $xmlPath = Join-Path $DossierRapports "battery-report.xml"
    try {
        powercfg /batteryreport /output $xmlPath /xml 2>&1 | Out-Null
        Start-Sleep -Milliseconds 800
        if (Test-Path $xmlPath) {
            [xml]$xmlBat = Get-Content $xmlPath -Encoding UTF8
            $bat = $xmlBat.BatteryReport.Batteries.Battery
            if ($bat -is [array]) { $bat = $bat[0] }

            $designCap = [int]$bat.DesignCapacity
            $fullCap   = [int]$bat.FullChargeCapacity
            $cycles    = [int]$bat.CycleCount
            $sante     = if ($designCap -gt 0) { [math]::Round(($fullCap / $designCap) * 100, 1) } else { 0 }

            $Batterie = [PSCustomObject]@{
                Fabricant           = $bat.Manufacturer
                Modele              = $bat.Id
                Chimie              = $bat.Chemistry
                CapaciteDesigneMWh  = $designCap
                CapaciteActuelleMWh = $fullCap
                Cycles              = $cycles
                SantePct            = $sante
                NiveauPct           = $InfosBat.EstimatedChargeRemaining
                EnCharge            = ($InfosBat.BatteryStatus -eq 2)
            }
        }
    } catch {
        $Batterie = [PSCustomObject]@{
            Fabricant           = "N/A"
            Modele              = "N/A"
            Chimie              = "N/A"
            CapaciteDesigneMWh  = 0
            CapaciteActuelleMWh = 0
            Cycles              = 0
            SantePct            = 0
            NiveauPct           = $InfosBat.EstimatedChargeRemaining
            EnCharge            = ($InfosBat.BatteryStatus -eq 2)
        }
    }
}

Write-Host "  - Mises a jour Windows (peut prendre 30-60s)" -ForegroundColor Gray
$Updates = @()
try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $resultat = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    $Updates  = @($resultat.Updates | ForEach-Object {
        [PSCustomObject]@{
            Titre    = $_.Title
            Severite = if ($_.MsrcSeverity) { $_.MsrcSeverity } else { "Standard" }
            TailleMo = [math]::Round(($_.MaxDownloadSize / 1MB), 1)
        }
    })
} catch {
    Write-Host "    (impossible d'interroger Windows Update)" -ForegroundColor Yellow
}

Write-Host "  - Programmes au demarrage" -ForegroundColor Gray
$ProgDemarrage = @()

$ProgDemarrage += Get-CimInstance Win32_StartupCommand | ForEach-Object {
    [PSCustomObject]@{
        Nom         = $_.Name
        Commande    = $_.Command
        Source      = $_.Location
        Utilisateur = $_.User
    }
}

$clesRegistre = @(
    @{ Chemin = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; Source = "Registre HKLM" },
    @{ Chemin = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; Source = "Registre HKCU" },
    @{ Chemin = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"; Source = "Registre HKLM (32 bits)" }
)
foreach ($cle in $clesRegistre) {
    if (Test-Path $cle.Chemin) {
        $props = Get-ItemProperty -Path $cle.Chemin
        $props.PSObject.Properties | Where-Object {
            $_.Name -notmatch '^PS' -and $_.Name -ne '(default)'
        } | ForEach-Object {
            $ProgDemarrage += [PSCustomObject]@{
                Nom         = $_.Name
                Commande    = $_.Value
                Source      = $cle.Source
                Utilisateur = ""
            }
        }
    }
}

$ProgDemarrage = $ProgDemarrage | Sort-Object Nom -Unique

Write-Host "Generation du rapport HTML..." -ForegroundColor Cyan

function Get-ClasseJauge {
    param([double]$Pct)
    if ($Pct -ge 90) { return "danger" }
    if ($Pct -ge 75) { return "warn" }
    return "ok"
}

$disquesHtml = ""
foreach ($d in $Disques) {
    $classe = Get-ClasseJauge -Pct $d.Pct
    $disquesHtml += @"
      <div class="card">
        <div class="row">
          <div class="title">$($d.Lettre) $($d.Nom)</div>
          <div class="value">$($d.UtilGo) / $($d.TotalGo) Go</div>
        </div>
        <div class="bar"><div class="fill $classe" style="width:$($d.Pct)%"></div></div>
        <div class="sub">Libre : $($d.LibreGo) Go - Utilise a $($d.Pct)%</div>
      </div>
"@
}

$ramClasse = Get-ClasseJauge -Pct $RamPct
$modulesHtml = ""
foreach ($m in $ModulesRam) {
    $modulesHtml += "<tr><td>$($m.Emplacement)</td><td>$($m.Capacite) Go</td><td>$($m.Vitesse) MHz</td><td>$($m.Fabricant)</td></tr>"
}

$batterieHtml = ""
if ($Batterie) {
    $santeClasse = if ($Batterie.SantePct -lt 70) { "danger" } elseif ($Batterie.SantePct -lt 85) { "warn" } else { "ok" }
    $cyclesClasse = if ($Batterie.Cycles -gt 500) { "danger" } elseif ($Batterie.Cycles -gt 300) { "warn" } else { "ok" }
    $etatCharge = if ($Batterie.EnCharge) { "En charge" } else { "Sur batterie" }
    $batterieHtml = @"
    <section class="section">
      <h2>Batterie</h2>
      <div class="grid">
        <div class="stat"><div class="stat-label">Niveau actuel</div><div class="stat-value">$($Batterie.NiveauPct)%</div><div class="sub">$etatCharge</div></div>
        <div class="stat"><div class="stat-label">Sante</div><div class="stat-value badge-$santeClasse">$($Batterie.SantePct)%</div><div class="sub">$($Batterie.CapaciteActuelleMWh) / $($Batterie.CapaciteDesigneMWh) mWh</div></div>
        <div class="stat"><div class="stat-label">Cycles de charge</div><div class="stat-value badge-$cyclesClasse">$($Batterie.Cycles)</div><div class="sub">Typique : ~500 avant perte notable</div></div>
        <div class="stat"><div class="stat-label">Modele</div><div class="stat-value small">$($Batterie.Modele)</div><div class="sub">$($Batterie.Fabricant) - $($Batterie.Chimie)</div></div>
      </div>
    </section>
"@
} else {
    $batterieHtml = @"
    <section class="section">
      <h2>Batterie</h2>
      <div class="card"><div class="sub">Aucune batterie detectee (PC fixe ?)</div></div>
    </section>
"@
}

$updatesHtml = ""
if ($Updates.Count -eq 0) {
    $updatesHtml = '<div class="card ok-card"><div class="title">Systeme a jour</div><div class="sub">Aucune mise a jour en attente.</div></div>'
} else {
    $lignes = ""
    foreach ($u in $Updates) {
        $lignes += "<tr><td>$($u.Titre)</td><td>$($u.Severite)</td><td class='num'>$($u.TailleMo) Mo</td></tr>"
    }
    $updatesHtml = @"
    <div class="card warn-card"><div class="title">$($Updates.Count) mise(s) a jour en attente</div></div>
    <table class="table">
      <thead><tr><th>Titre</th><th>Severite</th><th>Taille</th></tr></thead>
      <tbody>$lignes</tbody>
    </table>
"@
}

$startupHtml = ""
if ($ProgDemarrage.Count -eq 0) {
    $startupHtml = '<div class="card"><div class="sub">Aucun programme au demarrage detecte.</div></div>'
} else {
    $lignes = ""
    foreach ($p in $ProgDemarrage) {
        $cmd = if ($p.Commande) { $p.Commande } else { "" }
        $cmd = $cmd -replace '<', '&lt;' -replace '>', '&gt;'
        $lignes += "<tr><td><strong>$($p.Nom)</strong></td><td class='mono small'>$cmd</td><td class='small'>$($p.Source)</td></tr>"
    }
    $startupHtml = @"
    <div class="card"><div class="title">$($ProgDemarrage.Count) programme(s) au demarrage</div></div>
    <table class="table">
      <thead><tr><th>Nom</th><th>Commande</th><th>Source</th></tr></thead>
      <tbody>$lignes</tbody>
    </table>
"@
}

$pcNom = $env:COMPUTERNAME
$utilisateur = $env:USERNAME

$html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Rapport sante PC - $pcNom</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #f4f6f8; color: #222; margin: 0; padding: 32px; }
  .wrap { max-width: 1100px; margin: 0 auto; }
  header { background: linear-gradient(135deg, #4f46e5, #7c3aed); color: white; padding: 32px; border-radius: 16px; margin-bottom: 24px; box-shadow: 0 4px 20px rgba(79,70,229,0.2); }
  header h1 { margin: 0 0 8px 0; font-size: 28px; }
  header .meta { opacity: 0.9; font-size: 14px; }
  .section { background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
  .section h2 { margin: 0 0 16px 0; font-size: 20px; color: #111; border-bottom: 2px solid #eef; padding-bottom: 8px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; }
  .card { background: #fafbfc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; margin-bottom: 12px; }
  .ok-card { border-left: 4px solid #10b981; }
  .warn-card { border-left: 4px solid #f59e0b; }
  .row { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px; }
  .title { font-weight: 600; font-size: 15px; }
  .value { font-size: 14px; color: #555; }
  .sub { font-size: 13px; color: #666; margin-top: 4px; }
  .bar { background: #e5e7eb; height: 10px; border-radius: 5px; overflow: hidden; }
  .fill { height: 100%; border-radius: 5px; transition: width 0.3s; }
  .fill.ok { background: #10b981; }
  .fill.warn { background: #f59e0b; }
  .fill.danger { background: #ef4444; }
  .stat { background: #fafbfc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px; }
  .stat-label { font-size: 12px; text-transform: uppercase; color: #6b7280; letter-spacing: 0.5px; }
  .stat-value { font-size: 24px; font-weight: 700; margin-top: 4px; }
  .stat-value.small { font-size: 14px; font-weight: 500; }
  .badge-ok { color: #10b981; }
  .badge-warn { color: #f59e0b; }
  .badge-danger { color: #ef4444; }
  .table { width: 100%; border-collapse: collapse; margin-top: 8px; }
  .table th { background: #f3f4f6; padding: 10px 12px; text-align: left; font-size: 13px; color: #374151; border-bottom: 1px solid #e5e7eb; }
  .table td { padding: 10px 12px; border-bottom: 1px solid #f3f4f6; font-size: 14px; vertical-align: top; }
  .table tr:hover td { background: #fafbfc; }
  .mono { font-family: 'Consolas', 'Courier New', monospace; color: #555; }
  .small { font-size: 12px; }
  .num { text-align: right; font-variant-numeric: tabular-nums; }
  footer { text-align: center; padding: 24px; color: #9ca3af; font-size: 12px; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Rapport sante PC</h1>
    <div class="meta">$pcNom - $utilisateur - $DateAffichee</div>
  </header>

  <section class="section">
    <h2>Espace disque</h2>
    $disquesHtml
  </section>

  <section class="section">
    <h2>Memoire RAM</h2>
    <div class="card">
      <div class="row">
        <div class="title">Utilisation</div>
        <div class="value">$RamUtilGo / $RamTotalGo Go</div>
      </div>
      <div class="bar"><div class="fill $ramClasse" style="width:$RamPct%"></div></div>
      <div class="sub">Libre : $RamLibreGo Go - Utilisee a $RamPct%</div>
    </div>
    <table class="table">
      <thead><tr><th>Emplacement</th><th>Capacite</th><th>Vitesse</th><th>Fabricant</th></tr></thead>
      <tbody>$modulesHtml</tbody>
    </table>
  </section>

  $batterieHtml

  <section class="section">
    <h2>Mises a jour Windows</h2>
    $updatesHtml
  </section>

  <section class="section">
    <h2>Programmes au demarrage</h2>
    $startupHtml
  </section>

  <footer>Rapport genere automatiquement - $DateAffichee</footer>
</div>
</body>
</html>
"@

$html | Out-File -FilePath $FichierRapport -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "   Rapport genere avec succes" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "   $FichierRapport" -ForegroundColor Yellow
Write-Host ""

Start-Process $FichierRapport
