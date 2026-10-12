$ErrorActionPreference = 'Stop'

$dossierDeveloppement = $PSScriptRoot
$dossierComptabilite = Split-Path -Parent $dossierDeveloppement
$dossierTelechargements = Join-Path $env:USERPROFILE 'Downloads'

$fichierScript = Join-Path $dossierDeveloppement 'TraitementReleve.ps1'
$fichierTest = Join-Path $dossierDeveloppement 'Test.csv'
$fichierODS = Join-Path $dossierComptabilite 'Relevés banquaires traités analysables.ods'
$fichierResultat = Join-Path $dossierDeveloppement 'historique avec test.ods'

$processus = $null
$evenementSurveillanceReleveDemarree = $null
$fichierCSVInjecte = $null
$fichierSauvegarde = Join-Path $env:TEMP (
    'historique-comptabilite-' + [guid]::NewGuid().ToString('N') + '.ods'
)
$testReussi = $false
$restaurationReussie = $false

try {
    # Vérifications avant toute modification
    foreach ($chemin in @($fichierScript, $fichierTest, $fichierODS)) {
        if (-not (Test-Path -LiteralPath $chemin -PathType Leaf)) {
            throw "Fichier introuvable : $chemin"
        }
    }

    if (-not (Test-Path -LiteralPath $dossierTelechargements -PathType Container)) {
        throw "Dossier Téléchargements introuvable : $dossierTelechargements"
    }

    $premiereLigne = Get-Content -LiteralPath $fichierTest -Encoding Default -TotalCount 1
    $numeroCompte = $premiereLigne.Split(';', 2)[1].Trim()
    if ($numeroCompte -notmatch '^\d{7}[A-HJ-NPR-Z]\d{3}$') {
        throw "Le numéro de compte dans la première ligne de Test.csv n'a pas le format attendu : $numeroCompte"
    }

    # Génère un nom accepté par le script principal :
    # 7 chiffres + K + 16 chiffres, dont les 3 chiffres du numéro de compte.
    do {
        $suffixe = -join (1..13 | ForEach-Object { Get-Random -Minimum 0 -Maximum 10 })
        $nomCSV = $numeroCompte + $suffixe + '.csv'
        $fichierCSVInjecte = Join-Path $dossierTelechargements $nomCSV
    } while (Test-Path -LiteralPath $fichierCSVInjecte)

    Write-Host ''
    Write-Host '=== TEST DU TRAITEMENT DES RELEVÉS ===' -ForegroundColor Cyan
    Write-Host "Dossier de développement : $dossierDeveloppement"
    Write-Host "Historique réel           : $fichierODS"
    Write-Host "Résultat du test          : $fichierResultat"
    Write-Host ''
    Write-Host 'Ferme les documents ouverts dans LibreOffice, notamment l''historique et un ancien résultat de test.'
    [void](Read-Host 'Quand c''est fait, appuie sur Entrée pour démarrer')

    # Sauvegarde exacte de l'historique réel avant de lancer le traitement.
    Copy-Item -LiteralPath $fichierODS -Destination $fichierSauvegarde

    $nomEvenementSurveillanceReleveDemarree = 'Local\SurveillanceReleveDemarree-' + [guid]::NewGuid().ToString('N')
    $evenementCree = $false
    $evenementSurveillanceReleveDemarree = [System.Threading.EventWaitHandle]::new(
        $false,
        [System.Threading.EventResetMode]::ManualReset,
        $nomEvenementSurveillanceReleveDemarree,
        [ref]$evenementCree
    )

    $powershellExe = Join-Path $PSHOME 'powershell.exe'
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $fichierScript +
        '" -EvenementSurveillanceReleveDemarree "' + $nomEvenementSurveillanceReleveDemarree +
        '" -IgnorerVerificationNumeroCompteODS -ModeTestSansOuvertureApplications'

    $demarrageProcessus = [System.Diagnostics.ProcessStartInfo]::new()
    $demarrageProcessus.FileName = $powershellExe
    $demarrageProcessus.Arguments = $arguments
    $demarrageProcessus.WorkingDirectory = $dossierDeveloppement
    $demarrageProcessus.UseShellExecute = $false
    $demarrageProcessus.CreateNoWindow = $false

    $processus = [System.Diagnostics.Process]::new()
    $processus.StartInfo = $demarrageProcessus
    if (-not $processus.Start()) {
        throw 'Impossible de démarrer TraitementReleve.ps1.'
    }

    Write-Host ''
    Write-Host 'TraitementReleve.ps1 s''exécute dans cette fenêtre PowerShell.' -ForegroundColor Yellow
    Write-Host 'Attente du démarrage de la surveillance des relevés...'

    do {
        $signalRecu = $evenementSurveillanceReleveDemarree.WaitOne(250)
        $processus.Refresh()
        if ($processus.HasExited) {
            throw 'TraitementReleve.ps1 s''est terminé avant la copie du fichier de test. Aucun CSV de test n''a été injecté.'
        }
    } until ($signalRecu)

    Write-Host 'La surveillance est lancée.' -ForegroundColor Green

    $processus.Refresh()
    if ($processus.HasExited) {
        throw 'TraitementReleve.ps1 s''est terminé avant la copie du fichier de test. Aucun CSV de test n''a été injecté.'
    }

    # Copie le CSV seulement après le démarrage de la surveillance.
    Copy-Item -LiteralPath $fichierTest -Destination $fichierCSVInjecte
    Write-Host "Fichier de test copié dans Téléchargements : $nomCSV" -ForegroundColor Green
    Write-Host ''
    Write-Host 'Le traitement continue dans cette fenêtre.'
    Write-Host 'Quand TraitementReleve.ps1 ouvre l''historique dans LibreOffice, ferme ce document, puis appuie sur Entrée ici.'
    Write-Host 'Le test attendra la fin du script.'

    $processus.WaitForExit()
    $processus.Refresh()

    if ($processus.ExitCode -ne 0) {
        throw "TraitementReleve.ps1 s'est terminé avec le code $($processus.ExitCode). Le résultat ne sera pas archivé."
    }

    if (-not (Test-Path -LiteralPath $fichierODS -PathType Leaf)) {
        throw "Le script s'est terminé sans laisser l'historique attendu : $fichierODS"
    }

    # N'archive le résultat qu'après une exécution réussie.
    Copy-Item -LiteralPath $fichierODS -Destination $fichierResultat -Force
    $testReussi = $true
    Write-Host ''
    Write-Host "Résultat de test enregistré : $fichierResultat" -ForegroundColor Green
}
catch {
    Write-Host ''
    Write-Host "ERREUR : $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # En cas d'erreur ou d'interruption, évite qu'un processus de test encore actif
    # ne modifie l'historique après sa restauration.
    if ($null -ne $processus) {
        try {
            $processus.Refresh()
            if (-not $processus.HasExited) {
                Stop-Process -Id $processus.Id -Force -ErrorAction SilentlyContinue
                Wait-Process -Id $processus.Id -Timeout 10 -ErrorAction SilentlyContinue
            }
        }
        catch {
            # La restauration ci-dessous reste prioritaire.
        }
    }

    if ($null -ne $evenementSurveillanceReleveDemarree) {
        $evenementSurveillanceReleveDemarree.Dispose()
    }

    # Restaure toujours l'historique initial si la sauvegarde a été créée.
    if (Test-Path -LiteralPath $fichierSauvegarde -PathType Leaf) {
        try {
            Copy-Item -LiteralPath $fichierSauvegarde -Destination $fichierODS -Force
            $restaurationReussie = $true
            Remove-Item -LiteralPath $fichierSauvegarde -Force
            Write-Host ''
            Write-Host 'Historique réel initial restauré.' -ForegroundColor Green
        }
        catch {
            Write-Host ''
            Write-Host 'ATTENTION : la restauration automatique a échoué.' -ForegroundColor Red
            Write-Host "La sauvegarde de sécurité est conservée ici : $fichierSauvegarde" -ForegroundColor Yellow
            Write-Host 'Ferme LibreOffice si l''historique est encore ouvert, puis restaure cette copie manuellement.'
        }
    }

    # Supprime le CSV injecté seulement si le test s'est terminé avec succès.
    if ($testReussi -and $null -ne $fichierCSVInjecte -and
        (Test-Path -LiteralPath $fichierCSVInjecte -PathType Leaf)) {
        Remove-Item -LiteralPath $fichierCSVInjecte -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ''
if ($testReussi -and $restaurationReussie) {
    Write-Host 'Test terminé : le résultat est conservé dans developpement et l''historique réel a été restauré.' -ForegroundColor Green
    #[void](Read-Host 'Appuie sur Entrée pour fermer')
    exit 0
}
else {
    Write-Host 'Test non confirmé comme réussi. Vérifie les messages ci-dessus et l''état de l''historique.' -ForegroundColor Red
    #[void](Read-Host 'Appuie sur Entrée pour fermer')
    exit 1
}
