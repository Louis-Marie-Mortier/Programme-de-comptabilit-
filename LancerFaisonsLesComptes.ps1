$ErrorActionPreference = 'Stop'

$nomTache = 'Faisons les comptes'
$heureRappel = 8
$etape = 'lecture de la tâche planifiée'

try {
    $tache = Get-ScheduledTask -TaskName $nomTache -ErrorAction Stop

    if ($tache.State -eq 'Running') {
        throw "La tâche est déjà en cours. Réessayez lorsqu'elle sera terminée."
    }

    $declencheurs = @($tache.Triggers)

    if ($declencheurs.Count -ne 1) {
        throw 'La tâche doit avoir exactement un déclencheur calendrier pour être reprogrammée.'
    }

    if ($declencheurs[0].CimClass.CimClassName -ne 'MSFT_TaskDailyTrigger') {
        throw "Le déclencheur de la tâche n'est pas un déclencheur quotidien."
    }

    $intervalleJours = 0
    if (
        $null -eq $declencheurs[0].DaysInterval -or
        -not [int]::TryParse(
            [string]$declencheurs[0].DaysInterval,
            [ref]$intervalleJours
        ) -or
        $intervalleJours -lt 1 -or
        $intervalleJours -gt 90
    ) {
        throw "L'intervalle quotidien configuré pour la tâche est invalide."
    }

    $prochainRappel = [datetime]::Today.AddDays($intervalleJours).AddHours($heureRappel)
    $ancienneDateDebut = $declencheurs[0].StartBoundary
    if ([string]::IsNullOrWhiteSpace($ancienneDateDebut)) {
        throw 'Le déclencheur de la tâche ne contient pas de date de début.'
    }

    $declencheurs[0].StartBoundary = $prochainRappel.ToString(
        "yyyy-MM-dd'T'HH:mm:ss",
        [Globalization.CultureInfo]::InvariantCulture
    )

    $etape = 'mise à jour du déclencheur de la tâche'
    Set-ScheduledTask -InputObject $tache -ErrorAction Stop | Out-Null

    try {
        $etape = 'démarrage de la tâche planifiée'
        Start-ScheduledTask -TaskName $nomTache -ErrorAction Stop
    }
    catch {
        $erreurDemarrage = $_
        try {
            $declencheurs[0].StartBoundary = $ancienneDateDebut
            Set-ScheduledTask -InputObject $tache -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Le démarrage a échoué et l'ancien déclencheur n'a pas pu être restauré : $($_.Exception.Message)"
        }

        throw $erreurDemarrage
    }

    Write-Host "La tâche « $nomTache » a été lancée." -ForegroundColor Green
    Write-Host (
        'Le prochain rappel est prévu le {0} à {1} h ({2} jours).' -f
        $prochainRappel.ToString('dd/MM/yyyy'),
        $heureRappel,
        $intervalleJours
    )
}
catch {
    Write-Host "Impossible de terminer l'etape « $etape » pour la tâche « $nomTache »." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
