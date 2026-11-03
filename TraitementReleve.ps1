param(
    [string]$EvenementSurveillanceReleveDemarree,
    [switch]$IgnorerVerificationNumeroCompteODS,
    [switch]$ModeTestSansOuvertureApplications
)

# ============================================================
# ReleveAuto.ps1
# Orchestre automatiquement le téléchargement et le traitement
# du relevé bancaire.
# ============================================================
Write-Host "Démarrage du traitement..."
$ErrorActionPreference = "Stop"
# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------

# Dossier du programme
$dossierProgramme = $PSScriptRoot

# Dossier comptabilité contenant le fichier ODS
$dossierComptabilite = Split-Path -Parent $dossierProgramme

# Dossier de téléchargement de l'utilisateur
$dossierTelechargements = Join-Path $env:USERPROFILE "Downloads"

# Adresse à ouvrir dans le navigateur
$urlBanque = "https://www.labanquepostale.fr/particulier/connexion-espace-client.html"

# Nom attendu du fichier bancaire (attention ce regex n'est valide que pour la banque postale)
# voir https://irem.univ-grenoble-alpes.fr/medias/fichier/40x2_1568734553070-pdf, page 16/25 (ou 30/39)
# Exemple : 1234567A1234567812345678.csv
$motifNumeroCompte = '\d{7}[A-HJ-NPR-Z]\d{3}'
$motifFichier = '^' + $motifNumeroCompte + '\d{13}\.csv$'

# Nombre de secondes entre deux vérifications
$delaiVerification = 1

# Nombre de fois où la taille du fichier doit rester identique
# pour considérer que le téléchargement est terminé.
$nbVerificationsStables = 2

# Fichier ODS à ouvrir après traitement
$fichierODS = Join-Path `
    $dossierComptabilite `
    "Relevés banquaires traités analysables.ods"

$fichierModeleODS = Join-Path `
    $dossierProgramme `
    "Modele.ods"

# ------------------------------------------------------------
# VÉRIFICATIONS PRÉLIMINAIRES
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "       MISE À JOUR DU RELEVÉ BANCAIRE" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $dossierTelechargements)) {
    Write-Host "ERREUR : dossier de téléchargement introuvable :" -ForegroundColor Red
    Write-Host $dossierTelechargements
    Read-Host "Appuyez sur Entrée pour quitter"
    exit 1
}


# ------------------------------------------------------------
# REPÉRAGE DES FICHIERS EXISTANTS
# ------------------------------------------------------------
# On mémorise les fichiers correspondant au motif AVANT
# d'ouvrir le navigateur.
#
# Ainsi, si un ancien relevé existe déjà dans Téléchargements,
# il ne sera pas pris pour le nouveau.

$fichiersAvant = @(
    Get-ChildItem `
        -Path $dossierTelechargements `
        -File `
        -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match $motifFichier
    } |
    Select-Object -ExpandProperty FullName
)

Write-Host "Fichiers bancaires déjà présents : $($fichiersAvant.Count)"
Write-Host ""


# ------------------------------------------------------------
# OUVERTURE DU SITE BANCAIRE
# ------------------------------------------------------------
if (-not $ModeTestSansOuvertureApplications) {
    Write-Host "Le navigateur va s'ouvrir pour accéder au site bancaire." -ForegroundColor Yellow
    Write-Host "Veuillez vous connecter et télécharger le relevé bancaire." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Ouverture du site bancaire..." -ForegroundColor Yellow
    
    Start-Process $urlBanque
    
    Write-Host "Navigateur ouvert." -ForegroundColor Green
}
else {
    Write-Host "Mode test sans pop-up activé. Le navigateur ne sera pas ouvert." -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Le script attend un fichier correspondant à :"
Write-Host "1234567K1234567812345678.csv" -ForegroundColor Cyan
Write-Host ""

if (-not [string]::IsNullOrWhiteSpace($EvenementSurveillanceReleveDemarree)) {
    $signalPret = [System.Threading.EventWaitHandle]::OpenExisting($EvenementSurveillanceReleveDemarree)
    try {
        [void]$signalPret.Set()
    }
    finally {
        $signalPret.Dispose()
    }
}

# ------------------------------------------------------------
# ATTENTE DU NOUVEAU FICHIER
# ------------------------------------------------------------

$fichierDetecte = $null

while ($null -eq $fichierDetecte) {

    $fichiersActuels = @(
        Get-ChildItem `
            -Path $dossierTelechargements `
            -File `
            -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -match $motifFichier
        }
    )

    foreach ($fichier in $fichiersActuels) {

        # On ignore les fichiers qui existaient déjà.
        if ($fichiersAvant -contains $fichier.FullName) {
            continue
        }

        $fichierDetecte = $fichier
        break
    }

    if ($null -eq $fichierDetecte) {
        Write-Host "." -NoNewline
        Start-Sleep -Seconds $delaiVerification
    }
}

Write-Host ""
Write-Host ""
Write-Host "Nouveau fichier détecté :" -ForegroundColor Green
Write-Host $fichierDetecte.FullName -ForegroundColor Cyan
Write-Host ""


# ------------------------------------------------------------
# ATTENTE DE LA FIN DU TÉLÉCHARGEMENT
# ------------------------------------------------------------
# On vérifie que la taille du fichier ne change plus pendant
# plusieurs vérifications consécutives.

Write-Host "Vérification de la fin du téléchargement..." -ForegroundColor Yellow

$taillePrecedente = -1
$nbStables = 0

while ($nbStables -lt $nbVerificationsStables) {

    Start-Sleep -Seconds $delaiVerification

    # On recharge les informations du fichier
    $fichierActuel = Get-Item `
        -LiteralPath $fichierDetecte.FullName `
        -ErrorAction SilentlyContinue

    if ($null -eq $fichierActuel) {
        Write-Host "Le fichier a disparu. Reprise de l'attente..." -ForegroundColor Yellow
        $fichierDetecte = $null
        break
    }

    $tailleActuelle = $fichierActuel.Length

    if ($tailleActuelle -eq $taillePrecedente) {
        $nbStables++
    }
    else {
        $nbStables = 0
    }

    $taillePrecedente = $tailleActuelle

    Write-Host "Taille : $tailleActuelle octets - stabilité $nbStables/$nbVerificationsStables"
}

if ($null -eq $fichierDetecte) {
    Write-Host "Le téléchargement doit être relancé." -ForegroundColor Red
    Read-Host "Appuyez sur Entrée pour quitter"
    exit 1
}

$Fichier = $fichierDetecte.FullName

Write-Host ""
Write-Host "Téléchargement terminé." -ForegroundColor Green
Write-Host ""


# ------------------------------------------------------------
# TRAITEMENT DU RELEVÉ
# ------------------------------------------------------------

Write-Host "Lancement du traitement du relevé..." -ForegroundColor Yellow
Write-Host ""


# ------------------------------------------------------------
# Lecture du CSV
# ------------------------------------------------------------
$lignes = Get-Content -Path $Fichier -Encoding Default


# ------------------------------------------------------------
# Vérification de l'en-tête
# ------------------------------------------------------------

$enteteAttendue = @(
    "^Numéro Compte   ;" + $motifNumeroCompte + "$"
    "^Type         ;[A-Z]{3}$"
    "^Compte tenu en  ;[a-z]*$"
    "^Date            ;\d{2}/\d{2}/\d{4}$"
    "^Solde \(EUROS\)   ;\d+,\d{2}$"
)

$enteteTrouvee = $lignes[0..4]

$enteteValide = $true

for ($i = 0; $i -lt 5; $i++) {
    if ($enteteTrouvee[$i] -notmatch $enteteAttendue[$i]) {
        $enteteValide = $false
    }
}

if (-not $enteteValide) {

    Write-Host ""
    Write-Host "ERREUR : l'en-tête du fichier ne correspond pas au format attendu."
    Write-Host ""

    Write-Host "En-tête attendu :"
    $enteteAttendue | ForEach-Object {
        Write-Host "  $_"
    }

    Write-Host ""
    Write-Host "En-tête trouvée :"
    $enteteTrouvee | ForEach-Object {
        Write-Host "  $_"
    }

    Write-Host ""
    exit 1
}

$dossier = Split-Path -Parent $Fichier
$nomFichier = [System.IO.Path]::GetFileNameWithoutExtension($Fichier)

# Récupération du numéro de compte dans l'en-tête.
$numeroCompte = $enteteTrouvee[0].Split(";")[1].Trim()

if (-not $nomFichier.StartsWith($numeroCompte)) {
    throw "Le numéro de compte indiqué dans le nom du fichier ne correspond pas à celui indiqué dans l'en-tête du CSV."
}

$transactions = $lignes[6..($lignes.Count - 1)] |
    ConvertFrom-Csv -Delimiter ";"

# ------------------------------------------------------------
# Traitement
# ------------------------------------------------------------

$resultat = foreach ($ligne in $transactions) {

    $libelle = $ligne.Libellé.Trim()

    # Valeurs par défaut
    $type = ""
    # $moyenPaiement = ""
    $dateOperation = ""
    $destinataire = ""


    # ========================================================
    # ACHAT CB
    # ========================================================

    if ($libelle -match '^ACHAT CB\s(.*?)\s(\d{2}\.\d{2}\.\d{2})\sCARTE NUMERO\s*\d{3}\s*$' -or 
	$libelle -match '^ACHAT CB\s(.*?)\s(\d{2}\.\d{2}\.\d{2})\sEUR\s*\d+,\d{2}\s*CARTE NO\s*\d{3}\s*$') {

        $type = "ACHAT CB"

        # Nom du commerçant
        $destinataire = $matches[1].Trim()

        # Date de l'opération CB
        $dateOperation = $matches[2]

        # Conversion 08.09.26 -> 08/09/2026
        $dateOperation = $dateOperation -replace '\.', '/'

        $dateOperation = $dateOperation.Substring(0,6) + "20" + $dateOperation.Substring(6,2)
    }


    # ========================================================
    # PRELEVEMENT
    # ========================================================

    elseif ($libelle -match '^(VIREMENT INSTANTANE|VIREMENT|PRELEVEMENT) DE\s+(.+)$') {

    $type = $matches[1]
    $destinataire = $matches[2].Trim()
    }

    # ========================================================
    # RETRAIT DAB
    # ========================================================

    elseif ($libelle -match '^CARTE X\d{4}\s*(\d{2}\/\d{2}\/\d{2}).*RETRAIT DAB (\w*)$') {

        $type = "RETRAIT DAB"
        $dateOperation = $matches[1]
        $dateOperation = $dateOperation.Substring(0,6) + "20" + $dateOperation.Substring(6,2)
        $destinataire = $matches[2].Trim()
    }


    # ========================================================
    # LIBELLE NON RECONNU
    # ========================================================

    else {

        # On conserve le libellé complet plutôt que de perdre
        # une information lorsqu'un nouveau format apparaît.
        $type = "INCONNU"
        $destinataire = $libelle
	Write-Host "Ligne non reconnue :"
	Write-Host $libelle
	Write-Host ""
    }


    # ========================================================
    # Création de la nouvelle ligne
    # ========================================================

    [PSCustomObject]@{
        Date                = $ligne.Date
        Type                = $type
        # "Moyen de paiement" = $moyenPaiement
        "Date opération"    = $dateOperation
        "Débiteur/Destinataire" = $destinataire
        "Montant(EUROS)"    = $ligne.'Montant(EUROS)'
    }
}

# ============================================================
# MISE A JOUR DU FICHIER ODS HISTORIQUE
# ============================================================

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ------------------------------------------------------------
# Fonctions utilitaires
# ------------------------------------------------------------

function Get-OdsCellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $officeUri = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    $dateValue = $Cell.GetAttribute("date-value", $officeUri)
    if (-not [string]::IsNullOrWhiteSpace($dateValue)) {
        $date = [datetime]::ParseExact(
            $dateValue,
            "yyyy-MM-dd",
            [Globalization.CultureInfo]::InvariantCulture
        )
        return $date.ToString(
            "dd/MM/yyyy",
            [Globalization.CultureInfo]::InvariantCulture
        )
    }

    $p = $Cell.SelectSingleNode("./text:p", $Ns)

    if ($null -eq $p) {
        return ""
    }

    return $p.InnerText.Trim()
}


function Get-OdsRowCells {
    param(
        [System.Xml.XmlElement]$Row,
        [System.Xml.XmlNamespaceManager]$Ns
    )

    return @(
        $Row.SelectNodes("./table:table-cell | ./table:covered-table-cell", $Ns)
    )
}


function Get-TransactionKey {
    param(
        $Transaction
    )

    $date = [string]$Transaction.Date
    $type = [string]$Transaction.Type
    $dateOperation = [string]$Transaction.'Date opération'
    $destinataire = [string]$Transaction.'Débiteur/Destinataire'
    $montant = [string]$Transaction.'Montant(EUROS)'

    # On normalise uniquement le montant afin que
    # 12,86 et 12.86 soient considérés comme identiques.
    $montant = $montant.Trim().Replace(" ", "").Replace(",", ".")

    return "$date|$type|$dateOperation|$destinataire|$montant"
}


Write-Host ""
Write-Host "Traitement terminé avec succès." -ForegroundColor Green
Write-Host ""

# ------------------------------------------------------------
# Vérification / création de l'ODS
# ------------------------------------------------------------

$nouvelODS = -not (Test-Path $fichierODS)

if ($nouvelODS) {
    Write-Host ""
    Write-Host "Le fichier comptable n'existe pas encore."
    Write-Host ""
    Write-Host "Un fichier ODS va être créé à partir du modèle."
    Write-Host "Vous pourrez ensuite le modifier librement dans LibreOffice."
    Write-Host ""

    if (-not (Test-Path $fichierModeleODS)) {
        Write-Host "ERREUR : le fichier modèle ODS est introuvable."
        Write-Host ""
        Write-Host "Fichier attendu :"
        Write-Host $fichierModeleODS
        Write-Host ""
	    pause
        exit 1
    }

    Copy-Item `
        -Path $fichierModeleODS `
        -Destination $fichierODS `
        -Force

    Write-Host "Fichier ODS créé avec succès."
    Write-Host ""
}

# ------------------------------------------------------------
# Lecture du contenu XML de l'ODS
# ------------------------------------------------------------

$archive = [System.IO.Compression.ZipFile]::OpenRead($fichierODS)

try {

    $entreeContent = $archive.GetEntry("content.xml")

    if ($null -eq $entreeContent) {
        throw "Le fichier ODS ne contient pas content.xml."
    }

    $stream = $entreeContent.Open()
    $reader = New-Object System.IO.StreamReader(
        $stream,
        [System.Text.Encoding]::UTF8
    )

    $contenuXml = $reader.ReadToEnd()

    $reader.Close()
    $stream.Close()

}
finally {
    $archive.Dispose()
}


$xml = New-Object System.Xml.XmlDocument
$xml.PreserveWhitespace = $true
$xml.LoadXml($contenuXml)

# ------------------------------------------------------------
# Espaces de noms XML
# ------------------------------------------------------------

$tableUri = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
$textUri = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"


$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)

$ns.AddNamespace("table", $tableUri)
$ns.AddNamespace("text", $textUri)



# ------------------------------------------------------------
# Recherche de la feuille "données brutes"
# ------------------------------------------------------------

$table = $xml.SelectSingleNode(
    "//table:table[@table:name='données brutes']",
    $ns
)

if ($null -eq $table) {
    throw "La feuille 'données brutes' est introuvable dans l'ODS."
}

# ------------------------------------------------------------
# Récupération de toutes les lignes historiques
# ------------------------------------------------------------

$rows = @(
    $table.SelectNodes("./table:table-row", $ns)
)

$indiceLignePremiereTransaction = 8
if ($rows.Count -lt $indiceLignePremiereTransaction) {
    throw "La feuille 'données brutes' ne contient pas de transactions."
}

$cellulesInformationsNumeroCompte = Get-OdsRowCells -Row $rows[0] -Ns $ns

if (-not $nouvelODS) {
    if (-not $IgnorerVerificationNumeroCompteODS -and
        (Get-OdsCellText $cellulesInformationsNumeroCompte[1] $ns) -ne $numeroCompte) {
        throw "Le numéro de compte indiqué dans le nom du fichier et l'en-tête du CSV ne correspond pas à celui de l'historique."
    }
}

# ------------------------------------------------------------
# Construction de l'ensemble des transactions historiques
# ------------------------------------------------------------

$clesHistoriques = [System.Collections.Generic.HashSet[string]]::new()

for ($i = $indiceLignePremiereTransaction;
    -not $nouvelODS -and $i -lt $rows.Count;
    $i++) {

    $cellules = Get-OdsRowCells `
        -Row $rows[$i] `
        -Ns $ns

    if ($cellules.Count -lt 5) {
        continue
    }

    # Vérification que la ligne contient bien une transaction.
    $date = Get-OdsCellText $cellules[0] $ns

    if ([string]::IsNullOrWhiteSpace($date)) {
        continue
    }

    $transactionHistorique = [PSCustomObject]@{
        Date = Get-OdsCellText $cellules[0] $ns
        Type = Get-OdsCellText $cellules[1] $ns
        "Date opération" = Get-OdsCellText $cellules[2] $ns
        "Débiteur/Destinataire" = Get-OdsCellText $cellules[3] $ns
        "Montant(EUROS)" = Get-OdsCellText $cellules[4] $ns
    }

    $cle = Get-TransactionKey $transactionHistorique

    [void]$clesHistoriques.Add($cle)
}


Write-Host ""
Write-Host "Transactions historiques analysées : $($clesHistoriques.Count)"
Write-Host ""


# ------------------------------------------------------------
# Recherche des nouvelles transactions
# ------------------------------------------------------------

$resultatArray = @($resultat)

$nouvellesTransactions = @()

$nombreDoublons = 0

foreach ($transaction in $resultatArray) {

    $cle = Get-TransactionKey $transaction

    if ($clesHistoriques.Contains($cle)) {

        # La transaction existe déjà dans l'historique.
        $nombreDoublons++

    }
    else {

        # Nouvelle transaction.
        $nouvellesTransactions += $transaction
    }
}


Write-Host "Transactions du nouveau fichier : $($resultatArray.Count)"
Write-Host "Doublons ignorés : $nombreDoublons"
Write-Host "Nouvelles transactions : $($nouvellesTransactions.Count)"
Write-Host ""


# ------------------------------------------------------------
# Rien à ajouter
# ------------------------------------------------------------

if ($nouvellesTransactions.Count -eq 0) {

    Write-Host "Aucune nouvelle transaction."
    Write-Host "L'ODS n'a pas besoin d'être modifié."
    Write-Host ""

}
else {


    $fichierODSTemp = $fichierODS + ".tmp"

    if (Test-Path $fichierODSTemp) {
        Remove-Item $fichierODSTemp -Force
    }

    Copy-Item -LiteralPath $fichierODS -Destination $fichierODSTemp

    $fichierAssistantOds = Join-Path $PSScriptRoot "AjouterTransactionsOds.py"
    if (-not (Test-Path -LiteralPath $fichierAssistantOds -PathType Leaf)) {
        throw "Le script LibreOffice est introuvable : $fichierAssistantOds"
    }

    $dossiersLibreOffice = @()
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
        $dossiersLibreOffice += Join-Path $env:ProgramFiles "LibreOffice\program"
    }
    if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
        $dossiersLibreOffice += Join-Path ${env:ProgramFiles(x86)} "LibreOffice\program"
    }

    $pythonLibreOffice = $dossiersLibreOffice |
        ForEach-Object { Join-Path $_ "python.exe" } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1

    if ([string]::IsNullOrWhiteSpace($pythonLibreOffice)) {
        throw "Python intégré à LibreOffice est introuvable. Vérifiez l'installation de LibreOffice."
    }

    $fichierTransactionsTemp = [System.IO.Path]::GetTempFileName()
    try {
        $transactionsJson = ConvertTo-Json `
            -InputObject @($nouvellesTransactions) `
            -Depth 4

        [System.IO.File]::WriteAllText(
            $fichierTransactionsTemp,
            $transactionsJson,
            [System.Text.UTF8Encoding]::new($false)
        )

        & $pythonLibreOffice `
            $fichierAssistantOds `
            $fichierODSTemp `
            $fichierTransactionsTemp `
            ([string]$nouvelODS) `
            $numeroCompte

        if ($LASTEXITCODE -ne 0) {
            throw "L'insertion des transactions par LibreOffice a échoué (code $LASTEXITCODE)."
        }
    }
    finally {
        if (Test-Path -LiteralPath $fichierTransactionsTemp) {
            Remove-Item -LiteralPath $fichierTransactionsTemp -Force
        }
    }

    # --------------------------------------------------------
    # Sauvegarde de sécurité de l'ancien ODS
    # --------------------------------------------------------

    $fichierODSBackup = $fichierODS + ".bak"

    Copy-Item `
        -Path $fichierODS `
        -Destination $fichierODSBackup `
        -Force


    # --------------------------------------------------------
    # Remplacement de l'ancien ODS
    # --------------------------------------------------------

    Remove-Item $fichierODS -Force

    Move-Item `
        -Path $fichierODSTemp `
        -Destination $fichierODS


    Write-Host "ODS mis à jour avec succès."
    Write-Host "Transactions ajoutées : $($nouvellesTransactions.Count)"
    Write-Host ""
    Write-Host "Sauvegarde de sécurité :"
    Write-Host $fichierODSBackup
    Write-Host ""
}

# ------------------------------------------------------------
# Création du fichier de sortie
# ------------------------------------------------------------

$fichierSortie = Join-Path $dossier ($nomFichier + "_traite.csv")

$resultat | Export-Csv `
    -Path $fichierSortie `
    -Delimiter ";" `
    -NoTypeInformation `
    -Encoding UTF8


Write-Host ""
Write-Host "Traitement terminé."
Write-Host "Fichier créé :"
Write-Host $fichierSortie
Write-Host ""





# ------------------------------------------------------------
# OUVERTURE DE L'ODS
# ------------------------------------------------------------

if (Test-Path $fichierODS) {
    if ($ModeTestSansOuvertureApplications) {
        Write-Host "Le fichier ODS a été mis à jour, mais il ne sera pas ouvert automatiquement." -ForegroundColor Yellow
        Write-Host "Vous pouvez l'ouvrir manuellement dans LibreOffice Calc si vous le souhaitez." -ForegroundColor Yellow
        Write-Host ""
    }
    else {
        Write-Host "Ouverture de LibreOffice Calc..." -ForegroundColor Yellow

        Start-Process $fichierODS
    }
    Write-Host ""
    Write-Host "Le relevé est maintenant prêt à être catégorisé." -ForegroundColor Green
    Write-Host ""

}
else {

    Write-Host "ATTENTION : fichier ODS introuvable :" -ForegroundColor Red
    Write-Host $fichierODS
    Write-Host ""
}


# ------------------------------------------------------------
# FIN
# ------------------------------------------------------------

Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "             TRAITEMENT TERMINÉ" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# Read-Host "Appuyez sur Entrée pour fermer"
exit 0