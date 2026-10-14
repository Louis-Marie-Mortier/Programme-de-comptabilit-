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


function Set-OdsCellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [string]$Value,
        [System.Xml.XmlNamespaceManager]$Ns,
        [System.Xml.XmlDocument]$Xml
    )

    # Suppression des anciennes valeurs éventuelles
    $officeUri = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    $Cell.RemoveAttribute("value", $officeUri)
    $Cell.RemoveAttribute("date-value", $officeUri)
    $Cell.RemoveAttribute("time-value", $officeUri)
    $Cell.RemoveAttribute("boolean-value", $officeUri)
    $Cell.RemoveAttribute("string-value", $officeUri)

    [void]$Cell.SetAttribute(
        "value-type",
        $officeUri,
        "string"
    )

    # On supprime les anciens paragraphes
    $paragraphes = @(
        $Cell.SelectNodes("./text:p", $Ns)
    )

    foreach ($p in $paragraphes) {
        $Cell.RemoveChild($p) | Out-Null
    }

    # Création du nouveau paragraphe
    $p = $Xml.CreateElement(
        "text",
        "p",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    )

    $p.InnerText = $Value

    [void]$Cell.AppendChild($p)
}


function Set-OdsAmount {
    param(
        [System.Xml.XmlElement]$Cell,
        [string]$Value,
        [System.Xml.XmlNamespaceManager]$Ns,
        [System.Xml.XmlDocument]$Xml
    )

    $officeUri = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"

    $valeur = $Value.Trim().Replace(" ", "").Replace(",", ".")

    $nombre = 0.0

    if (-not [double]::TryParse(
        $valeur,
        [Globalization.NumberStyles]::Float,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$nombre
    )) {
        throw "Impossible de convertir le montant '$Value' en nombre."
    }

    # Suppression des anciennes valeurs
    $Cell.RemoveAttribute("value", $officeUri)
    $Cell.RemoveAttribute("date-value", $officeUri)
    $Cell.RemoveAttribute("time-value", $officeUri)
    $Cell.RemoveAttribute("boolean-value", $officeUri)
    $Cell.RemoveAttribute("string-value", $officeUri)

    [void]$Cell.SetAttribute(
        "value-type",
        $officeUri,
        "float"
    )

    [void]$Cell.SetAttribute(
        "value",
        $officeUri,
        $nombre.ToString(
            "0.################",
            [Globalization.CultureInfo]::InvariantCulture
        )
    )

    # Remplacement du texte affiché
    $paragraphes = @(
        $Cell.SelectNodes("./text:p", $Ns)
    )

    foreach ($p in $paragraphes) {
        $Cell.RemoveChild($p) | Out-Null
    }

    $p = $Xml.CreateElement(
        "text",
        "p",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    )

    $p.InnerText = $Value

    [void]$Cell.AppendChild($p)
}

function Set-OdsDate {
    param(
        [System.Xml.XmlElement]$Cell,
        [string]$Value,
        [System.Xml.XmlNamespaceManager]$Ns,
        [System.Xml.XmlDocument]$Xml
    )

    $officeUri = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    $calcextUri = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"

    # Supprime toutes les informations de valeur existantes
    $Cell.RemoveAttribute("value", $officeUri)
    $Cell.RemoveAttribute("date-value", $officeUri)
    $Cell.RemoveAttribute("time-value", $officeUri)
    $Cell.RemoveAttribute("boolean-value", $officeUri)
    $Cell.RemoveAttribute("string-value", $officeUri)
    $Cell.RemoveAttribute("value-type", $officeUri)

    # Attribut spécifique à LibreOffice
    $Cell.RemoveAttribute("value-type", $calcextUri)

    # Supprime les éventuels paragraphes existants
    $paragraphes = @(
        $Cell.SelectNodes("./text:p", $Ns)
    )

    foreach ($p in $paragraphes) {
        $Cell.RemoveChild($p) | Out-Null
    }

    # Si la date est vide : on laisse la cellule réellement vide
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return
    }

    # Vérifie et convertit la date
    $date = [datetime]::ParseExact(
        $Value.Trim(),
        "dd/MM/yyyy",
        [Globalization.CultureInfo]::InvariantCulture
    )

    # Crée une vraie cellule date ODS
    [void]$Cell.SetAttribute(
        "value-type",
        $officeUri,
        "date"
    )

    [void]$Cell.SetAttribute(
        "date-value",
        $officeUri,
        $date.ToString(
            "yyyy-MM-dd",
            [Globalization.CultureInfo]::InvariantCulture
        )
    )

    # Texte affiché dans Calc
    $p = $Xml.CreateElement(
        "text",
        "p",
        "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    )

    $p.InnerText = $date.ToString("dd/MM/yyyy")

    [void]$Cell.AppendChild($p)
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
$officeUri = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
$styleUri = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
$calcextUri = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"


$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)

$ns.AddNamespace("table", $tableUri)
$ns.AddNamespace("text", $textUri)
$ns.AddNamespace("office", $officeUri)
$ns.AddNamespace("style",$styleUri)
$ns.AddNamespace("calcext", $calcextUri)



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

$cellulesInformationsNumeroCompte = Get-OdsRowCells -Row $rows[0] -Ns $ns
$indiceLignePremiereTransaction = 8
$ligneModele = $rows[$indiceLignePremiereTransaction-1]
$ligneEntetes = $rows[$indiceLignePremiereTransaction-2]
if ($rows.Count -lt $indiceLignePremiereTransaction) {
    throw "La feuille 'données brutes' ne contient pas de transactions."
}

if ($nouvelODS) {
    # On supprime toutes les transactions du modèle
    for ($i = $rows.Count - 1; $i -ge $indiceLignePremiereTransaction; $i--) {
        $table.RemoveChild($rows[$i]) | Out-Null
    }
    Set-OdsCellText $cellulesInformationsNumeroCompte[1] $numeroCompte $ns $xml
}
else {
    if (-not $IgnorerVerificationNumeroCompteODS -and
        (Get-OdsCellText $cellulesInformationsNumeroCompte[1] $ns) -ne $numeroCompte) {
        throw "Le numéro de compte indiqué dans le nom du fichier et l'en-tête du CSV ne correspond pas à celui de l'historique."
    }
}

# ------------------------------------------------------------
# Construction de l'ensemble des transactions historiques
# ------------------------------------------------------------

$clesHistoriques = [System.Collections.Generic.HashSet[string]]::new()

for ($i = $indiceLignePremiereTransaction; $i -lt $rows.Count; $i++) {

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


    # --------------------------------------------------------
    # Insertion des nouvelles lignes
    # --------------------------------------------------------

    # On insère de la plus ancienne vers la plus récente.
    # Ainsi, chaque nouvelle ligne est placée avant la ligne
    # historique et l'ordre chronologique décroissant est conservé.

    for ($i = $nouvellesTransactions.Count - 1; $i -ge 0; $i--) {

        $transaction = $nouvellesTransactions[$i]
        # Copie complète de la première ligne historique.
        # Cela conserve notamment :
        # - le format des cellules
        # - le style
        # - la validation Catégorie
        # - la hauteur de ligne
        $nouvelleLigne = $ligneModele.CloneNode($true)

        $cellules = Get-OdsRowCells `
            -Row $nouvelleLigne `
            -Ns $ns

        if ($cellules.Count -lt 7) {
            throw "La ligne modèle ne contient pas 7 cellules."
        }


        # --------------------------------------------
        # Date
        # --------------------------------------------

        Set-OdsDate `
            -Cell $cellules[0] `
            -Value ([string]$transaction.Date) `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Type
        # --------------------------------------------

        Set-OdsCellText `
            -Cell $cellules[1] `
            -Value ([string]$transaction.Type) `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Date opération
        # --------------------------------------------

        Set-OdsDate `
            -Cell $cellules[2] `
            -Value ([string]$transaction.'Date opération') `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Débiteur / Destinataire
        # --------------------------------------------

        Set-OdsCellText `
            -Cell $cellules[3] `
            -Value ([string]$transaction.'Débiteur/Destinataire') `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Montant
        # --------------------------------------------

        Set-OdsAmount `
            -Cell $cellules[4] `
            -Value ([string]$transaction.'Montant(EUROS)') `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Catégorie
        # --------------------------------------------

        Set-OdsCellText `
            -Cell $cellules[5] `
            -Value "Non identifié" `
            -Ns $ns `
            -Xml $xml


        # --------------------------------------------
        # Tags
        # --------------------------------------------

        Set-OdsCellText `
            -Cell $cellules[6] `
            -Value "" `
            -Ns $ns `
            -Xml $xml

        # Insertion après la ligne des en-têtes
        [void]$table.InsertAfter(
            $nouvelleLigne,
            $ligneModele
        )
    }

    # --------------------------------------------------------
    # Allongement des plages de formatage conditionnel
    # --------------------------------------------------------

$colonnesAutorisees = @("A","B","C","D","E","F","G")
$derniereLigne = $rows.Count + $nouvellesTransactions.Count
$noeuds = @(
    $xml.SelectNodes(
        "//*[@calcext:target-range-address]",
        $ns
    )
)

foreach ($noeud in $noeuds) {

    $adresse = $noeud.GetAttribute(
        "target-range-address",
        $calcextUri
    )

    if ([string]::IsNullOrWhiteSpace($adresse)) {
        continue
    }

    foreach ($colonne in $colonnesAutorisees) {

        if ($adresse -match "^'données brutes'\.${colonne}6:'données brutes'\.${colonne}\d+$") {

            $nouvelleAdresse = $adresse -replace "\d+$", $derniereLigne

            [void]$noeud.SetAttribute(
                "target-range-address",
                $calcextUri,
                $nouvelleAdresse
            )

            break
        }
    }
}

    # --------------------------------------------------------
    # Sauvegarde XML temporaire
    # --------------------------------------------------------

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false

    $memoryStream = New-Object System.IO.MemoryStream

    $writer = [System.Xml.XmlWriter]::Create(
        $memoryStream,
        $settings
    )

    $xml.Save($writer)

    $writer.Flush()
    $writer.Close()

    $nouveauContentXml = $memoryStream.ToArray()

    $memoryStream.Close()


    # --------------------------------------------------------
    # Création d'un nouvel ODS
    # --------------------------------------------------------

    $fichierODSTemp = $fichierODS + ".tmp"

    if (Test-Path $fichierODSTemp) {
        Remove-Item $fichierODSTemp -Force
    }

    $archiveSource = [System.IO.Compression.ZipFile]::OpenRead(
        $fichierODS
    )

    $archiveDestination = [System.IO.Compression.ZipFile]::Open(
        $fichierODSTemp,
        [System.IO.Compression.ZipArchiveMode]::Create
    )

    try {

        # ----------------------------------------------------
        # Le fichier mimetype doit être en premier et
        # non compressé dans un fichier ODS.
        # ----------------------------------------------------

        $entreeMime = $archiveSource.GetEntry("mimetype")

        if ($null -ne $entreeMime) {

            $nouvelleEntree = $archiveDestination.CreateEntry(
                "mimetype",
                [System.IO.Compression.CompressionLevel]::NoCompression
            )

            $sourceStream = $entreeMime.Open()
            $destinationStream = $nouvelleEntree.Open()

            $sourceStream.CopyTo($destinationStream)

            $destinationStream.Close()
            $sourceStream.Close()
        }


        # ----------------------------------------------------
        # Copie des autres fichiers
        # ----------------------------------------------------

        foreach ($entree in $archiveSource.Entries) {

            if ($entree.FullName -eq "mimetype") {
                continue
            }

            $nouvelleEntree = $archiveDestination.CreateEntry(
                $entree.FullName,
                [System.IO.Compression.CompressionLevel]::Optimal
            )

            $destinationStream = $nouvelleEntree.Open()

            if ($entree.FullName -eq "content.xml") {

                $destinationStream.Write(
                    $nouveauContentXml,
                    0,
                    $nouveauContentXml.Length
                )

            }
            else {

                $sourceStream = $entree.Open()

                $sourceStream.CopyTo($destinationStream)

                $sourceStream.Close()
            }

            $destinationStream.Close()
        }

    }
    finally {

        $archiveDestination.Dispose()
        $archiveSource.Dispose()
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