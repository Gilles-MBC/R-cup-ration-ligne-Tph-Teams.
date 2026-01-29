# 1. Configuration de l'environnement
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force

$modules = @("Microsoft.Graph", "ImportExcel", "MicrosoftTeams")
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installation du module $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
    }
}

# --- CONFIGURATION DES TEXTES PRÉDÉFINIS ---
# Modifiez cette liste : "Nom_De_L_OU" = "Texte à afficher"
$mappingOU = @{
    "Agence Grands-Comptes" = "Grands Comptes"
    "Agence Centrale d'achats" = "Centrale d'achats"
    "Agence Ecole du Capital Toi(t)"    = "Ecole du Capital Toi(t)"
   
    # Ajoutez autant de lignes que nécessaire
}
# --------------------------------------------

# 2. Connexion aux services
Try {
    Write-Host "Connexion à Microsoft Teams et Microsoft Graph..." -ForegroundColor Cyan
    Connect-MicrosoftTeams -ErrorAction Stop
    Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All" -NoWelcome
} Catch {
    Write-Host "Erreur de connexion : $_" -ForegroundColor Red
    Exit
}

# 3. Récupération des données
Write-Host "Récupération des numéros..." -ForegroundColor Cyan
$phoneNumberAssignments = Get-CsPhoneNumberAssignment | Where-Object { $null -ne $_.AssignedPstnTargetId }

# 4. Traitement des données
$results = foreach ($assignment in $phoneNumberAssignments) {
    Try {
        $targetId = $assignment.AssignedPstnTargetId
        $identity = Get-MgUser -UserId $targetId -Property "DisplayName","UserPrincipalName","OnPremisesDistinguishedName" -ErrorAction SilentlyContinue
        
        $displayName = "Inconnu"
        $upn = "N/A"
        $ouBrute = "Cloud / Inconnu"
        $textePredefini = "Standard téléphonique"

        if ($identity) {
            $displayName = $identity.DisplayName
            $upn = $identity.UserPrincipalName
            
            if ($null -ne $identity.OnPremisesDistinguishedName) {
                $dnParts = $identity.OnPremisesDistinguishedName -split ","
                $ouPart = $dnParts | Where-Object { $_ -like "OU=*" } | Select-Object -First 1
                if ($ouPart) {
                    $ouBrute = $ouPart -replace "OU=", ""
                    
                    # Vérification si l'OU est dans notre table de correspondance
                    if ($mappingOU.ContainsKey($ouBrute)) {
                        $textePredefini = $mappingOU[$ouBrute]
                    } else {
                        $textePredefini = $ouBrute
                    }
                }
            }
        }

        $type = $assignment.AssignmentCategory
        if ($type -eq "ApplicationInstance") { $type = "Service (File/Standard)" }

        [PSCustomObject]@{
            #Type           = $type
            DisplayName    = $displayName
            Identifiant    = $upn
            Telephone      = $assignment.TelephoneNumber
            #OU_Originale   = $ouBrute
            Description_OU = $textePredefini
        }
    } Catch {
        Write-Warning "Erreur sur l'ID $targetId : $_"
    }
}

# 5. Exportation
if ($results) {
    # Génère un nom type : Inventaire_28_01_2026_14h30.xlsx
    $horodatage = Get-Date -Format "dd_MM_yyyy_HH\hmm"
    $nomFichier = "Inventaire_Teams_$horodatage.xlsx"
    
    $results | Export-Excel -Path $nomFichier -AutoSize -TableName "Inventaire" -WorksheetName "Telephonie" -Show
    
    Write-Host "Export terminé avec succès ! Fichier généré : $nomFichier" -ForegroundColor Green
}