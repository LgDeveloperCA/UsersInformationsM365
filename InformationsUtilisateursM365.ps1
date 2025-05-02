# Ce script permet de récuperer toutes les informations ENTRA de tous les utilisateurs du tenant
# Et d'exporter le tout dans un fichier Excel

# Connexion à Microsoft Graph (si pas déjà connecté)
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "AuditLog.Read.All", "Group.Read.All"

# Charger la liste des SKU une seule fois pour améliorer la performance
$skuMap = @{}
Get-MgSubscribedSku | ForEach-Object {
    $skuMap[$_.SkuId] = $_.SkuPartNumber
}

# Récupération des utilisateurs
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,GivenName,Surname,JobTitle,Department,CompanyName,Mail,OtherMails,MailNickname,OfficeLocation,BusinessPhones,MobilePhone,State,City,Country,PostalCode,AccountEnabled,CreatedDateTime,AssignedLicenses,UserType

#Construire une liste unique de toutes les licences et groupes
$allLicenses = [System.Collections.Generic.HashSet[string]]::new()
$allGroups   = [System.Collections.Generic.HashSet[string]]::new()

$usersData = foreach ($user in $users) {
    $userId = $user.Id

    # Licences
    $userLicenses = ($user.AssignedLicenses | ForEach-Object { $skuMap[$_.SkuId] }) | Where-Object { $_ }
    $userLicenses | ForEach-Object { $allLicenses.Add($_) | Out-Null }

    # Groupes
    $groupNames = ""
    try {
        $groupNames = (Get-MgUserMemberOf -UserId $userId | Where-Object { $_.AdditionalProperties.displayName } | Select-Object -ExpandProperty AdditionalProperties | ForEach-Object { $_["displayName"] })
        $groupNames | ForEach-Object { $allGroups.Add($_) | Out-Null }
    } catch {}

    # Création objet initial (sans licences/groupes encore)
    [PSCustomObject]@{
        Id             = $user.Id
        DisplayName    = $user.DisplayName
        Email          = $user.UserPrincipalName
        LicensesList   = $userLicenses
        GroupsList     = $groupNames
        GivenName      = $user.GivenName
        Surname        = $user.Surname
        JobTitle       = $user.JobTitle
        Department     = $user.Department
        OfficeLocation = $user.OfficeLocation
        Phone          = ($user.BusinessPhones -join ", ")
        Alias          = $user.MailNickname
        UserType       = $user.UserType
    }
}

# Étape 3 – Expansion des colonnes licences et groupes
$finalResult = foreach ($u in $usersData) {
    $obj = [ordered]@{
        ObjectId         = $u.Id
        NomComplet       = $u.DisplayName
        Courriel         = $u.Email
        Prenom           = $u.GivenName
        Nom              = $u.Surname
        Poste            = $u.JobTitle
        Service          = $u.Department
        Bureau           = $u.OfficeLocation
        Telephone        = $u.Phone
        Alias            = $u.Alias
        TypeUtilisateur  = $u.UserType
    }

    # Ajouter dynamiquement une colonne par licence
    foreach ($lic in $allLicenses) {
        $obj[$lic] = $u.LicensesList -contains $lic ? 1 : 0
    }

    # Ajouter dynamiquement une colonne par groupe
    foreach ($grp in $allGroups) {
        $obj["Groupe: $grp"] = $u.GroupsList -contains $grp ? 1 : 0
    }

    [PSCustomObject]$obj
}

# Export Excel
$chemin = "Utilisateurs_M365_Complet.xlsx"
$finalResult | Export-Excel -Path $chemin -AutoSize -TableName "UsersM365" -WorksheetName "UsersM365"

Write-Host "✅ Export terminé : $chemin"
