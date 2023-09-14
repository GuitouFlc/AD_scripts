Import-Module ActiveDirectory
Import-Module -Name ImportExcel

# Recuperation des "OU"
$ous = Get-ADOrganizationalUnit -Filter * -Properties Name, DistinguishedName
#On prepare un tableau vide que l'on va remplir par la suite
$emptyOUs = @()

foreach ($ou in $ous) {
    # On vérifie si l'OU est vide
    $isEmpty = !(Get-ADObject -SearchBase $ou.DistinguishedName -SearchScope OneLevel -Filter * -Property *)
    if ($isEmpty) {
        $emptyOU = [PSCustomObject]@{
            Name = $ou.Name
            DistinguishedName = $ou.DistinguishedName
        }
        #on ajoute nos OU vides dans notre tableau
        $emptyOUs += $emptyOU
    }
}

# Export Excel
$exportPath = "C:\script\empty_ous.xlsx"
# On exporte tout ça en *.xlsx
$emptyOUs | Export-Excel -Path $exportPath -AutoSize -AutoFilter
Write-Host "Les OU vides ont été exportées dans le fichier Excel : $exportPath."
