Import-Module ActiveDirectory
# chemin de l export Excel
$cheminFichierExcel = "C:\script\PCInactif.xlsx"

# date limite (12 mois)
$dateLimite = (Get-Date).AddMonths(-12)

# On recupere les ordinateurs de l'AD
$ordinateurs = Get-ADComputer -Filter * -Property Name, LastLogonDate, OperatingSystem, Description

# On les filtres avec notre $datelimite
$ordinateursInactifs = $ordinateurs | Where-Object { $_.LastLogonDate -lt $dateLimite }

# On vient utiliser Excel pour la mise en forme de nos données (voir les cmdlet Excel)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$classeur = $excel.Workbooks.Add()
$feuille = $classeur.Worksheets.Item(1)

$feuille.Cells.Item(1, 1) = "Nom de l'ordinateur"
$feuille.Cells.Item(1, 2) = "OS"
$feuille.Cells.Item(1, 3) = "Description"
$feuille.Cells.Item(1, 4) = "Dernière date de connexion"

$indiceLigne = 2
foreach ($ordinateur in $ordinateursInactifs) {
    $feuille.Cells.Item($indiceLigne, 1) = $ordinateur.Name
    $feuille.Cells.Item($indiceLigne, 2) = $ordinateur.OperatingSystem
    $feuille.Cells.Item($indiceLigne, 3) = $ordinateur.Description
    $feuille.Cells.Item($indiceLigne, 4) = $ordinateur.LastLogonDate
    $indiceLigne++
}

$classeur.SaveAs($cheminFichierExcel)
$excel.Quit()
# end Excel

