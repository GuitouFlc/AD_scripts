# Définir le chemin du fichier CSV
$CSVFile = "C:\temp\laps.csv"
# Ci-dessous on va se limiter aux machines ayant moins de 3ans
$DateLimite = (Get-Date).AddYears(-3)
# on filtre avec notre variable du dessus et on se limite au machines "actives"
# et on recupere la propriété qui nous interesse (ici le mot de passe du compte LAPS)
$Machines = Get-ADComputer -Filter {Enabled -eq $True -and WhenCreated -ge $DateLimite} -Properties ms-Mcs-AdmPwd
# on vient retrier nos machines en ne selectionnant que les machine dont la propriété "ms-Mcs-AdmPwd" est vide
$MachinesSansLAPS = $Machines | Where-Object { [string]::IsNullOrEmpty($_.'ms-Mcs-AdmPwd') }
#on stock tout ça dans un tableau en vu de l'utiliser plus tard (evolution de script)
$Rapport = @()
foreach ($Machine in $MachinesSansLAPS) {
    $NomMachine = $Machine.Name
    $Enabled = $machine.Enabled
    $Rapport += [PSCustomObject]@{
        'NomMachine' = $NomMachine
        #'Active' = $Enabled
    }
}

#on sort notre rapport ici en csv au chemin defini en début de script ($CSVFile)
$Rapport | Export-Csv -Path $CSVFile -NoTypeInformation
$Rapport | Format-Table -AutoSize

Write-Host "Le rapport des machines sans attribut LAPS a été enregistré dans $CSVFile."