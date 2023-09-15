
<# L'utilisation de se script necessite un acces aux profils wifi qui peuvent etre soit créé à la main,
soit extrait d'une machine via la commande netsh (attention le mot de passe doit etre extrait en clair de façon à pouvoir etre utilisé sur un autre pc)
commande d'extraction des profils > netsh wlan export profile "MON-WIFI" key=clear folder=C:\Temp
voir cet article https://www.it-connect.fr/import-et-export-de-profils-wi-fi-avec-netsh/ #>

function AddWiFi{

    #Define variable
    $EmplacementDesProfils = 'C:\Temp'
    $list = Get-ChildItem -path $fold
    
        #foreach loop
        foreach ($file in $list) {
            #add Wifi profile
            netsh wlan add profile filename=$EmplacementDesProfils$file user=all
        }
    
        Remove-Item $fold -Force -Recurse
        Write-Host "all xml files has been removed"
    }
    
    function main {
        AddWiFi    
    }
    
    main