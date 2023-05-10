$errors = "No error"

# Vérification de la version de PowerShell
if($PSVersionTable.PSVersion.Major -lt 4){
    echo 'PowerShell 4.0 ou superieur est requis pour executer ce script !'
}

# Vérification de l'espace disque disponible sur le lecteur C:
$freeSpace = (Get-PSDrive -Name C).Free
$requiredSpace = 5GB # Espace disque minimal nécessaire pour l'installation

if($freeSpace -lt $requiredSpace){
    echo "Espace disque insuffisant sur le lecteur C: $($freeSpace/1GB) Go disponibles sur $($requiredSpace/1GB) Go requis !" 
}

# Récupération de la valeur ID MySepteo / ICARE
$NumEtudeTmp = Select-String -Path 'C:\Program Files (x86)\CSiD\CSiD Update\paramgu.ini' -Pattern Numero 
$NumEtude = ($NumEtudeTmp -split '=')[1]

# Vérification de la variable $NumEtude
if([string]::IsNullOrEmpty($NumEtude)){
    echo 'Pas un serveur de production iNot !' 
    Exit
}

# Récupération du nom de l'étude
$nometude = (Get-Itemproperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\CSIDAlerts).NOMCLIENT

# Récupération de la version de l'OS
$OsVersion = (Get-WmiObject -Class Win32_OperatingSystem).Caption.Replace("Microsoft Windows", "Win")

# Récupération du dossier CSID Update
$csidUpdatePath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\CSiD\CSiDUpdate | Select-Object -ExpandProperty InstallLocation;
$csidArchivePath = Join-Path -Path $csidUpdatePath -ChildPath "Archives";

# Récupération de la version inot et inotBooks encodée dans le fichier XML
$xmlPath = Join-Path -Path $csidUpdatePath -ChildPath "Version_Etude.xml";
[xml]$xmlFile = Get-Content $xmlPath;
$inotProduit = $xmlFile.ConfigEtude.Produits.Produit | Where-Object id -eq 'iNot' | Select-Object -First 1;
$versionXML = $inotProduit.Versions.Version | Where-Object update -eq 'O' | Select-Object -First 1
$BooksProduit = $xmlFile.ConfigEtude.Produits.Produit | Where-Object id -eq 'Books' | Select-Object -First 1;
$versionBooksXML = $BooksProduit.Versions.Version | Where-Object update -eq 'O' | Select-Object -First 1
$versioninotGU = $versionXML.version;
$versionBooksGU = $versionBooksXML.version;

#Récupération du chemin de l'installation inot
$inotPath = $null;
if(Test-Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be)
{
$inotPath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be | Select-Object -ExpandProperty InstallLocation;

    if($null -eq $inotPath){
        $inotPath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be | Select-Object -ExpandProperty InstallLocationBak;
    }
}
if  ($null -eq $inotPath) {
        $drives = @()
        $drives += Get-PSDrive -PSProvider FileSystem | Foreach-Object { if(Test-Path -Path $_.Root){get-childitem $_.Root  | Where-Object {$_.PSIsContainer -eq $true -and $_.Name -match "FrameWork.CSID"} } }

        if($null -ne $drives)
        {
            $inotPath = $drives[0].FullName; 
            if($drives.Count -gt 1)
            {
                $errors = "Several disks with Framework.CSID_WARNING for install";
            }

            if(-Not (Test-Path -Path "$inotPath\San\i-Not\Builds\"))
            {
                if($drives.Count -gt 1)
                {
                    $inotPath = $drives[1].FullName;
                }
            }
        } else {
            $drives = Get-PSDrive -PSProvider FileSystem | Foreach-Object { if(Test-Path -Path $_.Root) { get-childitem $_.Root -recurse -ErrorAction SilentlyContinue  | Where-Object {$_.PSIsContainer -eq $true -and $_.Name -match "FrameWork.CSID"} } }
            if($null -ne $drives)
            {
                $inotPath = $drives[0].FullName; 
                if($drives.Count -gt 1)
                {
                    $errors = "Several disks with Framework.CSID_WARNING for install";
                }
    
                if(-Not (Test-Path -Path "$inotPath\San\i-Not\Builds\"))
                {
                    if($drives.Count -gt 1)
                    {
                        $inotPath = $drives[1].FullName;
                    }
                }
            }
        }
} else {
    $inotPath = Join-Path -Path $inotPath -ChildPath "Framework.CSID"; 
}
#Récupération du chemin de l'installation inotBooks
$inotBooksPath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\serveur | Select-Object -ExpandProperty RepertoireInstallation;

#Récupération du statut de la synchro
$svcName = "SynchroINot.exe"
$services = Get-WmiObject win32_service | Where-Object { $_.PathName -like "*$svcName*" } 

#Récéperation du dossier de la synchro
$synchroPath = $null;
$synchroService = Get-WmiObject win32_service | ?{$_.Name -like 'Synchronisation iNot Exchange'};
if($null -ne $synchroService)
{
    $synchroPath = $synchroService.PathName.Trim().Trim('"');
} else {
    $synchroService = Get-WmiObject win32_service | ?{$_.Name -like 'Synchro iNot Exchange'};
    if($null -ne $synchroService)
    {
        $synchroPath = $synchroService.PathName.Trim().Trim('"');
    }
} 

#récupération des versions des installations Inot, InotBooks et synchro présentes sur le serveur
if($null -ne $inotPath) {
    $inotversion = (Get-Item $inotPath\San\i-Not\Builds\GenApi.iNot.Client.exe).VersionInfo.ProductVersion;
} else {
    $inotversion = " No iNot Office";
}
if(Test-Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\Serveur)
{
    $booksversion = (Get-Itemproperty $inotBooksPath\Serveur\bin\GenApi.Finance.Serveur.Common.dll).VersionInfo.fileversion;
} else {
    $booksversion = " No Books";
}
if($null -ne $synchroPath) {
$synchroversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($synchroPath).FileVersion;
} else {
    $synchroversion = " No Sync";
}
$pendingupdate = (Get-ChildItem -path $csidArchivePath -Force -name);
if($null -eq $pendingupdate)
{
    $pendingupdate = " empty";
}
$ApplicationTmp = Select-String -Path $csidUpdatePath\paramgu.ini -Pattern Application;
$Application = ($ApplicationTmp -split '=')[1];


$output = "";
if($inotversion -eq $versioninotGU)
{
    $output = " $NumEtude ;Inot: $inotversion ;OK;;Books: $booksversion ;BooksXml: $versionBooksGU ;Sync: $synchroversion $($svc.State) ;Apps in GU: $Application ;Pending: $pendingupdate ; $errors ; $nometude ; $OsVersion";
} else {
    $output = " $NumEtude ;Inot: $inotversion ;NOK;GU: $versioninotGU ;Books: $booksversion ;BooksXml: $versionBooksGU ;Sync: $synchroversion $($svc.State) ;Apps in GU: $Application ;Pending: $pendingupdate ; $errors ; $nometude ; $OsVersion ";
}
# affichage des infos dans les remontées de RG
Write-Output $output;

# suppression du script et des anciennes versions
$scriptName = "precheckinstall"
$scriptPath = "C:\Windows\TEMP\rgsupv"
Get-ChildItem $scriptPath -Filter "$scriptName" -Recurse | Remove-Item -Force
exit
