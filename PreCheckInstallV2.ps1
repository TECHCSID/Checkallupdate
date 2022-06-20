$errors = "Pas d'erreurs"

#Récupréation de la valeur
$NumEtudeTmp = Select-String -Path 'C:\Program Files (x86)\CSiD\CSiD Update\paramgu.ini' -Pattern Numero 
$NumEtude = ($NumEtudeTmp -split '=')[1]

# Recuperation du dossier CSID Update
$csidUpdatePath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\CSiD\CSiDUpdate | Select-Object -ExpandProperty InstallLocation;
$csidArchivePath = Join-Path -Path $csidUpdatePath -ChildPath "Archives";

# Recuperation du dossier inot
$inotPath = $null;
if(Test-Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be)
{
$inotPath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be | Select-Object -ExpandProperty InstallLocation;

    if($null -eq $inotPath){
        $inotPath = Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\iNot.be | Select-Object -ExpandProperty InstallLocationBak;
    }
}
if  ($null -eq $inotPath) {
        $drives = Get-PSDrive -PSProvider FileSystem | Foreach-Object { if(Test-Path -Path $_.Root){get-childitem $_.Root | Where-Object {$_.PSIsContainer -eq $true -and $_.Name -match "FrameWork.CSID"} } }

        if($null -ne $drives)
        {
            $inotPath = Join-Path -Path $drives[0].Root -ChildPath "Framework.CSID"; 
            if($drives.Count -gt 1)
            {
                $errors = "Plusieurs disques trouves contenant Framework.CSID, attention aux problemes d installations";
            }

            if(-Not (Test-Path -Path "$inotPath\San\i-Not\Builds\"))
            {
                if($drives.Count -gt 1)
                {
                    $inotPath = Join-Path -Path $drives[1].Root -ChildPath "Framework.CSID";
                }
            }
        }
} else {
    $inotPath = Join-Path -Path $inotPath -ChildPath "Framework.CSID"; 
}

# Recuperation du dossier de la synchro
$synchroPath = $null;
$synchroService = Get-WmiObject win32_service | ?{$_.Name -like 'Synchronisation iNot Exchange'};
if($null -ne $synchroService)
{
    $synchroPath = $synchroService.PathName.Trim().Trim('"')
} 


$inotversion = (Get-Item $inotPath\San\i-Not\Builds\GenApi.iNot.Client.exe).VersionInfo.ProductVersion;
if(Test-Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\Serveur)
{
    $booksversion = (Get-Itemproperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\Serveur).version;
} else {
    $booksversion = "Pas de books";
}
if($null -ne $synchroPath) {
$synchroversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($synchroPath).FileVersion;
} else {
    $synchroversion = " Pas de Synchronisation Exchange";
}
$pendingupdate = (Get-ChildItem -path $csidArchivePath -Force -name);
$ApplicationTmp = Select-String -Path $csidUpdatePath\paramgu.ini -Pattern Application;
$Application = ($ApplicationTmp -split '=')[1];

Write-Output "$NumEtude Inot:$inotversion Books:$booksversion Synchro:$synchroversion List:$Application Pending:$pendingupdate Error:$errors ";
