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
if($null -eq $inotPath) {
        $drives = Get-PSDrive -PSProvider FileSystem | Foreach-Object { if(Test-Path -Path $_.Root){get-childitem $_.Root | Where-Object {$_.PSIsContainer -eq $true -and $_.Name -match "FrameWork.CSID"} } }

        if($null -ne $drives)
        {
            $inotPath = Join-Path -Path $drives.Root -ChildPath "Framework.CSID"; 
        }
} else {
    $inotPath = Join-Path -Path $inotPath -ChildPath "Framework.CSID"; 
}

# Recuperation du dossier de la synchro
$synchroService = Get-WmiObject win32_service | ?{$_.Name -like 'Synchronisation iNot Exchange'} 
$synchroPath = $synchroService.PathName.Trim().Trim('"')


$inotversion = (Get-Item $inotPath\San\i-Not\Builds\GenApi.iNot.Client.exe).VersionInfo.ProductVersion
$booksversion = (Get-Itemproperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\Serveur).version
$pendingupdate = (Get-ChildItem -path $csidArchivePath -Recurse -Force -name)
$synchroversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($synchroPath).FileVersion
$ApplicationTmp = Select-String -Path $csidUpdatePath\paramgu.ini -Pattern Application 
$Application = ($ApplicationTmp -split '=')[1]

Write-Output " Inot:$inotversion Books:$booksversion Synchro:$synchroversion List:$Application Pending:$pendingupdate  "
