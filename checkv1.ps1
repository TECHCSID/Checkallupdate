$inotversion = (Get-Item E:\FrameWork.CSID\San\i-Not\Builds\GenApi.iNot.Client.exe).VersionInfo.ProductVersion
$booksversion = (Get-Itemproperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\GenApi\Finance\Serveur).version
$pendingupdate = (Get-ChildItem -path C:\"Program Files (x86)"\CSID\"CSiD Update"\Archives\ -Recurse -Force -name)
$synchroversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Program Files (x86)\CSiD\Synchro iNot Exchange\SynchroInot.exe").FileVersion
$ApplicationTmp = Select-String -Path 'C:\Program Files (x86)\CSiD\CSiD Update\paramgu.ini' -Pattern Application 
$Application = ($ApplicationTmp -split '=')[1]

Write-Output " Inot:$inotversion Books:$booksversion Synchro:$synchroversion List:$Application Pending:$pendingupdate  " 
