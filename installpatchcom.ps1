$comobject = New-Object -ComObject Microsoft.Update.Session
$searcherresult = $comobject.CreateupdateSearcher().Search("isinstalled=0 and type='Software'").Updates
$download = $comobject.CreateUpdateDownloader() 
$download.Updates = $searcherresult 
$downloadresult = $download.download()


$install = $comobject.CreateUpdateInstaller()
$install.updates = $searcherresult
$installresult = $install.install()
