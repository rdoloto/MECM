function Get-OfficeC2RCom{ 
$comAdmin = New-Object -com ("COMAdmin.COMAdminCatalog")

$applications = $comAdmin.GetCollection("Applications")
$applications.Populate()

foreach ($application in $applications)
{
 if ($application.Name -eq "OfficeC2RCom")
 { $components = $applications.GetCollection("Components",$application.key) 
  $components.Populate()
  $componentscount= $components.Count

if ($componentscount -eq 2){
    write-host $false

 
 }
 if($componentscount -eq 1)
 {   write-host $true}

}
 
}
}
Get-OfficeC2RCom