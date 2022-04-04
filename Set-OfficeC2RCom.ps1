function Set-OfficeC2RCom{ 
$comAdmin = New-Object -com ("COMAdmin.COMAdminCatalog")

$applications = $comAdmin.GetCollection("Applications")
$applications.Populate()

foreach ($application in $applications)
{
 if ($application.Name -eq "OfficeC2RCom")
 { $components = $applications.GetCollection("Components",$application.key) 
  $components.Populate()

 $comp=$components | Where-Object {$_.Name -eq "UpdateNotify.Object.2"}
 
 for( $i=0; $i -lt $components.Count; $i++ )
 {
  if ($components.item($i).name -eq "UpdateNotify.Object.2")
  {$components.Remove($i)
   $components.SaveChanges()
  }
 }
  $components = $applications.GetCollection("Components",$application.key) 
  $components.Populate()
  $componentscount= $components.Count
  if($componentscount -eq 1)
 { cls 
  write-host $true}
  }
}
}
Set-OfficeC2RCom