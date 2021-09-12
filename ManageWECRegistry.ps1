$subscription = 'Default'
$Folder = 'C:\scripts'
"Test to see if folder [$Folder]  exists"
if (Test-Path -Path $Folder) {
} else {
    md $Folder > $null
}
if(Test-Path -Path C:\scripts\ManageWECRegistry.log){
    Move-Item C:\scripts\ManageWECRegistry.log C:\scripts\ManageWECRegistry.log.bak -Force
}
Get-Date | Out-File -FilePath C:\scripts\ManageWECRegistry.log -Append -Encoding ascii
$computers = @()
$domains = (Get-ADForest).domains
"Domains found:"  | Out-File -FilePath C:\scripts\ManageWECRegistry.log -Append -Encoding ascii
$domains  | Out-File -FilePath C:\scripts\ManageWECRegistry.log -Append -Encoding ascii
foreach ($result in $domains){
    $dc = Get-ADDomainController -DomainName $result -Discover -Service PrimaryDC
    $computers += (get-adcomputer -filter * -Server $dc).DNSHostName
}
$subscribed_computers = Get-childitem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$subscription\EventSources -Name

$subscribed_computers | Where-Object { $computers -notcontains $_ } | ForEach-Object {
    $computer_name = $_
    Write-Host "Removing computer $computer_name from WEC registry"
    "Removing computer $computer_name from WEC registry" | Out-File -FilePath C:\scripts\ManageWECRegistry.log -Append -Encoding ascii
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$subscription\EventSources\$computer_name" -Force -Verbose
}
Get-Date | Out-File -FilePath C:\scripts\ManageWECRegistry.log -Append -Encoding ascii
