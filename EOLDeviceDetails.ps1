$credentials = Get-Credential
    Write-Output "Getting the Exchange Online cmdlets"
 
    $session = New-PSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
        -ConfigurationName Microsoft.Exchange -Credential $credentials `
        -Authentication Basic -AllowRedirection
    Import-PSSession $session
 
$csv = "C:\CE-Audit\EOLDevices.csv"
$results = @()
$mailboxUsers = get-mailbox -resultsize unlimited
$mobileDevice = @()
 
foreach($user in $mailboxUsers)
{
$UPN = $user.UserPrincipalName
$displayName = $user.DisplayName
 
$mobileDevices = Get-MobileDeviceStatistics -Mailbox $UPN
       
      foreach($mobileDevice in $mobileDevices)
      {
          Write-Output "Getting info about a device for $displayName"
          $properties = @{
          Name = $user.name
          UPN = $UPN
          DisplayName = $displayName
          ClientType = $mobileDevice.ClientType
          ClientVersion = $mobileDevice.ClientVersion
          DeviceId = $mobileDevice.DeviceId
          DeviceMobileOperator = $mobileDevice.DeviceMobileOperator
          DeviceModel = $mobileDevice.DeviceModel
          DeviceOS = $mobileDevice.DeviceOS
          DeviceTelephoneNumber = $mobileDevice.DeviceTelephoneNumber
          DeviceType = $mobileDevice.DeviceType
          LastSuccessSync = $mobileDevice.LastSuccessSync
          UserDisplayName = $mobileDevice.UserDisplayName
          }
          $results += New-Object psobject -Property $properties
      }
}
 
$results | Select-Object Name,UPN,DisplayName,ClientType,ClientVersion,DeviceId,DeviceMobileOperator,DeviceModel,DeviceOS,DeviceTelephoneNumber,DeviceType,LastSuccessSync,UserDisplayName | Export-Csv -notypeinformation -Path $csv
 pause
#Remove-PSSession $session
