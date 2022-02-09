Param(
    [string]$MainDomain,
    [switch]$UseExisting,
    [switch]$ForceRenew
)

function Logging {
    param([string]$Message)
    Write-Host $Message
    $Message >> $LogFile
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Import-Module PKI
Import-Module Posh-Acme
$LogFile = '.\UpdateWAC.log'
Get-Date | Out-File $LogFile -Append
if($UseExisting) {
    Logging -Message "Using Existing Certificate"
    $cert = get-pacertificate -MainDomain $MainDomain
}
else {
    if($ForceRenew) {
        Logging -Message "Starting Forced Certificate Renewal"
        $cert = Submit-Renewal -MainDomain $MainDomain -Force
    }
    else {
        Logging -Message "Starting Certificate Renewal"
        $cert = Submit-Renewal -MainDomain $MainDomain
    }
    Logging -Message "...Renew Complete!"
}

if($cert){
    $wac = get-wmiobject Win32_Product | select IdentifyingNumber, Name, LocalPackage | Where Name -eq "Windows Admin Center"

     if ($wac -ne $null)
      {
          # Bind new certificate to the service
          Logging -Message "Updating WAC Certificate"
          Start-Process msiexec.exe -Wait -ArgumentList "/i $($wac.LocalPackage) /qn /L*v c:\script\log.txt SME_PORT=1080 SME_THUMBPRINT=$($c.Thumbprint) SSL_CERTIFICATE_OPTION=installed"

          # When upgrading WAC, the firewall rule may be deleted. If so create a new rule after upgrade.
          Logging -Message "Recreate firewall rule"
          New-NetFirewallRule -DisplayName "SmeInboundOpenException" -Description "Windows Admin Center inbound port exception" -LocalPort 1080 -RemoteAddress Any -Protocol TCP

          # Restart Windows Admin Center
          Logging -Message "Restarting WAC service"
          Restart-Service ServerManagementGateway -Force
      }

    # Remove old certs
    ls Cert:\LocalMachine\My | ? Subject -eq "CN=$MainDomain" | ? NotAfter -lt $(get-date) | remove-item -Force
}else{
    Logging -Message "No need to update WAC certifcate" 
}
