function Ping-IpRange 
{
  <#
      .SYNOPSIS
      Tests a range of IP addresses in a class "C" network

      .DESCRIPTION
      Tests a range of IP addresses using "Test-Netconnetconnection -InformationLevel Detailed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue"


      .PARAMETER FirstIpAddress
      Starting IP Address

      .PARAMETER LastIpAddress
      Ending IP Address

      .PARAMETER Class
      Future Use

      .EXAMPLE
      Ping-IpRange -FirstIpAddress 192.168.1.21 -LastIpAddress 192.168.1.23
    
      OUTPUT:
      Computer Name: 192.168.1.22
      Computer Name: 192.168.1.23

      .EXAMPLE
      Ping-IpRange -FirstIpAddress 192.168.1.21 -LastIpAddress 192.168.1.23 -Verbose
    
      OUTPUT:
      VERBOSE: Full Address: 192.168.1.21
      VERBOSE: Ping Failed: 192.168.1.21
      VERBOSE: Full Address: 192.168.1.22
      Ping Completed: 192.168.1.22
      VERBOSE: Full Address: 192.168.1.23
      Ping Completed: 192.168.1.23

      .NOTES
      This will currently only handle IP address ranges where the 4th octet is changed

      .LINK
      https://github.com/OgJAkFy8/AssetManagentapp
  #>


  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $FirstIpAddress = '192.168.1.41',
    [Parameter(Mandatory = $true, Position = 1)]
    [string]
    $LastIpAddress = '192.168.1.50',
    [Parameter(Mandatory = $false, Position = 2)]
    [ValidateSet('A','B','C')]
    [String]$Class
  )
  [hashtable]$IpRange = @{}
  function Test-Ipaddress
  {
    param
    (
      [Parameter(Mandatory = $true, Position = 0)]
    [Object]$IpAddress)
    if(([ipaddress]$IpAddress).AddressFamily -eq 'InterNetwork')
    {
      return 'IPv4'
    }
    elseif(([ipaddress]$IpAddress).AddressFamily -eq 'InterNetworkV6')
    {
      return 'Ipv6'
    }
  }
  function Resolve-IpRange
  {
    param([Parameter(Mandatory = $true)]
      [Object]$FirstIpAddress, 
      [Parameter(Mandatory = $true)]
    [Object]$LastIpAddress)

    $StartIp = [int]([ipaddress]$FirstIpAddress).GetAddressBytes()[3] 
    $EndIp = [int]([ipaddress]$LastIpAddress).GetAddressBytes()[3] 
    if($StartIp -lt $EndIp)
    {
      $First3Oct = $FirstIpAddress.TrimEnd([string]$StartIp)
    }
    else
    {
      $First3Oct = $LastIpAddress.TrimEnd([string]$EndIp)
      $StartIp = [int]([ipaddress]$LastIpAddress).GetAddressBytes()[3] 
      $EndIp = [int]([ipaddress]$FirstIpAddress).GetAddressBytes()[3]
    }
    [hashtable]$IpRange = @{
      First3Oct = $First3Oct
      StartIp   = $StartIp
      EndIp     = $EndIp
    }
    return $IpRange
  }
  
  if( ((Test-Ipaddress -IpAddress $FirstIpAddress) -eq 'ipv4') -and ((Test-Ipaddress -IpAddress $LastIpAddress) -eq 'ipv4'))
  {
    $IpRange = Resolve-IpRange -FirstIpAddress $FirstIpAddress -LastIpAddress $LastIpAddress
    for ($i = $IpRange.StartIp;$i -le $IpRange.EndIp;$i++)
    {
      $FullIpAdd = '{0}{1}' -f $IpRange.First3Oct, $i
      Write-Verbose -Message ('Full Address: {0}' -f $FullIpAdd)
      $ping = Test-NetConnection -ComputerName $FullIpAdd -InformationLevel Detailed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
      if($ping.PingSucceeded -eq 'True')
      {
        Write-Output -InputObject ('Ping Completed: {0}' -f $ping.ComputerName)
      }
      else
      {
        Write-Verbose -Message ('Ping Failed: {0}' -f $FullIpAdd)
      }
    }
  }
  else
  {
    Write-Error -Message 'Not v4 Address'
  }
}

