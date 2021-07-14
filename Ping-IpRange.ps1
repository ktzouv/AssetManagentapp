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


  [CmdletBinding(SupportsShouldProcess = $true,ConfirmImpact = 'Low',DefaultParameterSetName = 'Connection')]

  param
  (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
          $ptrn = '^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$'
          If($_ -match $ptrn)
          {
            $true
          }
          Else
          {
            Throw 'v4 IPAddress required. (ex: 192.168.10.123)'
          }
    })][String]$FirstIpAddress = '192.168.1.41',
    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateScript({
          $ptrn = '^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$'
          If($_ -match $ptrn)
          {
            $true
          }
          Else
          {
            Throw 'v4 IPAddress required. (ex: 192.168.10.123)'
          }
    })][String]$LastIpAddress = '192.168.1.50',
    [Parameter(Mandatory = $false, Position = 2)]
    [ValidateSet('A','B','C')]
    [String]$Class,
    [Parameter(ParameterSetName = 'NetConnection')]
    [Switch]$useNetconnetion ,
    [Parameter(ParameterSetName = 'Workflow')]
    [Switch]$useWorkflow
  )
  BEGIN
  {
    [hashtable]$IpRange = @{} # For storing the address numbers that will be needed for the "ping"
    $FullIpList = @() # For storing the full list of IP addresses.  This was put in after the other for use with the workflow
    function Test-Ipaddress # 
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

    workflow Test-WFConnection 
    {
      param(

        [string[]]$Computers

      )

      foreach -parallel ($computer in $Computers) 
      {
        Test-Connection -ComputerName $computer -Count 1 -ErrorAction SilentlyContinue
      }
    }
  }
  PROCESS
  {
    if( ((Test-Ipaddress -IpAddress $FirstIpAddress) -eq 'ipv4') -and ((Test-Ipaddress -IpAddress $LastIpAddress) -eq 'ipv4'))
    {
      $IpRange = Resolve-IpRange -FirstIpAddress $FirstIpAddress -LastIpAddress $LastIpAddress
      for ($i = $IpRange.StartIp;$i -le $IpRange.EndIp;$i++)
      {
        $FullIpList += '{0}{1}' -f $IpRange.First3Oct, $i # Builds the list of IP Addresses.  This was put in after the fact for use with the workflow.
      }
        
      if($useWorkflow)
      {
        Test-WFConnection -Computers $FullIpList
      }
      else
      {
        foreach($FullIpAdd in $FullIpList)
        {
          Write-Verbose -Message ('Full Address: {0}' -f $FullIpAdd)
          if($useNetconnetion)
          {
            $ping = Test-NetConnection -ComputerName $FullIpAdd -InformationLevel Detailed -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $PingResult = $ping.PingSucceeded
          }
          else 
          {
            $ping = Test-Connection -ComputerName $FullIpAdd -Count 1 -Quiet
            $PingResult = $ping
          }
          if($PingResult -eq 'True')
          {
            Write-Output -InputObject ('Ping Completed: {0}' -f $FullIpAdd)
          }
          else
          {
            Write-Verbose -Message ('Ping Failed: {0}' -f $FullIpAdd)
          }
        }
      }
    }
    else
    {
      Write-Error -Message 'Not v4 Address'
    }
  }
  END
  {}
}



# This can be deleted, it only is used for testing.  
# Finds the local gateway and machine IP and pings that range.
# What it really shows is how long it takes.
$Workstation = $env:COMPUTERNAME
$NicServiceName = (Get-WmiObject -Class win32_networkadapter -Filter 'netconnectionstatus = 2' | Where-Object -FilterScript {
    $_.Description -notmatch 'virtual'
}).ServiceName
$NIC = Get-WmiObject -Class Win32_NetworkAdapterConfiguration |
Where-Object -Property ServiceName -EQ -Value $NicServiceName |
Select-Object -Property *  

Ping-IpRange -FirstIpAddress $($NIC.DefaultIPGateway[0]) -LastIpAddress $($NIC.IPAddress[0]) -Verbose -useWorkflow

