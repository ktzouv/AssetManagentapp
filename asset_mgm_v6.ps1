#requires -Version 2.0 -Modules CimCmdlets
#Build the GUI with WPF
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window" Title="Initial Window" WindowStartupLocation = "CenterScreen" ResizeMode="NoResize"
    Width="1400" Height ="600" ShowInTaskbar = "True" Background = "lightgray"> 

<StackPanel x:Name="Main">

<StackPanel Orientation="Horizontal" HorizontalAlignment="Left"  Margin="20,30,0,0">
   <Label Content="First 3 Parts" Width="100"   />
   <Label Content="Start IP" Width="100" />
   <Label Content="End IP" Width="100" />
</StackPanel>

<StackPanel Orientation="Horizontal" HorizontalAlignment="Left"  Margin="20,10,0,0">
   <TextBox x:Name="first3part"  Width="100" Height="20" />
   <Label Content="."  Width="15" Margin="0,20,0,10" />
   <TextBox x:Name="startip"  Width="60" Height="20" />
   <Label Content="-"  Width="15" Margin="0,15,0,10" />
   <TextBox x:Name="endip"  Width="60" Height="20" />
   <Button x:Name="scan"  Content="Scan" Width="60" Height="20" Margin="20"/>
</StackPanel>

<StackPanel Orientation="Horizontal" HorizontalAlignment="Left"  Margin="10,10,0,0">
   <Button x:Name="exportbutton"  Content="Export to Excel" Width="100" Height="30" Margin="20"/>
</StackPanel>

<ListView Name="datagrid" Grid.Column="10" Grid.Row="0" Margin="30,30,30,30" Height="300" ScrollViewer.VerticalScrollBarVisibility="Visible"  >

            <ListView.View>
            
                <GridView>
                    <GridViewColumn Header="IP Address" Width="150" DisplayMemberBinding="{Binding IPAddress}" />
                    <GridViewColumn Header="Computername" Width="150" DisplayMemberBinding="{Binding Computername}" />
                    <GridViewColumn Header="Hardware" Width="150" DisplayMemberBinding="{Binding Hardware}"/>
                    <GridViewColumn Header="OS" Width="50" DisplayMemberBinding="{Binding OS}" />
                    <GridViewColumn Header="OS Version" Width="150" DisplayMemberBinding="{Binding OSVersion}"/>
                    <GridViewColumn Header="CPU" Width="50" DisplayMemberBinding="{Binding CPU}" />
                    <GridViewColumn Header="Total Physica lMemory (GB)" Width="150" DisplayMemberBinding="{Binding TotalPhysicalMemory}" />
                    <GridViewColumn Header="Free Physical Memory (GB)" Width="150" DisplayMemberBinding="{Binding FreePhysicalMemory}"/>
                     <GridViewColumn Header="Disk C (GB)" Width="150" DisplayMemberBinding="{Binding Disk_C}" />
                    <GridViewColumn Header="Free Space-Disk C" Width="200" DisplayMemberBinding="{Binding Free_Disk_C}"/>
                </GridView>
               
            </ListView.View>
        </ListView>


</StackPanel>



</Window>
"@ 
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml)
$Window = [Windows.Markup.XamlReader]::Load( $reader )
$scan = $Window.FindName('scan')
$datagrid = $Window.FindName('datagrid')
$exportbutton = $Window.FindName('exportbutton')
$startip = $Window.FindName('startip')
$endip = $Window.FindName('endip')
$first3part = $Window.FindName('first3part')
$scan.Add_Click({
    $a = [int]$startip.text
    $b = [int]$endip.text
    $c = [string]$first3part.text
    $cred = Get-Credential domain\user
    $arr1 = @()
    #Scan one by one all ip addresses to retrieve Computername and save it in object
    #Before start to scan use Test-Connection to identify if PC/Server is online and check if Ws-Man is enable.
    #In any case Catch the error and add to an object 
    $a..$b |  ForEach-Object -Process {
      $address = "$c.$_"
      if (Test-Connection -Cn $address -Count 1 -Quiet )
      {
        try
        {
          if (Test-WSMan -cn $address -ErrorAction Stop)
          {
            $o = New-Object  -TypeName psobject
            $null = $o |
            Add-Member -MemberType noteproperty -Name ops -Value (Get-WmiObject -Class Win32_Computersystem -ComputerName $address).Caption
            $null = $o |
            Add-Member -MemberType noteproperty -Name ip -Value $address
          }
        }
        catch
        {
          $o = New-Object  -TypeName psobject
          $null = $o |
          Add-Member -MemberType noteproperty -Name ops -Value "Wsman isn't enable"
          $null = $o |
          Add-Member -MemberType noteproperty -Name ip -Value $address
        }
      }
      else
      {
        $o = New-Object  -TypeName psobject
        $o | Add-Member -MemberType noteproperty -Name ops -Value 'Check if Pc/Server is online or firewall block the scan'
        $null = $o |
        Add-Member -MemberType noteproperty -Name ip -Value $address
      }
      $arr1 += $o
    }
    #Scan one by one all the results from the above array to get the information from the pc
    #Use Invoke-Command for faster access and Get-Cim to retrieve details
    $result = @()
    Foreach($pc in $arr1)
    {
      $ip = $pc.ops
      $Scriptblock = {
        $properties = @{
          IPAddress           = (Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE).IPAddress
          Computername        = (Get-CimInstance Win32_ComputerSystem ).Caption
          Hardware            = (Get-WmiObject -Class win32_computersystem).model
          OS                  = (Get-CimInstance Win32_Operatingsystem ).Caption
          OSVersion           = (Get-CimInstance Win32_operatingSystem).Version
          CPU                 = (Get-CimInstance Win32_processor ).Caption
          TotalPhysicalMemory = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory
          FreePhysicalMemory  = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory
          Disk_C              = (Get-CimInstance Win32_Logicaldisk -filter "deviceid='C:'").Size
          Free_Disk_C         = (Get-CimInstance Win32_Logicaldisk -filter "deviceid='C:'").FreeSpace
        }
        $r = New-Object -TypeName psobject -Property $properties
        $r
      }
      Try
      {
        $result += Invoke-Command  -cn $ip -Credential $cred -ArgumentList $ip -ScriptBlock $Scriptblock -ErrorAction Stop | Select-Object -Property IPAddress, Computername, hardware, OS, OSVersion, CPU, @{
          Name = 'TotalPhysicalMemory'
          e    = {
            $_.TotalPhysicalMemory /1GB -as [int]
          }
        }, @{
          Name = 'FreePhysicalMemory'
          e    = {
            [math]::Round($_.FreePhysicalMemory /1MB , 2)
          }
        }, @{
          Name = 'Disk_C'
          e    = {
            $_.Disk_C /1GB -as [int]
          }
        }, @{
          Name = 'Free_Disk_C'
          e    = {
            $_.Free_Disk_C /1GB -as [int]
          }
        } -ExcludeProperty PSComputername, RunspaceID
      }
      catch
      {
        $properties = @{
          IPAddress           = $pc.ip
          Computername        = $pc.ops
          hardware            = $pc.ops
          OS                  = $pc.ops
          OSVersion           = $pc.ops
          CPU                 = $pc.ops
          TotalPhysicalMemory = $pc.ops
          FreePhysicalMemory  = $pc.ops
          Disk_C              = $pc.ops
          Free_Disk_C         = $pc.ops
        }
        $r = New-Object -TypeName psobject -Property $properties
        $result += $r
      }
    }
    $result| Select-Object -Property IPAddress, Computername, hardware, OS, OSVersion, CPU, TotalPhysicalMemory, FreePhysicalMemory, HDD_C, Free_Disk_C -ExcludeProperty PSComputername, RunspaceID
    $datagrid.ItemsSource = $result
})
#Export Csv file all the results in Datagridview in path that you will select
$exportbutton.Add_Click({
    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.filter = 'CSV (*.csv)| *.csv'
    $null = $OpenFileDialog.ShowDialog()
    $datagrid.Items | Export-Csv -Path $OpenFileDialog.FileName -NoTypeInformation
})
$Window.ShowDialog()