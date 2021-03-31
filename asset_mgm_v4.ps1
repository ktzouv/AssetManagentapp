[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


$Form = New-Object Windows.Forms.Form

#Create Form
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 1000
$System_Drawing_Size.Height = 500
$form.ClientSize = $System_Drawing_Size
$form.text="Reports"
$form.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 75
$System_Drawing_Point.Y = 85


 #Textbox
$labelstartip=New-Object System.Windows.Forms.label
$labelstartip.Location = New-Object System.Drawing.Size(100,20) 
$labelstartip.Size = New-Object System.Drawing.Size(60,20) 
$labelstartip.text="Start Ip"
$form.Controls.Add($labelstartip) 

 #Textbox
$labelendip=New-Object System.Windows.Forms.label
$labelendip.Location = New-Object System.Drawing.Size(170,20) 
$labelendip.Size = New-Object System.Drawing.Size(120,20) 
$labelendip.text="End Ip"
$form.Controls.Add($labelendip) 

$label3parts=New-Object System.Windows.Forms.label
$label3parts.Location = New-Object System.Drawing.Size(10,20) 
$label3parts.Size = New-Object System.Drawing.Size(120,20) 
$label3parts.text="First 3 parts"
$form.Controls.Add($label3parts) 


#Textbox
$textbox1=New-Object System.Windows.Forms.TextBox
$textbox1.Location = New-Object System.Drawing.Size(10,40) 
$textbox1.Size = New-Object System.Drawing.Size(80,20) 
#$textbox1.text="000.000.000"
$form.Controls.Add($textbox1) 

#Textbox
$startip=New-Object System.Windows.Forms.TextBox
$startip.Location = New-Object System.Drawing.Size(100,40) 
$startip.Size = New-Object System.Drawing.Size(40,20) 
$form.Controls.Add($startip) 

$line=New-Object System.Windows.Forms.label
$line.Location = New-Object System.Drawing.Size(150,40) 
$line.Size = New-Object System.Drawing.Size(20,40) 
$line.text="-"
$form.Controls.Add($line) 

$endip=New-Object System.Windows.Forms.TextBox
$endip.Location = New-Object System.Drawing.Size(170,40) 
$endip.Size = New-Object System.Drawing.Size(40,20) 
$form.Controls.Add($endip) 


#Datagrid
$datagridview= New-Object System.Windows.Forms.DataGridView
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 905
$System_Drawing_Size.Height = 250
$datagridview.Size = $System_Drawing_Size
$datagridview.DataBindings.DefaultDataSourceUpdateMode = 0
$datagridview.Name = "dataGrid1"
$datagridview.DataMember = ""
$datagridview.TabIndex = 0
$datagridview.datasource=$null

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 200
$datagridview.Location = $System_Drawing_Point

$form.Controls.Add($datagridview)


#Button
$button=New-Object System.Windows.Forms.Button
$button.Text="Scan"
$button.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 100
$System_Drawing_Size.Height = 24
$button.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 10
$System_Drawing_Point.Y = 130
$button.Location = $System_Drawing_Point
$form.Controls.Add($button) 


#Button
$exportbutton=New-Object System.Windows.Forms.Button
$exportbutton.Text="Export to CSV"
$exportbutton.UseVisualStyleBackColor = $True
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 100
$System_Drawing_Size.Height = 24
$exportbutton.Size = $System_Drawing_Size
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 130
$exportbutton.Location = $System_Drawing_Point
$form.Controls.Add($exportbutton) 








$button.Add_Click({




$a=[int]$startip.text
$b=[int]$endip.text
$c=[string]$textbox1.text

$cred = Get-Credential domain\user

$arr1=@()

#Scan one by one all ip addresses to retrieve Computername and save it in object
#Before start to scan use Test-Connection to identify if PC/Server is online and check if Ws-Man is enable.
#In any case Catch the error and add to an object 

$a..$b |  ForEach {  $address="$c.$_"



        if (Test-Connection -Cn $address -Count 1 -Quiet )
            {

            try{

               if (Test-WSMan -cn $address -ErrorAction Stop)
                    {

                     $o = new-object  psobject
                     $o | add-member -membertype noteproperty -name ops -value (gwmi Win32_Computersystem -computername $address).Caption | Out-Null
                     $o | add-member -membertype noteproperty -name ip -value $address | Out-Null


                     }
    
            }

        

    catch
    
    {
 
         $o = new-object  psobject
         $o | add-member -membertype noteproperty -name ops -value "Wsman isn't enable"| Out-Null
         $o | add-member -membertype noteproperty -name ip -value $address | Out-Null
    }


    }

 else
 {
     $o = new-object  psobject
     $o | add-member -membertype noteproperty -name ops -value "Check if Pc/Server is online or firewall block the scan"
     $o | add-member -membertype noteproperty -name ip -value $address | Out-Null

}




   $arr1 += $o
}



#Scan one by one all the results from the above array to get the information from the pc
#Use Invoke-Command for faster access and Get-Cim to retrieve details

$result=@()


 Foreach($pc in $arr1)
 {
   $ip=$pc.ops

    $Scriptblock= {
             
     $properties=@{
     IPAddress=(Get-CIMInstance -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE).IPAddress
     Computername=(Get-CIMInstance Win32_ComputerSystem ).Caption
     OS=(Get-CIMInstance Win32_Operatingsystem ).Caption
     CPU=(Get-CIMInstance Win32_processor ).Caption
     TotalPhysicalMemory=(Get-CIMInstance Win32_ComputerSystem).TotalPhysicalMemory 
     Disk_C=(Get-CIMInstance Win32_Logicaldisk -filter "deviceid='C:'").Size
     }
      
    $r=New-Object psobject -Property $properties
    $r
         
    
    }


Try
    {

       $result+=Invoke-Command  -cn $ip -Credential $cred -ArgumentList $ip -ScriptBlock $Scriptblock -ErrorAction Stop | Select-Object -property IPAddress,Computername,OS,CPU,TotalPhysicalMemory,Disk_C -ExcludeProperty PSComputername,RunspaceID


   }
   
   catch
   {
     $properties=@{
     IPAddress=$pc.ip
     Computername=$pc.ops
     OS=$pc.ops
     CPU=$pc.ops
     TotalPhysicalMemory=$pc.ops
     Disk_C=$pc.ops

   }
      
    $r=New-Object psobject -Property $properties
    $result+=$r
         
    
  }

 
 }
   
 
$result| Select-Object -property IPAddress,Computername,OS,CPU,TotalPhysicalMemory,HDD -ExcludeProperty PSComputername,RunspaceID
$arrlist=new-object System.Collections.ArrayList
$arrlist.AddRange($result)

$datagridview.DataSource=$arrlist 
$datagridview.AutoResizeColumns()
#$datagridview.AutoResizeRow()

})



#Export Csv file all the results in Datagridview in path that you will select

$exportbutton.Add_Click({

$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null

$datagridview.Rows |
     select -expand DataBoundItem |
     export-csv $OpenFileDialog.FileName -NoType




})





$form.ShowDialog() | Out-Null
