<#	
    .NOTES
    ===========================================================================
    Created with: 	Powershell
    Created on:   	16.11.2022
    Created by:   	 Martin Heinc
    Organization: 	ACP IT Solutions GmbH
    ===========================================================================
    .DESCRIPTION
    
#>

### Init Script
##### load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
$TempPath = "X:\temp"
If (!(Test-Path -Path $TempPath)) {
  New-Item -Path $TempPath -Force -ItemType Directory
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if((Test-NetConnection -ComputerName "hhviesmp01.HOUSING.HOSTING").PingSucceeded -eq $True){
    $Server = "HHVIESMP01.housing.hosting"
}
else {
    $Server = "127.0.0.1:59100"
}
$Webserver = "https://$Server/Altiris/NS/NSCap/Bin/Deployment/ACP/Repo/TS/TSData"
$xaml = Invoke-WebRequest -Uri "$webserver/GUI.xaml" -UseBasicParsing
$Settingfile = Invoke-RestMethod -Uri "$webserver/settings.json" -UseBasicParsing
Invoke-WebRequest -Uri "$webserver/ACPTools.png" -UseBasicParsing -OutFile "$tempPath\ACPTools.png"

<#
$headerlogo = New-Object -TypeName System.Windows.Media.Imaging.BitmapImage
$headerlogo.BeginInit()
$headerlogo.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($headerlogo_base64)
$headerlogo.EndInit()
$headerlogo.Freeze()
#>



function Convert-XAMLtoWindow {
  param
  (
    [Parameter(Mandatory = $true)]
    [string]
    $xaml
  )
  
  Add-Type -AssemblyName PresentationFramework
  
  $reader = [XML.XMLReader]::Create([IO.StringReader]$xaml)
  $result = [Windows.Markup.XAMLReader]::Load($reader)
  $reader.Close()
  $reader = [XML.XMLReader]::Create([IO.StringReader]$xaml)
  while ($reader.Read()) {
    $name = $reader.GetAttribute('Name')
    if (!$name) {
      $name = $reader.GetAttribute('x:Name') 
    }
    if ($name) {
      $result | Add-Member -MemberType NoteProperty -Name $name -Value $result.FindName($name) -Force
    }
  }
  $reader.Close()
  $result
}


function Show-WPFWindow {
  param
  (
    [Parameter(Mandatory = $true)]
    [Windows.Window]
    $Window
  )

  $result = $null
  $null = $Window.Dispatcher.InvokeAsync{
    $result = $Window.ShowDialog()
    Set-Variable -Name result -Value $result -Scope 1
  }.Wait()
  $result
}

$Window = Convert-XAMLtoWindow -XAML $xaml 

# add click handlers
$Window.ButCancel.add_Click{
  # close window
  $Window.DialogResult = $false
}

$Window.ButOk.add_Click{
  ###Check if Textbox has Path
  $CheckClientname = $window.Clientname.Text
   
 
  if ($CheckClientname.Length -le '1') {
    [System.Windows.Forms.MessageBox]::Show('Please enter Clientname', 'There is no Clientname!', 0, 'Error')
  }
  elseif ($CheckClientname.Length -gt '15') {
    [System.Windows.Forms.MessageBox]::Show('Clientname is to long, max. 15 Chars.', 'There is no Clientname!', 0, 'Error')
  }
  Else {  
    $SelectedKunde = $window.Kundenname.text
    $SelectedOS = $window.Winversion.text
    
    $Settings = $Settingfile.Kunden | Where-Object { $PSItem.Name -like $SelectedKunde }
    $OSlang = if ($Settings.OS_Lang -eq 'de') {
      'de-de'
    }
    elseif ($Settings.OS_Lang -eq 'en') {
      'en-us'
    }
    elseif ($Settings.OS_Lang -eq 'gb') {
      'en-gb'
    }
    
    $properties = @{'ComputerName' = $CheckClientname;
      'KUNDENNAME'                 = $Settings.Kundenname;
      'KUNDENShort'                = $Settings.Kundenshort;
      'KUNDENNUMMER'               = $Settings.Kundennummer;
      'Webrootgroup'               = $Settings.Webrootgroup;
      'WebrootKey'                 = $Settings.WebrootKey;
      '7Zip'                       = $Settings.'7Zip';
      'Office'                     = $Settings.Office;
      'Chrome'                     = $Settings.Chrome;
      'FireFox'                    = $Settings.FireFox;
      'AdobeReaderDC'              = $Settings.AdobeReaderDC;
      'Itunes'                     = $Settings.Itunes;
      'VLC'                        = $Settings.VLC;
      'WorkspaceApp'               = $Settings.WorkspaceApp;
      'Exitcode'                   = $Settings.Exitcode;
      'Language'                   = $OSlang;
      'WindowsVersion'             = $SelectedOS
    }
    $output = New-Object -TypeName PSObject -Property $properties
    $Exportpath = "x:\Temp"
    if (!(Test-Path -Path $Exportpath)) {
      New-Item -Path $Exportpath -Force -Confirm:$false -ItemType Directory
    }
    Export-Csv -InputObject $output -Path "$Exportpath\Settings.csv" -NoTypeInformation -Force
        
    $window.DialogResult = $true 
  }       
}

# fill the combobox Kundenname
$Kunden = $Settingfile.Kunden.Name
$OsVersions = $Settingfile.OSVersion.Name

$Window.Kundenname.ItemsSource = $Kunden
$Window.Kundenname.SelectedIndex = 0

$Window.Winversion.ItemsSource = $OsVersions
$Window.Winversion.SelectedIndex = 0

$Window.headerLogo.Source = "$tempPath\ACPTools.png"
#$Window.WinLogo.Source = $script:Winlogo
Show-WPFWindow -Window $Window
#endregion