$XAML   = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:MailBoxDelegation"
        x:Name="Window" Title="Délégation de boîte aux lettres" WindowStartupLocation = "CenterScreen" SizeToContent = "WidthAndHeight" ShowInTaskbar = "True" Background = "lightyellow">
    <Grid Margin="0,0,0,33">
        <GroupBox Header="Boîtes aux lettres" HorizontalAlignment="Left" Height="76" Margin="125,10,0,0" VerticalAlignment="Top" Width="691"/>
        <GroupBox Header="Actions" HorizontalAlignment="Left" Height="160" Margin="10,10,0,0" VerticalAlignment="Top" Width="100"/>
        <GroupBox Header="Autorisations" HorizontalAlignment="Left" Height="79" Margin="131,91,0,0" VerticalAlignment="Top" Width="690"/>
        <GroupBox Header="Permissions Effectives" HorizontalAlignment="Left" Height="423" Margin="10,175,0,0" VerticalAlignment="Top" Width="811"/>
        <Label Content="Propriétaire:" HorizontalAlignment="Left" Margin="160,27,0,0" VerticalAlignment="Top" Height="23" Width="70"/>
        <Label Content="Délégué:" HorizontalAlignment="Left" Margin="455,27,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.364,0.509" Height="23" Width="53"/>
        <Label Content="Courriel:" HorizontalAlignment="Left" Margin="147,109,0,0" VerticalAlignment="Top" Height="23" Width="51"/>
        <Label Content="Calendrier:" HorizontalAlignment="Left" Margin="510,109,0,0" VerticalAlignment="Top" RenderTransformOrigin="-3.234,0.81" Height="23" Width="63"/>
        <Label Content="Contacts:" HorizontalAlignment="Left" Margin="329,109,0,0" VerticalAlignment="Top" RenderTransformOrigin="-5.003,0.332" Height="23" Width="55"/>
        <TextBox x:Name="MailBox_TextBox" HorizontalAlignment="Left" Height="23" Margin="160,50,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="250"/>
        <TextBox x:Name="Delegated_Textbox" HorizontalAlignment="Left" Height="23" Margin="455,50,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="250"/>
        <Button x:Name="Get_button" Content="Afficher" HorizontalAlignment="Left" Margin="725,50,0,0" VerticalAlignment="Top" Width="75" Height="19"/>
        <Button x:Name="Set_button" Content="Enregistrer" HorizontalAlignment="Left" Margin="23,57,0,0" VerticalAlignment="Top" Width="75" Height="19"/>
        <Button x:Name="Quit_button" Content="Quitter" HorizontalAlignment="Left" Margin="23,137,0,0" VerticalAlignment="Top" Width="75" Height="19"/>
        <Button x:Name="Connect_button" Content="Connecter" HorizontalAlignment="Left" Margin="23,32,0,0" VerticalAlignment="Top" Width="75" Height="19"/>
        <ComboBox x:Name="Contact_comboBox" HorizontalAlignment="Left" Margin="329,132,0,0" VerticalAlignment="Top" Width="160" Height="19"/>
        <ComboBox x:Name="Calendar_comboBox" HorizontalAlignment="Left" Margin="510,132,0,0" VerticalAlignment="Top" Width="160" Height="19"/>
        <ComboBox x:Name="MailBox_comboBox" HorizontalAlignment="Left" Margin="150,132,0,0" VerticalAlignment="Top" Width="160" Height="19"/>
        <DataGrid x:Name="Result_Grid" AlternatingRowBackground = 'LightBlue'  AlternationCount='2' CanUserAddRows='False' HorizontalAlignment="Right" Margin="0,195,21,23" Width="787"/>
        <StatusBar x:Name="Info_StatusBar" HorizontalAlignment="Left" Height="24" Margin="23,610,0,-26" VerticalAlignment="Top" Width="787"/>
        <CheckBox x:Name="CheckBox_SendAs" Content="Send As" HorizontalAlignment="Left" Margin="718,114,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="CheckBox_SendOnBehalf" Content="Send On Behalf" HorizontalAlignment="Left" Margin="718,137,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
'@
 
function Convert-XAMLtoWindow
{
  param
  (
    [Parameter(Mandatory)]
    [string]
    $XAML,
   
    [string[]]
    $NamedElement=$null,
   
    [switch]
    $PassThru
  )
 
  Add-Type -AssemblyName PresentationFramework
 
  $reader = [XML.XMLReader]::Create([System.IO.StringReader]$XAML)
  $result = [Windows.Markup.XAMLReader]::Load($reader)
  foreach($Name in $NamedElement)
  {
    $result | Add-Member NoteProperty -Name $Name -Value $result.FindName($Name) -Force
  }
 
  if ($PassThru)
  {
    $result
  }
  else
  {
    $null = $window.Dispatcher.InvokeAsync{
      $result = $window.ShowDialog()
      Set-Variable -Name result -Value $result -Scope 1
    }.Wait()
    $result
  }
}
 
 
function Show-WPFWindow
{
  param
  (
    [Parameter(Mandatory)]
    [Windows.Window]
    $Window
  )
 
  $result = $null
  $null = $window.Dispatcher.InvokeAsync{
    $result = $window.ShowDialog()
    Set-Variable -Name result -Value $result -Scope 1
  }.Wait()
  $result
}
 
$window = Convert-XAMLtoWindow -XAML $XAML -NamedElement 'Calendar_comboBox', 'CheckBox_SendAs', 'CheckBox_SendOnBehalf', 'Connect_button', 'Contact_comboBox', 'Delegated_Textbox', 'Get_button', 'Info_StatusBar', 'MailBox_comboBox', 'MailBox_TextBox', 'Quit_button', 'Result_Grid', 'Set_button', 'Window' -PassThru
 
 
 
 
#===========================================================================
# Connect to Controls
#===========================================================================
## CheckBox
$CheckBox_SendAs = $Window.FindName('CheckBox_SendAs')
$CheckBox_SendOnBehalf = $Window.FindName('CheckBox_SendOnBehalf')
 
## StatusBar
$Info_StatusBar = $Window.FindName('Info_StatusBar')
 
## DataGrid
$Result_Grid = $Window.FindName('Result_Grid')
 
## ComboBox
$MailBox_comboBox = $Window.FindName('MailBox_comboBox')
$Calendar_comboBox = $Window.FindName('Calendar_comboBox')
$Contact_comboBox = $Window.FindName('Contact_comboBox')
 
## Button
$Get_button = $Window.FindName('Get_button')
 
## $Set_button = $Window.FindName('Set_button')
$Quit_button = $Window.FindName('Quit_button')
$Connect_button = $Window.FindName('Connect_button')
 
## TextBox
$MailBox_TextBox = $Window.FindName('MailBox_TextBox')
$Delegated_Textbox = $Window.FindName('Delegated_Textbox')
 
 
#===========================================================================
## Populate DropDown Lists
#===========================================================================
 
#list Mailbox Permissions
$MailboxAccess=@("None","Author","Contributor","Editor","Reviewer","FullAccess")
$MailboxAccess|ForEach-object {$MailBox_comboBox.AddChild($_)}
 
# List Calendar and Contacts permissions
$CalendarAccess=@("None","Author","Contributor","Editor","Reviewer","Owner")
$CalendarAccess|ForEach-object {$Calendar_comboBox.AddChild($_)}
$CalendarAccess|ForEach-object {$Contact_comboBox.AddChild($_)}
 
#===========================================================================
# Events
#===========================================================================
 
## Checked Events
[System.Windows.RoutedEventHandler]$Script:CheckBoxChecked = {
  $Window.Content.Children | Where {
    $_ -is [System.Windows.Controls.CheckBox] -AND $This.Name  -ne $_.Name
  } | ForEach {
    $_.IsChecked = $False
  }
}
$Window.Content.Children | Where {
  $_ -is [System.Windows.Controls.CheckBox]
} | ForEach {
  $_.AddHandler([System.Windows.Controls.CheckBox]::CheckedEvent, $CheckBoxChecked)
}
 
## Get Button
$window.Get_button.add_Click({
    $Delegate = $Delegated_Textbox.Text
    $MBXOwner = $MailBox_TextBox.Text
   
    
    If (-NOT ([string]::IsNullOrEmpty($MBXOwner))){
          If ([string]::IsNullOrEmpty($Delegate))  {
            Try 
                    {
                      $window.Info_StatusBar.items.Clear()       
                      $MBXResults = Get-MailBoxPermission ($MBXOwner) | Where{$_.IsInherited -eq $false} | Select-Object Identity,User,AccessRights,Deny,IsInherited
                      $Result_Grid.Clear()
                      $window.Info_StatusBar.AddText("Permissions sur la boite de courriels $($MBXOwner)")
                   }
 
             Catch  {
                     $info = Write-Warning  $_
                     $Info_StatusBar.Items.Clear()
                     $Info_StatusBar.AddText("$($info)")
 
                   }
                                                  }
                                                 }
                                      
                                            
                                            
                                            
   
    If (-NOT ([string]::IsNullOrEmpty($MBXOwner))){ 
              If (-NOT ([string]::IsNullOrEmpty($Delegate)))  {
                                Try  {
                                        $window.Info_StatusBar.items.Clear()       
                                        $MBXResults = Get-MailBoxPermission -Identity $MBXOwner -User $Delegate |Select-Object Identity,User,AccessRights
                                        $Result_Grid.Clear()
                                        $window.Info_StatusBar.AddText("Permissions de $($Delegate) sur la boite de courriels $($MBXOwner)")
 
                                     }
 
                                Catch  {
                                        $info = Write-Warning  $_
                                        $Info_StatusBar.Items.Clear()
                                        $Info_StatusBar.AddText("$($info)")
 
                                       }
                                                          }
                                                      }
          [PSCustomObject]$paramList = @()
          Foreach ($MBXResult in $MBXResults){
                                                $param=[PSCustomObject]@{
                                                Propriétaire=$MBXResult.Identity
                                                Délégué = $MBXResult.User
                                                Accés = $MBXResult.AccessRights
                                                                        }
          $paramList+=$param
                                             }
                           
          $Result_Grid.ItemsSource =  $paramList
   
  })
 
 
     
 
## Set Button
$Window.Set_Button.Add_Click({
    $Info_StatusBar.Items.Clear()
    If ((-NOT [string]::IsNullOrEmpty($MailBox_comboBox.text))) {
      $Info_StatusBar.Items.Add($MailBox_comboBox.text)
     
    }
})
 
## Connect Button
$Connect_Button.Add_Click({
    Get-PSSession | Remove-PSSession
    $UserCredential = Get-Credential -Credential $env:username'@fondsftq.com'
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession -Session $Session
    $window.Activate()
 
    if($Session){
      $Info_StatusBar.Items.Clear()
      $Info_StatusBar.AddText("Connecté à Office 365")
    }
    Else{
      $Info_StatusBar.Items.Clear()
      $Info_StatusBar.AddText("La Connection à Office 365 a échouée. Utilisateur ou mot de passe incorrecte")
    }
  }  
)
## Quit Button
$Quit_Button.Add_Click({
    $session = Get-PSSession
    Remove-PSSession -Session $session
    $Window.Close()
    })
 
 
#===========================================================================
# Shows the form
#===========================================================================
Show-WPFWindow -Window $window
