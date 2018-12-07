## Code to hide the powershell command window when GUI is running
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

   #ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Class="BlogPostIII.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:BlogPostIII"
        mc:Ignorable="d"
        Title="MRG User Account Creation Tool" Height="350" Width="616.976">
    <Grid x:Name="background" Background="#FF1D3245">
        <Button x:Name="MakeUserbutton" Content="Create User" HorizontalAlignment="Left" Height="34" Margin="10,277,0,0" VerticalAlignment="Top" Width="155" FontSize="14.667"/>
        <CheckBox x:Name="checkBox" Content="Temporary User?" HorizontalAlignment="Left" Height="36" Margin="10,176,0,0" VerticalAlignment="Top" Width="140" FontSize="14.667" Foreground="White"/>
        <Label x:Name="label" Content="Use this tool to create a new user&#xD;&#xA;Note: The new user will have to change their password on first login" HorizontalAlignment="Left" Height="47" Margin="10,0,0,0" VerticalAlignment="Top" Width="415" Foreground="White"/>
        <TextBox x:Name="firstName" HorizontalAlignment="Left" Height="25" Margin="165,47,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="119" FontSize="14.667"/>
        <TextBlock x:Name="firstName_label" HorizontalAlignment="Left" Height="25" Margin="19,47,0,0" TextWrapping="Wrap" Text="First Name" VerticalAlignment="Top" Width="118" Background="#FF98BCD4" FontSize="16"/>
        <TextBox x:Name="lastName" HorizontalAlignment="Left" Height="25" Margin="165,89,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="119" FontSize="14.667"/>
        <TextBlock x:Name="lastName_label" HorizontalAlignment="Left" Height="25" Margin="19,89,0,0" TextWrapping="Wrap" Text="Last Name" VerticalAlignment="Top" Width="118" Background="#FF98BCD4" FontSize="16"/>
        <TextBox x:Name="logonName" HorizontalAlignment="Left" Height="25" Margin="165,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="119" FontSize="14.667"/>
        <TextBlock x:Name="logonName_label" HorizontalAlignment="Left" Height="25" Margin="19,130,0,0" TextWrapping="Wrap" Text="Logon Name" VerticalAlignment="Top" Width="118" Background="#FF98BCD4" FontSize="16"/>
        <Separator HorizontalAlignment="Left" Height="30" Margin="19,155,0,0" VerticalAlignment="Top" Width="265" RenderTransformOrigin="-1.016,-0.225"/>
        <RadioButton x:Name="radioButton_7" Content="7 Days" HorizontalAlignment="Left" Height="20" Margin="42,200,0,0" VerticalAlignment="Top" Width="108" RenderTransformOrigin="0.861,0.107" FontSize="14.667" Background="#FFF9F2F2" Foreground="White" IsEnabled="False"/>
        <RadioButton x:Name="radioButton_30" Content="30 Days" HorizontalAlignment="Left" Height="20" Margin="42,225,0,0" VerticalAlignment="Top" Width="108" RenderTransformOrigin="0.861,0.107" FontSize="14.667" Background="#FFF9F2F2" Foreground="White" IsEnabled="False"/>
        <RadioButton x:Name="radioButton_90" Content="90 Days" HorizontalAlignment="Left" Height="20" Margin="42,250,0,0" VerticalAlignment="Top" Width="108" RenderTransformOrigin="0.861,0.107" FontSize="14.667" Background="#FFF9F2F2" Foreground="White" IsEnabled="False"/>
        <TextBox x:Name="password" HorizontalAlignment="Left" Height="25" Margin="441,47,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="119" FontSize="14.667" RenderTransformOrigin="0.5,0.5"/>
        <TextBlock x:Name="password_label" HorizontalAlignment="Left" Height="25" Margin="303,47,0,0" TextWrapping="Wrap" Text="Temp Password" VerticalAlignment="Top" Width="122" Background="#FF98BCD4" FontSize="16"/>
        <TextBlock x:Name="targetOU_label" HorizontalAlignment="Left" Height="25" Margin="303,89,0,0" TextWrapping="Wrap" Text="Target OU" VerticalAlignment="Top" Width="122" Background="#FF98BCD4" FontSize="16"/>
        <ComboBox x:Name="targetOU_comboBox" HorizontalAlignment="Left" Margin="441,89,0,0" VerticalAlignment="Top" Width="120"/>
        <TextBlock x:Name="DefaultOUMsg" Text ="If TargetOu not specified, user will be placed in the @anchor OU" HorizontalAlignment="Left" Margin="303,130,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Foreground="#FFFBFAFA" FontSize="14.667"/>
    </Grid>
</Window>
"@        

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'


[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML

    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}

Get-FormVariables

#===========================================================================
# Actually make the objects work
#===========================================================================
##Create a Windows Scripting host shell instance to support interactive popups.
$wshell = New-Object -ComObject Wscript.Shell
#Resolve the default OU to show where the user will end up
$defaultOU = (get-adobject -filter 'ObjectClass -eq "domain"' -Properties wellKnownObjects).wellknownobjects.Split("`n")[-1].Split(':') | select -Last 1
 $WPFDefaultOUMsg.Text = $WPFDefaultOUMsg.Text -replace "@anchor",$defaultOU

#gather all of the settings the user specifies, needed to splat to the New-ADUser Cmd later
function Get-FormFields {
$TargetOU = if ($WPFtargetOU_comboBox.Text -ne $null){$WPFtargetOU_comboBox.Text}else{$defaultOU}
if ($WPFcheckBox.IsChecked){
    $ExpirationDate = if ($WPFradioButton_7.IsChecked -eq $true){7}`
                elseif ($WPFradioButton_30.IsChecked -eq $true){30}`
                elseif ($WPFradioButton_90.IsChecked -eq $true){90}
    
    $ExpirationDate = (get-date).AddDays($ExpirationDate)
    
    $HashArguments = 
        @{ Name = $WPFlogonName.Text;
           GivenName=$WPFfirstName.Text;
           SurName = $WPFlastName.Text;
           AccountPassword=($WPFpassword.text | ConvertTo-SecureString -AsPlainText -Force);
           #Added must change password at next logon atributre
           ChangePasswordAtLogon = $true;
           AccountExpirationDate = $ExpirationDate;
           Path=$TargetOU;
            }
        }
    else{
    $HashArguments = 
       @{ Name = $WPFlogonName.Text;
          GivenName=$WPFfirstName.Text;
          SurName = $WPFlastName.Text;
          AccountPassword=($WPFpassword.text | ConvertTo-SecureString -AsPlainText -Force);
          #Added must change password at next logon atribute
          ChangePasswordAtLogon = $true;
          Path=$TargetOU;
          }
    }
$HashArguments
}
function formValidate{
    if( $WPFlogonName.Text -eq "" -or $WPFfirstName.Text -eq "" -or $WPFlastName.Text -eq "" -or $WPFpassword.text -eq ""){
        return $false
    }
    else{
        return $true
    }
}

$defaultOU,"OU=SBSUsers,OU=Users,OU=MyBusiness,DC=medfordradiology,DC=local" | ForEach-object {$WPFtargetOU_comboBox.AddChild($_)}

#Add logic to the checkbox to enable items when checked
$WPFcheckBox.Add_Checked({
    $WPFradioButton_7.IsEnabled=$true
   $WPFradioButton_30.IsEnabled=$true
   $WPFradioButton_90.IsEnabled=$true
    $WPFradioButton_7.IsChecked=$true
    })

$WPFcheckBox.Add_UnChecked({
    $WPFradioButton_7.IsEnabled=$false
   $WPFradioButton_30.IsEnabled=$false
   $WPFradioButton_90.IsEnabled=$false
    $WPFradioButton_7.IsChecked,$WPFradioButton_30.IsChecked,$WPFradioButton_90.IsChecked=$false,$false,$false})

#$WPFMakeUserbutton.Add_Click({(Get-FormFields)})
$WPFMakeUserbutton.Add_Click({
    #Check form is filled
   if (formValidate){
        #Resolve Form Settings
        $hash = Get-FormFields
        New-ADUser @hash -PassThru

        #create and resolve spiceworks ticket w/ 1hr worked.
        $mycredentials = Get-Credential -Message "Please provide the email address and password accociated with your ticketing purposes"
        $SmtpServer = 'smtp.office365.com'
        $MailtTo = 'helpdesk@medfordradiology.com'
        $MailFrom = $mycredentials.UserName 
        $MailSubject = "Created New User: " + $WPFlogonName.Text
        $emailbody = "New User account for: " + $WPFlogonName.Text + " Created. Please contact IT for more infomraiton `n #add 1h `n #close"

        Send-MailMessage -To "$MailtTo" -from "$MailFrom" -Subject $MailSubject -Body $emailbody -SmtpServer $SmtpServer -UseSsl -Port 587 -Credential $mycredentials 
        $wshell.Popup("Success: Account created, is SBSUsers OU was selected please allow upto three hours for Office 365 to sync the account", 5, "MRG New Employee Acount Creator", 0x30)
        #add code here to contact HR to set up payroll information.
        $Form.Close()
    }
    else{
        $wshell.Popup("ERROR Not all fields are filled.", 5, "MRG New Employee Acount Creator", 0x30)
    }
    })


#===========================================================================
# Shows the form
#===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan

function Show-Form{
$Form.ShowDialog() | out-null

}

Show-Form