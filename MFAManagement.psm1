function Get-MFAUserInfo {

    param([string]$UserPrincipalName)

    if ($UserPrincipalName -ne "") {
        try {
            $msolUser = (Get-MsolUser -SearchString $UserPrincipalName -ErrorAction Stop)[0]
            if ($msolUser) {
                try { 
                    $emailAddress = (($msolUser.ProxyAddresses | Where-Object {$_ -cmatch "SMTP:"}) -replace "SMTP:","").ToLower()
                } catch { $emailAddress = "" }

                $mfaInfo  = [PSCustomObject]@{
                    Name              = $msolUser.DisplayName
                    Department        = $msolUser.Department
                    JobTitle          = $msolUser.Title
                    MobilePhone       = $msolUser.MobilePhone
                    EmailAddress      = $emailAddress
                    UserPrincipalName = $msolUser.UserPrincipalName
                    DefaultMethod     = ($msolUser.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true}).MethodType
                    State             = $msolUser.StrongAuthenticationRequirements.State
                    UserDetails       = $msolUser.StrongAuthenticationUserDetails
                    PhoneAppDetails   = $msolUser.StrongAuthenticationPhoneAppDetails
                    LastResetDate     = $msolUser.StrongAuthenticationRequirements.RememberDevicesNotIssuedBefore
                }
                if (!$mfaInfo.State) { $mfaInfo.State = "Disabled" }
                return $mfaInfo
            }
        } catch {
            Show-Message -Message "Couldn't found the user in MSOL" -Title "User not found" -ButtonSet OKCancel -Icon Warning
        }
    } else { return }
}

function Reset-MFA {

    param([string]$UserPrincipalName)

    try {
        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationMethods @() -ErrorAction Stop
    } catch {
        $errorMessage = $_.Exception.Message
        Show-Message -Message $errorMessage -Title "Reset MFA" -ButtonSet OKCancel -Icon Error
    }
}

function Enable-MFA {

    param([string]$UserPrincipalName)

    $mfa = @(
        [Microsoft.Online.Administration.StrongAuthenticationRequirement]@{
            RelyingParty = "*"
            State = "Enabled"
        }
    )
    try {
        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements $mfa -ErrorAction Stop
    } catch {
        $errorMessage = $_.Exception.Message
        Show-Message -Message $errorMessage -Title "Enable MFA" -ButtonSet OKCancel -Icon Error
    }
}

function Disable-MFA {

    param([string]$UserPrincipalName)

    try {
        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements @() -ErrorAction Stop
    } catch {
        $errorMessage = $_.Exception.Message
        Show-Message -Message $errorMessage -Title "Disable MFA" -ButtonSet OKCancel -Icon Error
    }
}

function Set-MFADefaultMethod {

    param(
        [string]$UserPrincipalName,
        [ValidateSet("PhoneAppOTP","PhoneAppNotification","OneWaySMS","TwoWayVoiceMobile",$null)][string]$Method
    )

    if (!$Method) { return }

    $mfa = @(
        [Microsoft.Online.Administration.StrongAuthenticationMethod]@{
            IsDefault  = $true
            MethodType = $Method
        }
    )

    try {
        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationMethods $mfa -ErrorAction Stop
    } catch {
        $errorMessage = $_.Exception.Message
        Show-Message -Message $errorMessage -Title "Default MFA method" -ButtonSet OKCancel -Icon Error
    }
}

function Get-MFAUserInterface {

    [xml]$Xml = @'
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MFA" Height="700" Width="550" MinWidth="480" MinHeight="390">
    <Grid Name="gridMain">

        <!-- Grid colomn & row definition -->
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="100"></ColumnDefinition>
            <ColumnDefinition Width="213*"></ColumnDefinition>
            <ColumnDefinition Width="8*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="120"></RowDefinition>
            <RowDefinition/>
            <RowDefinition Height="42"></RowDefinition>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Label Name="labelTitle" Foreground="Black" Content="Multi-factor authentication (MFA)" Grid.Column="1" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" FontSize="24" Height="42" Width="Auto" FontWeight="Bold"/>
        <TextBlock Grid.Column="1" HorizontalAlignment="Stretch" Margin="15,55,15,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="48"><Run Text="Multi-factor authentication is a process where a user is prompted during the sign-in process for an additional form of identification, such as to enter a code on their cellphone or to provide a fingerprint scan."/></TextBlock>
        <Separator HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Grid.Column="1"/>

        <!-- Side pannel -->
        <Rectangle Name="rectangleSidePannel" Fill="#E6E6E6" Height="Auto" VerticalAlignment="Stretch" Grid.RowSpan="3" Margin="2,0,1,0"/>
        <!-- <Image Height="70" Width="70" Source="C:\temp\logo_roullier_menu.jpg" VerticalAlignment="Center" HorizontalAlignment="Center" Grid.Column="0" Grid.RowSpan="1"/> -->
        <Viewbox Width="70" Height="70" Grid.Column="0" Grid.RowSpan="1">
            <Canvas Width="24" Height="24">
                <Path Fill="DodgerBlue" Data="M12,1L3,5V11C3,16.55 6.84,21.74 12,23C17.16,21.74 21,16.55 21,11V5L12,1M12,7C13.4,7 14.8,8.1 14.8,9.5V11C15.4,11 16,11.6 16,12.3V15.8C16,16.4 15.4,17 14.7,17H9.2C8.6,17 8,16.4 8,15.7V12.2C8,11.6 8.6,11 9.2,11V9.5C9.2,8.1 10.6,7 12,7M12,8.2C11.2,8.2 10.5,8.7 10.5,9.5V11H13.5V9.5C13.5,8.7 12.8,8.2 12,8.2Z" />
            </Canvas>
        </Viewbox>

        <!-- Body-->
        <ScrollViewer Grid.Row="1" Grid.Column="1" HorizontalAlignment="Stretch" Height="Auto" VerticalAlignment="Stretch" Width="Auto" Grid.ColumnSpan="2">
            <Grid>

                <!-- Body grid colomn & row definition -->
                <Grid.ColumnDefinitions>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="70"></RowDefinition>
                    <RowDefinition Height="150"></RowDefinition>
                    <RowDefinition></RowDefinition>
                    <RowDefinition Height="70"></RowDefinition>
                </Grid.RowDefinitions>

                <!-- Search user -->
                <Grid Grid.Row="0" Margin="5">

                    <Grid.ColumnDefinitions>
                        <ColumnDefinition></ColumnDefinition>
                        <ColumnDefinition Width="26"></ColumnDefinition>
                        <ColumnDefinition Width="138"></ColumnDefinition>
                    </Grid.ColumnDefinitions>

                    <TextBox Name="textboxSearch" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" VerticalContentAlignment="Center" Width="Auto" Grid.Column="0" Grid.ColumnSpan="2" Margin="15,15,0,15"/>
                    <Button Name="buttonClearSearch" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Right" VerticalAlignment="Center" Height="26" Grid.Column="1" Grid.Row="0">
                        <Path Fill="LightGray"  Stretch="Fill" Margin="3" Data="M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z"/>
                    </Button>
                    <Button Name="buttonSearch" Background="DodgerBlue" Foreground="White" BorderThickness="0" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="2" Margin="15">
                        <Path Fill="White" Stretch="Uniform" Margin="5" Data="M9.5,3A6.5,6.5 0 0,1 16,9.5C16,11.11 15.41,12.59 14.44,13.73L14.71,14H15.5L20.5,19L19,20.5L14,15.5V14.71L13.73,14.44C12.59,15.41 11.11,16 9.5,16A6.5,6.5 0 0,1 3,9.5A6.5,6.5 0 0,1 9.5,3M9.5,5C7,5 5,7 5,9.5C5,12 7,14 9.5,14C12,14 14,12 14,9.5C14,7 12,5 9.5,5Z"/>
                    </Button>
                    <Separator HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Grid.ColumnSpan="3" Height="1"/>

                </Grid>

                <!-- User info -->
                <Grid Name="gridUserInfo" Grid.Row="1" Margin="5">

                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"></ColumnDefinition>
                        <ColumnDefinition></ColumnDefinition>
                        <ColumnDefinition Width="26"></ColumnDefinition>
                        <ColumnDefinition Width="26"></ColumnDefinition>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="30"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                    </Grid.RowDefinitions>

                    <!-- DisplayName & Refresh button-->
                    <Label Name="labelDisplayName" Content="User information" HorizontalAlignment="Stretch" VerticalAlignment="Bottom" Grid.ColumnSpan="2" Grid.Row="0" FontWeight="Bold" FontSize="14"/>
                    <Button Name="buttonRefresh" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="3" Grid.Row="0">
                        <Path Fill="LightGray" Stretch="Uniform" Margin="3" Data="M17.65,6.35C16.2,4.9 14.21,4 12,4A8,8 0 0,0 4,12A8,8 0 0,0 12,20C15.73,20 18.84,17.45 19.73,14H17.65C16.83,16.33 14.61,18 12,18A6,6 0 0,1 6,12A6,6 0 0,1 12,6C13.66,6 15.14,6.69 16.22,7.78L13,11H20V4L17.65,6.35Z"/>
                    </Button>

                    <!-- Company -->
                    <Label Content="Department" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Row="1"/>
                    <Label Name="labelDepartment" Content="" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Row="1" Grid.Column="1"/>

                    <!-- Job title -->
                    <Label Content="Job title" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Row="2"/>
                    <Label Name="labelJobTitle" Content="" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="1" Grid.Row="2"/>

                    <!-- Mobile phone -->
                    <Label Content="Mobile phone" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Row="3"/>
                    <Label Name="labelPhoneNumber" Content="" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="1" Grid.Row="3"/>
                    <Button Name="buttonPhoneTo" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="2" Grid.Row="3">
                        <Path Fill="LightGray" Stretch="Fill" Margin="3" Data="M6.62,10.79C8.06,13.62 10.38,15.94 13.21,17.38L15.41,15.18C15.69,14.9 16.08,14.82 16.43,14.93C17.55,15.3 18.75,15.5 20,15.5A1,1 0 0,1 21,16.5V20A1,1 0 0,1 20,21A17,17 0 0,1 3,4A1,1 0 0,1 4,3H7.5A1,1 0 0,1 8.5,4C8.5,5.25 8.7,6.45 9.07,7.57C9.18,7.92 9.1,8.31 8.82,8.59L6.62,10.79Z"/>
                    </Button>
                    <Button Name="buttonPhoneNumber" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="3" Grid.Row="3">
                        <Path Fill="LightGray" Stretch="Fill" Margin="3" Data="M19,21H8V7H19M19,5H8A2,2 0 0,0 6,7V21A2,2 0 0,0 8,23H19A2,2 0 0,0 21,21V7A2,2 0 0,0 19,5M16,1H4A2,2 0 0,0 2,3V17H4V3H16V1Z"/>
                    </Button>

                    <!-- Email address -->
                    <Label Content="Email address" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Row="4"/>
                    <Label Name="labelEmailAddress" Content="" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="1" Grid.Row="4" Foreground="#006CBE"/>
                    <Button Name="buttonMailTo" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="2" Grid.Row="4">
                        <Path Fill="LightGray" Stretch="Uniform" Margin="3" Data="M22 6C22 4.9 21.1 4 20 4H4C2.9 4 2 4.9 2 6V18C2 19.1 2.9 20 4 20H20C21.1 20 22 19.1 22 18V6M20 6L12 11L4 6H20M20 18H4V8L12 13L20 8V18Z"/>
                    </Button>
                    <Button Name="buttonEmailAddress" Background="Transparent" BorderBrush="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Grid.Column="3" Grid.Row="4">
                        <Path Fill="LightGray" Stretch="Fill" Margin="3" Data="M19,21H8V7H19M19,5H8A2,2 0 0,0 6,7V21A2,2 0 0,0 8,23H19A2,2 0 0,0 21,21V7A2,2 0 0,0 19,5M16,1H4A2,2 0 0,0 2,3V17H4V3H16V1Z"/>
                    </Button>


                </Grid>

                <!-- MFA Configuration -->
                <Grid Name="gridMFAconfiguration" Grid.Row="2" Margin="5">

                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"></ColumnDefinition>
                        <ColumnDefinition Width="19"></ColumnDefinition>
                        <ColumnDefinition></ColumnDefinition>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="30"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="26"></RowDefinition>
                        <RowDefinition Height="*"></RowDefinition>
                    </Grid.RowDefinitions>

                    <Label Content="MFA configuration" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.ColumnSpan="3" Grid.Row="0" FontSize="14" FontWeight="Bold" Height="30"/>

                    <!-- State -->
                    <Label Content="State" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="0" Grid.Row="1" Height="26"/>
                    <Viewbox Grid.Column="1" Grid.Row="1">
                        <Canvas Width="24" Height="24">
                            <Path Name="pathState" Fill="Transparent" Data=""/>
                        </Canvas>
                    </Viewbox>
                    <Label Name="labelState" Content="" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="2" Grid.Row="1" Height="26"/>

                    <!-- Default method -->
                    <Label Content="Default method" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="0" Grid.Row="2" Height="26"/>
                    <ComboBox Name="comboboxDefaultMethod" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="2" Margin="2">
                        <ComboBoxItem Content="PhoneAppOTP"></ComboBoxItem>
                        <ComboBoxItem Content="PhoneAppNotification"></ComboBoxItem>
                    </ComboBox>

                    <!-- Last reset -->
                    <Label Content="Last reset" HorizontalAlignment="Stretch" VerticalAlignment="Center" Grid.Column="0" Grid.Row="3"/>
                    <Label Name="labelLastReset" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="3"/>

                    <!-- Devices -->
                    <Label Content="Devices" HorizontalAlignment="Stretch" VerticalAlignment="Top" Grid.Column="0" Grid.Row="4"/>
                    <Rectangle Fill="#E6E6E6" StrokeThickness="0" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="4" Margin="2"/>
                    <TextBlock Name="textblocPhoneAppDetails" Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="4" Margin="10" FontFamily="Consolas" TextWrapping="Wrap" Text=""></TextBlock>


                </Grid>

                <!-- Controls -->
                <Grid Grid.Row="3">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="33*"></ColumnDefinition>
                        <ColumnDefinition Width="33*"></ColumnDefinition>
                        <ColumnDefinition Width="33*"></ColumnDefinition>
                    </Grid.ColumnDefinitions>
                    <!-- Disable -->
                    <Button Name="buttonDisable" ToolTip="DisableMFA" Background="#d9534f" BorderThickness="0" Grid.Column="0" Margin="15">
                        <Path Fill="White" Stretch="Uniform" Margin="8" Data="M18 1C15.24 1 13 3.24 13 6V8H4C2.9 8 2 8.89 2 10V20C2 21.11 2.9 22 4 22H16C17.11 22 18 21.11 18 20V10C18 8.9 17.11 8 16 8H15V6C15 4.34 16.34 3 18 3C19.66 3 21 4.34 21 6V8H23V6C23 3.24 20.76 1 18 1M10 13C11.1 13 12 13.89 12 15C12 16.11 11.11 17 10 17C8.9 17 8 16.11 8 15C8 13.9 8.9 13 10 13Z"/>
                    </Button>
                    <!-- Enable -->
                    <Button Name="buttonEnable" ToolTip="EnableMFA" Background="#5cb85c" BorderThickness="0" Grid.Column="1" Margin="15">
                        <Path Fill="White" Stretch="Uniform" Margin="8" Data="M12,17A2,2 0 0,0 14,15C14,13.89 13.1,13 12,13A2,2 0 0,0 10,15A2,2 0 0,0 12,17M18,8A2,2 0 0,1 20,10V20A2,2 0 0,1 18,22H6A2,2 0 0,1 4,20V10C4,8.89 4.9,8 6,8H7V6A5,5 0 0,1 12,1A5,5 0 0,1 17,6V8H18M12,3A3,3 0 0,0 9,6V8H15V6A3,3 0 0,0 12,3Z"/>
                    </Button>
                    <!-- Reset -->
                    <Button Name="buttonReset" ToolTip="ResetMFA" Background="DodgerBlue" BorderThickness="0" Grid.Column="2" Margin="15">
                        <Path Fill="White" Stretch="Uniform" Margin="8" Data="M12.63,2C18.16,2 22.64,6.5 22.64,12C22.64,17.5 18.16,22 12.63,22C9.12,22 6.05,20.18 4.26,17.43L5.84,16.18C7.25,18.47 9.76,20 12.64,20A8,8 0 0,0 20.64,12A8,8 0 0,0 12.64,4C8.56,4 5.2,7.06 4.71,11H7.47L3.73,14.73L0,11H2.69C3.19,5.95 7.45,2 12.63,2M15.59,10.24C16.09,10.25 16.5,10.65 16.5,11.16V15.77C16.5,16.27 16.09,16.69 15.58,16.69H10.05C9.54,16.69 9.13,16.27 9.13,15.77V11.16C9.13,10.65 9.54,10.25 10.04,10.24V9.23C10.04,7.7 11.29,6.46 12.81,6.46C14.34,6.46 15.59,7.7 15.59,9.23V10.24M12.81,7.86C12.06,7.86 11.44,8.47 11.44,9.23V10.24H14.19V9.23C14.19,8.47 13.57,7.86 12.81,7.86Z"/>
                    </Button>
                </Grid>

            </Grid>
        </ScrollViewer>

        <!-- Footer -->
        <Button Name="buttonHelp" Background="Transparent" HorizontalAlignment="Center" VerticalAlignment="Center" Height="24" Width="24" BorderThickness="0" Grid.Row="2">
            <Path Fill="DimGray" Data="M15.07,11.25L14.17,12.17C13.45,12.89 13,13.5 13,15H11V14.5C11,13.39 11.45,12.39 12.17,11.67L13.41,10.41C13.78,10.05 14,9.55 14,9C14,7.89 13.1,7 12,7A2,2 0 0,0 10,9H8A4,4 0 0,1 12,5A4,4 0 0,1 16,9C16,9.88 15.64,10.67 15.07,11.25M13,19H11V17H13M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12C22,6.47 17.5,2 12,2Z" Stretch="Fill" />
        </Button>

        <Label Name="labelSignature" Foreground="LightGray" FontSize="11" HorizontalAlignment="Right" Margin="0,9,10,9" VerticalAlignment="Center" Grid.Column="1" Grid.Row="2"/>
        <Separator HorizontalAlignment="Stretch" VerticalAlignment="Top" Grid.Column="1" Grid.Row="2"/>

    </Grid>
</Window>
'@

    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms
    $Global:XmlWPF = $Xml
    $Global:XamGUI = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Global:XmlWPF))
    $Global:XmlWPF.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Global:XamGUI.FindName($_.Name) -Scope Global }
    $labelSignature.Content = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("TWFkZSB3aXRoIOKdpCBieSBMw6lvIEJvdWFyZA=="))
}

function Show-MFAUserInterface {
    
    $Global:XamGUI.ShowDialog() | Out-Null
}

function Update-MFAUserInterface {

    param([psobject]$User)

    # User information
    $textboxSearch.Text        = $user.UserPrincipalName
    $labelDisplayName.Content  = $user.Name
    $labelDepartment.Content   = $user.Department
    $labelJobTitle.Content     = $user.JobTitle
    $labelPhoneNumber.Content  = $user.MobilePhone
    $labelEmailAddress.Content = $user.EmailAddress

    # MFA configuration
    $mfaInfo = Get-MFAUserInfo -UserPrincipalName $user.UserPrincipalName

    $labelState.Content           = $mfaInfo.State
    $comboboxDefaultMethod.Text   = $mfaInfo.DefaultMethod
    $labelLastReset.Content       = $mfaInfo.LastResetDate
    $textblocPhoneAppDetails.Text = ($mfaInfo.PhoneAppDetails | Select-Object DeviceName,AuthenticationType | Format-List | Out-String).Trim()

    switch -Wildcard ($mfaInfo.State) {
        "Disabled" { 
            $buttonDisable.IsEnabled = $false
            $buttonEnable.IsEnabled  = $true  
            $pathState.Fill          = "#d9534f"
            $pathState.Data          = "M18 1C15.24 1 13 3.24 13 6V8H4C2.9 8 2 8.89 2 10V20C2 21.11 2.9 22 4 22H16C17.11 22 18 21.11 18 20V10C18 8.9 17.11 8 16 8H15V6C15 4.34 16.34 3 18 3C19.66 3 21 4.34 21 6V8H23V6C23 3.24 20.76 1 18 1M10 13C11.1 13 12 13.89 12 15C12 16.11 11.11 17 10 17C8.9 17 8 16.11 8 15C8 13.9 8.9 13 10 13Z"
        }
        "En*" {
            $buttonDisable.IsEnabled = $true
            $buttonEnable.IsEnabled  = $false
            $pathState.Fill          = "#5cb85c"
            $pathState.Data          = "M12,17A2,2 0 0,0 14,15C14,13.89 13.1,13 12,13A2,2 0 0,0 10,15A2,2 0 0,0 12,17M18,8A2,2 0 0,1 20,10V20A2,2 0 0,1 18,22H6A2,2 0 0,1 4,20V10C4,8.89 4.9,8 6,8H7V6A5,5 0 0,1 12,1A5,5 0 0,1 17,6V8H18M12,3A3,3 0 0,0 9,6V8H15V6A3,3 0 0,0 12,3Z"
        }
        default {
            $buttonDisable.IsEnabled = $true
            $buttonEnable.IsEnabled  = $true
            $pathState.Fill          = "Transparent"
            $pathState.Data          = $null
        }
    }

    $xamGUI.LayoutTransform | Out-Null
}

function Get-MFAHelp {

    Start-Process "https://www.microsoft.com/security/business/identity-access-management/mfa-multi-factor-authentication"

}

function Reset-MFAInterface {

    $labelDisplayName.Content     = "User information"
    $labelDepartment.Content      = $null
    $labelJobTitle.Content        = $null
    $labelPhoneNumber.Content     = $null
    $labelEmailAddress.Content    = $null
    $textboxSearch.Text           = $null
    $labelState.Content           = $null
    $comboboxDefaultMethod.Text   = $null
    $labelLastReset.Content       = $null
    $textblocPhoneAppDetails.Text = $null
    $buttonDisable.IsEnabled      = $true
    $buttonEnable.IsEnabled       = $true
    $pathState.Fill               = "Transparent"
    $pathState.Data               = $null

    $xamGUI.LayoutTransform | Out-Null
}

function Show-Message {

    param (
        [string]$Message,
        [string]$Title,
        [ValidateSet("OKCancel","AbortRetryIgnore","YesNoCancel","YesNo","RetryCancel",$null)][string]$ButtonSet,
        [ValidateSet("Error","Question","Warning","Information",$null)][string]$Icon
    )

    switch ($ButtonSet) {
        "OKCancel"         { $Btn = 1 }
        "AbortRetryIgnore" { $Btn = 2 }
        "YesNoCancel"      { $Btn = 3 }
        "YesNo"            { $Btn = 4 }
        "RetryCancel"      { $Btn = 5 }
        default            { $Btn = 0 }
    }

    switch ($Icon) {
        "Error"       {$Ico = 16 }
        "Question"    {$Ico = 32 }
        "Warning"     {$Ico = 48 }
        "Information" {$Ico = 64 }
        default       {$Ico = 0  }
    }
    
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $reponse = [System.Windows.Forms.MessageBox]::Show($Message,$Title,$Btn,$Ico)
    
    return $reponse
}

function Test-MFAMsolConnection {

    try { 
        Get-MsolAccountSku -ErrorAction Stop | Out-Null
        return "Connected"
    } catch {
        return "Disconnected"
    }
}

function Get-MFAManagementUI {

    if ((Test-MFAMsolConnection) -eq "Disconnected") { Connect-MsolService }
    
    Get-MFAUserInterface

    $buttonSearch.Add_Click({
        $Global:User = Get-MFAUserInfo -UserPrincipalName $textboxSearch.Text
        if ($null -ne $Global:User) { Update-MFAUserInterface -User $user } else { Reset-MFAInterface }
    })

    $textboxSearch.Add_KeyDown({
        if ($_.Key -eq "Return") {
            $Global:User = Get-MFAUserInfo -UserPrincipalName $textboxSearch.Text
            if ($null -ne $Global:User) { Update-MFAUserInterface -User $user } else { Reset-MFAInterface }
        }
    })

    $buttonRefresh.Add_Click({
        if ($null -ne $Global:User) { Update-MFAUserInterface -User $user } else { Reset-MFAInterface }
    })

    $buttonDisable.Add_Click({
        Disable-MFA -UserPrincipalName $textboxSearch.Text
        Update-MFAUserInterface -User $Global:User
    })

    $buttonEnable.Add_Click({
        Enable-MFA -UserPrincipalName $textboxSearch.Text
        Update-MFAUserInterface -User $Global:User
    })

    $buttonReset.Add_Click({
        Reset-MFA -UserPrincipalName $textboxSearch.Text
    })

    $buttonHelp.Add_Click({
        Get-MFAHelp
    })

    $buttonClearSearch.Add_Click({
        $textboxSearch.Text = $null
        Reset-MFAInterface
    })

    $buttonPhoneNumber.Add_Click({
        $labelPhoneNumber.Content | Set-Clipboard
    })

    $buttonEmailAddress.Add_Click({
        $labelEmailAddress.Content | Set-Clipboard
    })

    $comboboxDefaultMethod.Add_DropDownClosed({
        Set-MFADefaultMethod -UserPrincipalName $textboxSearch.Text -Method $comboboxDefaultMethod.Text
    })

    $buttonMailTo.Add_Click({
        Start-Process "mailto:$($labelEmailAddress.Content)"
    })

    $buttonPhoneTo.Add_Click({
        Start-Process "tel:$($labelEmailAddress.Content)"
    })

    Show-MFAUserInterface

}