function Get-CredentialGUI {
    <#
    .SYNOPSIS
        Displays a graphical credential prompt dialog

    .DESCRIPTION
        Shows a WPF-based credential dialog for collecting username and password.
        Returns a PSCredential object if credentials are entered, or null if canceled.

    .PARAMETER TargetName
        Name of the target system or service (displayed in window title).

    .PARAMETER Message
        Optional informational message to display above the credential fields.

    .OUTPUTS
        [PSCredential]. Credential object with username and secure password, or $null if canceled.

    .EXAMPLE
        $cred = Get-CredentialGUI -TargetName "SQL Server"

    .EXAMPLE
        $cred = Get-CredentialGUI -TargetName "Remote Server" -Message "Enter your domain credentials"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetName,
        
        [Parameter(Mandatory=$false)]
        [string]$Message
    )
    
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    
    # Définition du XAML pour l'interface
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Connexion - $TargetName"
        Height="$(if($Message) { '250' } else { '201' })"
        Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Topmost="True"
        ShowInTaskbar="True">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            $(if($Message) { '<RowDefinition Height="Auto"/><RowDefinition Height="10"/>' })
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="20"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="120"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        $(if($Message) { "<Label Grid.Row=`"0`" Grid.Column=`"0`" Grid.ColumnSpan=`"2`" Content=`"$Message`" FontStyle=`"Italic`" Foreground=`"DarkBlue`" Margin=`"0,0,0,5`"/>" })
        
        <!-- Nom d'utilisateur -->
        <Label Grid.Row="$(if($Message) { '2' } else { '0' })" Grid.Column="0" Content="Nom d'utilisateur:" FontWeight="Bold" VerticalAlignment="Center"/>
        <TextBox Grid.Row="$(if($Message) { '2' } else { '0' })" Grid.Column="1" Name="txtUsername" Height="25" FontSize="12" VerticalAlignment="Center" VerticalContentAlignment="Center"/>
        
        <!-- Mot de passe -->
        <Label Grid.Row="$(if($Message) { '4' } else { '2' })" Grid.Column="0" Content="Mot de passe:" FontWeight="Bold" VerticalAlignment="Center"/>
        <PasswordBox Grid.Row="$(if($Message) { '4' } else { '2' })" Grid.Column="1" Name="txtPassword" Height="25" FontSize="12" VerticalAlignment="Center" VerticalContentAlignment="Center"/>
        
        <!-- Boutons -->
        <StackPanel Grid.Row="$(if($Message) { '6' } else { '4' })" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="OK" Width="75" Height="30" Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="btnCancel" Content="Annuler" Width="75" Height="30" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        # Création de la fenêtre à partir du XAML
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Récupération des contrôles
        $txtUsername = $window.FindName("txtUsername")
        $txtPassword = $window.FindName("txtPassword")
        $btnOK = $window.FindName("btnOK")
        $btnCancel = $window.FindName("btnCancel")
        
        # Variables pour stocker le résultat
        $script:DialogResult = $false
        $script:Username = ""
        $script:Password = ""
        
        # Gestionnaire d'événement pour le bouton OK
        $btnOK.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtUsername.Text) -or [string]::IsNullOrWhiteSpace($txtPassword.Password)) {
                [System.Windows.MessageBox]::Show("Le nom d'utilisateur et le mot de passe sont requis.", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $script:DialogResult = $true
            $script:Username = $txtUsername.Text
            $script:Password = $txtPassword.Password
            $window.Close()
        })
        
        # Gestionnaire d'événement pour le bouton Annuler
        $btnCancel.Add_Click({
            $script:DialogResult = $false
            $window.Close()
        })
        
        # Gestionnaire pour la touche Entrée dans le champ mot de passe
        $txtPassword.Add_KeyDown({
            if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
                $btnOK.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
        })
        
        # Gestionnaire pour la touche Entrée dans le champ nom d'utilisateur
        $txtUsername.Add_KeyDown({
            if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
                $txtPassword.Focus()
            }
        })
        
        # Focus sur le champ nom d'utilisateur au chargement
        $window.Add_Loaded({
            $txtUsername.Focus()
        })
        
        # Afficher la fenêtre
        $null = $window.ShowDialog()
        
        # Retourner le résultat
        if ($script:DialogResult) {
            $securePassword = ConvertTo-SecureString $script:Password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($script:Username, $securePassword)
            
            # Nettoyer les variables sensibles
            $script:Password = ""
            
            return $credential
        }
        else {
            return $null
        }
    }
    catch {
        Write-Error "Erreur lors de la création de l'interface: $($_.Exception.Message)"
        return $null
    }
}