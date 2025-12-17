<#
.SYNOPSIS
    Main Window Logic
.DESCRIPTION
    PowerShell code-behind for MainWindow.xaml
    Handles UI events and data binding
#>

#region Window Loading
function Show-MainWindow {
    <#
    .SYNOPSIS
        Load and display the main window
    #>
    
    # Load XAML
    $xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
    
    if (-not (Test-Path $xamlPath)) {
        Write-M365Log -Message "MainWindow.xaml not found: $xamlPath" -Level Error -Component "UI"
        throw "UI definition file not found"
    }
    
    $xaml = Get-Content $xamlPath -Raw
    
    # Remove x:Class attribute if present (not supported in PowerShell)
    $xaml = $xaml -replace 'x:Class="[^"]*"', ''
    
    # Load XAML
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get named elements
    $script:UI = @{}
    $namedElements = @(
        'TenantSelector', 'BtnAddTenant', 'BtnRefresh', 'BtnSettings',
        'StatusText', 'AuthStatus', 'VersionText',
        'NavTree', 'ContentHeader', 'ContentPanel',
        'DashboardView', 'UsersView', 'PlaceholderView', 'PlaceholderText',
        'TenantCount', 'UserCount', 'LicensePercent', 'AlertCount',
        'TenantGrid', 'UsersGrid',
        'UserSearchBox', 'ChkShowGuests', 'ChkShowDisabled',
        'BtnExport', 'BtnHelp'
    )
    
    foreach ($name in $namedElements) {
        $element = $window.FindName($name)
        if ($element) {
            $script:UI[$name] = $element
        }
    }
    
    # Wire up events
    Initialize-UIEvents -Window $window
    
    # Load initial data
    Initialize-UIData
    
    # Show window
    $window.ShowDialog() | Out-Null
}
#endregion

#region Event Handlers
function Initialize-UIEvents {
    param($Window)
    
    # Tenant selector change
    $script:UI.TenantSelector.Add_SelectionChanged({
        $selected = $script:UI.TenantSelector.SelectedItem
        if ($selected) {
            Set-TenantContext -TenantId $selected.id
        }
    })
    
    # Add tenant button
    $script:UI.BtnAddTenant.Add_Click({
        Show-AddTenantDialog
    })
    
    # Refresh button
    $script:UI.BtnRefresh.Add_Click({
        Refresh-CurrentView
    })
    
    # Settings button
    $script:UI.BtnSettings.Add_Click({
        Show-SettingsDialog
    })
    
    # Navigation tree selection
    $script:UI.NavTree.Add_SelectedItemChanged({
        $selected = $script:UI.NavTree.SelectedItem
        if ($selected -and $selected.Tag) {
            Navigate-ToView -ViewName $selected.Tag
        }
    })
    
    # Export button
    $script:UI.BtnExport.Add_Click({
        Start-ExportWizard
    })
    
    # Help button
    $script:UI.BtnHelp.Add_Click({
        Show-HelpDialog
    })
    
    # Window closing
    $Window.Add_Closing({
        Write-M365Log -Message "Application closing" -Level Information -Component "UI"
    })
}
#endregion

#region Data Loading
function Initialize-UIData {
    <#
    .SYNOPSIS
        Load initial data into UI
    #>
    
    # Set version
    $script:UI.VersionText.Text = "v$(Get-M365ManagerVersion)"
    
    # Load registered tenants
    $tenants = Get-M365RegisteredTenants
    
    # Add status color property
    $tenantsWithStatus = $tenants | ForEach-Object {
        $t = $_
        Add-Member -InputObject $t -MemberType NoteProperty -Name "StatusColor" -Value "#27AE60" -Force
        $t
    }
    
    $script:UI.TenantSelector.ItemsSource = $tenantsWithStatus
    $script:UI.TenantGrid.ItemsSource = $tenantsWithStatus
    
    # Update summary cards
    $script:UI.TenantCount.Text = $tenants.Count.ToString()
    $script:UI.UserCount.Text = "---"
    $script:UI.LicensePercent.Text = "---"
    $script:UI.AlertCount.Text = "---"
    
    # Select first tenant if available
    if ($tenants.Count -gt 0) {
        $script:UI.TenantSelector.SelectedIndex = 0
    }
    
    Update-Status "Ready - $($tenants.Count) tenants registered"
}

function Refresh-CurrentView {
    <#
    .SYNOPSIS
        Refresh data in current view
    #>
    Update-Status "Refreshing..."
    
    try {
        # Reload tenant list
        Initialize-UIData
        
        # Refresh current tenant data if connected
        if (Test-M365Connection) {
            $context = Get-M365AuthContext
            Update-Status "Connected to: $($context.TenantId)"
            $script:UI.AuthStatus.Text = "Connected"
            $script:UI.AuthStatus.Foreground = [System.Windows.Media.Brushes]::Green
        }
        else {
            $script:UI.AuthStatus.Text = "Not Connected"
            $script:UI.AuthStatus.Foreground = [System.Windows.Media.Brushes]::Gray
        }
    }
    catch {
        Update-Status "Error: $_"
    }
}
#endregion

#region Navigation
function Navigate-ToView {
    param([string]$ViewName)
    
    Write-M365Log -Message "Navigating to: $ViewName" -Level Debug -Component "UI"
    
    # Hide all views
    $script:UI.DashboardView.Visibility = "Collapsed"
    $script:UI.UsersView.Visibility = "Collapsed"
    $script:UI.PlaceholderView.Visibility = "Collapsed"
    
    # Update header
    $headers = @{
        "Dashboard"        = "Dashboard"
        "TenantsOverview"  = "Tenant Overview"
        "TenantsHealth"    = "Tenant Health Status"
        "UsersAll"         = "All Users"
        "UsersGuest"       = "Guest Users"
        "GroupsAll"        = "All Groups"
        "LicensesOverview" = "License Overview"
        "SecureScore"      = "Secure Score"
        "ConditionalAccess"= "Conditional Access Policies"
    }
    
    $script:UI.ContentHeader.Text = if ($headers[$ViewName]) { $headers[$ViewName] } else { $ViewName }
    
    # Show appropriate view
    switch -Wildcard ($ViewName) {
        "Dashboard" {
            $script:UI.DashboardView.Visibility = "Visible"
        }
        "Users*" {
            $script:UI.UsersView.Visibility = "Visible"
            Load-UsersData
        }
        default {
            $script:UI.PlaceholderView.Visibility = "Visible"
            $script:UI.PlaceholderText.Text = "View: $ViewName`n`nThis view is under development."
        }
    }
}

function Set-TenantContext {
    param([string]$TenantId)
    
    Update-Status "Switching to tenant: $TenantId"
    
    try {
        Switch-M365Tenant -TenantId $TenantId
        Update-M365TenantAccess -TenantId $TenantId
        
        $script:UI.AuthStatus.Text = "Connected"
        $script:UI.AuthStatus.Foreground = [System.Windows.Media.Brushes]::Green
        
        Update-Status "Connected to tenant"
        
        # Refresh current view
        Refresh-CurrentView
    }
    catch {
        Update-Status "Connection failed: $_"
        $script:UI.AuthStatus.Text = "Connection Failed"
        $script:UI.AuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
    }
}
#endregion

#region Data Views
function Load-UsersData {
    <#
    .SYNOPSIS
        Load users into the users grid
    #>
    if (-not (Test-M365Connection)) {
        Update-Status "Not connected - please select a tenant"
        return
    }
    
    Update-Status "Loading users..."
    
    try {
        $users = Get-M365GraphAllPages -Uri "/users?`$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses&`$top=100" -MaxPages 5
        
        # Add license count
        $usersWithCount = $users | ForEach-Object {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name "licenseCount" -Value $_.assignedLicenses.Count -Force
            $_
        }
        
        $script:UI.UsersGrid.ItemsSource = $usersWithCount
        $script:UI.UserCount.Text = $users.Count.ToString()
        
        Update-Status "Loaded $($users.Count) users"
    }
    catch {
        Update-Status "Error loading users: $_"
    }
}
#endregion

#region Dialogs
function Show-AddTenantDialog {
    <#
    .SYNOPSIS
        Show dialog to add a new tenant
    #>
    $dialog = New-Object System.Windows.Window
    $dialog.Title = "Add Tenant"
    $dialog.Width = 400
    $dialog.Height = 200
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.ResizeMode = "NoResize"
    
    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = 20
    
    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = "Enter Tenant ID or Domain:"
    $label.Margin = "0,0,0,10"
    $panel.Children.Add($label)
    
    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Height = 28
    $textBox.Margin = "0,0,0,20"
    $panel.Children.Add($textBox)
    
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    
    $okButton = New-Object System.Windows.Controls.Button
    $okButton.Content = "Add"
    $okButton.Width = 80
    $okButton.Margin = "0,0,10,0"
    $okButton.Add_Click({
        if ($textBox.Text) {
            try {
                Register-M365Tenant -TenantId $textBox.Text
                $dialog.DialogResult = $true
                $dialog.Close()
                Refresh-CurrentView
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to add tenant: $_", "Error")
            }
        }
    })
    $buttonPanel.Children.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 80
    $cancelButton.Add_Click({ $dialog.Close() })
    $buttonPanel.Children.Add($cancelButton)
    
    $panel.Children.Add($buttonPanel)
    
    $dialog.Content = $panel
    $dialog.ShowDialog() | Out-Null
}

function Show-SettingsDialog {
    [System.Windows.MessageBox]::Show(
        "Settings dialog coming soon.`n`nCurrent settings are stored in:`n$(Get-M365SettingsPath)",
        "Settings",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
}

function Show-HelpDialog {
    $helpText = @"
M365 Multi-Tenant Manager

Quick Start:
1. Click '+' next to tenant selector to add a tenant
2. Authenticate with your admin credentials
3. Select tenant from dropdown to switch context
4. Use navigation tree to explore data

Keyboard Shortcuts:
• Ctrl+T - Focus tenant selector
• F5 - Refresh current view
• Ctrl+E - Export data

For more help, see the documentation.
"@
    
    [System.Windows.MessageBox]::Show($helpText, "Help", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}

function Start-ExportWizard {
    [System.Windows.MessageBox]::Show(
        "Export wizard coming soon.`n`nUse PowerShell commands for now:`n`nInvoke-M365BatchJobs -Jobs `$jobs",
        "Export",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
}
#endregion

#region Utilities
function Update-Status {
    param([string]$Message)
    
    if ($script:UI.StatusText) {
        $script:UI.StatusText.Text = $Message
    }
    
    Write-M365Log -Message $Message -Level Information -Component "UI"
}
#endregion

# Export handled by dot-sourcing
