<#
.SYNOPSIS
    SecurePassGenerator - A Comprehensive Password Generation and Sharing Tool

.DESCRIPTION
    A PowerShell script that creates a modern WPF GUI for generating secure passwords
    and pushing them to Password Pusher's API. Features include password customization,
    strength assessment, breach checking, and secure sharing options.

    Key Features:
    - Multiple password generation methods (random or memorable)
    - Built-in and custom password presets with management capabilities
    - Real-time password strength assessment with entropy calculation
    - Integration with Have I Been Pwned to check for compromised passwords
    - Secure password sharing via Password Pusher with configurable settings
    - QR code generation for mobile access
    - Integrated update system with support for stable and pre-release versions
    - Rate-limited API interactions to respect service limitations
    
    Note: This script implements cooldown timers between API requests to respect
    rate limits for external services. Please do not modify these limiting features.

.NOTES
    File Name      : SecurePassGenerator.ps1
    Version        : 1.1.0
    Author         : onlyalex1984
    Copyright      : (C) 2025 onlyalex1984
    License        : GPL v3 - Full license text available in the project root directory
    Created        : Februari 2025
    Last Modified  : April 2025
    Prerequisite   : PowerShell 5.1 or higher, Windows environment with .NET Framework
    GitHub         : https://github.com/onlyalex1984/securepassgenerator-powershell
#>

# Script configuration
$script:Version = "1.1.0"
$script:ScriptDisplayName = "SecurePassGenerator"

# Application paths
$script:AppDataPath = [System.Environment]::GetFolderPath('ApplicationData')
$script:InstallDir = Join-Path -Path $script:AppDataPath -ChildPath "securepassgenerator-ps"
$script:InstallerPath = Join-Path -Path $script:InstallDir -ChildPath "Installer.ps1"

# Password links history collection
$script:PasswordLinks = New-Object System.Collections.ArrayList

# Password presets collection and file path
$script:PasswordPresets = New-Object System.Collections.ArrayList
# Get the directory where the script is located
$script:ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
# Store presets in the InstallDir for consistency across different execution locations
$script:PresetsFilePath = Join-Path -Path $script:InstallDir -ChildPath "presets.json"

# Password preset model class
class PasswordPresetModel {
    [string]$Name
    [int]$Length
    [bool]$IncludeUppercase
    [bool]$IncludeLowercase = $true  # Always true as lowercase is always included
    [bool]$IncludeNumbers
    [bool]$IncludeSpecial
    [bool]$IsDefault = $false  # Indicates if this is a built-in preset
    [bool]$Enabled = $true     # Indicates if this preset is enabled
    [bool]$IsSelectedByDefault = $false  # Indicates if this preset should be selected by default on startup

    PasswordPresetModel([string]$name, [int]$length, [bool]$includeUppercase, [bool]$includeNumbers, [bool]$includeSpecial) {
        $this.Name = $name
        $this.Length = $length
        $this.IncludeUppercase = $includeUppercase
        $this.IncludeNumbers = $includeNumbers
        $this.IncludeSpecial = $includeSpecial
    }
    
    # Clone method to create a copy of the preset
    [PasswordPresetModel] Clone() {
        $clone = [PasswordPresetModel]::new($this.Name, $this.Length, $this.IncludeUppercase, $this.IncludeNumbers, $this.IncludeSpecial)
        $clone.IsDefault = $this.IsDefault
        $clone.Enabled = $this.Enabled
        $clone.IsSelectedByDefault = $this.IsSelectedByDefault
        return $clone
    }
    
    # Convert to JSON-friendly hashtable
    [hashtable] ToHashtable() {
        return @{
            Name = $this.Name
            Length = $this.Length
            IncludeUppercase = $this.IncludeUppercase
            IncludeLowercase = $this.IncludeLowercase
            IncludeNumbers = $this.IncludeNumbers
            IncludeSpecial = $this.IncludeSpecial
            IsDefault = $this.IsDefault
            Enabled = $this.Enabled
            IsSelectedByDefault = $this.IsSelectedByDefault
        }
    }
}

# Loads password presets from JSON file or creates default presets if file doesn't exist
function Import-PasswordPresets {
    # Create directory if it doesn't exist
    $presetsDir = Split-Path -Path $script:PresetsFilePath -Parent
    if (-not (Test-Path -Path $presetsDir)) {
        try {
            New-Item -Path $presetsDir -ItemType Directory -Force | Out-Null
            $global:controls.LogOutput.AppendText("Created presets directory: $presetsDir`n")
        }
        catch {
            $global:controls.LogOutput.AppendText("Error creating presets directory: $($_.Exception.Message)`n")
        }
    }
    
    # Clear existing presets
    $script:PasswordPresets.Clear()
    
    # Check if presets file exists
    if (Test-Path -Path $script:PresetsFilePath) {
        try {
            # Read and parse JSON file
            $presetsJson = Get-Content -Path $script:PresetsFilePath -Raw
            $presetsData = $presetsJson | ConvertFrom-Json
            
            # Add each preset to the collection
            foreach ($presetData in $presetsData) {
                $preset = [PasswordPresetModel]::new(
                    $presetData.Name,
                    $presetData.Length,
                    $presetData.IncludeUppercase,
                    $presetData.IncludeNumbers,
                    $presetData.IncludeSpecial
                )
                
                # Set IsDefault property if it exists in the JSON
                if (Get-Member -InputObject $presetData -Name "IsDefault" -MemberType Properties) {
                    $preset.IsDefault = $presetData.IsDefault
                }
                
                # Set Enabled property if it exists in the JSON
                if (Get-Member -InputObject $presetData -Name "Enabled" -MemberType Properties) {
                    $preset.Enabled = $presetData.Enabled
                }
                
                # Set IsSelectedByDefault property if it exists in the JSON
                if (Get-Member -InputObject $presetData -Name "IsSelectedByDefault" -MemberType Properties) {
                    $preset.IsSelectedByDefault = $presetData.IsSelectedByDefault
                }
                
                [void]$script:PasswordPresets.Add($preset)
            }
            
            $global:controls.LogOutput.AppendText("Loaded $($script:PasswordPresets.Count) presets from $script:PresetsFilePath`n")
        }
        catch {
            $global:controls.LogOutput.AppendText("Error loading presets: $($_.Exception.Message)`n")
            $global:controls.LogOutput.AppendText("Loading default presets instead`n")
            Add-DefaultPresets
        }
    }
    else {
        # File doesn't exist, create default presets
        $global:controls.LogOutput.AppendText("Presets file not found. Creating default presets`n")
        Add-DefaultPresets
        
        # Don't save the default presets to file until changes are made
        # This prevents creating the file unnecessarily on startup
    }
    
    # Update the UI with the loaded presets
    Update-PresetsDropdown
}

# Adds default presets to the collection
function Add-DefaultPresets {
    # Clear existing presets
    $script:PasswordPresets.Clear()
    
    # Medium Password (10 chars)
    $mediumPreset = [PasswordPresetModel]::new("Medium Password", 10, $true, $true, $true)
    $mediumPreset.IsDefault = $true
    [void]$script:PasswordPresets.Add($mediumPreset)
    
    # Strong Password (15 chars)
    $strongPreset = [PasswordPresetModel]::new("Strong Password", 15, $true, $true, $true)
    $strongPreset.IsDefault = $true
    $strongPreset.IsSelectedByDefault = $true
    [void]$script:PasswordPresets.Add($strongPreset)
    
    # Very Strong Password (20 chars)
    $veryStrongPreset = [PasswordPresetModel]::new("Very Strong Password", 20, $true, $true, $true)
    $veryStrongPreset.IsDefault = $true
    [void]$script:PasswordPresets.Add($veryStrongPreset)
    
    # NIST Compliant (12 chars)
    $nistPreset = [PasswordPresetModel]::new("NIST Compliant", 12, $true, $true, $true)
    $nistPreset.IsDefault = $true
    [void]$script:PasswordPresets.Add($nistPreset)
    
    # SOC 2 Compliant (14 chars)
    $soc2Preset = [PasswordPresetModel]::new("SOC 2 Compliant", 14, $true, $true, $true)
    $soc2Preset.IsDefault = $true
    [void]$script:PasswordPresets.Add($soc2Preset)
    
    # Financial Compliant (16 chars)
    $financialPreset = [PasswordPresetModel]::new("Financial Compliant", 16, $true, $true, $true)
    $financialPreset.IsDefault = $true
    [void]$script:PasswordPresets.Add($financialPreset)
    
    $global:controls.LogOutput.AppendText("Added default presets`n")
}

# Saves password presets to JSON file
function Save-PasswordPresets {
    try {
        # Create directory if it doesn't exist
        $presetsDir = Split-Path -Path $script:PresetsFilePath -Parent
        if (-not (Test-Path -Path $presetsDir)) {
            New-Item -Path $presetsDir -ItemType Directory -Force | Out-Null
        }
        
        # Create backup of existing file if it exists
        if (Test-Path -Path $script:PresetsFilePath) {
            $backupPath = "$script:PresetsFilePath.bak"
            Copy-Item -Path $script:PresetsFilePath -Destination $backupPath -Force
            $global:controls.LogOutput.AppendText("Created backup of presets file: $backupPath`n")
        }
        
        # Convert presets to array of hashtables for JSON serialization
        $presetsData = @()
        foreach ($preset in $script:PasswordPresets) {
            $presetsData += $preset.ToHashtable()
        }
        
        # Convert to JSON and save to file
        $presetsJson = $presetsData | ConvertTo-Json -Depth 5
        Set-Content -Path $script:PresetsFilePath -Value $presetsJson -Force
        
        $global:controls.LogOutput.AppendText("Saved $($script:PasswordPresets.Count) presets to $script:PresetsFilePath`n")
        return $true
    }
    catch {
        $global:controls.LogOutput.AppendText("Error saving presets: $($_.Exception.Message)`n")
        return $false
    }
}

# Updates the presets dropdown with current presets
function Update-PresetsDropdown {
    if (-not $global:controls -or -not $global:controls.PasswordPresets) {
        return
    }
    
    # Clear existing items
    $global:controls.PasswordPresets.Items.Clear()
    
    # Maximum length for preset names in dropdown
    $maxDisplayLength = 32
    
    # Add only enabled presets to the dropdown
    $enabledPresets = $script:PasswordPresets | Where-Object { $_.Enabled }
    $defaultPresetIndex = -1
    $currentIndex = 0
    
    foreach ($preset in $enabledPresets) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        
        # Truncate long preset names
        if ($preset.Name.Length -gt $maxDisplayLength) {
            $displayName = $preset.Name.Substring(0, $maxDisplayLength) + "..."
        } else {
            $displayName = $preset.Name
        }
        
        $item.Content = $displayName
        $item.Tag = $preset
        $item.ToolTip = $preset.Name  # Show full name in tooltip
        [void]$global:controls.PasswordPresets.Items.Add($item)
        
        # Check if this is the default preset to select
        if ($preset.IsSelectedByDefault) {
            $defaultPresetIndex = $currentIndex
        }
        
        $currentIndex++
    }
    
    # Select the default preset if one is marked, otherwise select the first item
    if ($global:controls.PasswordPresets.Items.Count -gt 0) {
        if ($defaultPresetIndex -ge 0) {
            $global:controls.PasswordPresets.SelectedIndex = $defaultPresetIndex
            $global:controls.LogOutput.AppendText("Selected default preset: $($enabledPresets[$defaultPresetIndex].Name)`n")
        } else {
            $global:controls.PasswordPresets.SelectedIndex = 0
        }
    } else {
        # If no enabled presets, enable the first preset to ensure at least one is always enabled
        if ($script:PasswordPresets.Count -gt 0) {
            $script:PasswordPresets[0].Enabled = $true
            # Recursively call this function to update the dropdown with the newly enabled preset
            Update-PresetsDropdown
        }
    }
}

# Adds a new preset to the collection
function Add-PasswordPreset {
    param (
        [string]$Name,
        [int]$Length,
        [bool]$IncludeUppercase,
        [bool]$IncludeNumbers,
        [bool]$IncludeSpecial,
        [bool]$IsSelectedByDefault = $false
    )
    
    # Check if a preset with this name already exists
    $existingPreset = $script:PasswordPresets | Where-Object { $_.Name -eq $Name }
    if ($existingPreset) {
        $global:controls.LogOutput.AppendText("A preset with the name '$Name' already exists`n")
        return $false
    }
    
    # Create and add the new preset
    $preset = [PasswordPresetModel]::new($Name, $Length, $IncludeUppercase, $IncludeNumbers, $IncludeSpecial)
    $preset.IsSelectedByDefault = $IsSelectedByDefault
    
    # If this preset is being set as default, unset any other presets
    if ($preset.IsSelectedByDefault) {
        foreach ($otherPreset in $script:PasswordPresets) {
            $otherPreset.IsSelectedByDefault = $false
        }
    }
    
    [void]$script:PasswordPresets.Add($preset)
    
    # Save presets to file
    $result = Save-PasswordPresets
    
    # Update the UI
    if ($result) {
        Update-PresetsDropdown
        $global:controls.LogOutput.AppendText("Added new preset: $Name`n")
    }
    
    return $result
}

# Removes a preset from the collection
function Remove-PasswordPreset {
    param (
        [string]$Name
    )
    
    # Find the preset to remove
    $presetToRemove = $script:PasswordPresets | Where-Object { $_.Name -eq $Name }
    if (-not $presetToRemove) {
        $global:controls.LogOutput.AppendText("Preset '$Name' not found`n")
        return $false
    }
    
    # Check if it's a default preset
    if ($presetToRemove.IsDefault) {
        $global:controls.LogOutput.AppendText("Cannot remove default preset: $Name`n")
        return $false
    }
    
    # Check if the preset being removed was selected by default
    $wasSelectedByDefault = $presetToRemove.IsSelectedByDefault
    
    # Remove the preset
    [void]$script:PasswordPresets.Remove($presetToRemove)
    
    # If the removed preset was selected by default, set the Strong Password preset as default
    if ($wasSelectedByDefault) {
        $strongPreset = $script:PasswordPresets | Where-Object { $_.Name -eq "Strong Password" }
        if ($strongPreset) {
            $strongPreset.IsSelectedByDefault = $true
            $global:controls.LogOutput.AppendText("Fallback to Strong Password as default`n")
        }
    }
    
    # Save presets to file
    $result = Save-PasswordPresets
    
    # Update the UI
    if ($result) {
        Update-PresetsDropdown
        $global:controls.LogOutput.AppendText("Removed preset: $Name`n")
    }
    
    return $result
}

# Edits an existing preset
function Edit-PasswordPreset {
    param (
        [string]$OriginalName,
        [string]$NewName,
        [int]$Length,
        [bool]$IncludeUppercase,
        [bool]$IncludeNumbers,
        [bool]$IncludeSpecial,
        [bool]$IsSelectedByDefault
    )
    
    # Find the preset to edit
    $presetToEdit = $script:PasswordPresets | Where-Object { $_.Name -eq $OriginalName }
    if (-not $presetToEdit) {
        $global:controls.LogOutput.AppendText("Preset '$OriginalName' not found`n")
        return $false
    }
    
    # Check if it's a default preset and name is being changed
    if ($presetToEdit.IsDefault -and $OriginalName -ne $NewName) {
        $global:controls.LogOutput.AppendText("Cannot rename default preset: $OriginalName`n")
        return $false
    }
    
    # Check if new name conflicts with existing preset
    if ($OriginalName -ne $NewName) {
        $existingPreset = $script:PasswordPresets | Where-Object { $_.Name -eq $NewName }
        if ($existingPreset) {
            $global:controls.LogOutput.AppendText("A preset with the name '$NewName' already exists`n")
            return $false
        }
    }
    
    # Update the preset properties
    $presetToEdit.Name = $NewName
    $presetToEdit.Length = $Length
    $presetToEdit.IncludeUppercase = $IncludeUppercase
    $presetToEdit.IncludeNumbers = $IncludeNumbers
    $presetToEdit.IncludeSpecial = $IncludeSpecial
    $presetToEdit.IsSelectedByDefault = $IsSelectedByDefault
    
    # If this preset is being set as default, unset any other presets
    if ($presetToEdit.IsSelectedByDefault) {
        foreach ($otherPreset in $script:PasswordPresets) {
            if ($otherPreset -ne $presetToEdit) {
                $otherPreset.IsSelectedByDefault = $false
            }
        }
    }
    
    # Save presets to file
    $result = Save-PasswordPresets
    
    # Update the UI
    if ($result) {
        Update-PresetsDropdown
        $global:controls.LogOutput.AppendText("Updated preset: $OriginalName -> $NewName`n")
    }
    
    return $result
}

# Gets current UI settings as a preset model
function Get-CurrentSettings {
    $name = "Custom Preset"
    $length = 12
    $includeUppercase = $true
    $includeNumbers = $true
    $includeSpecial = $true
    
    if ($global:controls) {
        if ($global:controls.RandomPasswordType.IsChecked) {
            $length = [int]$global:controls.PasswordLength.Value
        }
        else {
            # For memorable passwords, use a default length
            $length = 12
        }
        
        $includeUppercase = $global:controls.IncludeUppercase.IsChecked
        $includeNumbers = $global:controls.IncludeNumbers.IsChecked
        $includeSpecial = $global:controls.IncludeSpecial.IsChecked
    }
    
    return [PasswordPresetModel]::new($name, $length, $includeUppercase, $includeNumbers, $includeSpecial)
}

# Password link model class
class PasswordLinkModel {
    [string]$Url
    [DateTime]$CreatedAt
    [System.Security.SecureString]$Password
    [bool]$IsExpired = $false  # New property to track expiration status

    PasswordLinkModel([string]$url, [System.Security.SecureString]$password) {
        $this.Url = $url
        $this.CreatedAt = Get-Date
        $this.Password = $password
    }
    
    # Method to extract token from URL
    [string] GetToken() {
        # Extract token from URL (e.g., https://pwpush.com/p/abcdef123456)
        # More flexible pattern to match different URL formats
        if ($this.Url -match '(?:https?://)?(?:www\.)?pwpush\.com/p/([a-zA-Z0-9_-]+)') {
            $token = $matches[1]
            return $token
        }
        
        # Log the failure for debugging with null check
        if ($global:controls -and $global:controls.LogOutput) {
            $global:controls.LogOutput.AppendText("Failed to extract token from URL: $($this.Url)`n")
        }
        return ""
    }
}

# Displays password preset management window
function Show-PasswordPresets {
    # Create a new window for the password presets
    $presetsWindow = New-Object System.Windows.Window
    $presetsWindow.Title = "Password Presets"
    $presetsWindow.Width = 580
    $presetsWindow.Height = 400
    $presetsWindow.WindowStartupLocation = "CenterScreen"
    $presetsWindow.ResizeMode = "CanResize"
    $presetsWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Create a grid for the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    
    # Define row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    
    # Add a title
    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = "Password Presets"
    $titleBlock.FontWeight = "Bold"
    $titleBlock.FontSize = 14
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($titleBlock, 0)
    $grid.Children.Add($titleBlock)
    
    # Create a ListView for the presets
    $listView = New-Object System.Windows.Controls.ListView
    $listView.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($listView, 1)
    
    # Create GridView columns
    $gridView = New-Object System.Windows.Controls.GridView
    
    # Name column
    $nameColumn = New-Object System.Windows.Controls.GridViewColumn
    $nameColumn.Header = "Name"
    $nameColumn.Width = 150
    $nameColumn.DisplayMemberBinding = New-Object System.Windows.Data.Binding("Name")
    $gridView.Columns.Add($nameColumn)
    
    # Length column
    $lengthColumn = New-Object System.Windows.Controls.GridViewColumn
    $lengthColumn.Header = "Length"
    $lengthColumn.Width = 60
    $lengthColumn.DisplayMemberBinding = New-Object System.Windows.Data.Binding("Length")
    $gridView.Columns.Add($lengthColumn)
    
    # Enabled column
    $enabledColumn = New-Object System.Windows.Controls.GridViewColumn
    $enabledColumn.Header = "Enabled"
    $enabledColumn.Width = 70
    
    # Create a factory for the enabled column cells
    $enabledFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
    $enabledFactory.SetBinding([System.Windows.Controls.CheckBox]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("Enabled")))
    $enabledFactory.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    
    # Add click handler for the enabled checkbox
    $enabledFactory.AddHandler(
        [System.Windows.Controls.CheckBox]::ClickEvent,
        [System.Windows.RoutedEventHandler]{
            param($buttonSender, $e)
            
            # Get the preset from the DataContext
            $preset = $buttonSender.DataContext
            
                # If trying to uncheck (disable)
                if (-not $buttonSender.IsChecked) {
                    # Count how many presets are currently enabled
                    $enabledCount = ($script:PasswordPresets | Where-Object { $_.Enabled }).Count
                    
                    # If this is the last enabled preset, prevent unchecking
                    if ($enabledCount -le 0) {
                        $buttonSender.IsChecked = $true
                        [System.Windows.MessageBox]::Show(
                            "At least one preset must remain enabled.",
                            "Cannot Disable All Presets",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Information
                        )
                        $e.Handled = $true
                        return
                    }
                }
            
            # Update the preset's Enabled property
            $preset.Enabled = $buttonSender.IsChecked
            
            # Save the changes to the presets file
            Save-PasswordPresets
            
            # Update the presets dropdown to reflect the changes
            Update-PresetsDropdown
        }
    )
    
    # Create the DataTemplate for the enabled column
    $enabledTemplate = New-Object System.Windows.DataTemplate
    $enabledTemplate.VisualTree = $enabledFactory
    $enabledColumn.CellTemplate = $enabledTemplate
    
    $gridView.Columns.Add($enabledColumn)
    
    # Options columns
    $uppercaseColumn = New-Object System.Windows.Controls.GridViewColumn
    $uppercaseColumn.Header = "Uppercase"
    $uppercaseColumn.Width = 80
    
    # Create a factory for the uppercase column cells
    $uppercaseFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
    $uppercaseFactory.SetBinding([System.Windows.Controls.CheckBox]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("IncludeUppercase")))
    $uppercaseFactory.SetValue([System.Windows.Controls.CheckBox]::IsEnabledProperty, $false)
    $uppercaseFactory.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    
    # Create the DataTemplate for the uppercase column
    $uppercaseTemplate = New-Object System.Windows.DataTemplate
    $uppercaseTemplate.VisualTree = $uppercaseFactory
    $uppercaseColumn.CellTemplate = $uppercaseTemplate
    
    $gridView.Columns.Add($uppercaseColumn)
    
    # Numbers column
    $numbersColumn = New-Object System.Windows.Controls.GridViewColumn
    $numbersColumn.Header = "Numbers"
    $numbersColumn.Width = 80
    
    # Create a factory for the numbers column cells
    $numbersFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
    $numbersFactory.SetBinding([System.Windows.Controls.CheckBox]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("IncludeNumbers")))
    $numbersFactory.SetValue([System.Windows.Controls.CheckBox]::IsEnabledProperty, $false)
    $numbersFactory.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    
    # Create the DataTemplate for the numbers column
    $numbersTemplate = New-Object System.Windows.DataTemplate
    $numbersTemplate.VisualTree = $numbersFactory
    $numbersColumn.CellTemplate = $numbersTemplate
    
    $gridView.Columns.Add($numbersColumn)
    
    # Special column
    $specialColumn = New-Object System.Windows.Controls.GridViewColumn
    $specialColumn.Header = "Special"
    $specialColumn.Width = 80
    
    # Create a factory for the special column cells
    $specialFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.CheckBox])
    $specialFactory.SetBinding([System.Windows.Controls.CheckBox]::IsCheckedProperty, (New-Object System.Windows.Data.Binding("IncludeSpecial")))
    $specialFactory.SetValue([System.Windows.Controls.CheckBox]::IsEnabledProperty, $false)
    $specialFactory.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    
    # Create the DataTemplate for the special column
    $specialTemplate = New-Object System.Windows.DataTemplate
    $specialTemplate.VisualTree = $specialFactory
    $specialColumn.CellTemplate = $specialTemplate
    
    $gridView.Columns.Add($specialColumn)
    
    # Set the GridView as the View for the ListView
    $listView.View = $gridView
    
    # Add items to the ListView
    foreach ($preset in $script:PasswordPresets) {
        [void]$listView.Items.Add($preset)
    }
    
    $grid.Children.Add($listView)
    
    # Add buttons panel
    $buttonsPanel = New-Object System.Windows.Controls.StackPanel
    $buttonsPanel.Orientation = "Horizontal"
    $buttonsPanel.HorizontalAlignment = "Right"
    $buttonsPanel.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    [System.Windows.Controls.Grid]::SetRow($buttonsPanel, 2)
    
    # Add button
    $addButton = New-Object System.Windows.Controls.Button
    $addButton.Content = "Add"
    $addButton.Width = 80
    $addButton.Height = 30
    $addButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $addButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $addButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $addButton.Add_Click({
        # Get current settings as a starting point
        $preset = Get-CurrentSettings
        $preset.Name = "New Preset"
        
        # Show edit dialog
        $result = Show-PresetEditDialog -Preset $preset -IsNew $true
        if ($result) {
        # Add the new preset
        $success = Add-PasswordPreset -Name $preset.Name -Length $preset.Length -IncludeUppercase $preset.IncludeUppercase -IncludeNumbers $preset.IncludeNumbers -IncludeSpecial $preset.IncludeSpecial -IsSelectedByDefault $preset.IsSelectedByDefault
            
            if ($success) {
                # Refresh the ListView
                $listView.Items.Clear()
                foreach ($p in $script:PasswordPresets) {
                    [void]$listView.Items.Add($p)
                }
            }
        }
    })
    $buttonsPanel.Children.Add($addButton)
    
    # Edit button
    $editButton = New-Object System.Windows.Controls.Button
    $editButton.Content = "Edit"
    $editButton.Width = 80
    $editButton.Height = 30
    $editButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $editButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $editButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $editButton.Add_Click({
        $selectedPreset = $listView.SelectedItem
        if ($selectedPreset) {
            # Check if it's a default preset
            if ($selectedPreset.IsDefault) {
                [System.Windows.MessageBox]::Show(
                    "Default presets cannot be edited. You can create a new preset based on this one.",
                    "Cannot Edit Default Preset",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                return
            }
            
            # Clone the preset to avoid modifying the original until confirmed
            $presetCopy = $selectedPreset.Clone()
            
            # Show edit dialog
            $result = Show-PresetEditDialog -Preset $presetCopy -IsNew $false
            if ($result) {
                # Edit the existing preset
                $success = Edit-PasswordPreset -OriginalName $selectedPreset.Name -NewName $presetCopy.Name -Length $presetCopy.Length -IncludeUppercase $presetCopy.IncludeUppercase -IncludeNumbers $presetCopy.IncludeNumbers -IncludeSpecial $presetCopy.IncludeSpecial -IsSelectedByDefault $presetCopy.IsSelectedByDefault
                
                if ($success) {
                    # Refresh the ListView
                    $listView.Items.Clear()
                    foreach ($p in $script:PasswordPresets) {
                        [void]$listView.Items.Add($p)
                    }
                }
            }
        }
        else {
            [System.Windows.MessageBox]::Show(
                "Please select a preset to edit.",
                "No Preset Selected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
    })
    $buttonsPanel.Children.Add($editButton)
    
    # Delete button
    $deleteButton = New-Object System.Windows.Controls.Button
    $deleteButton.Content = "Delete"
    $deleteButton.Width = 80
    $deleteButton.Height = 30
    $deleteButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $deleteButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(220, 53, 69))
    $deleteButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $deleteButton.Add_Click({
        $selectedPreset = $listView.SelectedItem
        if ($selectedPreset) {
            # Check if it's a default preset
            if ($selectedPreset.IsDefault) {
                [System.Windows.MessageBox]::Show(
                    "Default presets cannot be deleted.",
                    "Cannot Delete Default Preset",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                return
            }
            
            # Confirm deletion
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Are you sure you want to delete the preset '$($selectedPreset.Name)'?",
                "Confirm Deletion",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
                # Delete the preset
                $success = Remove-PasswordPreset -Name $selectedPreset.Name
                
                if ($success) {
                    # Refresh the ListView
                    $listView.Items.Clear()
                    foreach ($p in $script:PasswordPresets) {
                        [void]$listView.Items.Add($p)
                    }
                }
            }
        }
        else {
            [System.Windows.MessageBox]::Show(
                "Please select a preset to delete.",
                "No Preset Selected",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
    })
    $buttonsPanel.Children.Add($deleteButton)
    
    # Close button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Width = 80
    $closeButton.Height = 30
    $closeButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $closeButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $closeButton.Add_Click({ $presetsWindow.Close() })
    $buttonsPanel.Children.Add($closeButton)
    
    $grid.Children.Add($buttonsPanel)
    
    # Set the content and show the window
    $presetsWindow.Content = $grid
    $presetsWindow.ShowDialog() | Out-Null
}

# Shows a dialog for editing a password preset
function Show-PresetEditDialog {
    param (
        [PasswordPresetModel]$Preset,
        [bool]$IsNew = $false
    )
    
    # Create a new window for editing the preset
    $editWindow = New-Object System.Windows.Window
    $editWindow.Title = if ($IsNew) { "Add Preset" } else { "Edit Preset" }
    $editWindow.Width = 400
    $editWindow.SizeToContent = "Height"
    $editWindow.WindowStartupLocation = "CenterScreen"
    $editWindow.ResizeMode = "NoResize"
    $editWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Create a grid for the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    
    # Define row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = [System.Windows.GridLength]::Auto
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = [System.Windows.GridLength]::Auto
    $row4 = New-Object System.Windows.Controls.RowDefinition
    $row4.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    $grid.RowDefinitions.Add($row4)
    
    # Name field
    $nameLabel = New-Object System.Windows.Controls.Label
    $nameLabel.Content = "Preset Name:"
    $nameLabel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    [System.Windows.Controls.Grid]::SetRow($nameLabel, 0)
    $grid.Children.Add($nameLabel)
    
    $nameTextBox = New-Object System.Windows.Controls.TextBox
    $nameTextBox.Text = $Preset.Name
    $nameTextBox.Height = 30
    $nameTextBox.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $nameTextBox.VerticalContentAlignment = "Center"
    [System.Windows.Controls.Grid]::SetRow($nameTextBox, 0)
    $nameTextBox.Margin = New-Object System.Windows.Thickness(100, 0, 0, 10)
    $grid.Children.Add($nameTextBox)
    
    # Length field
    $lengthLabel = New-Object System.Windows.Controls.Label
    $lengthLabel.Content = "Length:"
    $lengthLabel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    [System.Windows.Controls.Grid]::SetRow($lengthLabel, 1)
    $grid.Children.Add($lengthLabel)
    
    $lengthGrid = New-Object System.Windows.Controls.Grid
    $lengthGrid.Margin = New-Object System.Windows.Thickness(100, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($lengthGrid, 1)
    
    $lengthCol1 = New-Object System.Windows.Controls.ColumnDefinition
    $lengthCol1.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $lengthCol2 = New-Object System.Windows.Controls.ColumnDefinition
    $lengthCol2.Width = New-Object System.Windows.GridLength(40)
    $lengthGrid.ColumnDefinitions.Add($lengthCol1)
    $lengthGrid.ColumnDefinitions.Add($lengthCol2)
    
    $lengthSlider = New-Object System.Windows.Controls.Slider
    $lengthSlider.Minimum = 8
    $lengthSlider.Maximum = 32
    $lengthSlider.Value = $Preset.Length
    $lengthSlider.TickFrequency = 1
    $lengthSlider.IsSnapToTickEnabled = $true
    [System.Windows.Controls.Grid]::SetColumn($lengthSlider, 0)
    $lengthGrid.Children.Add($lengthSlider)
    
    $lengthValue = New-Object System.Windows.Controls.TextBlock
    $lengthValue.Text = $Preset.Length.ToString()
    $lengthValue.VerticalAlignment = "Center"
    $lengthValue.HorizontalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($lengthValue, 1)
    $lengthGrid.Children.Add($lengthValue)
    
    $lengthSlider.Add_ValueChanged({
        $lengthValue.Text = [int]$lengthSlider.Value
    })
    
    $grid.Children.Add($lengthGrid)
    
    # Character options
    $optionsLabel = New-Object System.Windows.Controls.Label
    $optionsLabel.Content = "Options:"
    $optionsLabel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    [System.Windows.Controls.Grid]::SetRow($optionsLabel, 2)
    $grid.Children.Add($optionsLabel)
    
    $optionsPanel = New-Object System.Windows.Controls.StackPanel
    $optionsPanel.Orientation = "Vertical"
    $optionsPanel.Margin = New-Object System.Windows.Thickness(100, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($optionsPanel, 2)
    
    $uppercaseCheck = New-Object System.Windows.Controls.CheckBox
    $uppercaseCheck.Content = "Include Uppercase"
    $uppercaseCheck.IsChecked = $Preset.IncludeUppercase
    $uppercaseCheck.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    $optionsPanel.Children.Add($uppercaseCheck)
    
    $numbersCheck = New-Object System.Windows.Controls.CheckBox
    $numbersCheck.Content = "Include Numbers"
    $numbersCheck.IsChecked = $Preset.IncludeNumbers
    $numbersCheck.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    $optionsPanel.Children.Add($numbersCheck)
    
    $specialCheck = New-Object System.Windows.Controls.CheckBox
    $specialCheck.Content = "Include Special Characters"
    $specialCheck.IsChecked = $Preset.IncludeSpecial
    $specialCheck.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    $optionsPanel.Children.Add($specialCheck)
    
    # Add default selection checkbox
    $defaultSelectionCheck = New-Object System.Windows.Controls.CheckBox
    $defaultSelectionCheck.Content = "Select by default on startup"
    $defaultSelectionCheck.IsChecked = $Preset.IsSelectedByDefault
    $defaultSelectionCheck.Margin = New-Object System.Windows.Thickness(0, 5, 0, 5)
    $defaultSelectionCheck.ToolTip = "When checked, this preset will be automatically selected when the application starts"
    $optionsPanel.Children.Add($defaultSelectionCheck)
    
    $grid.Children.Add($optionsPanel)
    
    # Buttons
    $buttonsPanel = New-Object System.Windows.Controls.StackPanel
    $buttonsPanel.Orientation = "Horizontal"
    $buttonsPanel.HorizontalAlignment = "Right"
    $buttonsPanel.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    [System.Windows.Controls.Grid]::SetRow($buttonsPanel, 3)
    
    $saveButton = New-Object System.Windows.Controls.Button
    $saveButton.Content = "Save"
    $saveButton.Width = 80
    $saveButton.Height = 30
    $saveButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
    $saveButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $saveButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    # Dialog result for returning to the caller
    $dialogResult = $false
    
    $saveButton.Add_Click({
        # Validate input
        if ([string]::IsNullOrWhiteSpace($nameTextBox.Text)) {
            [System.Windows.MessageBox]::Show(
                "Preset name cannot be empty.",
                "Validation Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
        
        # Check if at least one character set is selected
        if (-not ($uppercaseCheck.IsChecked -or $numbersCheck.IsChecked -or $specialCheck.IsChecked)) {
            [System.Windows.MessageBox]::Show(
                "At least one character set must be selected.",
                "Validation Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return
        }
        
        # Update the preset with the new values
        $Preset.Name = $nameTextBox.Text
        $Preset.Length = [int]$lengthSlider.Value
        $Preset.IncludeUppercase = $uppercaseCheck.IsChecked
        $Preset.IncludeNumbers = $numbersCheck.IsChecked
        $Preset.IncludeSpecial = $specialCheck.IsChecked
        $Preset.IsSelectedByDefault = $defaultSelectionCheck.IsChecked
        
        # Set the dialog result to true
        $script:dialogResult = $true
        
        # Close the dialog
        $editWindow.Close()
    })
    $buttonsPanel.Children.Add($saveButton)
    
    $cancelButton = New-Object System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Width = 80
    $cancelButton.Height = 30
    $cancelButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
    $cancelButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $cancelButton.Add_Click({ $editWindow.Close() })
    $buttonsPanel.Children.Add($cancelButton)
    
    $grid.Children.Add($buttonsPanel)
    
    # Set the content and show the window
    $editWindow.Content = $grid
    $editWindow.ShowDialog() | Out-Null
    
    # Return the dialog result
    return $dialogResult
}

# Gets information about available updates from GitHub
function Get-UpdateInformation {
    [CmdletBinding()]
    param()

    try {
        # GitHub API URL for releases
        $apiUrl = "https://api.github.com/repos/onlyalex1984/securepassgenerator-powershell/releases"
        $global:controls.LogOutput.AppendText("Checking for updates...`n")

        # Check if GitHub API is available
        if (-not (Test-ServiceAvailability -ServiceUrl "https://api.github.com")) {
            $global:controls.LogOutput.AppendText("GitHub API: Service unavailable`n")
            return @{
                UpdateAvailable = $false
                PreReleaseUpdateAvailable = $false
                Error = "GitHub API: Service unavailable"
            }
        }

        $global:controls.LogOutput.AppendText("Connecting to GitHub API: $apiUrl`n")

        # Create a web request with specific headers and timeout
        $webRequest = [System.Net.WebRequest]::Create($apiUrl)
        $webRequest.Method = "GET"
        $webRequest.Timeout = 10000 # 10 seconds timeout
        $webRequest.UserAgent = "PowerShell-SecurePassGenerator"
        
        # Add Accept header using the proper method
        $webRequest.Accept = "application/vnd.github.v3+json"

        # Set proxy settings
        $webRequest.Proxy = [System.Net.WebRequest]::DefaultWebProxy
        $webRequest.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

        $global:controls.LogOutput.AppendText("Sending request to GitHub API...`n")
        
        try {
            # Get the response
            $response = $webRequest.GetResponse()
            $responseStream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($responseStream)
            $responseContent = $reader.ReadToEnd()
            $reader.Close()
            $response.Close()
            
            # Parse the JSON response
            $releases = $responseContent | ConvertFrom-Json
            $global:controls.LogOutput.AppendText("Successfully retrieved releases from GitHub`n")
        }
        catch [System.Net.WebException] {
            $global:controls.LogOutput.AppendText("WebException: $($_.Exception.Message)`n")
            
            # Check if there's a response with error details
            if ($_.Exception.Response) {
                $errorResponse = $_.Exception.Response
                $global:controls.LogOutput.AppendText("Status code: $($errorResponse.StatusCode)`n")
                
                # Try to read the error response body
                try {
                    $errorStream = $errorResponse.GetResponseStream()
                    $errorReader = New-Object System.IO.StreamReader($errorStream)
                    $errorContent = $errorReader.ReadToEnd()
                    $errorReader.Close()
                    $global:controls.LogOutput.AppendText("Error details: $errorContent`n")
                }
                catch {
                    $global:controls.LogOutput.AppendText("Could not read error response: $($_.Exception.Message)`n")
                }
            }
            
            # Try alternative approach with Invoke-WebRequest
            $global:controls.LogOutput.AppendText("Trying alternative approach with Invoke-WebRequest...`n")
            try {
                $releases = Invoke-RestMethod -Uri $apiUrl -Method Get -UserAgent "PowerShell-SecurePassGenerator" -TimeoutSec 10
                $global:controls.LogOutput.AppendText("Successfully retrieved releases using Invoke-RestMethod`n")
            }
            catch {
                $global:controls.LogOutput.AppendText("Alternative approach failed: $($_.Exception.Message)`n")
                throw
            }
        }

        # Find latest stable release
        $latestRelease = $releases | Where-Object { -not $_.prerelease } | Select-Object -First 1
        if (-not $latestRelease) {
            $global:controls.LogOutput.AppendText("No stable releases found`n")
            throw "No stable releases found"
        }

        # Find latest pre-release
        $latestPreRelease = $releases | Where-Object { $_.prerelease } | Select-Object -First 1

        # Parse version numbers for comparison
        $currentVersion = [Version]$script:Version
        $global:controls.LogOutput.AppendText("Current version: $currentVersion`n")
        
        # Check if tag_name exists and has a valid format
        if (-not $latestRelease.tag_name) {
            $global:controls.LogOutput.AppendText("Error: Latest release has no tag_name`n")
            throw "Latest release has no tag_name"
        }
        
        # Remove 'v' prefix if present and parse version
        $latestVersionString = $latestRelease.tag_name -replace '^v', ''
        try {
            $latestVersion = [Version]$latestVersionString
            $global:controls.LogOutput.AppendText("Latest stable version: $latestVersion`n")
        }
        catch {
            $global:controls.LogOutput.AppendText("Error parsing version from tag: $latestVersionString`n")
            throw "Invalid version format in tag_name: $($latestRelease.tag_name)"
        }
        
        # Parse pre-release version if available
        $preReleaseVersion = $null
        if ($latestPreRelease -and $latestPreRelease.tag_name) {
            $preReleaseVersionString = $latestPreRelease.tag_name -replace '^v', ''
            
            # Handle pre-release version tags like "1.1.0-pre.1"
            if ($preReleaseVersionString -match '(\d+\.\d+\.\d+)(-\w+\.\d+)?') {
                $versionPart = $matches[1]
                try {
                    $preReleaseVersion = [Version]$versionPart
                    $global:controls.LogOutput.AppendText("Latest pre-release version: $preReleaseVersion (from tag: $preReleaseVersionString)`n")
                }
                catch {
                    $global:controls.LogOutput.AppendText("Error parsing pre-release version part: $versionPart from tag: $preReleaseVersionString`n")
                    # Don't throw here, just ignore the pre-release
                    $preReleaseVersion = $null
                }
            } else {
                try {
                    # Fallback to direct parsing if the regex doesn't match
                    $preReleaseVersion = [Version]$preReleaseVersionString
                    $global:controls.LogOutput.AppendText("Latest pre-release version: $preReleaseVersion`n")
                }
                catch {
                    $global:controls.LogOutput.AppendText("Error parsing pre-release version from tag: $preReleaseVersionString`n")
                    # Don't throw here, just ignore the pre-release
                    $preReleaseVersion = $null
                }
            }
        }

        # Determine if updates are available
        $updateAvailable = $latestVersion -gt $currentVersion
        $preReleaseUpdateAvailable = $preReleaseVersion -and ($preReleaseVersion -gt $currentVersion)

        $global:controls.LogOutput.AppendText("Update available: $updateAvailable`n")
        $global:controls.LogOutput.AppendText("Pre-release update available: $preReleaseUpdateAvailable`n")

        # Return the update information
        return @{
            CurrentVersion = $currentVersion
            LatestVersion = $latestVersion
            LatestRelease = $latestRelease
            PreReleaseVersion = $preReleaseVersion
            PreRelease = $latestPreRelease
            UpdateAvailable = $updateAvailable
            PreReleaseUpdateAvailable = $preReleaseUpdateAvailable
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        $global:controls.LogOutput.AppendText("Error checking for updates: $errorMessage`n")
        
        # Check if the error is related to connectivity issues
        if ($errorMessage -like "*Det gick inte att matcha fjärrnamnet*" -or 
            $errorMessage -like "*Could not resolve host*" -or 
            $errorMessage -like "*No such host is known*" -or
            $errorMessage -like "*Unable to connect*" -or
            $errorMessage -like "*The remote name could not be resolved*") {
            return @{
                UpdateAvailable = $false
                PreReleaseUpdateAvailable = $false
                Error = "GitHub API: Service unavailable"
            }
        }
        
        $global:controls.LogOutput.AppendText("Stack trace: $($_.ScriptStackTrace)`n")
        return @{
            UpdateAvailable = $false
            PreReleaseUpdateAvailable = $false
            Error = $errorMessage
        }
    }
}

# Shows changelog information for a specific version by downloading from GitHub
function Show-Changelog {
    param (
        [string]$Version
    )
    
    # Create a new window for the changelog
    $changelogWindow = New-Object System.Windows.Window
    $changelogWindow.Title = "Changelog for $Version"
    $changelogWindow.Width = 600
    $changelogWindow.Height = 500
    $changelogWindow.WindowStartupLocation = "CenterScreen"
    $changelogWindow.ResizeMode = "CanResize"
    $changelogWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Create a grid for the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    
    # Define row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    
    # Add a title
    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = "Changelog for $Version"
    $titleBlock.FontWeight = "Bold"
    $titleBlock.FontSize = 16
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $titleBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 120, 215))
    [System.Windows.Controls.Grid]::SetRow($titleBlock, 0)
    $grid.Children.Add($titleBlock)
    
    # Create a ScrollViewer for the changelog content
    $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = "Auto"
    $scrollViewer.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($scrollViewer, 1)
    
    # Create a TextBlock for the changelog content
    $changelogContent = New-Object System.Windows.Controls.TextBlock
    $changelogContent.TextWrapping = "Wrap"
    $changelogContent.Margin = New-Object System.Windows.Thickness(5)
    
    # Set initial loading message
    $changelogContent.Text = "Loading changelog from GitHub..."
    
    # Add the TextBlock to the ScrollViewer
    $scrollViewer.Content = $changelogContent
    $grid.Children.Add($scrollViewer)
    
    # Add a close button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Width = 80
    $closeButton.Height = 30
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    $closeButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $closeButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $closeButton.Add_Click({ $changelogWindow.Close() })
    [System.Windows.Controls.Grid]::SetRow($closeButton, 2)
    $grid.Children.Add($closeButton)
    
    # Set the content and show the window
    $changelogWindow.Content = $grid
    
    # Start a background job to download and process the changelog
    $dispatcher = $changelogWindow.Dispatcher
    $dispatcher.InvokeAsync([System.Action]{
        try {
            # Determine the appropriate branch based on the version
            $branch = "main"
            # Check if the version contains a pre-release indicator (like -pre, -alpha, -beta, etc.)
            if ($Version -match "-") {
                $branch = "prerelease"
                $global:controls.LogOutput.AppendText("Detected pre-release version, using prerelease branch`n")
            }
            
            # Download the CHANGELOG.md file from GitHub
            $changelogUrl = "https://raw.githubusercontent.com/onlyalex1984/securepassgenerator-powershell/$branch/CHANGELOG.md"
            $global:controls.LogOutput.AppendText("Downloading changelog from: $changelogUrl`n")
            
            # Create a web request with specific headers and timeout
            $webRequest = [System.Net.WebRequest]::Create($changelogUrl)
            $webRequest.Method = "GET"
            $webRequest.Timeout = 10000 # 10 seconds timeout
            $webRequest.UserAgent = "PowerShell-SecurePassGenerator"
            
            # Set proxy settings
            $webRequest.Proxy = [System.Net.WebRequest]::DefaultWebProxy
            $webRequest.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            
            # Get the response
            $response = $webRequest.GetResponse()
            $responseStream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($responseStream)
            $changelogText = $reader.ReadToEnd()
            $reader.Close()
            $response.Close()
            
            $global:controls.LogOutput.AppendText("Successfully downloaded changelog from GitHub`n")
            
            # Log a simple message about showing the changelog
            $global:controls.LogOutput.AppendText("Showing changelog for version $Version`n")
            
            # Try different regex patterns to find the version
            $versionPattern = "## \[$Version\].*?(?=## \[|$)"
            
            if ($changelogText -match $versionPattern) {
                $versionChangelog = $matches[0]
                # No need to log the length of the changelog section
                
                # Format the changelog text - ensure proper encoding of special characters
                # First replace section headers, then replace bullet points with proper Unicode character
                $formattedChangelog = $versionChangelog -replace "### (.*)", "`n`${1}:" -replace "- ", ([char]0x2022 + " ")
                
                # Update the UI on the UI thread
                $dispatcher.Invoke([System.Action]{
                    # Set the text with proper encoding
                    $changelogContent.Text = $formattedChangelog.Trim()
                })
            } 
            # Try a simpler pattern as fallback
            elseif ($changelogText -match "(?s)## \[$Version\](.*?)(?=## \[|$)") {
                $versionChangelog = $matches[0]
                # No need to log the fallback pattern details
                
                # Format the changelog text - ensure proper encoding of special characters
                # First replace section headers, then replace bullet points with proper Unicode character
                $formattedChangelog = $versionChangelog -replace "### (.*)", "`n`${1}:" -replace "- ", ([char]0x2022 + " ")
                
                # Update the UI on the UI thread
                $dispatcher.Invoke([System.Action]{
                    # Set the text with proper encoding
                    $changelogContent.Text = $formattedChangelog.Trim()
                })
            }
            # Try an exact string search as a last resort
            elseif ($changelogText.Contains("## [$Version]")) {
                $global:controls.LogOutput.AppendText("Found version header but couldn't extract section. Using string search.`n")
                
                # Split the changelog by sections and find the one for our version
                $sections = $changelogText -split "## \["
                foreach ($section in $sections) {
                    if ($section.StartsWith("$Version]")) {
                        $versionChangelog = "## [$section"
                        # Truncate at the next section if present
                        $nextSectionIndex = $versionChangelog.IndexOf("## [", 5)
                        if ($nextSectionIndex -gt 0) {
                            $versionChangelog = $versionChangelog.Substring(0, $nextSectionIndex)
                        }
                        
                        # No need to log extraction details
                        
                # Format the changelog text - ensure proper encoding of special characters
                # First replace section headers, then replace bullet points with proper Unicode character
                $formattedChangelog = $versionChangelog -replace "### (.*)", "`n`${1}:" -replace "- ", ([char]0x2022 + " ")
                
                # Update the UI on the UI thread
                $dispatcher.Invoke([System.Action]{
                    # Set the text with proper encoding
                    $changelogContent.Text = $formattedChangelog.Trim()
                })
                        break
                    }
                }
            }
            else {
                $global:controls.LogOutput.AppendText("No changelog information found for version $Version`n")
                
                # Update the UI on the UI thread
                $dispatcher.Invoke([System.Action]{
                    $changelogContent.Text = "No changelog information found for version $Version."
                })
            }
        } catch {
            # Log a simple error message
            $global:controls.LogOutput.AppendText("Error downloading changelog: $($_.Exception.Message)`n")
            
            # Update the UI on the UI thread
            $dispatcher.Invoke([System.Action]{
                $changelogContent.Text = "Error downloading changelog: $($_.Exception.Message)"
            })
            
            # Try alternative approach with Invoke-WebRequest
            try {
                $global:controls.LogOutput.AppendText("Retrying download...`n")
                $changelogText = Invoke-WebRequest -Uri $changelogUrl -UseBasicParsing | Select-Object -ExpandProperty Content
                
                # Extract the section for the specified version
                $versionPattern = "## \[$Version\].*?(?=## \[|$)"
                if ($changelogText -match $versionPattern) {
                    $versionChangelog = $matches[0]
                    
                    # Format the changelog text
                    $formattedChangelog = $versionChangelog -replace "### (.*)", "`n`${1}:" -replace "- ", "• "
                    
                    # Update the UI on the UI thread
                    $dispatcher.Invoke([System.Action]{
                        $changelogContent.Text = $formattedChangelog.Trim()
                    })
                } else {
                    # Update the UI on the UI thread
                    $dispatcher.Invoke([System.Action]{
                        $changelogContent.Text = "No changelog information found for version $Version."
                    })
                }
                
                $global:controls.LogOutput.AppendText("Successfully downloaded changelog for version $Version`n")
            } catch {
                # Log a simple error message for the retry failure
                $global:controls.LogOutput.AppendText("Retry failed: $($_.Exception.Message)`n")
                
                # Update the UI on the UI thread
                $dispatcher.Invoke([System.Action]{
                    $changelogContent.Text = "Error downloading changelog: $($_.Exception.Message)"
                })
            }
        }
    })
    
    # Show the window
    $changelogWindow.ShowDialog() | Out-Null
}

# Shows a dialog with update information and options
function Show-UpdateDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$UpdateInfo
    )

    # Create update dialog window
    $updateWindow = New-Object System.Windows.Window
    $updateWindow.Title = "SecurePassGenerator Update"
    $updateWindow.Width = 450
    $updateWindow.Height = 400
    $updateWindow.WindowStartupLocation = "CenterScreen"
    $updateWindow.ResizeMode = "NoResize"
    $updateWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))

    # Create content
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = New-Object System.Windows.Thickness(15)

    # Add title
    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = "SecurePassGenerator Update"
    $titleBlock.FontSize = 16
    $titleBlock.FontWeight = "Bold"
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $titleBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 120, 215))
    $stackPanel.Children.Add($titleBlock)

    # Add current version info
    $currentVersionText = New-Object System.Windows.Controls.TextBlock
    $currentVersionText.Text = "Current Version: $($UpdateInfo.CurrentVersion)"
    $currentVersionText.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $currentVersionText.FontWeight = "SemiBold"
    $stackPanel.Children.Add($currentVersionText)

    # Add latest stable version info
    $latestVersionText = New-Object System.Windows.Controls.TextBlock
    $latestVersionText.Text = "Latest Stable Version: $($UpdateInfo.LatestVersion)"
    $latestVersionText.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
    $stackPanel.Children.Add($latestVersionText)
    
        # Add latest pre-release version info if available
        if ($UpdateInfo.PreReleaseVersion) {
            # Get the full pre-release tag (including -pre.X suffix)
            $fullPreReleaseTag = $UpdateInfo.PreRelease.tag_name -replace '^v', ''
            
            # Create a grid for pre-release version info with changelog button
            $preReleaseGrid = New-Object System.Windows.Controls.Grid
            $preReleaseGrid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
            
            # Define columns for the grid
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Auto)
            $preReleaseGrid.ColumnDefinitions.Add($col1)
            $preReleaseGrid.ColumnDefinitions.Add($col2)
            
            # Add pre-release version text
            $preReleaseVersionText = New-Object System.Windows.Controls.TextBlock
            $preReleaseVersionText.Text = "Latest Pre-Release Version: $fullPreReleaseTag"
            $preReleaseVersionText.VerticalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($preReleaseVersionText, 0)
            $preReleaseGrid.Children.Add($preReleaseVersionText)
            
            # Add Show Changelog button
            $showChangelogButton = New-Object System.Windows.Controls.Button
            $showChangelogButton.Content = "Show Changelog"
            $showChangelogButton.Padding = New-Object System.Windows.Thickness(5, 2, 5, 2)
            $showChangelogButton.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
            $showChangelogButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
            $showChangelogButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
            $showChangelogButton.Add_Click({
                Show-Changelog -Version $fullPreReleaseTag
            })
            [System.Windows.Controls.Grid]::SetColumn($showChangelogButton, 1)
            $preReleaseGrid.Children.Add($showChangelogButton)
            
            # Add the grid to the main stack panel
            $stackPanel.Children.Add($preReleaseGrid)
        }

    # Determine version status
    $runningNewerThanStable = $UpdateInfo.CurrentVersion -gt $UpdateInfo.LatestVersion

    # Add update button for stable version if update is available
    if ($UpdateInfo.UpdateAvailable) {
        $updateButton = New-Object System.Windows.Controls.Button
        $updateButton.Content = "Update to Latest Stable Version ($($UpdateInfo.LatestVersion))"
        $updateButton.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
        $updateButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
        $updateButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
        $updateButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        $updateButton.Add_Click({
            $updateWindow.DialogResult = $true
            $updateWindow.Tag = "Latest"
            $updateWindow.Close()
        })
        $stackPanel.Children.Add($updateButton)
    }
    elseif ($runningNewerThanStable) {
    # Add option to install stable version
        $installButton = New-Object System.Windows.Controls.Button
        $installButton.Content = "Install Latest Stable Version ($($UpdateInfo.LatestVersion))"
        $installButton.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
        $installButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
        $installButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(40, 167, 69)) # Green color
        $installButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        $installButton.Add_Click({
            $updateWindow.DialogResult = $true
            $updateWindow.Tag = "Latest"
            $updateWindow.Close()
        })
        $stackPanel.Children.Add($installButton)
    }
    else {
        $upToDateText = New-Object System.Windows.Controls.TextBlock
        $upToDateText.Text = "You have the latest stable version."
        $upToDateText.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
        $upToDateText.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 128, 0))
        $stackPanel.Children.Add($upToDateText)
        
        # Add option to reinstall stable version
        $reinstallButton = New-Object System.Windows.Controls.Button
        $reinstallButton.Content = "Reinstall Latest Stable Version ($($UpdateInfo.LatestVersion))"
        $reinstallButton.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
        $reinstallButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
        $reinstallButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(40, 167, 69)) # Green color
        $reinstallButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        $reinstallButton.Add_Click({
            $updateWindow.DialogResult = $true
            $updateWindow.Tag = "Latest"
            $updateWindow.Close()
        })
        $stackPanel.Children.Add($reinstallButton)
    }

    # Add separator
    $separator = New-Object System.Windows.Controls.Separator
    $separator.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $stackPanel.Children.Add($separator)

    # Add pre-release update options if available
    if ($UpdateInfo.PreReleaseVersion) {
        # Get the full pre-release tag (including -pre.X suffix)
        $fullPreReleaseTag = $UpdateInfo.PreRelease.tag_name -replace '^v', ''

        # Add update button for pre-release if newer than current
        if ($UpdateInfo.PreReleaseUpdateAvailable) {
            $preReleaseUpdateButton = New-Object System.Windows.Controls.Button
            $preReleaseUpdateButton.Content = "Update to Pre-Release Version ($fullPreReleaseTag)"
            $preReleaseUpdateButton.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
            $preReleaseUpdateButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
            $preReleaseUpdateButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0))
            $preReleaseUpdateButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
            $preReleaseUpdateButton.Add_Click({
                $updateWindow.DialogResult = $true
                $updateWindow.Tag = "PreRelease"
                $updateWindow.Close()
            })
            $stackPanel.Children.Add($preReleaseUpdateButton)

            # Add warning about pre-release versions
            $preReleaseWarningText = New-Object System.Windows.Controls.TextBlock
            $preReleaseWarningText.Text = "Warning: Pre-release versions may contain bugs or incomplete features."
            $preReleaseWarningText.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $preReleaseWarningText.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0))
            $preReleaseWarningText.TextWrapping = "Wrap"
            $stackPanel.Children.Add($preReleaseWarningText)
        }
        else {
            # If pre-release is not newer, still offer option to switch to it
            $switchToPreReleaseButton = New-Object System.Windows.Controls.Button
            $switchToPreReleaseButton.Content = "Switch to Pre-Release Version ($fullPreReleaseTag)"
            $switchToPreReleaseButton.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
            $switchToPreReleaseButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
            $switchToPreReleaseButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255)) # Blue color
            $switchToPreReleaseButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
            $switchToPreReleaseButton.Add_Click({
                $updateWindow.DialogResult = $true
                $updateWindow.Tag = "PreRelease"
                $updateWindow.Close()
            })
            $stackPanel.Children.Add($switchToPreReleaseButton)
            
            # Add warning about pre-release versions
            $preReleaseWarningText = New-Object System.Windows.Controls.TextBlock
            $preReleaseWarningText.Text = "Warning: Pre-release versions may contain bugs or incomplete features."
            $preReleaseWarningText.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $preReleaseWarningText.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0))
            $preReleaseWarningText.TextWrapping = "Wrap"
            $stackPanel.Children.Add($preReleaseWarningText)
        }
    }
    else {
        $noPreReleaseText = New-Object System.Windows.Controls.TextBlock
        $noPreReleaseText.Text = "No pre-release versions available."
        $noPreReleaseText.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
        $stackPanel.Children.Add($noPreReleaseText)
    }

    # Add close button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    $closeButton.Padding = New-Object System.Windows.Thickness(10, 5, 10, 5)
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255)) # Blue color
    $closeButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $closeButton.Add_Click({
        $updateWindow.DialogResult = $false
        $updateWindow.Close()
    })
    $stackPanel.Children.Add($closeButton)

    # Set content and show dialog
    $updateWindow.Content = $stackPanel
    $result = $updateWindow.ShowDialog()
    
    # Return both the dialog result and the selected release type
    return @{
        Result = $result
        ReleaseType = $updateWindow.Tag
    }
}

# Downloads and executes the installer with appropriate parameters
function Invoke-Update {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Latest", "PreRelease")]
        [string]$ReleaseType
    )

    try {
        # Variable to track if we're using a temporary installer
        $usingTempInstaller = $false
        $tempInstallerPath = $null
        
        # 1. Check if installer exists in the same directory as the script
        $currentDirInstallerPath = Join-Path -Path $PSScriptRoot -ChildPath "Installer.ps1"
        if (Test-Path -Path $currentDirInstallerPath) {
            $installerPath = $currentDirInstallerPath
            $global:controls.LogOutput.AppendText("Found installer in current directory: $installerPath`n")
        }
        # 2. Check in AppData directory
        elseif (Test-Path -Path $script:InstallerPath) {
            $installerPath = $script:InstallerPath
            $global:controls.LogOutput.AppendText("Found installer in AppData directory: $installerPath`n")
        }
        # 3. Check tools\installer directory
        elseif (Test-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath "tools\installer\Installer.ps1")) {
            $installerPath = Join-Path -Path $PSScriptRoot -ChildPath "tools\installer\Installer.ps1"
            $global:controls.LogOutput.AppendText("Found installer in tools\installer directory: $installerPath`n")
        }
        # 4. If not found in any location, download from GitHub to temp directory
        else {
            $global:controls.LogOutput.AppendText("Installer not found locally, downloading from GitHub to temp directory...`n")

            # Create a dedicated temp directory for the installer
            $tempDir = Join-Path -Path $env:TEMP -ChildPath "securepassgenerator"
            if (-not (Test-Path -Path $tempDir)) {
                New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
                $global:controls.LogOutput.AppendText("Created temporary directory: $tempDir`n")
            }
            
            # Use the original file name in the temp directory
            $tempInstallerPath = Join-Path -Path $tempDir -ChildPath "Installer.ps1"
            $installerPath = $tempInstallerPath
            $usingTempInstaller = $true
            
            # Download installer from GitHub to temp location
            $installerUrl = "https://raw.githubusercontent.com/onlyalex1984/securepassgenerator-powershell/main/tools/installer/Installer.ps1"
            $global:controls.LogOutput.AppendText("Downloading installer from: $installerUrl`n")
            Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

            # Unblock the file to prevent security warnings
            Unblock-File -Path $installerPath
            $global:controls.LogOutput.AppendText("Installer downloaded and unblocked successfully to: $installerPath`n")
        }

        # Prepare the update command
        $updateCommand = "powershell.exe -ExecutionPolicy Bypass -File `"$installerPath`" -InstallType DirectAPI -ReleaseType $ReleaseType"
        $global:controls.LogOutput.AppendText("Update command: $updateCommand`n")

        # Show confirmation dialog
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Are you sure you want to update to the $ReleaseType version? The application will close during the update process.",
            "Confirm Update",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
            # Execute the installer
            $global:controls.LogOutput.AppendText("Starting update process...`n")
            Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$installerPath`" -InstallType DirectAPI -ReleaseType $ReleaseType" -Wait

            # Clean up the temporary installer file if we used one
            if ($usingTempInstaller -and (Test-Path -Path $tempInstallerPath)) {
                Remove-Item -Path $tempInstallerPath -Force -ErrorAction SilentlyContinue
                $global:controls.LogOutput.AppendText("Temporary installer file removed`n")
            }

            # Inform the user that the application will close
            [System.Windows.MessageBox]::Show(
                "Update process completed. The application will now close. Please restart SecurePassGenerator after the update.",
                "Update Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Close the application
            $window.Close()
        }
        else {
            $global:controls.LogOutput.AppendText("Update cancelled by user`n")
            
            # Clean up the temporary installer file if we used one and the update was cancelled
            if ($usingTempInstaller -and (Test-Path -Path $tempInstallerPath)) {
                Remove-Item -Path $tempInstallerPath -Force -ErrorAction SilentlyContinue
                $global:controls.LogOutput.AppendText("Temporary installer file removed`n")
            }
        }
    }
    catch {
        $global:controls.LogOutput.AppendText("Error during update: $($_.Exception.Message)`n")
        
        # Clean up the temporary installer file if we used one and there was an error
        if ($usingTempInstaller -and $tempInstallerPath -and (Test-Path -Path $tempInstallerPath)) {
            Remove-Item -Path $tempInstallerPath -Force -ErrorAction SilentlyContinue
            $global:controls.LogOutput.AppendText("Temporary installer file removed`n")
        }
        
        [System.Windows.MessageBox]::Show(
            "An error occurred during the update process: $($_.Exception.Message)",
            "Update Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Displays password links history in a popup window
function Show-PasswordLinksHistory {
    # Create a new window for the password links history
    $linksWindow = New-Object System.Windows.Window
    $linksWindow.Title = "Password Links History"
    $linksWindow.Width = 600
    $linksWindow.Height = 400
    $linksWindow.WindowStartupLocation = "CenterScreen"
    $linksWindow.ResizeMode = "CanResize"
    $linksWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Create a grid for the content
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = New-Object System.Windows.Thickness(15)
    
    # Define row definitions
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = [System.Windows.GridLength]::Auto
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = [System.Windows.GridLength]::Auto
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    
    # Add a title
    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = "Password Links - Current Session"
    $titleBlock.FontWeight = "Bold"
    $titleBlock.FontSize = 14
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($titleBlock, 0)
    $grid.Children.Add($titleBlock)
    
    # Create a ListView for the links
    $listView = New-Object System.Windows.Controls.ListView
    $listView.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    [System.Windows.Controls.Grid]::SetRow($listView, 1)
    
    # Create GridView columns
    $gridView = New-Object System.Windows.Controls.GridView
    
    # Time column
    $timeColumn = New-Object System.Windows.Controls.GridViewColumn
    $timeColumn.Header = "Time"
    $timeColumn.Width = 120
    $timeBinding = New-Object System.Windows.Data.Binding("CreatedAt")
    $timeBinding.StringFormat = "HH:mm:ss"
    $timeColumn.DisplayMemberBinding = $timeBinding
    $gridView.Columns.Add($timeColumn)
    
    # URL column with truncation
    $urlColumn = New-Object System.Windows.Controls.GridViewColumn
    $urlColumn.Header = "URL"
    $urlColumn.Width = 180
    
    # Create proper IValueConverter implementation for URL truncation
    Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Data;
    using System.Globalization;

    public class UrlTruncateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string url = value as string;
            if (string.IsNullOrEmpty(url))
                return string.Empty;
                
            // Extract token from URL (e.g., https://pwpush.com/p/abcdef123456)
            if (url.Contains("/p/"))
            {
                string[] parts = url.Split(new string[] { "/p/" }, StringSplitOptions.None);
                if (parts.Length > 1)
                {
                    return "pwpush.com/p/" + parts[1];
                }
            }
            
            // Fallback if URL doesn't match expected format
            if (url.Length > 30)
                return url.Substring(0, 27) + "...";
                
            return url;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
"@ -ReferencedAssemblies PresentationFramework

    # Create an instance of the converter
    $urlTruncateConverter = New-Object UrlTruncateConverter
    
    # Create binding with converter
    $urlBinding = New-Object System.Windows.Data.Binding("Url")
    $urlBinding.Converter = $urlTruncateConverter
    $urlColumn.DisplayMemberBinding = $urlBinding
    
    $gridView.Columns.Add($urlColumn)
    
    # Status column
    $statusColumn = New-Object System.Windows.Controls.GridViewColumn
    $statusColumn.Header = "Status"
    $statusColumn.Width = 80
    
    # Create proper IValueConverter implementation for status text
    Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Data;
    using System.Globalization;

    public class StatusToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isExpired = false;
            if (value is bool)
            {
                isExpired = (bool)value;
            }
            
            if (isExpired)
                return "Expired";
            else
                return "Active";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value as string == "Expired";
        }
    }
"@ -ReferencedAssemblies PresentationFramework

    # Create proper IValueConverter implementation for status color
    Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Data;
    using System.Windows.Media;
    using System.Globalization;

    public class StatusToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isExpired = false;
            if (value is bool)
            {
                isExpired = (bool)value;
            }
            
            if (isExpired)
                return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 53, 69)); // Red for expired
            else
                return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(40, 167, 69)); // Green for active
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
"@ -ReferencedAssemblies PresentationFramework, PresentationCore, WindowsBase

    # Create instances of the converters
    $statusToTextConverter = New-Object StatusToTextConverter
    $statusToColorConverter = New-Object StatusToColorConverter
    
    # Create a factory for the status column cells
    $statusFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.TextBlock])
    $statusBinding = New-Object System.Windows.Data.Binding("IsExpired")
    $statusBinding.Converter = $statusToTextConverter
    $statusFactory.SetBinding([System.Windows.Controls.TextBlock]::TextProperty, $statusBinding)
    
    # Set color based on status
    $colorBinding = New-Object System.Windows.Data.Binding("IsExpired")
    $colorBinding.Converter = $statusToColorConverter
    $statusFactory.SetBinding([System.Windows.Controls.TextBlock]::ForegroundProperty, $colorBinding)
    
    # Create the DataTemplate for the status column
    $statusTemplate = New-Object System.Windows.DataTemplate
    $statusTemplate.VisualTree = $statusFactory
    $statusColumn.CellTemplate = $statusTemplate
    
    $gridView.Columns.Add($statusColumn)
    
    # Actions column
    $actionsColumn = New-Object System.Windows.Controls.GridViewColumn
    $actionsColumn.Header = "Actions"
    $actionsColumn.Width = 180
    
    # Create a factory for the actions column cells
    $actionsFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.StackPanel])
    $actionsFactory.SetValue([System.Windows.Controls.StackPanel]::OrientationProperty, [System.Windows.Controls.Orientation]::Horizontal)
    
    # Create Copy button
    $copyButtonFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Button])
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::ContentProperty, "Copy")
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::MarginProperty, (New-Object System.Windows.Thickness(0, 0, 5, 0)))
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::WidthProperty, [System.Windows.Controls.Button]::WidthProperty.DefaultMetadata.DefaultValue)
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::HeightProperty, [System.Windows.Controls.Button]::HeightProperty.DefaultMetadata.DefaultValue)
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::BackgroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))))
    $copyButtonFactory.SetValue([System.Windows.Controls.Button]::ForegroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))))
    
    # Create proper IValueConverter implementation for inverting boolean values
    Add-Type -TypeDefinition @"
    using System;
    using System.Windows.Data;
    using System.Globalization;

    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool)
            {
                bool boolValue = (bool)value;
                return !boolValue;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool)
            {
                bool boolValue = (bool)value;
                return !boolValue;
            }
            return false;
        }
    }
"@ -ReferencedAssemblies PresentationFramework, PresentationCore, WindowsBase

    # Create an instance of the converter
    $inverseBoolConverter = New-Object InverseBoolConverter
    
    # Add IsEnabled binding to disable button when password is expired
    $isEnabledBinding = New-Object System.Windows.Data.Binding("IsExpired")
    $isEnabledBinding.Converter = $inverseBoolConverter
    $copyButtonFactory.SetBinding([System.Windows.Controls.Button]::IsEnabledProperty, $isEnabledBinding)
    
    $copyButtonFactory.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($buttonSender, $e)
        $item = $buttonSender.DataContext
        try {
            [System.Windows.Forms.Clipboard]::SetText($item.Url)
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Link copied to clipboard`n")
            }
        }
        catch {
            # Safely access Exception.Message with a fallback message if it's null
            $errorMessage = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { "Unknown error" }
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Error copying link: $errorMessage`n")
            }
        }
    })
    $actionsFactory.AppendChild($copyButtonFactory)
    
    # Create Browse button
    $browseButtonFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Button])
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::ContentProperty, "Browse")
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::MarginProperty, (New-Object System.Windows.Thickness(0, 0, 5, 0)))
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::WidthProperty, [System.Windows.Controls.Button]::WidthProperty.DefaultMetadata.DefaultValue)
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::HeightProperty, [System.Windows.Controls.Button]::HeightProperty.DefaultMetadata.DefaultValue)
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::BackgroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))))
    $browseButtonFactory.SetValue([System.Windows.Controls.Button]::ForegroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))))
    
    # Add IsEnabled binding to disable browse button when password is expired
    $browseIsEnabledBinding = New-Object System.Windows.Data.Binding("IsExpired")
    $browseIsEnabledBinding.Converter = $inverseBoolConverter
    $browseButtonFactory.SetBinding([System.Windows.Controls.Button]::IsEnabledProperty, $browseIsEnabledBinding)
    
    $browseButtonFactory.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($buttonSender, $e)
        $item = $buttonSender.DataContext
        try {
            Start-Process $item.Url
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Opening link in browser`n")
            }
        }
        catch {
            # Safely access Exception.Message with a fallback message if it's null
            $errorMessage = if ($_.Exception -and $_.Exception.Message) { $_.Exception.Message } else { "Unknown error" }
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Error opening link: $errorMessage`n")
            }
        }
    })
    $actionsFactory.AppendChild($browseButtonFactory)
    
    # Create Expire button
    $expireButtonFactory = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Button])
    $expireButtonFactory.SetValue([System.Windows.Controls.Button]::ContentProperty, "Expire")
    $expireButtonFactory.SetValue([System.Windows.Controls.Button]::WidthProperty, [System.Windows.Controls.Button]::WidthProperty.DefaultMetadata.DefaultValue)
    $expireButtonFactory.SetValue([System.Windows.Controls.Button]::HeightProperty, [System.Windows.Controls.Button]::HeightProperty.DefaultMetadata.DefaultValue)
    $expireButtonFactory.SetValue([System.Windows.Controls.Button]::BackgroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(220, 53, 69))))
    $expireButtonFactory.SetValue([System.Windows.Controls.Button]::ForegroundProperty, (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))))
    
    # Add IsEnabled binding to disable expire button when password is expired
    $expireIsEnabledBinding = New-Object System.Windows.Data.Binding("IsExpired")
    $expireIsEnabledBinding.Converter = $inverseBoolConverter
    $expireButtonFactory.SetBinding([System.Windows.Controls.Button]::IsEnabledProperty, $expireIsEnabledBinding)
    
$expireButtonFactory.AddHandler([System.Windows.Controls.Button]::ClickEvent, [System.Windows.RoutedEventHandler]{
        param($buttonSender, $e)
        $item = $buttonSender.DataContext
        
        # Skip if already expired
        if ($item.IsExpired) {
            return
        }
        
        # Extract token from URL
        $token = $item.GetToken()
        
        if (-not $token) {
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Error: Could not extract token from URL`n")
            }
            return
        }
        
        # Show confirmation dialog
        $confirmResult = [System.Windows.MessageBox]::Show(
            "Are you sure you want to expire this password? This action cannot be undone.",
            "Confirm Password Expiration",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
            # Call the expire function
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText("Attempting to expire password with token: $token...`n")
            }
            $result = Remove-Password -Token $token
            
            # Add null check before calling AppendText
            if ($global:controls -and $global:controls.LogOutput) {
                $global:controls.LogOutput.AppendText($result.Log + "`n")
            }
            
            if ($result.Success) {
                # Update the item's expired status
                $item.IsExpired = $true
                
                # Update the UI to reflect the expired status
                # Add null check before calling AppendText
                if ($global:controls -and $global:controls.LogOutput) {
                    $global:controls.LogOutput.AppendText("Password link expired successfully`n")
                    $global:controls.LogOutput.AppendText("Status: Expired`n")
                }
                
                # Find the parent StackPanel that contains all buttons
                $stackPanel = $buttonSender.Parent
                if ($null -ne $stackPanel) {
                    # Update all buttons in the StackPanel
                    foreach ($child in $stackPanel.Children) {
                        if ($child -is [System.Windows.Controls.Button]) {
                            # Force UI update for each button
                            $child.GetBindingExpression([System.Windows.Controls.Button]::IsEnabledProperty).UpdateTarget()
                        }
                    }
                }
                
                # Find the ListView and refresh it
                $listView = $buttonSender.FindName("listView")
                if ($null -eq $listView) {
                    # Try to find the ListView by traversing up the visual tree
                    $parent = $buttonSender.Parent
                    while ($null -ne $parent -and $null -eq $listView) {
                        if ($parent -is [System.Windows.Controls.ListView]) {
                            $listView = $parent
                        }
                        $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($parent)
                    }
                }
                
                # Refresh the ListView to show updated status
                if ($null -ne $listView) {
                    $listView.Items.Refresh()
                }
                
                # Force a more comprehensive UI update
                [System.Windows.Data.BindingOperations]::GetBindingExpression($buttonSender, [System.Windows.Controls.Button]::IsEnabledProperty).UpdateSource()
                
                # Explicitly update the Copy and Browse buttons by finding them in the same row
                $row = $buttonSender.DataContext
                $copyButton = $stackPanel.Children | Where-Object { $_.Content -eq "Copy" -and $_.DataContext -eq $row }
                $browseButton = $stackPanel.Children | Where-Object { $_.Content -eq "Browse" -and $_.DataContext -eq $row }
                
                if ($null -ne $copyButton) {
                    $copyButton.IsEnabled = $false
                }
                
                if ($null -ne $browseButton) {
                    $browseButton.IsEnabled = $false
                }
            }
            else {
                $global:controls.LogOutput.AppendText("Failed to expire password. Please try again.`n")
            }
        }
    })
    $actionsFactory.AppendChild($expireButtonFactory)
    
    # Create the DataTemplate for the actions column
    $actionsTemplate = New-Object System.Windows.DataTemplate
    $actionsTemplate.VisualTree = $actionsFactory
    $actionsColumn.CellTemplate = $actionsTemplate
    
    $gridView.Columns.Add($actionsColumn)
    
    # Set the GridView as the View for the ListView
    $listView.View = $gridView
    
    # Add items to the ListView
    foreach ($link in $script:PasswordLinks) {
        [void]$listView.Items.Add($link)
    }
    
    $grid.Children.Add($listView)
    
    # Add a close button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Padding = New-Object System.Windows.Thickness(20, 5, 20, 5)
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
    $closeButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $closeButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $closeButton.Add_Click({ $linksWindow.Close() })
    [System.Windows.Controls.Grid]::SetRow($closeButton, 2)
    $grid.Children.Add($closeButton)
    
    # Set the content and show the window
    $linksWindow.Content = $grid
    $linksWindow.ShowDialog() | Out-Null
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# Swedish character definitions using Unicode escape sequences
$smallA_ring = [char]0x00E5  # å
$smallA_umlaut = [char]0x00E4  # ä
$smallO_umlaut = [char]0x00F6  # ö
$capitalA_ring = [char]0x00C5  # Å
$capitalA_umlaut = [char]0x00C4  # Ä
$capitalO_umlaut = [char]0x00D6  # Ö

# Calculates password entropy based on character sets used
function Get-PasswordEntropy {
    param (
        [System.Security.SecureString]$Password
    )
    
    # Convert SecureString to plain text for entropy calculation
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    # Analyze the actual password content to determine which character sets are used
    $hasLowercase = $PlainPassword -cmatch '[a-z]'
    $hasUppercase = $PlainPassword -cmatch '[A-Z]'
    $hasNumbers = $PlainPassword -cmatch '[0-9]'
    $hasSpecial = $PlainPassword -match '[^a-zA-Z0-9]'
    
    # Calculate pool size based on character sets actually used
    $poolSize = 0
    if ($hasLowercase) { $poolSize += 26 } # Lowercase letters
    if ($hasUppercase) { $poolSize += 26 } # Uppercase letters
    if ($hasNumbers) { $poolSize += 10 }   # Numbers
    if ($hasSpecial) { $poolSize += 6 }    # Special characters (matching C# version)
    
    # Default to lowercase if no character sets detected
    if ($poolSize -eq 0) { $poolSize = 26 }
    
    $entropy = [Math]::Log($poolSize, 2) * $PlainPassword.Length
    return $entropy
}

# Generates random password with customizable options
function New-RandomPassword {
    param (
        [int]$Length = 12,
        [bool]$IncludeUppercase = $true,
        [bool]$IncludeNumbers = $true,
        [bool]$IncludeSpecial = $true
    )
    
    $lowercase = "abcdefghijklmnopqrstuvwxyz"
    $uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $numbers = "0123456789"
    $special = "!@#-?_"
    
    $charSet = $lowercase
    if ($IncludeUppercase) { $charSet += $uppercase }
    if ($IncludeNumbers) { $charSet += $numbers }
    if ($IncludeSpecial) { $charSet += $special }
    
    $random = New-Object System.Random
    $result = ""
    
    # Ensure at least one character from each selected character set
    if ($IncludeUppercase) {
        $result += $uppercase[$random.Next(0, $uppercase.Length)]
    }
    if ($IncludeNumbers) {
        $result += $numbers[$random.Next(0, $numbers.Length)]
    }
    if ($IncludeSpecial) {
        $result += $special[$random.Next(0, $special.Length)]
    }
    
    # Fill the rest of the password
    for ($i = $result.Length; $i -lt $Length; $i++) {
        $result += $charSet[$random.Next(0, $charSet.Length)]
    }
    
    # Shuffle the password
    $resultArray = $result.ToCharArray()
    $n = $resultArray.Length
    while ($n -gt 1) {
        $n--
        $k = $random.Next(0, $n + 1)
        $temp = $resultArray[$k]
        $resultArray[$k] = $resultArray[$n]
        $resultArray[$n] = $temp
    }
    
    return -join $resultArray
}

# Word lists for generating memorable passwords
$script:EnglishWords = @("time", "year", "people", "way", "day", "man", "thing", "woman", "life", "child", "world", "school", "state", "family", "student", 
               "group", "country", "problem", "hand", "part", "place", "case", "week", "company", "system", "question", "work", "government", "number", "night", 
               "point", "home", "water", "room", "mother", "area", "money", "story", "fact", "month", "book", "eye", "job", "word", "business", 
               "issue", "side", "kind", "head", "house", "service", "friend", "father", "power", "hour", "game", "line", "end", "member", "law", 
               "car", "city", "community", "name", "president", "team", "minute", "idea", "kid", "body", "information", "back", "parent", "face", "others", 
               "level", "office", "door", "health", "person", "art", "war", "history", "party", "result", "change", "morning", "reason", "research", "girl", 
               "guy", "moment", "air", "teacher", "force", "education", "foot", "boy", "food", "energy", "table", "chair", "window", "phone", "computer", 
               "document", "music", "movie", "garden", "street", "road", "building", "sun", "moon", "star", "cloud", "rain", "snow", "tree", "flower", "grass", 
               "river", "mountain", "beach", "ocean", "island", "forest", "animal", "bird", "fish", "dog", "cat", "horse", "bear", "lion", "tiger", 
               "rabbit", "mouse", "chicken", "duck", "coffee", "tea", "milk", "juice", "bread", "cheese", "meat", "fruit", "apple", "orange", "banana", 
               "potato", "tomato", "carrot", "rice", "pasta", "soup", "salad", "cake", "sugar", "salt", "pepper", "kitchen", "bathroom", "bedroom", "living", "wall", 
               "floor", "ceiling", "roof", "key", "lock", "box", "bag", "clock", "watch", "picture", "camera", "radio", "television", "internet", "email", 
               "message", "letter", "card", "paper", "pencil", "pen", "color", "paint", "shirt", "pants", "shoes", "hat", "coat", "dress", "ring", 
               "glass", "bottle", "cup", "plate", "bowl", "knife", "fork", "spoon", "market", "store", "shop", "mall", "bank", "hospital", "doctor", 
               "nurse", "police", "fire", "library", "museum", "park", "shore", "pool", "gym", "sport", "ball", "race", "jump", "run", "walk", 
               "sleep", "dream", "think", "speak", "listen", "read", "write", "learn", "teach", "help", "labor", "play", "dance", "sing", "laugh", 
               "smile", "cry", "love", "hate", "feel", "touch", "see", "hear", "taste", "smell", "give", "take", "make", "break", "open", "close", 
               "start", "halt", "move", "stay", "come", "go", "arrive", "leave", "enter", "exit", "push", "pull", "carry", "drop", "catch", "throw", 
               "buy", "sell", "pay", "save", "spend", "find", "lose", "win", "fail", "try", "test", "pass", "stop", "continue", "begin", "finish", 
               "live", "die", "grow", "shrink", "rise", "fall", "increase", "decrease", "expand", "reduce")
# Swedish words without å, ä, ö
$script:SwedishWords = @("tid", "dag", "man", "barn", "liv", "hand", "del", "plats", "vecka", "grupp", 
                 "land", "problem", "fall", "system", "arbete", "morgon", "punkt", "hem", "vatten", "rum", 
                 "pengar", "bok", "ord", "sida", "hus", "bil", "stad", "namn", "minut", "kropp", 
                 "skola", "student", "klass", "kontor", "chef", "kund", "projekt", "rapport", "schema", "test", 
                 "data", "dator", "telefon", "mobil", "mail", "program", "modul", "kurs", "lektion", "prov", 
                 "bord", "stol", "madrass", "lampa", "spegel", "ur", "radio", "teve", "soffa", "matta", 
                 "kudde", "tavla", "gardin", "hylla", "bokhylla", "matbord", "diskho", "spis", "ugn", "kyl", 
                 "mat", "bulle", "dryck", "kaffe", "te", "vatten", "juice", "frukt", "glass", "godis", 
                 "lunch", "middag", "soppa", "sallad", "pasta", "ris", "protein", "fisk", "salt", "socker", 
                 "sol", "moln", "regn", "frost", "vind", "storm", "himmel", "jord", "berg", "dal", 
                 "skog", "strand", "hav", "vik", "flod", "sten", "sand", "grus", "plant", "blad", 
                 "hund", "katt", "duva", "svan", "zebra", "orm", "mus", "myra", "spindel", "insekt", 
                 "lejon", "tiger", "varg", "ekorre", "panter", "sal", "val", "delfin", "elefant", "giraff", 
                 "svart", "vit", "purpur", "grann", "ljus", "gul", "rosa", "lila", "brun", "silver", 
                 "guld", "metall", "plast", "glas", "kristall", "tyg", "papper", "granit", "mineral", "lera", 
                 "buss", "trafik", "pendel", "plan", "skepp", "cykel", "moped", "lastbil", "vagn", "taxi", 
                 "tunnel", "bro", "gata", "torg", "park", "station", "hamn", "flygplats", "garage", "terminal", 
                 "minut", "timme", "natt", "vecka", "helg", "period", "sommar", "vinter", "termin", "epok", 
                 "ett", "par", "tre", "fyra", "fem", "sex", "sju", "nio", "tio", "hundra", 
                 "gilla", "tycka", "tanka", "saga", "skriva", "studera", "springa", "hoppa", "sitta", "ligga", 
                 "sova", "vakna", "dricka", "prata", "lyssna", "titta", "vinka", "komma", "rusa", "starta", 
                 "stor", "liten", "hel", "kort", "bred", "smal", "tjock", "tunn", "varm", "kall", 
                 "stark", "svag", "snabb", "smart", "glad", "ledsen", "tyst", "vacker", "ren", "smutsig", 
                 "fot", "arm", "ben", "huvud", "hals", "mage", "rygg", "puls", "torso", "handled", 
                 "finger", "tand", "tunga", "panna", "mun", "haka", "lugg", "nagel", "kind", "axel", 
                 "byxa", "linne", "jacka", "rock", "skjorta", "slips", "dress", "kostym", "mossa", "vante", 
                 "sax", "hammare", "skruv", "spik", "penna", "block", "mobil", "kamera", "verktyg", "knapp", 
                 "doktor", "polis", "mentor", "kock", "pilot", "artist", "snickare", "konsult", "bagare", "bilist", 
                 "bank", "butik", "hotell", "teater", "bio", "kyrka", "slott", "fabrik", "verkstad", "kontor", 
                 "sak", "form", "bild", "text", "film", "ljud", "ljus", "kraft", "plan", "ide")

# Generates memorable password using word lists
function New-MemorablePassword {
    param (
        [int]$WordCount = 3,
        [string]$Language = "Swedish",
        [bool]$IncludeUppercase = $true,
        [bool]$IncludeNumbers = $true,
        [bool]$IncludeSpecial = $true
    )
    
    # Select word list based on language
    $words = switch ($Language) {
        "English" { $script:EnglishWords }
        "Swedish" { $script:SwedishWords }
        default { $script:SwedishWords }
    }
    
    $random = New-Object System.Random
    
    # Generate words for the password
    $selectedWords = @()
    for ($i = 0; $i -lt $WordCount; $i++) {
        $word = $words[$random.Next(0, $words.Length)]
        
        # Apply uppercase if enabled
        if ($IncludeUppercase) {
            $word = $word.Substring(0, 1).ToUpper() + $word.Substring(1)
        }
        
        $selectedWords += $word
    }
    
    # Prepare extras (numbers and special characters)
    $extras = @()
    
    # Add a number if enabled
    if ($IncludeNumbers) {
        $extras += $random.Next(0, 10).ToString()
    }
    
    # Add a special character if enabled
    if ($IncludeSpecial) {
        $specialChars = "!@#-?_"
        $extras += $specialChars[$random.Next(0, $specialChars.Length)].ToString()
    }
    
    # Insert extras at word boundaries (before first word, between words, or after last word)
    return Add-ExtrasAtWordBoundaries -Words $selectedWords -Extras $extras
}

# Adds special characters and numbers at word boundaries
function Add-ExtrasAtWordBoundaries {
    param (
        [string[]]$Words,
        [string[]]$Extras
    )
    
    # If no extras to add, just join the words
    if ($Extras.Count -eq 0) {
        return [string]::Join("", $Words)
    }
    
    $random = New-Object System.Random
    
    # Shuffle the extras to randomize their order
    $Extras = $Extras | Sort-Object { $random.Next() }
    
    # Create a list of possible positions (before first word, between words, after last word)
    $positions = @()
    for ($i = 0; $i -le $Words.Count; $i++) {
        $positions += $i
    }
    
    # Shuffle the positions to randomize where extras are inserted
    $positions = $positions | Sort-Object { $random.Next() }
    
    # Take only as many positions as we have extras
    $positions = $positions[0..($Extras.Count-1)] | Sort-Object
    
    # Build the result
    $result = ""
    
    for ($i = 0; $i -le $Words.Count; $i++) {
        # If this position is selected for an extra, add the next extra
        if ($positions -contains $i -and $Extras.Count -gt 0) {
            $result += $Extras[0]
            $Extras = $Extras[1..($Extras.Count-1)]
        }
        
        # Add the word if we're not past the last word
        if ($i -lt $Words.Count) {
            $result += $Words[$i]
        }
    }
    
    return $result
}

# Expires a password via Password Pusher API
function Remove-Password {
    param (
        [string]$Token
    )

    $apiUrl = "https://pwpush.com/p/$Token.json"
    $logOutput = "Attempting to expire password with token: $Token..."

    # Check if the Password Pusher service is available
    if (-not (Test-ServiceAvailability -ServiceUrl "https://pwpush.com")) {
        return @{
            Success = $false
            Log = "Password Pusher API: Service unavailable"
        }
    }

    try {
        # Send DELETE request to expire the password
        $headers = @{
            "Accept" = "application/json"
            "Content-Type" = "application/json"
        }
        
        # Use Invoke-RestMethod without fallback to curl
        Invoke-RestMethod -Uri $apiUrl -Method Delete -Headers $headers | Out-Null
        $logOutput += "`nPassword successfully expired using Invoke-RestMethod."
        
        $logOutput += "`nPassword successfully expired."
        $logOutput += "`nStatus: Expired"

        return @{
            Success = $true
            Log = $logOutput
        }
    }
    catch {
        # Improved error handling
        $errorMessage = $_.Exception.Message
        $logOutput += "`nError: $errorMessage"
        
        # Check if it's a 404 Not Found error, which might mean the password is already expired
        if ($errorMessage -like "*404*" -or $errorMessage -like "*Not Found*") {
            $logOutput += "`nThe password may already be expired or not found."
            # Return success anyway to update the UI
            return @{
                Success = $true
                Log = $logOutput
            }
        }
        
        # For other errors, return failure
        return @{
            Success = $false
            Log = $logOutput
        }
    }
}

# Verifies if a service endpoint is available
function Test-ServiceAvailability {
    param (
        [string]$ServiceUrl
    )

    try {
        # For HIBP API, we'll use a different approach since it doesn't respond to HEAD requests
        if ($ServiceUrl -like "*pwnedpasswords.com*") {
            $controls.LogOutput.AppendText("Testing HIBP API availability with a direct GET request`n")
            
            # Use a direct GET request to the API with a known prefix
            # This is a more reliable way to check if the API is available
            $testPrefix = "00000" # A simple prefix that will return minimal data
            $testUrl = "https://api.pwnedpasswords.com/range/$testPrefix"
            
            try {
                # Create a WebRequest with default proxy
                $request = [System.Net.WebRequest]::Create($testUrl)
                $request.Method = "GET"
                $request.Timeout = 10000 # 10 seconds
                $request.UserAgent = "PowerShell-PasswordGenerator"
                
                # Set proxy and credentials
                $request.Proxy = [System.Net.WebRequest]::DefaultWebProxy
                $request.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
                
                # Get the response
                $webResponse = $request.GetResponse()
                $webResponse.Close()
                
                $controls.LogOutput.AppendText("HIBP API is available`n")
                return $true
            }
            catch {
                $controls.LogOutput.AppendText("HIBP API is not available: $($_.Exception.Message)`n")
                return $false
            }
        }
        # For Password Pusher API, also use a direct GET request approach
        elseif ($ServiceUrl -like "*pwpush.com*") {
            $controls.LogOutput.AppendText("Testing Password Pusher API availability with a direct GET request`n")
            
            # Use a direct GET request to the main website instead of the API endpoint
            # This is more reliable as the main site should always respond to GET requests
            $testUrl = "https://pwpush.com"
            
            try {
                # Create a WebRequest with default proxy
                $request = [System.Net.WebRequest]::Create($testUrl)
                $request.Method = "GET"
                $request.Timeout = 10000 # 10 seconds
                $request.UserAgent = "PowerShell-PasswordGenerator"
                
                # Set proxy and credentials
                $request.Proxy = [System.Net.WebRequest]::DefaultWebProxy
                $request.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
                
                # Get the response
                $webResponse = $request.GetResponse()
                $webResponse.Close()
                
                $controls.LogOutput.AppendText("Password Pusher API is available`n")
                return $true
            }
            catch {
                $controls.LogOutput.AppendText("Password Pusher API is not available: $($_.Exception.Message)`n")
                return $false
            }
        }
        # For GitHub API, also use a direct GET request approach to handle corporate proxies
        elseif ($ServiceUrl -like "*api.github.com*") {
            $controls.LogOutput.AppendText("Testing GitHub API availability with a direct GET request`n")
            
            # Use a direct GET request to the API endpoint
            # This is more reliable than ping tests which are often blocked by corporate proxies
            $testUrl = "https://api.github.com"
            
            try {
                # Create a WebRequest with default proxy
                $request = [System.Net.WebRequest]::Create($testUrl)
                $request.Method = "GET"
                $request.Timeout = 10000 # 10 seconds
                $request.UserAgent = "PowerShell-PasswordGenerator"
                
                # Add Accept header for GitHub API
                $request.Accept = "application/vnd.github.v3+json"
                
                # Set proxy and credentials
                $request.Proxy = [System.Net.WebRequest]::DefaultWebProxy
                $request.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
                
                # Get the response
                $webResponse = $request.GetResponse()
                $webResponse.Close()
                
                $controls.LogOutput.AppendText("GitHub API is available`n")
                return $true
            }
            catch {
                $controls.LogOutput.AppendText("GitHub API is not available: $($_.Exception.Message)`n")
                return $false
            }
        }
        else {
            # For other services, use the original method
            # Extract the hostname from the URL
            $uri = [System.Uri]$ServiceUrl
            $hostname = $uri.Host

            # Try to resolve the hostname and check connectivity
            $result = Test-Connection -ComputerName $hostname -Count 1 -Quiet
            
            # If Test-Connection succeeds, try a more direct HTTP check
            if ($result) {
                try {
                    # Try a HEAD request with a short timeout
                    $request = [System.Net.WebRequest]::Create($ServiceUrl)
                    $request.Method = "HEAD"
                    $request.Timeout = 5000 # 5 seconds timeout
                    $request.GetResponse().Close()
                    return $true
                }
                catch {
                    # Log the error but don't fail yet - the ping test passed
                    Write-Debug "HTTP check failed: $_"
                    # Still return true since the ping test passed
                    return $true
                }
            }
            
            return $result
        }
    }
    catch {
        $controls.LogOutput.AppendText("Error checking service availability: $($_.Exception.Message)`n")
        return $false
    }
}

# Checks if password exists in HIBP database
function Test-PwnedPassword {
    param (
        [System.Security.SecureString]$Password
    )

    # Check if the HIBP service is available
    if (-not (Test-ServiceAvailability -ServiceUrl "https://api.pwnedpasswords.com")) {
        return @{
            Found = $false
            Error = "Have I Been Pwned API: Service unavailable"
        }
    }

    # Convert SecureString to plain text for API check
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    try {
        # Convert the password to a SHA-1 hash
        $sha1 = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider
        $passwordBytes = [System.Text.Encoding]::UTF8.GetBytes($PlainPassword)
        $hashBytes = $sha1.ComputeHash($passwordBytes)
        $hash = [System.BitConverter]::ToString($hashBytes).Replace("-", "")
        
        # Get the first 5 characters of the hash (prefix)
        $prefix = $hash.Substring(0, 5)
        $suffix = $hash.Substring(5).ToUpper()
        
        # Query the HIBP API with the prefix
        $response = Invoke-RestMethod -Uri "https://api.pwnedpasswords.com/range/$prefix" -Method Get -UserAgent "PowerShell-PasswordGenerator"
        
        # Check if the suffix is in the response
        $lines = $response -split "`r`n"
        foreach ($line in $lines) {
            $parts = $line -split ":"
            if ($parts[0] -eq $suffix) {
                return @{
                    Found = $true
                    Count = [int]$parts[1]
                }
            }
        }
        
        return @{
            Found = $false
            Count = 0
        }
    }
    catch {
        return @{
            Found = $false
            Error = $_.Exception.Message
        }
    }
}
# Returns phonetic code for a character in NATO or Swedish system
function Get-PhoneticCode {
    param (
        [char]$Character,
        [string]$PhoneticLanguage = "NATO"
    )
    
    # Ensure we're using UTF-8 encoding
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    
    $lowerChar = [char]::ToLower($Character)
    $isUppercase = [char]::IsUpper($Character)
    
    # NATO phonetic alphabet mapping
    $natoPhonetic = @{
        'a' = 'Alpha'; 'b' = 'Bravo'; 'c' = 'Charlie'; 'd' = 'Delta'; 'e' = 'Echo';
        'f' = 'Foxtrot'; 'g' = 'Golf'; 'h' = 'Hotel'; 'i' = 'India'; 'j' = 'Juliet';
        'k' = 'Kilo'; 'l' = 'Lima'; 'm' = 'Mike'; 'n' = 'November'; 'o' = 'Oscar';
        'p' = 'Papa'; 'q' = 'Quebec'; 'r' = 'Romeo'; 's' = 'Sierra'; 't' = 'Tango';
        'u' = 'Uniform'; 'v' = 'Victor'; 'w' = 'Whiskey'; 'x' = 'X-ray'; 'y' = 'Yankee';
        'z' = 'Zulu';
        '0' = 'Zero'; '1' = 'One'; '2' = 'Two'; '3' = 'Three'; '4' = 'Four';
        '5' = 'Five'; '6' = 'Six'; '7' = 'Seven'; '8' = 'Eight'; '9' = 'Nine';
        '!' = 'Exclamation Mark'; '@' = 'At Sign'; '#' = 'Hash'; '-' = 'Dash';
        '?' = 'Question Mark'; '_' = 'Underscore'; ' ' = 'Space';
        '/' = 'Forward Slash'; '\' = 'Backslash'; '"' = 'Double Quote';
        '&' = 'Ampersand'; '$' = 'Dollar Sign'; '*' = 'Asterisk'; '=' = 'Equals Sign';
        '%' = 'Percent'; '^' = 'Caret'; '+' = 'Plus';
        '[' = 'Left Square Bracket'; ']' = 'Right Square Bracket';
        '{' = 'Left Curly Brace'; '}' = 'Right Curly Brace';
        '|' = 'Vertical Bar'; '<' = 'Less Than'; '>' = 'Greater Than'
    }
    
    # Define Swedish phonetic codes using the Unicode variables
    $swedishA = "$capitalA_ring" + "ke"      # Åke
    $swedishAE = "$capitalA_umlaut" + "rlig"   # Ärlig
    $swedishO = "$capitalO_umlaut" + "sten"    # Östen
    
    # Swedish military phonetic alphabet (with proper Swedish characters)
    $swedishPhonetic = @{
        'a' = 'Adam'; 'b' = 'Bertil'; 'c' = 'Cesar'; 'd' = 'David'; 'e' = 'Erik';
        'f' = 'Filip'; 'g' = 'Gustav'; 'h' = 'Helge'; 'i' = 'Ivar'; 'j' = 'Johan';
        'k' = 'Kalle'; 'l' = 'Ludvig'; 'm' = 'Martin'; 'n' = 'Niklas'; 'o' = 'Olof';
        'p' = 'Petter'; 'q' = 'Qvintus'; 'r' = 'Rudolf'; 's' = 'Sigurd'; 't' = 'Tore';
        'u' = 'Urban'; 'v' = 'Viktor'; 'w' = 'Wilhelm'; 'x' = 'Xerxes'; 'y' = 'Yngve';
        'z' = 'Z' + $smallA_umlaut + 'ta';
        '0' = 'Noll'; '1' = 'Ett'; '2' = 'Tv' + $smallA_ring; '3' = 'Tre'; '4' = 'Fyra';
        '5' = 'Fem'; '6' = 'Sex'; '7' = 'Sju'; '8' = $capitalA_ring + 'tta'; '9' = 'Nio';
        '!' = 'Utropstecken'; '@' = 'Snabel-a'; '#' = 'Br' + $smallA_umlaut + 'dg' + $smallA_ring + 'rd'; '-' = 'Streck';
        '?' = 'Fr' + $smallA_ring + 'getecken'; '_' = 'Understreck'; ' ' = 'Mellanslag';
        '/' = 'Snedstreck'; '\' = 'Omv' + $smallA_umlaut + 'nt Snedstreck'; '´' = 'Akut Accent'; '"' = 'Citattecken';
        '&' = 'Et-tecken'; '$' = 'Dollartecken'; '*' = 'Asterisk'; '=' = 'Likhetstecken';
        '%' = 'Procent'; '^' = 'Cirkumflex'; '+' = 'Plus';
        '[' = 'V' + $smallA_umlaut + 'nster Hakparentes'; ']' = 'H' + $smallO_umlaut + 'ger Hakparentes';
        '{' = 'V' + $smallA_umlaut + 'nster Krullparentes'; '}' = 'H' + $smallO_umlaut + 'ger Krullparentes';
        '|' = 'Vertikalstreck'; '<' = 'Mindre ' + $smallA_umlaut + 'n'; '>' = 'St' + $smallO_umlaut + 'rre ' + $smallA_umlaut + 'n'
    }
    
    # Select the appropriate phonetic alphabet
    $phoneticAlphabet = if ($PhoneticLanguage -eq "NATO") { $natoPhonetic } else { $swedishPhonetic }
    
    # Get phonetic code
    $phonetic = $null
    
    # Special handling for Swedish characters
    if ($PhoneticLanguage -eq "Swedish") {
        if ($lowerChar -eq $smallA_ring -or $lowerChar -eq $capitalA_ring) {
            $phonetic = $swedishA
        }
        elseif ($lowerChar -eq $smallA_umlaut -or $lowerChar -eq $capitalA_umlaut) {
            $phonetic = $swedishAE
        }
        elseif ($lowerChar -eq $smallO_umlaut -or $lowerChar -eq $capitalO_umlaut) {
            $phonetic = $swedishO
        }
    }
    # Special handling for Swedish characters in NATO phonetic
    elseif ($PhoneticLanguage -eq "NATO") {
        if ($lowerChar -eq $smallA_ring -or $lowerChar -eq $capitalA_ring) {
            $phonetic = "Alpha with Ring"
        }
        elseif ($lowerChar -eq $smallA_umlaut -or $lowerChar -eq $capitalA_umlaut) {
            $phonetic = "Alpha with Umlaut"
        }
        elseif ($lowerChar -eq $smallO_umlaut -or $lowerChar -eq $capitalO_umlaut) {
            $phonetic = "Oscar with Umlaut"
        }
    }
    
    # If not a special Swedish character, check for other special characters
    if ($null -eq $phonetic) {
        # Handle special characters directly
        $charCode = [int]$lowerChar
        
        # Check for specific Unicode code points and common punctuation/symbols
        if ($charCode -eq 0x00B4) { # ´ (Acute Accent)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Acute Accent" } else { "Akut Accent" }
        }
        elseif ($charCode -eq 0x0060) { # ` (Grave Accent)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Grave Accent" } else { "Grav Accent" }
        }
        elseif ($charCode -eq 0x2018) { # ' (Left Single Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Left Single Quote" } else { "V" + $smallA_umlaut + "nster Enkelt Citattecken" }
        }
        elseif ($charCode -eq 0x2019) { # ' (Right Single Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Right Single Quote" } else { "H" + $smallO_umlaut + "ger Enkelt Citattecken" }
        }
        elseif ($charCode -eq 0x201C) { # " (Left Double Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Left Double Quote" } else { "V" + $smallA_umlaut + "nster Dubbelt Citattecken" }
        }
        elseif ($charCode -eq 0x201D) { # " (Right Double Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Right Double Quote" } else { "H" + $smallO_umlaut + "ger Dubbelt Citattecken" }
        }
        elseif ($charCode -eq 0x0028) { # ( (Left Parenthesis)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Left Parenthesis" } else { "V" + $smallA_umlaut + "nster Parentes" }
        }
        elseif ($charCode -eq 0x0029) { # ) (Right Parenthesis)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Right Parenthesis" } else { "H" + $smallO_umlaut + "ger Parentes" }
        }
        elseif ($charCode -eq 0x0027) { # ' (Single Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Single Quote" } else { "Enkelt Citattecken" }
        }
        elseif ($charCode -eq 0x002A) { # * (Asterisk)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Asterisk" } else { "Asterisk" }
        }
        elseif ($charCode -eq 0x003B) { # ; (Semicolon)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Semicolon" } else { "Semikolon" }
        }
        elseif ($charCode -eq 0x002C) { # , (Comma)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Comma" } else { "Komma" }
        }
        elseif ($charCode -eq 0x002E) { # . (Period)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Period" } else { "Punkt" }
        }
        elseif ($charCode -eq 0x002F) { # / (Forward Slash)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Forward Slash" } else { "Snedstreck" }
        }
        elseif ($charCode -eq 0x003A) { # : (Colon)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Colon" } else { "Kolon" }
        }
        elseif ($charCode -eq 0x0022) { # " (Double Quote)
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Double Quote" } else { "Citattecken" }
        }
        elseif ($charCode -eq 0x0020) { # Space
            $phonetic = if ($PhoneticLanguage -eq "NATO") { "Space" } else { "Mellanslag" }
        }
        # If not a special character, look it up in the dictionary
        else {
            $key = $lowerChar.ToString()
            if ($null -eq $phoneticAlphabet -or $null -eq $key) {
                $phonetic = "Symbol"
            }
            elseif ($phoneticAlphabet.ContainsKey($key)) {
                $phonetic = $phoneticAlphabet[$key]
            } else {
                # Just use "Symbol" without the debugging code
                $phonetic = "Symbol"
            }
        }
    }
    
    # Add Capital/Stor prefix for uppercase letters
    if ($isUppercase -and [char]::IsLetter($Character)) {
        $prefix = if ($PhoneticLanguage -eq "NATO") { "Capital " } else { "Stor " }
        return "$prefix$phonetic"
    } else {
        return $phonetic
    }
}

# Displays phonetic codes for password in a popup window
function Show-PhoneticCodes {
    param (
        [System.Security.SecureString]$Password,
        [string]$PhoneticLanguage = "NATO"
    )
    
    # Convert SecureString to plain text for display
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    # Create a new window for the phonetic codes
    $phoneticWindow = New-Object System.Windows.Window
    $phoneticWindow.Title = if ($PhoneticLanguage -eq "NATO") { "NATO Phonetic Pronunciation" } else { "Swedish Phonetic Pronunciation" }
    $phoneticWindow.SizeToContent = "WidthAndHeight"
    $phoneticWindow.WindowStartupLocation = "CenterScreen"
    $phoneticWindow.ResizeMode = "NoResize"
    $phoneticWindow.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Create a stack panel for the content
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = New-Object System.Windows.Thickness(15)
    
    # Add a title
    $titleBlock = New-Object System.Windows.Controls.TextBlock
    $titleBlock.Text = "Phonetic Pronunciation for Password:"
    $titleBlock.FontWeight = "Bold"
    $titleBlock.FontSize = 14
    $titleBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
    $stackPanel.Children.Add($titleBlock)
    
    # Add the password
    $passwordBlock = New-Object System.Windows.Controls.TextBlock
    $passwordBlock.Text = $PlainPassword
    $passwordBlock.FontFamily = "Consolas"
    $passwordBlock.FontSize = 16
    $passwordBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $stackPanel.Children.Add($passwordBlock)
    
    # Add a separator
    $separator = New-Object System.Windows.Controls.Separator
    $separator.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
    $stackPanel.Children.Add($separator)
    
    # Define the threshold for splitting into two columns
    $splitThreshold = 10
    
    # Check if we need to split the password into two columns
    if ($PlainPassword.Length -gt $splitThreshold) {
        # Create a grid for two-column layout
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
        
        # Define columns for the grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $grid.ColumnDefinitions.Add($col1)
        $grid.ColumnDefinitions.Add($col2)
        
        # Create stack panels for each group
        $leftPanel = New-Object System.Windows.Controls.StackPanel
        $leftPanel.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
        [System.Windows.Controls.Grid]::SetColumn($leftPanel, 0)
        
        $rightPanel = New-Object System.Windows.Controls.StackPanel
        $rightPanel.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
        [System.Windows.Controls.Grid]::SetColumn($rightPanel, 1)
        
        # Add group headers
        $leftHeader = New-Object System.Windows.Controls.TextBlock
        $leftHeader.Text = "Group 1:"
        $leftHeader.FontWeight = "Bold"
        $leftHeader.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
        $leftPanel.Children.Add($leftHeader)
        
        $rightHeader = New-Object System.Windows.Controls.TextBlock
        $rightHeader.Text = "Group 2:"
        $rightHeader.FontWeight = "Bold"
        $rightHeader.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
        $rightPanel.Children.Add($rightHeader)
        
        # Calculate midpoint for splitting (ensures even split with odd numbers)
        $midpoint = [Math]::Ceiling($PlainPassword.Length / 2)
        
        # Add characters to left panel (Group 1)
        for ($i = 0; $i -lt $midpoint; $i++) {
            $char = $PlainPassword[$i]
            $phoneticCode = Get-PhoneticCode -Character $char -PhoneticLanguage $PhoneticLanguage
            
            # Create a grid for this character
            $charGrid = New-Object System.Windows.Controls.Grid
            $charGrid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
            
            # Define columns
            $charCol1 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol1.Width = New-Object System.Windows.GridLength(30)
            $charCol2 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol2.Width = New-Object System.Windows.GridLength(20)
            $charCol3 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol3.Width = New-Object System.Windows.GridLength(1, "Star")
            $charGrid.ColumnDefinitions.Add($charCol1)
            $charGrid.ColumnDefinitions.Add($charCol2)
            $charGrid.ColumnDefinitions.Add($charCol3)
            
            # Character
            $charBlock = New-Object System.Windows.Controls.TextBlock
            $charBlock.Text = $char
            $charBlock.FontFamily = "Consolas"
            $charBlock.FontWeight = "Bold"
            $charBlock.FontSize = 14
            $charBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($charBlock, 0)
            $charGrid.Children.Add($charBlock)
            
            # Arrow
            $arrowBlock = New-Object System.Windows.Controls.TextBlock
            $arrowBlock.Text = "=>"
            $arrowBlock.FontSize = 14
            $arrowBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($arrowBlock, 1)
            $charGrid.Children.Add($arrowBlock)
            
            # Phonetic code
            $codeBlock = New-Object System.Windows.Controls.TextBlock
            $codeBlock.Text = $phoneticCode
            $codeBlock.FontSize = 14
            [System.Windows.Controls.Grid]::SetColumn($codeBlock, 2)
            $charGrid.Children.Add($codeBlock)
            
            $leftPanel.Children.Add($charGrid)
        }
        
        # Add characters to right panel (Group 2)
        for ($i = $midpoint; $i -lt $PlainPassword.Length; $i++) {
            $char = $PlainPassword[$i]
            $phoneticCode = Get-PhoneticCode -Character $char -PhoneticLanguage $PhoneticLanguage
            
            # Create a grid for this character
            $charGrid = New-Object System.Windows.Controls.Grid
            $charGrid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
            
            # Define columns
            $charCol1 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol1.Width = New-Object System.Windows.GridLength(30)
            $charCol2 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol2.Width = New-Object System.Windows.GridLength(20)
            $charCol3 = New-Object System.Windows.Controls.ColumnDefinition
            $charCol3.Width = New-Object System.Windows.GridLength(1, "Star")
            $charGrid.ColumnDefinitions.Add($charCol1)
            $charGrid.ColumnDefinitions.Add($charCol2)
            $charGrid.ColumnDefinitions.Add($charCol3)
            
            # Character
            $charBlock = New-Object System.Windows.Controls.TextBlock
            $charBlock.Text = $char
            $charBlock.FontFamily = "Consolas"
            $charBlock.FontWeight = "Bold"
            $charBlock.FontSize = 14
            $charBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($charBlock, 0)
            $charGrid.Children.Add($charBlock)
            
            # Arrow
            $arrowBlock = New-Object System.Windows.Controls.TextBlock
            $arrowBlock.Text = "=>"
            $arrowBlock.FontSize = 14
            $arrowBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($arrowBlock, 1)
            $charGrid.Children.Add($arrowBlock)
            
            # Phonetic code
            $codeBlock = New-Object System.Windows.Controls.TextBlock
            $codeBlock.Text = $phoneticCode
            $codeBlock.FontSize = 14
            [System.Windows.Controls.Grid]::SetColumn($codeBlock, 2)
            $charGrid.Children.Add($codeBlock)
            
            $rightPanel.Children.Add($charGrid)
        }
        
        # Add panels to the grid
        $grid.Children.Add($leftPanel)
        $grid.Children.Add($rightPanel)
        
        # Add the grid to the main stack panel
        $stackPanel.Children.Add($grid)
    }
    else {
        # For shorter passwords, display in a single column
        for ($i = 0; $i -lt $PlainPassword.Length; $i++) {
            $char = $PlainPassword[$i]
            $phoneticCode = Get-PhoneticCode -Character $char -PhoneticLanguage $PhoneticLanguage
            
            # Create a grid for this character
            $grid = New-Object System.Windows.Controls.Grid
            $grid.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
            
            # Define columns
            $col1 = New-Object System.Windows.Controls.ColumnDefinition
            $col1.Width = New-Object System.Windows.GridLength(30)
            $col2 = New-Object System.Windows.Controls.ColumnDefinition
            $col2.Width = New-Object System.Windows.GridLength(20)
            $col3 = New-Object System.Windows.Controls.ColumnDefinition
            $col3.Width = New-Object System.Windows.GridLength(1, "Star")
            $grid.ColumnDefinitions.Add($col1)
            $grid.ColumnDefinitions.Add($col2)
            $grid.ColumnDefinitions.Add($col3)
            
            # Character
            $charBlock = New-Object System.Windows.Controls.TextBlock
            $charBlock.Text = $char
            $charBlock.FontFamily = "Consolas"
            $charBlock.FontWeight = "Bold"
            $charBlock.FontSize = 14
            $charBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($charBlock, 0)
            $grid.Children.Add($charBlock)
            
            # Arrow
            $arrowBlock = New-Object System.Windows.Controls.TextBlock
            $arrowBlock.Text = "=>"
            $arrowBlock.FontSize = 14
            $arrowBlock.HorizontalAlignment = "Center"
            [System.Windows.Controls.Grid]::SetColumn($arrowBlock, 1)
            $grid.Children.Add($arrowBlock)
            
            # Phonetic code
            $codeBlock = New-Object System.Windows.Controls.TextBlock
            $codeBlock.Text = $phoneticCode
            $codeBlock.FontSize = 14
            [System.Windows.Controls.Grid]::SetColumn($codeBlock, 2)
            $grid.Children.Add($codeBlock)
            
            $stackPanel.Children.Add($grid)
        }
    }
    
    # Add a close button
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "Close"
    $closeButton.Padding = New-Object System.Windows.Thickness(20, 5, 20, 5)
    $closeButton.Margin = New-Object System.Windows.Thickness(0, 15, 0, 0)
    $closeButton.HorizontalAlignment = "Center"
    $closeButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $closeButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $closeButton.Add_Click({ $phoneticWindow.Close() })
    $stackPanel.Children.Add($closeButton)
    
    # Set the content and show the window
    $phoneticWindow.Content = $stackPanel
    $phoneticWindow.ShowDialog() | Out-Null
}

# Adds NATO and Swedish phonetic buttons to the UI
function Add-PhoneticButtons {
    param (
        [System.Windows.Controls.Grid]$PassGrid,
        [System.Windows.Controls.TextBox]$PassTextBox
    )
    
    # Add two new column definitions for the NATO and Swedish buttons
    $col4 = New-Object System.Windows.Controls.ColumnDefinition
    $col4.Width = New-Object System.Windows.GridLength("Auto")
    $col5 = New-Object System.Windows.Controls.ColumnDefinition
    $col5.Width = New-Object System.Windows.GridLength("Auto")
    $PassGrid.ColumnDefinitions.Add($col4)
    $PassGrid.ColumnDefinitions.Add($col5)
    
    # Create NATO phonetic button
    $natoButton = New-Object System.Windows.Controls.Button
    $natoButton.Content = "NATO Phonetic"
    $natoButton.Width = 110
    $natoButton.Height = 30
    $natoButton.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
    $natoButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $natoButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    [System.Windows.Controls.Grid]::SetColumn($natoButton, 3)
    $PassGrid.Children.Add($natoButton)
    
    # Create Swedish phonetic button
    $swedishButton = New-Object System.Windows.Controls.Button
    $swedishButton.Content = "Swedish Phonetic"
    $swedishButton.Width = 120
    $swedishButton.Height = 30
    $swedishButton.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
    $swedishButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 153, 0))
    $swedishButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    [System.Windows.Controls.Grid]::SetColumn($swedishButton, 4)
    $PassGrid.Children.Add($swedishButton)
    
    # Add click event handlers
    $natoButton.Add_Click({
        if ($PassTextBox.Text) {
            # Convert the password to a SecureString
            $securePassword = New-Object System.Security.SecureString
            foreach ($char in $PassTextBox.Text.ToCharArray()) {
                $securePassword.AppendChar($char)
            }
            Show-PhoneticCodes -Password $securePassword -PhoneticLanguage "NATO"
        }
    })
    
    $swedishButton.Add_Click({
        if ($PassTextBox.Text) {
            # Convert the password to a SecureString
            $securePassword = New-Object System.Security.SecureString
            foreach ($char in $PassTextBox.Text.ToCharArray()) {
                $securePassword.AppendChar($char)
            }
            Show-PhoneticCodes -Password $securePassword -PhoneticLanguage "Swedish"
        }
    })
}

# Sends password to Password Pusher API with configurable options
function Push-Password {
    param (
        [System.Security.SecureString]$Password,
        [int]$ExpireDays,
        [int]$ExpireViews,
        [bool]$DeletableByViewer,
        [bool]$RetrievalStep,
        [System.Security.SecureString]$Passphrase = [System.Security.SecureString]::new(),
        [bool]$UseQRCode = $false,
        [bool]$UseCurl = $false
    )
    
    # Convert SecureString to plain text for API
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    # Convert Passphrase if provided
    $PlainPassphrase = ""
    if ($Passphrase.Length -gt 0) {
        $PassphraseBSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Passphrase)
        $PlainPassphrase = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($PassphraseBSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($PassphraseBSTR)
    }
    
    $apiUrl = "https://pwpush.com/p.json"
    $logOutput = "Attempting to push password to Password Pusher..."
    
    # Check if the Password Pusher service is available
    if (-not (Test-ServiceAvailability -ServiceUrl "https://pwpush.com")) {
        return @{
            Success = $false
            Url = ""
            IsQRCode = $false
            Log = "Password Pusher API: Service unavailable"
        }
    }
    
    try {
        if ($UseCurl) {
            $logOutput += "`nUsing curl method..."
            
            $curlParams = "password[payload]=$PlainPassword&password[expire_after_days]=$ExpireDays&password[expire_after_views]=$ExpireViews&password[deletable_by_viewer]=$($DeletableByViewer.ToString().ToLower())&password[retrieval_step]=$($RetrievalStep.ToString().ToLower())"
            
            if ($PlainPassphrase) {
                $curlParams += "&password[passphrase]=$PlainPassphrase"
            }
            
            # Add kind=qr parameter if QR code is requested
            if ($UseQRCode) {
                $curlParams += "&password[kind]=qr"
                $logOutput += "`nGenerating QR code..."
            }
            
            $result = curl.exe -X POST $apiUrl -H "Content-Type: application/x-www-form-urlencoded" -d $curlParams
            
            # Parse the JSON response
            $jsonStart = $result.IndexOf("{")
            $jsonEnd = $result.LastIndexOf("}")
            
            if ($jsonStart -ge 0 -and $jsonEnd -gt $jsonStart) {
                $jsonResponse = $result.Substring($jsonStart, $jsonEnd - $jsonStart + 1)
                $response = $jsonResponse | ConvertFrom-Json
                
            if ($response.url_token) {
                $url = "https://pwpush.com/p/" + $response.url_token
                $logOutput += "`nPassword successfully pushed using direct API method."
                
                if ($UseQRCode) {
                    $logOutput += "`nThis URL contains a QR code that can be scanned to access the password."
                }
                
                # Add the password link to the history collection
                # Convert plain password to SecureString for the PasswordLinkModel
                $securePasswordForHistory = New-Object System.Security.SecureString
                foreach ($char in $PlainPassword.ToCharArray()) {
                    $securePasswordForHistory.AppendChar($char)
                }
                $passwordLink = [PasswordLinkModel]::new($url, $securePasswordForHistory)
                [void]$script:PasswordLinks.Add($passwordLink)
                $logOutput += "`nAdded to password links history."
                
                return @{
                    Success = $true
                    Url = $url
                    IsQRCode = $UseQRCode
                    Log = $logOutput
                }
                }
                else {
                    $logOutput += "`nError: Failed to extract URL token from response."
                    return @{
                        Success = $false
                        Url = ""
                        IsQRCode = $false
                        Log = $logOutput
                    }
                }
            }
            else {
                $logOutput += "`nError: Invalid JSON response from curl: $result"
                return @{
                    Success = $false
                    Url = ""
                    IsQRCode = $false
                    Log = $logOutput
                }
            }
        }
        else {
            $logOutput += "`nUsing Invoke-RestMethod..."
            
            $body = @{
                'password[payload]' = $PlainPassword
                'password[expire_after_days]' = $ExpireDays
                'password[expire_after_views]' = $ExpireViews
                'password[deletable_by_viewer]' = $DeletableByViewer.ToString().ToLower()
                'password[retrieval_step]' = $RetrievalStep.ToString().ToLower()
            }
            
            if ($PlainPassphrase) {
                $body['password[passphrase]'] = $PlainPassphrase
            }
            
            # Add kind=qr parameter if QR code is requested
            if ($UseQRCode) {
                $body['password[kind]'] = 'qr'
                $logOutput += "`nGenerating QR code..."
            }
            
            $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $body
            
            if ($response.url_token) {
                $url = "https://pwpush.com/p/" + $response.url_token
                $logOutput += "`nPassword successfully pushed using curl method."
                
                if ($UseQRCode) {
                    $logOutput += "`nThis URL contains a QR code that can be scanned to access the password."
                }
                
                # Add the password link to the history collection
                # Convert plain password to SecureString for the PasswordLinkModel
                $securePasswordForHistory = New-Object System.Security.SecureString
                foreach ($char in $PlainPassword.ToCharArray()) {
                    $securePasswordForHistory.AppendChar($char)
                }
                $passwordLink = [PasswordLinkModel]::new($url, $securePasswordForHistory)
                [void]$script:PasswordLinks.Add($passwordLink)
                $logOutput += "`nAdded to password links history."
                
                return @{
                    Success = $true
                    Url = $url
                    IsQRCode = $UseQRCode
                    Log = $logOutput
                }
            }
            else {
                $logOutput += "`nError: Failed to extract URL token from response."
                return @{
                    Success = $false
                    Url = ""
                    IsQRCode = $false
                    Log = $logOutput
                }
            }
        }
    }
    catch {
        $logOutput += "`nError: $_"
        
        # If direct API call failed, try curl as fallback
        if (-not $UseCurl) {
            $logOutput += "`nAttempting fallback to curl method..."
            return Push-Password -Password $Password -ExpireDays $ExpireDays -ExpireViews $ExpireViews -DeletableByViewer $DeletableByViewer -RetrievalStep $RetrievalStep -Passphrase $Passphrase -UseQRCode $UseQRCode -UseCurl $true
        }
        
        return @{
            Success = $false
            Url = ""
            IsQRCode = $false
            Log = $logOutput
        }
    }
}

# WPF GUI definition in XAML
[xml]$xaml = @'
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SecurePassGenerator" Height="650" Width="600"
    WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize">
    <Window.Resources>
        <BooleanToVisibilityConverter x:Key="BooleanToVisibilityConverter"/>
    </Window.Resources>
    <Grid Margin="12">
        <TabControl>
            <!-- Main Tab -->
            <TabItem Header="Password Generator">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Password Generation Section -->
                    <GroupBox Grid.Row="0" Header="Password Generation" Margin="0,8,0,8" Padding="8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Password Type Selection with Language Dropdown and Presets -->
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Password Type:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <RadioButton Grid.Column="1" x:Name="RandomPasswordType" Content="Random" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
                    <RadioButton Grid.Column="2" x:Name="MemorablePasswordType" Content="Memorable" Margin="0,0,10,0" VerticalAlignment="Center"/>
                    <TextBlock Grid.Column="3" Text="Language:" VerticalAlignment="Center" Margin="0,0,10,0" Visibility="{Binding ElementName=MemorablePasswordType, Path=IsChecked, Converter={StaticResource BooleanToVisibilityConverter}}"/>
                    <ComboBox Grid.Column="4" x:Name="LanguageSelector" Width="100" HorizontalAlignment="Left" SelectedIndex="1" Visibility="{Binding ElementName=MemorablePasswordType, Path=IsChecked, Converter={StaticResource BooleanToVisibilityConverter}}">
                        <ComboBoxItem Content="English"/>
                        <ComboBoxItem Content="Swedish"/>
                    </ComboBox>
                    <TextBlock Grid.Column="5" Text="Presets:" VerticalAlignment="Center" Margin="10,0,10,0" Visibility="{Binding ElementName=RandomPasswordType, Path=IsChecked, Converter={StaticResource BooleanToVisibilityConverter}}"/>
                    <Grid Grid.Column="6" Visibility="{Binding ElementName=RandomPasswordType, Path=IsChecked, Converter={StaticResource BooleanToVisibilityConverter}}">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox Grid.Column="0" x:Name="PasswordPresets" Width="150" HorizontalAlignment="Left" SelectedIndex="1">
                            <ComboBoxItem Content="Medium Password"/>
                            <ComboBoxItem Content="Strong Password"/>
                            <ComboBoxItem Content="Very Strong Password"/>
                        </ComboBox>
                        <Button Grid.Column="1" x:Name="PresetsSettingsButton" Content="(...)" Width="40" Height="24" Margin="5,0,0,0" 
                                ToolTip="Manage Password Presets" Background="#0078D7" Foreground="White" FontWeight="Normal" FontSize="12"/>
                    </Grid>
                </Grid>
                
                <!-- Password Length/Word Count -->
                <Grid Grid.Row="1" Margin="0,0,0,10">
                    <!-- Random Password Settings -->
                    <StackPanel x:Name="RandomPasswordSettings" Orientation="Horizontal">
                        <TextBlock Text="Length:" VerticalAlignment="Center" Margin="0,0,10,0" Width="60"/>
                        <Slider x:Name="PasswordLength" Minimum="8" Maximum="32" Value="12" Width="200" VerticalAlignment="Center" TickFrequency="1" IsSnapToTickEnabled="True"/>
                        <TextBlock x:Name="PasswordLengthValue" Text="12" VerticalAlignment="Center" Margin="5,0,0,0" Width="30"/>
                    </StackPanel>
                    
                    <!-- Memorable Password Settings -->
                    <StackPanel x:Name="MemorablePasswordSettings" Orientation="Horizontal" Visibility="Collapsed">
                        <TextBlock Text="Words:" VerticalAlignment="Center" Margin="0,0,10,0" Width="60"/>
                        <Slider x:Name="WordCount" Minimum="1" Maximum="5" Value="3" Width="200" VerticalAlignment="Center" TickFrequency="1" IsSnapToTickEnabled="True"/>
                        <TextBlock x:Name="WordCountValue" Text="3" VerticalAlignment="Center" Margin="5,0,0,0" Width="30"/>
                    </StackPanel>
                </Grid>
                
                <!-- Character Options -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
                    <CheckBox x:Name="IncludeUppercase" Content="Uppercase" IsChecked="True" Margin="0,0,10,0"/>
                    <CheckBox x:Name="IncludeNumbers" Content="Numbers" IsChecked="True" Margin="0,0,10,0"/>
                    <CheckBox x:Name="IncludeSpecial" Content="Special Characters" IsChecked="True"/>
                </StackPanel>
                
                <!-- Generated Password with Generate and Copy Buttons to the right -->
                <Grid Grid.Row="3" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="GeneratedPassword" Grid.Column="0" Height="30" FontSize="14" IsReadOnly="False" VerticalContentAlignment="Center" Padding="5,0"/>
                    <Button x:Name="GenerateButton" Grid.Column="1" Content="Generate" Width="80" Height="30" Margin="10,0,0,0"/>
                    <Button x:Name="CopyPasswordButton" Grid.Column="2" Content="Copy" Width="50" Height="30" Margin="10,0,0,0"/>
                    <Grid Grid.Column="3" Margin="10,0,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="PhoneticButton" Grid.Column="0" Content="(P)" Width="30" Height="30" Background="#0078D7" Foreground="White"/>
                        <Button x:Name="PhoneticDropdownButton" Grid.Column="1" Content="v" Width="15" Height="30" Background="#0078D7" Foreground="White" Padding="0"/>
                    </Grid>
                </Grid>
                
                <!-- Password Strength with Entropy Display -->
                <Grid Grid.Row="4" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Vertical">
                        <TextBlock Text="Password Strength:" Margin="0,0,0,5"/>
                        <Grid>
                            <ProgressBar x:Name="StrengthBar" Height="20" Maximum="100" Value="0"/>
                            <TextBlock x:Name="StrengthText" Text="Weak" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Grid>
                    </StackPanel>
                    <TextBlock x:Name="EntropyValue" Grid.Column="1" Text="0 bits" VerticalAlignment="Bottom" Margin="10,0,0,0" FontWeight="SemiBold"/>
                </Grid>
                
                <!-- Have I Been Pwned Section -->
                <Grid Grid.Row="5" Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Vertical">
                        <TextBlock Text="Have I Been Pwned:" Margin="0,0,0,5"/>
                        <TextBlock x:Name="PwnedStatus" Text="Not checked" Foreground="Gray" FontWeight="SemiBold"/>
                    </StackPanel>
                    <Button x:Name="CheckPwnedButton" Grid.Column="1" Content="Check" Width="100" Height="30" Margin="10,0,0,0"/>
                </Grid>
            </Grid>
        </GroupBox>
        
                    <!-- Password Pusher Section -->
                    <GroupBox Grid.Row="1" Header="Password Pusher" Margin="0,0,0,8" Padding="8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- API Method Selection -->
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="API Method:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <RadioButton x:Name="DirectApiMethod" Content="Direct API Call" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
                    <RadioButton x:Name="CurlApiMethod" Content="Curl" VerticalAlignment="Center"/>
                </StackPanel>
                
                <!-- Expiration Settings (Days and Views side by side) -->
                <Grid Grid.Row="1" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Expiration Days -->
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="Expire Days:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <ComboBox x:Name="ExpireDays" Width="80" SelectedIndex="1">
                            <ComboBoxItem Content="1 Day" Tag="1"/>
                            <ComboBoxItem Content="3 Days" Tag="3"/>
                            <ComboBoxItem Content="7 Days" Tag="7"/>
                            <ComboBoxItem Content="14 Days" Tag="14"/>
                            <ComboBoxItem Content="30 Days" Tag="30"/>
                        </ComboBox>
                    </StackPanel>
                    
                    <!-- Expiration Views -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <TextBlock Text="Expire Views:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <ComboBox x:Name="ExpireViews" Width="80" SelectedIndex="0">
                            <ComboBoxItem Content="1 View" Tag="1"/>
                            <ComboBoxItem Content="2 Views" Tag="2"/>
                            <ComboBoxItem Content="5 Views" Tag="5"/>
                            <ComboBoxItem Content="10 Views" Tag="10"/>
                            <ComboBoxItem Content="20 Views" Tag="20"/>
                        </ComboBox>
                    </StackPanel>
                </Grid>
                
                <!-- Additional Options -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
                    <CheckBox x:Name="DeletableByViewer" Content="Deletable by Viewer" IsChecked="True" Margin="0,0,20,0"/>
                    <CheckBox x:Name="RetrievalStep" Content="Add Retrieval Step" IsChecked="True" Margin="0,0,20,0"/>
                    <CheckBox x:Name="UseQRCode" Content="Generate QR Code" IsChecked="False"/>
                </StackPanel>
                
                <!-- Passphrase Protection -->
                <Grid Grid.Row="3" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox x:Name="UsePassphrase" Content="Passphrase Protection:" Grid.Column="0" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <PasswordBox x:Name="Passphrase" Grid.Column="1" Height="30" IsEnabled="{Binding ElementName=UsePassphrase, Path=IsChecked}"/>
                </Grid>
                
                <!-- Push Button -->
                <Button x:Name="PushButton" Grid.Row="4" Content="Push to Password Pusher" Height="35" FontSize="14"/>
            </Grid>
        </GroupBox>
        
                    <!-- Result Section -->
                    <GroupBox Grid.Row="2" Header="Result" Margin="0,0,0,8" Padding="8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Result URL and Buttons -->
                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="75*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="ResultUrl" Grid.Column="0" Height="30" FontSize="14" IsReadOnly="True" VerticalContentAlignment="Center" Padding="5,0"/>
                    <Button x:Name="CopyUrlButton" Grid.Column="1" Content="Copy" Width="60" Height="30" Margin="10,0,0,0"/>
                    <Button x:Name="OpenInBrowserButton" Grid.Column="2" Content="Browse" Width="60" Height="30" Margin="10,0,0,0"/>
                    <Button x:Name="HistoryButton" Grid.Column="3" Content="History" Width="60" Height="30" Margin="10,0,0,0"/>
                </Grid>
            </Grid>
        </GroupBox>
        
                    <!-- Empty row for spacing -->
                    <Border Grid.Row="3" Height="0"/>
                </Grid>
            </TabItem>
            
            <!-- Log Tab -->
            <TabItem Header="Log">
                <Grid Margin="0,8,0,0">
                    <TextBox x:Name="LogOutput" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" 
                             IsReadOnly="True" FontFamily="Consolas" Margin="0" 
                             Height="Auto" Width="Auto"/>
                </Grid>
            </TabItem>
            
            <!-- About Tab -->
            <TabItem Header="About" Background="#E6F2FF">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock x:Name="AboutTitle" Grid.Row="0" Text="About SecurePassGenerator" FontSize="18" FontWeight="Bold" Margin="0,0,0,15" Foreground="#0078D7"/>
                    
                    <Border Grid.Row="1" Background="#F8F8F8" BorderBrush="#CCCCCC" BorderThickness="1" Padding="15" Margin="0,0,0,2">
                        <StackPanel>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                                This application is a secure Password Generator with Password Pusher integration. It provides a modern interface for creating strong, customizable passwords and securely sharing them.
                            </TextBlock>
                            
                            <TextBlock Text="Features include:" TextWrapping="Wrap" Margin="0,0,0,5"/>
                            <StackPanel Margin="10,0,0,15">
                                <TextBlock Text="- Random or memorable password generation" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Customizable password length and complexity" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Predefined and custom password presets for different security levels" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Password strength assessment with entropy calculation" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- NATO and Swedish phonetic pronunciation support" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Have I Been Pwned database integration" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Secure password sharing via Password Pusher" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- QR code generation for easy mobile access" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Password links history tracking and management" TextWrapping="Wrap" Margin="0,2,0,2"/>
                                <TextBlock Text="- Manual password expiration capability" TextWrapping="Wrap" Margin="0,2,0,2"/>
                            </StackPanel>
                            
                            <!-- No GitHub button here anymore -->
                        </StackPanel>
                    </Border>
                    
                    <!-- Third-Party Credits Section - Compact Version -->
                    <Border Grid.Row="2" Background="#F8F8F8" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" Margin="0,0,0,2" Height="120">
                        <StackPanel>
                            <TextBlock Text="Credits" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,6" FontSize="11">
                                This application uses the following third-party services:
                            </TextBlock>
                            
                            <Grid Margin="0,0,0,4">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="Password Pusher" FontWeight="SemiBold" FontSize="11"/>
                                    <TextBlock Text="Secure password sharing service" TextWrapping="Wrap" FontSize="11"/>
                                </StackPanel>
                                <Button x:Name="PasswordPusherButton" Grid.Column="1" Content="Visit" Width="50" Height="22"
                                        Background="#0078D7" Foreground="White" Padding="2,1" FontSize="10"/>
                            </Grid>
                            
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="Have I Been Pwned" FontWeight="SemiBold" FontSize="11"/>
                                    <TextBlock Text="Data breach search service" TextWrapping="Wrap" FontSize="11"/>
                                </StackPanel>
                                <Button x:Name="HibpButton" Grid.Column="1" Content="Visit" Width="50" Height="22"
                                        Background="#0078D7" Foreground="White" Padding="2,1" FontSize="10"/>
                            </Grid>
                        </StackPanel>
                    </Border>
                    
                    <!-- Author and Version Information -->
                    <Border Grid.Row="3" Background="#F0F8FF" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" Margin="0,0,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0">
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,5">
                                    <TextBlock Text="Author: " FontWeight="SemiBold"/>
                                    <TextBlock Text="onlyalex1984"/>
                                </StackPanel>
                                
                            <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                                <TextBlock Text="Version: " FontWeight="SemiBold"/>
                                <TextBlock x:Name="VersionText" Text="1.0"/>
                            </StackPanel>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="CheckForUpdatesButton" Content="Check for Updates" Width="90" Height="22"
                                        Background="#0078D7" Foreground="White" Padding="2,1" FontSize="10" Margin="0,0,10,0"/>
                                <Button x:Name="GitHubButton" Content="Project on GitHub" Width="90" Height="22"
                                        Background="#0078D7" Foreground="White" Padding="2,1" FontSize="10"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@

# Initialize main application window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Set window title and About tab title from script variables
$window.Title = $script:ScriptDisplayName
$aboutTitle = $window.FindName("AboutTitle")
if ($aboutTitle) {
    $aboutTitle.Text = "About $script:ScriptDisplayName"
}

# Set version text from script variable
$versionText = $window.FindName("VersionText")
if ($versionText) {
    $versionText.Text = $script:Version
}

# Reference UI controls for event handling
$global:controls = @{
    RandomPasswordType = $window.FindName("RandomPasswordType")
    MemorablePasswordType = $window.FindName("MemorablePasswordType")
    LanguageSelector = $window.FindName("LanguageSelector")
    PasswordPresets = $window.FindName("PasswordPresets")
    PresetsSettingsButton = $window.FindName("PresetsSettingsButton")
    PasswordLength = $window.FindName("PasswordLength")
    PasswordLengthValue = $window.FindName("PasswordLengthValue")
    WordCount = $window.FindName("WordCount")
    WordCountValue = $window.FindName("WordCountValue")
    IncludeUppercase = $window.FindName("IncludeUppercase")
    IncludeNumbers = $window.FindName("IncludeNumbers")
    IncludeSpecial = $window.FindName("IncludeSpecial")
    GeneratedPassword = $window.FindName("GeneratedPassword")
    CopyPasswordButton = $window.FindName("CopyPasswordButton")
    PhoneticButton = $window.FindName("PhoneticButton")
    PhoneticDropdownButton = $window.FindName("PhoneticDropdownButton")
    StrengthBar = $window.FindName("StrengthBar")
    StrengthText = $window.FindName("StrengthText")
    EntropyValue = $window.FindName("EntropyValue")
    GenerateButton = $window.FindName("GenerateButton")
    PwnedStatus = $window.FindName("PwnedStatus")
    CheckPwnedButton = $window.FindName("CheckPwnedButton")
    DirectApiMethod = $window.FindName("DirectApiMethod")
    CurlApiMethod = $window.FindName("CurlApiMethod")
    ExpireDays = $window.FindName("ExpireDays")
    ExpireViews = $window.FindName("ExpireViews")
    DeletableByViewer = $window.FindName("DeletableByViewer")
    RetrievalStep = $window.FindName("RetrievalStep")
    UseQRCode = $window.FindName("UseQRCode")
    UsePassphrase = $window.FindName("UsePassphrase")
    Passphrase = $window.FindName("Passphrase")
    PushButton = $window.FindName("PushButton")
    ResultUrl = $window.FindName("ResultUrl")
    CopyUrlButton = $window.FindName("CopyUrlButton")
    OpenInBrowserButton = $window.FindName("OpenInBrowserButton")
    LogOutput = $window.FindName("LogOutput")
    CheckForUpdatesButton = $window.FindName("CheckForUpdatesButton")
}

# Applies light theme styling to UI elements
function Set-LightTheme {
    # Set window background
    $window.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(245, 245, 245))
    
    # Set specific control colors
    $controls.GeneratedPassword.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $controls.GeneratedPassword.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51, 51, 51))
    $controls.GeneratedPassword.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(204, 204, 204))
    
    $controls.ResultUrl.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $controls.ResultUrl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51, 51, 51))
    $controls.ResultUrl.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(204, 204, 204))
    
    $controls.LogOutput.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    $controls.LogOutput.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51, 51, 51))
    $controls.LogOutput.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(204, 204, 204))
    
    # Set button colors
    $controls.GenerateButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $controls.GenerateButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    $controls.CopyPasswordButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $controls.CopyPasswordButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    # Explicitly set foreground to white for the Check button
    $controls.CheckPwnedButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $controls.PushButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    $controls.CopyUrlButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $controls.CopyUrlButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    $controls.OpenInBrowserButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
    $controls.OpenInBrowserButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
    
    # Set text colors for labels
    $controls.PasswordLengthValue.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51, 51, 51))
    $controls.StrengthText.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(51, 51, 51))
}

# Updates password strength indicator based on entropy
function Update-PasswordStrength {
    param (
        [System.Security.SecureString]$Password
    )
    
    $entropy = Get-PasswordEntropy -Password $Password
    
    # Set strength bar value (0-100)
    $strengthPercentage = [Math]::Min(100, ($entropy / 128) * 100)
    $controls.StrengthBar.Value = $strengthPercentage
    
    # Set strength text and color
    if ($entropy -lt 40) {
        $controls.StrengthText.Text = "Weak"
        $controls.StrengthBar.Foreground = "Red"
    }
    elseif ($entropy -lt 60) {
        $controls.StrengthText.Text = "Moderate"
        $controls.StrengthBar.Foreground = "Orange"
    }
    elseif ($entropy -lt 80) {
        $controls.StrengthText.Text = "Strong"
        $controls.StrengthBar.Foreground = "YellowGreen"
    }
    else {
        $controls.StrengthText.Text = "Very Strong"
        $controls.StrengthBar.Foreground = "Green"
    }
    
    # Update entropy value in the GUI
    $roundedEntropy = [Math]::Round($entropy, 2)
    $controls.EntropyValue.Text = "$roundedEntropy bits"
    
    # Also log to the log tab
    $controls.LogOutput.AppendText("Password entropy: $roundedEntropy bits`n")
}

# Evaluates custom password strength and updates UI
function Measure-CustomPassword {
    param (
        [System.Security.SecureString]$Password
    )
    
    # Convert SecureString to plain text for evaluation
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    
    # If password is empty, set zero entropy and disable buttons
    if ([string]::IsNullOrWhiteSpace($PlainPassword)) {
        $controls.StrengthBar.Value = 0
        $controls.StrengthText.Text = "Weak"
        $controls.StrengthBar.Foreground = "Red"
        $controls.EntropyValue.Text = "0 bits"
        
        # Disable Copy and Phonetic buttons
        $controls.CopyPasswordButton.IsEnabled = $false
        $controls.PhoneticButton.IsEnabled = $false
        $controls.PhoneticDropdownButton.IsEnabled = $false
        
        # Disable Check and Push buttons when password is empty
        $controls.CheckPwnedButton.IsEnabled = $false
        $controls.PushButton.IsEnabled = $false
        
        # Disable Copy and Browse buttons for password push result
        $controls.CopyUrlButton.IsEnabled = $false
        $controls.OpenInBrowserButton.IsEnabled = $false
        
        # Reset pwned status
        $controls.PwnedStatus.Text = "Not checked"
        $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(128, 128, 128))
        
        # Log that password field is empty
        $controls.LogOutput.AppendText("Password field is empty`n")
        
        return
    }
    
    # Log that a custom password is being evaluated
    $controls.LogOutput.AppendText("Evaluated custom password`n")
    
    # Enable buttons if password is not empty
    $controls.CopyPasswordButton.IsEnabled = $true
    $controls.PhoneticButton.IsEnabled = $true
    $controls.PhoneticDropdownButton.IsEnabled = $true
    
    # Reset the password flags when the password is modified
    $script:currentPasswordPushed = $false
    $script:currentPasswordChecked = $false
    
    # Reset the Password Pusher result
    $controls.ResultUrl.Text = ""
    $controls.ResultUrl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0))
    
    # Only enable the push button if the push timer is not running
    if (-not $script:pushTimer.IsEnabled) {
        $controls.PushButton.IsEnabled = $true
        $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
        $controls.PushButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        $controls.PushButton.Content = "Push to Password Pusher"
    }
    
    # Only enable the check button if the hibp timer is not running
    if (-not $script:hibpTimer.IsEnabled) {
        $controls.CheckPwnedButton.IsEnabled = $true
        $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
        $controls.CheckPwnedButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))
        $controls.CheckPwnedButton.Content = "Check"
    }
    
    # Calculate entropy and update UI using the SecureString directly
    Update-PasswordStrength -Password $Password
}

# Generates password based on current UI settings
function New-Password {
    $includeUppercase = $controls.IncludeUppercase.IsChecked
    $includeNumbers = $controls.IncludeNumbers.IsChecked
    $includeSpecial = $controls.IncludeSpecial.IsChecked
    
    if ($controls.RandomPasswordType.IsChecked) {
        $length = [int]$controls.PasswordLength.Value
        $password = New-RandomPassword -Length $length -IncludeUppercase $includeUppercase -IncludeNumbers $includeNumbers -IncludeSpecial $includeSpecial
        $controls.LogOutput.AppendText("Generated random password of length $length`n")
    }
    else {
        $wordCount = [int]$controls.WordCount.Value
        $language = $controls.LanguageSelector.SelectedItem.Content
        $password = New-MemorablePassword -WordCount $wordCount -Language $language -IncludeUppercase $includeUppercase -IncludeNumbers $includeNumbers -IncludeSpecial $includeSpecial
        $controls.LogOutput.AppendText("Generated memorable password with $wordCount $language words`n")
    }
    
    $controls.GeneratedPassword.Text = $password
    
    # Convert the password to a SecureString for strength calculation
    $securePassword = New-Object System.Security.SecureString
    foreach ($char in $password.ToCharArray()) {
        $securePassword.AppendChar($char)
    }
    
    Update-PasswordStrength -Password $securePassword
    
    # Reset the pwned status
    $controls.PwnedStatus.Text = "Not checked"
    $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(128, 128, 128))
    
    # Reset the Password Pusher result
    $controls.ResultUrl.Text = ""
    $controls.ResultUrl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0))
    
    # Keep the copy and open buttons disabled until a password is pushed
    $controls.CopyUrlButton.IsEnabled = $false
    $controls.OpenInBrowserButton.IsEnabled = $false
    
    # Reset the password flags
    $script:currentPasswordPushed = $false
    $script:currentPasswordChecked = $false
    
    # Only enable the push button if the push timer is not running
    if (-not $script:pushTimer.IsEnabled) {
        $controls.PushButton.IsEnabled = $true
        $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
        $controls.PushButton.Content = "Push to Password Pusher"
    } else {
        # Keep the button disabled during cooldown
        $controls.LogOutput.AppendText("Note: Push button remains disabled during cooldown period`n")
    }
    
    # Only enable the check button if the hibp timer is not running
    if (-not $script:hibpTimer.IsEnabled) {
        $controls.CheckPwnedButton.IsEnabled = $true
        $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
        $controls.CheckPwnedButton.Content = "Check"
    } else {
        # Keep the button disabled during cooldown
        $controls.LogOutput.AppendText("Note: Check button remains disabled during cooldown period`n")
    }
}

# UI event handlers

# Get the grid panels for password settings
$randomPasswordSettings = $window.FindName("RandomPasswordSettings")
$memorablePasswordSettings = $window.FindName("MemorablePasswordSettings")

# Password type radio buttons
$controls.RandomPasswordType.Add_Checked({
    $randomPasswordSettings.Visibility = [System.Windows.Visibility]::Visible
    $memorablePasswordSettings.Visibility = [System.Windows.Visibility]::Collapsed
})

$controls.MemorablePasswordType.Add_Checked({
    $randomPasswordSettings.Visibility = [System.Windows.Visibility]::Collapsed
    $memorablePasswordSettings.Visibility = [System.Windows.Visibility]::Visible
})

# Sliders
$controls.PasswordLength.Add_ValueChanged({
    $controls.PasswordLengthValue.Text = [int]$controls.PasswordLength.Value
})

$controls.WordCount.Add_ValueChanged({
    $controls.WordCountValue.Text = [int]$controls.WordCount.Value
})

$controls.GenerateButton.Add_Click({
    New-Password
})

$controls.CopyPasswordButton.Add_Click({
    if ($controls.GeneratedPassword.Text) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($controls.GeneratedPassword.Text)
            $controls.LogOutput.AppendText("Password copied to clipboard`n")
        }
        catch {
            $controls.LogOutput.AppendText("Error copying to clipboard: $($_.Exception.Message)`n")
        }
    }
})

# Countdown timers and state tracking for API rate limiting
$script:hibpCountdown = 10
$script:pushCountdown = 10
$script:currentPasswordPushed = $false
$script:currentPasswordChecked = $false

# Timers for API rate limiting with visual feedback
$script:hibpTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:hibpTimer.Interval = [TimeSpan]::FromSeconds(1) # Update every second for countdown
$script:hibpTimer.Add_Tick({
    $script:hibpCountdown--
    if ($script:hibpCountdown -le 0) {
        # Reset when countdown reaches 0
        $script:hibpTimer.Stop()
        $script:hibpCountdown = 10 # Reset countdown for next time
        
        # Only enable the button if the current password hasn't been checked yet
        if (-not $script:currentPasswordChecked) {
            $controls.CheckPwnedButton.IsEnabled = $true
            $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
            $controls.CheckPwnedButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255)) # Reset to white text
            $controls.CheckPwnedButton.FontWeight = "Normal" # Reset to normal font weight
            $controls.CheckPwnedButton.Content = "Check"
        } else {
            # Keep the button disabled but update its appearance and text
            $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(200, 200, 200)) # Gray
            $controls.CheckPwnedButton.FontWeight = "Normal" # Normal font weight (not bold)
            $controls.CheckPwnedButton.Content = "Already Checked"
        }
    } else {
        # Just display "Wait" without the countdown
        $controls.CheckPwnedButton.Content = "Wait"
    }
})

$script:pushTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:pushTimer.Interval = [TimeSpan]::FromSeconds(1) # Update every second for countdown
# Initialize the timer as stopped
$script:pushTimer.IsEnabled = $false
$script:pushTimer.Add_Tick({
    $script:pushCountdown--
    if ($script:pushCountdown -le 0) {
        # Reset when countdown reaches 0
        $script:pushTimer.Stop()
        $script:pushCountdown = 10 # Reset countdown for next time
        
        # Only enable the button if the current password hasn't been pushed yet
        if (-not $script:currentPasswordPushed) {
            $controls.PushButton.IsEnabled = $true
            $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
            $controls.PushButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255)) # Reset to white text
            $controls.PushButton.FontWeight = "Normal" # Reset to normal font weight
            $controls.PushButton.Content = "Push to Password Pusher"
        } else {
            # Keep the button disabled but update its appearance and text
            $controls.PushButton.IsEnabled = $false
            $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(200, 200, 200)) # Gray
            $controls.PushButton.FontWeight = "Normal" # Normal font weight (not bold)
            $controls.PushButton.Content = "Already Pushed"
        }
    } else {
        # Just display "Wait" without the countdown
        $controls.PushButton.Content = "Wait"
    }
})

$controls.CheckPwnedButton.Add_Click({
    if (-not $controls.GeneratedPassword.Text) {
        $controls.LogOutput.AppendText("Error: Please generate a password first`n")
        return
    }
    
    $controls.LogOutput.AppendText("Checking password against Have I Been Pwned database...`n")
    $controls.PwnedStatus.Text = "Checking..."
    $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(128, 128, 128))
    
    # Convert the password to a SecureString
    $securePassword = New-Object System.Security.SecureString
    foreach ($char in $controls.GeneratedPassword.Text.ToCharArray()) {
        $securePassword.AppendChar($char)
    }
    
    # Check if the password has been pwned
    $result = Test-PwnedPassword -Password $securePassword
    
    if ($result.Error) {
        $controls.LogOutput.AppendText("Error checking password: $($result.Error)`n")
        
        # If the error contains "Service unavailable", display a more specific message
        if ($result.Error -like "*Service unavailable*") {
            $controls.PwnedStatus.Text = "Have I Been Pwned API: Service unavailable"
            $controls.PwnedStatus.ClearValue([System.Windows.Controls.TextBlock]::ForegroundProperty)
            $controls.LogOutput.AppendText("Have I Been Pwned API: Service unavailable`n")
        } else {
            $controls.PwnedStatus.Text = "Error checking"
            $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0))
        }
    }
    elseif ($result.Found) {
        $controls.LogOutput.AppendText("Password found in $($result.Count) data breaches!`n")
        $controls.PwnedStatus.Text = "Found in $($result.Count) breaches!"
        $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 0, 0))
        
        # Disable the push button and update its text to indicate the password is compromised
        $controls.PushButton.IsEnabled = $false
        $controls.PushButton.Content = "Compromised Password - Not Secure to Share"
        $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 0, 0)) # Red background
        $controls.PushButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0)) # Black text for better visibility
    }
    else {
        $controls.LogOutput.AppendText("Password not found in any known data breaches.`n")
        $controls.PwnedStatus.Text = "Not found in any known breaches"
        $controls.PwnedStatus.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 128, 0))
    }
    
    # Mark the current password as checked
    $script:currentPasswordChecked = $true
    
    # Disable the button and start countdown timer
    $controls.CheckPwnedButton.IsEnabled = $false
    $controls.CheckPwnedButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0)) # Orange
    $controls.CheckPwnedButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0)) # Black text for better visibility
    $controls.CheckPwnedButton.FontWeight = "Normal"
    $controls.CheckPwnedButton.Content = "Wait"
    $controls.LogOutput.AppendText("Rate limit applied: Check button disabled for 10 seconds`n")
    $script:hibpCountdown = 10
    $script:hibpTimer.Start()
})

$controls.PushButton.Add_Click({
    if (-not $controls.GeneratedPassword.Text) {
        $controls.LogOutput.AppendText("Error: Please generate a password first`n")
        return
    }
    
    # Check if the password has been pwned and the push button is disabled
    if (-not $controls.PushButton.IsEnabled) {
        $controls.LogOutput.AppendText("Error: Cannot push a password that has been found in data breaches`n")
        return
    }
    
    # Convert the password to a SecureString
    $securePassword = New-Object System.Security.SecureString
    foreach ($char in $controls.GeneratedPassword.Text.ToCharArray()) {
        $securePassword.AppendChar($char)
    }
    
    $expireDays = $controls.ExpireDays.SelectedItem.Tag
    $expireViews = $controls.ExpireViews.SelectedItem.Tag
    $deletableByViewer = $controls.DeletableByViewer.IsChecked
    $retrievalStep = $controls.RetrievalStep.IsChecked
    $useCurl = $controls.CurlApiMethod.IsChecked
    
    # Convert the passphrase to a SecureString if provided
    $securePassphrase = New-Object System.Security.SecureString
    if ($controls.UsePassphrase.IsChecked) {
        foreach ($char in $controls.Passphrase.Password.ToCharArray()) {
            $securePassphrase.AppendChar($char)
        }
    }
    
    $useQRCode = $controls.UseQRCode.IsChecked
    
    $controls.LogOutput.AppendText("Pushing password to Password Pusher...`n")
    $controls.LogOutput.AppendText("Settings: Expire after $expireDays days or $expireViews views`n")
    
    $result = Push-Password -Password $securePassword -ExpireDays $expireDays -ExpireViews $expireViews -DeletableByViewer $deletableByViewer -RetrievalStep $retrievalStep -Passphrase $securePassphrase -UseQRCode $useQRCode -UseCurl $useCurl
    
    $controls.LogOutput.AppendText($result.Log + "`n")
    
    if ($result.Success) {
        $controls.ResultUrl.Text = $result.Url
        $controls.ResultUrl.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0)) # Reset to black text
        
        # If QR code was generated, add additional information
        if ($result.IsQRCode) {
            $controls.LogOutput.AppendText("A QR code has been generated. Click 'Browse' to view and scan it.`n")
        }
        
        # Enable the copy and open buttons
        $controls.CopyUrlButton.IsEnabled = $true
        $controls.OpenInBrowserButton.IsEnabled = $true
        
        # Mark the current password as pushed
        $script:currentPasswordPushed = $true
    } 
    else {
        # Check if the error is due to service unavailability
        if ($result.Log -like "*Service unavailable*") {
            $controls.LogOutput.AppendText("Password Pusher service is not accessible. Please check your internet connection.`n")
            
            # Update the UI to show service unavailable status
            $controls.ResultUrl.Text = "Password Pusher API: Service unavailable"
            $controls.ResultUrl.ClearValue([System.Windows.Controls.TextBlock]::ForegroundProperty)
            
            # Disable the copy and open buttons since there's no valid URL
            $controls.CopyUrlButton.IsEnabled = $false
            $controls.OpenInBrowserButton.IsEnabled = $false
            
            # Mark the current password as pushed so the button remains disabled after countdown
            $script:currentPasswordPushed = $true
        }
    }
    
    # Disable the button and start countdown timer
    $controls.PushButton.IsEnabled = $false
    $controls.PushButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 165, 0)) # Orange
    $controls.PushButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 0, 0)) # Black text for better visibility
    $controls.PushButton.FontWeight = "Normal"
    $controls.PushButton.Content = "Wait"
    $controls.LogOutput.AppendText("Rate limit applied: Push button disabled for 10 seconds`n")
    $script:pushCountdown = 10
    $script:pushTimer.Start()
})

$controls.CopyUrlButton.Add_Click({
    if ($controls.ResultUrl.Text) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($controls.ResultUrl.Text)
            $controls.LogOutput.AppendText("URL copied to clipboard`n")
        }
        catch {
            $controls.LogOutput.AppendText("Error copying to clipboard: $($_.Exception.Message)`n")
        }
    }
})

$controls.OpenInBrowserButton.Add_Click({
    if ($controls.ResultUrl.Text) {
        try {
            Start-Process $controls.ResultUrl.Text
            $controls.LogOutput.AppendText("Opening URL in browser`n")
        }
        catch {
            $controls.LogOutput.AppendText("Error opening URL: $($_.Exception.Message)`n")
        }
    }
})

# Handles password preset selection changes
$global:controls.PasswordPresets.Add_SelectionChanged({
    $selectedItem = $global:controls.PasswordPresets.SelectedItem
    
    # Get the preset object from the selected item's Tag property
    $preset = $selectedItem.Tag
    
    if ($preset) {
        # Update UI controls based on the preset's properties
        $global:controls.PasswordLength.Value = $preset.Length
        $global:controls.IncludeUppercase.IsChecked = $preset.IncludeUppercase
        $global:controls.IncludeNumbers.IsChecked = $preset.IncludeNumbers
        $global:controls.IncludeSpecial.IsChecked = $preset.IncludeSpecial
        
        $global:controls.LogOutput.AppendText("Applied preset: $($preset.Name)`n")
        $global:controls.LogOutput.AppendText("  - Length: $($preset.Length)`n")
        $global:controls.LogOutput.AppendText("  - Uppercase: $($preset.IncludeUppercase)`n")
        $global:controls.LogOutput.AppendText("  - Numbers: $($preset.IncludeNumbers)`n")
        $global:controls.LogOutput.AppendText("  - Special: $($preset.IncludeSpecial)`n")
        
        # Generate a new password with the updated settings
        New-Password
    }
})

# GitHub button click handler
$gitHubButton = $window.FindName("GitHubButton")
$gitHubButton.Add_Click({
    Start-Process "https://github.com/onlyalex1984/securepassgenerator-powershell"
    $controls.LogOutput.AppendText("Opening GitHub page in browser`n")
})

# Password Pusher website link handler
$passwordPusherButton = $window.FindName("PasswordPusherButton")
$passwordPusherButton.Add_Click({
    Start-Process "https://pwpush.com"
    $controls.LogOutput.AppendText("Opening Password Pusher website in browser`n")
})

# HIBP website link handler
$hibpButton = $window.FindName("HibpButton")
$hibpButton.Add_Click({
    Start-Process "https://haveibeenpwned.com"
    $controls.LogOutput.AppendText("Opening Have I Been Pwned website in browser`n")
})

# Initial application startup log
$controls.LogOutput.AppendText("Application started`n")

# Apply UI theme and initial settings
Set-LightTheme

# Disable result buttons until password is pushed
$controls.CopyUrlButton.IsEnabled = $false
$controls.OpenInBrowserButton.IsEnabled = $false
$controls.ResultUrl.Text = ""

# Configure history button appearance
$historyButton = $window.FindName("HistoryButton")
$historyButton.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 123, 255))
$historyButton.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 255, 255))

New-Password

# Phonetic button click handler
$controls.PhoneticButton.Add_Click({
    if ($controls.GeneratedPassword.Text) {
        # Convert the password to a SecureString
        $securePassword = New-Object System.Security.SecureString
        foreach ($char in $controls.GeneratedPassword.Text.ToCharArray()) {
            $securePassword.AppendChar($char)
        }
        # Log showing phonetic codes with system name
        $controls.LogOutput.AppendText("Showing Swedish phonetic codes`n")
        Show-PhoneticCodes -Password $securePassword -PhoneticLanguage "Swedish"
    }
})

# Phonetic options context menu
$contextMenu = New-Object System.Windows.Controls.ContextMenu

# NATO phonetic menu option
$natoMenuItem = New-Object System.Windows.Controls.MenuItem
$natoMenuItem.Header = "NATO Phonetic"
$natoMenuItem.Add_Click({
    if ($controls.GeneratedPassword.Text) {
        # Convert the password to a SecureString
        $securePassword = New-Object System.Security.SecureString
        foreach ($char in $controls.GeneratedPassword.Text.ToCharArray()) {
            $securePassword.AppendChar($char)
        }
        # Log showing phonetic codes with system name
        $controls.LogOutput.AppendText("Showing NATO phonetic codes`n")
        Show-PhoneticCodes -Password $securePassword -PhoneticLanguage "NATO"
    }
}) | Out-Null
[void]$contextMenu.Items.Add($natoMenuItem)

# Swedish phonetic menu option
$swedishMenuItem = New-Object System.Windows.Controls.MenuItem
$swedishMenuItem.Header = "Swedish Phonetic"
$swedishMenuItem.Add_Click({
    if ($controls.GeneratedPassword.Text) {
        # Convert the password to a SecureString
        $securePassword = New-Object System.Security.SecureString
        foreach ($char in $controls.GeneratedPassword.Text.ToCharArray()) {
            $securePassword.AppendChar($char)
        }
        # Log showing phonetic codes with system name
        $controls.LogOutput.AppendText("Showing Swedish phonetic codes`n")
        Show-PhoneticCodes -Password $securePassword -PhoneticLanguage "Swedish"
    }
}) | Out-Null
[void]$contextMenu.Items.Add($swedishMenuItem)

# Phonetic dropdown button handler
$controls.PhoneticDropdownButton.Add_Click({
    if ($controls.GeneratedPassword.Text) {
        $contextMenu.IsOpen = $true
        $contextMenu.PlacementTarget = $controls.PhoneticDropdownButton
        $contextMenu.Placement = "Bottom"
    }
})

# History button click handler
$historyButton = $window.FindName("HistoryButton")
$historyButton.Add_Click({
    Show-PasswordLinksHistory
    $controls.LogOutput.AppendText("Opened password links history window`n")
})

# Real-time password strength evaluation on text change
$controls.GeneratedPassword.Add_TextChanged({
    # Get the current text from the TextBox
    $password = $controls.GeneratedPassword.Text
    
    # Create a SecureString from the plain text password
    $securePassword = New-Object System.Security.SecureString
    if (-not [string]::IsNullOrWhiteSpace($password)) {
        foreach ($char in $password.ToCharArray()) {
            $securePassword.AppendChar($char)
        }
    }
    
    # Measure the password strength and update UI
    Measure-CustomPassword -Password $securePassword
})

# Check for Updates button click handler
$checkForUpdatesButton = $window.FindName("CheckForUpdatesButton")
$checkForUpdatesButton.Add_Click({
    $global:controls.LogOutput.AppendText("Checking for updates...`n")

    # Show a progress indicator
    $global:controls.CheckForUpdatesButton.IsEnabled = $false
    $global:controls.CheckForUpdatesButton.Content = "Checking..."

    try {
        # Get update information
        $updateInfo = Get-UpdateInformation

        # Reset button state
        $global:controls.CheckForUpdatesButton.IsEnabled = $true
        $global:controls.CheckForUpdatesButton.Content = "Check for Updates"

        # Check if there was an error with the update check
        if ($null -eq $updateInfo) {
            [System.Windows.MessageBox]::Show(
                "Unable to check for updates. Please check your internet connection and try again.",
                "Update Check Failed",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Check if there's an error message in the update info
        if ($updateInfo.Error) {
            [System.Windows.MessageBox]::Show(
                "Cannot connect to GitHub: $($updateInfo.Error)`n`nPlease check your internet connection or proxy settings and try again.",
                "Update Check Failed",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }

        # Show update dialog only if we have valid update information
        $dialogResult = Show-UpdateDialog -UpdateInfo $updateInfo

        # Process update if user confirmed
        if ($dialogResult.Result -eq $true) {
            $releaseType = $dialogResult.ReleaseType
            Invoke-Update -ReleaseType $releaseType
        }
    }
    catch {
        # Reset button state
        $global:controls.CheckForUpdatesButton.IsEnabled = $true
        $global:controls.CheckForUpdatesButton.Content = "Check for Updates"

        $global:controls.LogOutput.AppendText("Error checking for updates: $($_.Exception.Message)`n")
        [System.Windows.MessageBox]::Show(
            "An error occurred while checking for updates: $($_.Exception.Message)",
            "Update Check Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# Window loaded event to ensure visibility
$window.Add_Loaded({
    # Set window as topmost to bring it to the front
    $window.Topmost = $true
    # Request focus
    $window.Activate()
    # Revert topmost setting to allow normal behavior
    $window.Topmost = $false
    
    # Load custom presets on startup
    Import-PasswordPresets
    $global:controls.LogOutput.AppendText("Loaded custom presets on startup`n")
    
    # Debug: Check if PresetsSettingsButton exists
    if ($global:controls.PresetsSettingsButton) {
        $global:controls.LogOutput.AppendText("PresetsSettingsButton found in controls`n")
    } else {
        $global:controls.LogOutput.AppendText("ERROR: PresetsSettingsButton not found in controls!`n")
    }
    
    # Add settings button click handler here, after controls are fully initialized
    $global:controls.PresetsSettingsButton.Add_Click({
        $global:controls.LogOutput.AppendText("Settings button clicked`n")
        Show-PasswordPresets
        $global:controls.LogOutput.AppendText("Opened password presets management window`n")
    })
})

# Display application window
$window.ShowDialog() | Out-Null
        $selectedWords += $word
