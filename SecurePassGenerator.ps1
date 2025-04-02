<#
.SYNOPSIS
    SecurePassGenerator - A Comprehensive Password Generation and Sharing Tool

.DESCRIPTION
    A PowerShell script that creates a modern WPF GUI for generating secure passwords
    and pushing them to Password Pusher's API. Features include password customization,
    strength assessment, breach checking, and secure sharing options.

    Key Features:
    - Multiple password generation methods (random or memorable)
    - Real-time password strength assessment with entropy calculation
    - Integration with Have I Been Pwned to check for compromised passwords
    - Secure password sharing via Password Pusher with configurable settings
    - QR code generation for mobile access
    - Rate-limited API interactions to respect service limitations
    
    Note: This script implements cooldown timers between API requests to respect
    rate limits for external services. Please do not modify these limiting features.

.NOTES
    File Name      : SecurePassGenerator.ps1
    Version        : 1.0
    Author         : onlyalex1984
    Copyright      : (C) 2025 onlyalex1984
    License        : GPL v3 - Full license text available in the project root directory
    Created        : Februari 2025
    Last Modified  : April 2025
    Prerequisite   : PowerShell 5.1 or higher, Windows environment with .NET Framework
    GitHub         : https://github.com/onlyalex1984/securepassgenerator-powershell
#>

# Script configuration
$script:Version = "1.0"
$script:ScriptDisplayName = "SecurePassGenerator"

# Password links history collection
$script:PasswordLinks = New-Object System.Collections.ArrayList

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
    catch {
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
                    <ComboBox Grid.Column="6" x:Name="PasswordPresets" Width="150" HorizontalAlignment="Right" SelectedIndex="1" Visibility="{Binding ElementName=RandomPasswordType, Path=IsChecked, Converter={StaticResource BooleanToVisibilityConverter}}">
                        <ComboBoxItem Content="Medium Password"/>
                        <ComboBoxItem Content="Strong Password"/>
                        <ComboBoxItem Content="Very Strong Password"/>
                    </ComboBox>
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
                                <TextBlock Text="- Predefined password presets for different security levels" TextWrapping="Wrap" Margin="0,2,0,2"/>
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
                                
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Version: " FontWeight="SemiBold"/>
                                    <TextBlock x:Name="VersionText" Text="1.0"/>
                                </StackPanel>
                            </StackPanel>
                            
                            <Button x:Name="GitHubButton" Grid.Column="1" Content="Project on GitHub" Width="150" Height="30"
                                    Background="#0078D7" Foreground="White" Padding="5,3" VerticalAlignment="Center"/>
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
$controls = @{
    RandomPasswordType = $window.FindName("RandomPasswordType")
    MemorablePasswordType = $window.FindName("MemorablePasswordType")
    LanguageSelector = $window.FindName("LanguageSelector")
    PasswordPresets = $window.FindName("PasswordPresets")
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
$controls.PasswordPresets.Add_SelectionChanged({
    $selectedPreset = $controls.PasswordPresets.SelectedItem.Content
    
    # Set all checkboxes to checked for all presets
    $controls.IncludeUppercase.IsChecked = $true
    $controls.IncludeNumbers.IsChecked = $true
    $controls.IncludeSpecial.IsChecked = $true
    
    # Adjust password length based on the selected preset
    switch ($selectedPreset) {
        "Medium Password" {
            $controls.PasswordLength.Value = 10
            $controls.LogOutput.AppendText("Applied Medium Password preset (10 characters, uppercase, numbers, special)`n")
        }
        "Strong Password" {
            $controls.PasswordLength.Value = 15
            $controls.LogOutput.AppendText("Applied Strong Password preset (15 characters, uppercase, numbers, special)`n")
        }
        "Very Strong Password" {
            $controls.PasswordLength.Value = 20
            $controls.LogOutput.AppendText("Applied Very Strong Password preset (20 characters, uppercase, numbers, special)`n")
        }
    }
    
    # Generate a new password with the updated settings
    New-Password
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

# Window loaded event to ensure visibility
$window.Add_Loaded({
    # Set window as topmost to bring it to the front
    $window.Topmost = $true
    # Request focus
    $window.Activate()
    # Revert topmost setting to allow normal behavior
    $window.Topmost = $false
})

# Display application window
$window.ShowDialog() | Out-Null
