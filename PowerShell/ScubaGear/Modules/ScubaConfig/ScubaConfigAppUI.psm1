Function Invoke-SCuBAConfigAppUI {
    <#
    .SYNOPSIS
    Opens the ScubaConfig UI for configuring Scuba settings.
    .DESCRIPTION
    This function opens a WPF-based UI for configuring Scuba settings.
    .EXAMPLE
    Invoke-SCuBAConfigAppUI
    # Opens the ScubaConfig UI.
    .PARAMETER YAMLConfigFile
    Specifies the YAML configuration file to load. If not provided, the default configuration will be used.
    .PARAMETER Language
    Specifies the language for the UI. Default is 'en-US'.
    .PARAMETER Online
    If specified, connects to Microsoft Graph to retrieve additional configuration data.
    .PARAMETER M365Environment
    Specifies the M365 environment to use. Valid values are 'commercial', 'dod', 'gcc', 'gcchigh'. Default is 'commercial'.
    .PARAMETER Passthru
    If specified, returns the configuration object after loading.
    .NOTES
    This function requires the ScubaConfig module to be loaded and the ConvertFrom-Yaml function to be available.
    .LINK
    https://github.com/cisagov/ScubaGear
    #>

    [CmdletBinding(DefaultParameterSetName = 'Offline')]
    Param(
        $YAMLConfigFile,

        [ValidateSet('en-US')]
        $Language = 'en-US',

        [Parameter(Mandatory = $false,ParameterSetName = 'Online')]
        [switch]$Online,

        [Parameter(Mandatory = $true,ParameterSetName = 'Online')]
        [ValidateSet('commercial', 'dod', 'gcc', 'gcchigh')]
        [string]$M365Environment,

        [switch]$Passthru
    )

    [string]${CmdletName} = $MyInvocation.MyCommand
    Write-Verbose ("{0}: Sequencer started" -f ${CmdletName})

    switch($M365Environment){
        'commercial' {
            $GraphEndpoint = "https://graph.microsoft.com"
            $graphEnvironment = "Global"
        }
        'gcc' {
            $GraphEndpoint = "https://graph.microsoft.com"
            $graphEnvironment = "Global"
        }
        'gcchigh' {
            # Set GCC High environment variables
            $GraphEndpoint = "https://graph.microsoft.us"
            $graphEnvironment = "UsGov"
        }
        'dod' {
            # Set DoD environment variables
            $GraphEndpoint = "https://dod-graph.microsoft.us"
            $graphEnvironment = "UsGovDoD"
        }
        default {
            $GraphEndpoint = "https://graph.microsoft.com"
            $graphEnvironment = "Global"
        }

    }

    # Connect to Microsoft Graph if Online parameter is used
    if ($Online) {
        try {
            Write-Output "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Policy.Read.All", "Organization.Read.All" -NoWelcome -Environment $graphEnvironment -ErrorAction Stop | Out-Null
            $Online = $true
            Write-Output "Successfully connected to Microsoft Graph"
        }
        catch {
            Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
            $Online = $false
        }
    } else {
        $Online = $false
    }

    # build a hash table with locale data to pass to runspace
    $syncHash = [hashtable]::Synchronized(@{})
    $Runspace = [runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $Runspace
    $syncHash.GraphConnected = $Online
    $syncHash.XamlPath = "$PSScriptRoot\ScubaConfigAppUI.xaml"
    $syncHash.UIConfigPath = "$PSScriptRoot\ScubaConfig_$Language.json"
    $syncHash.YAMLImport = $YAMLConfigFile
    $syncHash.GraphEndpoint = $GraphEndpoint
    $syncHash.M365Environment = $M365Environment
    $syncHash.Exclusions = @{}
    $syncHash.Omissions = @{}
    $syncHash.Annotations = @{}
    $syncHash.GeneralSettings = @{}
    $syncHash.BulkUpdateInProgress = $false
    $syncHash.Placeholder = @{}
    $syncHash.DebugOutputQueue = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
    $syncHash.DebugFlushTimer = $null
    #$syncHash.Theme = $Theme
    #build runspace
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open() | Out-Null
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $PowerShellCommand = [PowerShell]::Create().AddScript({

        #Load assembies to display UI
        [System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework') | out-null
        [System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      | out-null
        #Load additional assemblies for folder browser and certificate selection
        [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | out-null
        [System.Reflection.Assembly]::LoadWithPartialName('System.Security') | out-null

        #need to replace compile code in xaml and x:Class and xmlns needs to be removed
        #$xaml = $xaml -replace 'xmlns:x="http://schemas.microsoft

        [string]$XAML = (Get-Content $syncHash.XamlPath -ReadCount 0) -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window' -replace 'Click=".*','/>'
        [xml]$UIXML = $XAML
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$UIXML)
        $syncHash.window = [Windows.Markup.XamlReader]::Load($reader)
        $syncHash.UIXML = $UIXML
        # Store Form Objects In PowerShell
        $UIXML.SelectNodes("//*[@Name]") | ForEach-Object{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

        function Write-DebugOutput {
            <#
            .SYNOPSIS
            Writes debug output messages to the debug queue when debug mode is enabled.
            .DESCRIPTION
            This function adds timestamped debug messages to the syncHash debug queue for troubleshooting and monitoring UI operations.
            #>
            param(
                [string]$Message,
                [string]$Source = "General",
                [string]$Level = "Info"
            )

            if ($syncHash.DebugMode -ne 'None') {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                $logEntry = "[$timestamp] [$Level] [$Source] $Message"

                $syncHash.DebugOutputQueue.Enqueue($logEntry)
            }
        }

        #Import UI configuration file
        $syncHash.UIConfigs = Get-Content -Path $syncHash.UIConfigPath -Raw | ConvertFrom-Json
        Write-DebugOutput -Message "UIConfigs loaded: $($syncHash.UIConfigPath)" -Source "UI Launch" -Level "Info"

        $syncHash.DebugMode = $syncHash.UIConfigs.DebugMode

        # If YAMLImport is specified, load the YAML configuration
        If($syncHash.YAMLImport){
            $syncHash.YAMLConfig = Get-Content -Path $syncHash.YAMLImport -Raw | ConvertFrom-Yaml
            Write-DebugOutput -Message "YAMLConfig loaded: $($syncHash.YAMLImport)" -Source "UI Launch" -Level "Info"
        }

        $syncHash.Window.Dispatcher.Invoke([Action]{
            try {
                $syncHash.Debug_TextBox.AppendText("UI START`r`n")
                $syncHash.Debug_TextBox.ScrollToEnd()
            } catch {
                Write-Error "Dispatcher error: $($_.Exception.Message)"
            }
        })

        # Create a DispatcherTimer for periodic UI updates and debug log flushing
        $syncHash.UIUpdateTimer = New-Object System.Windows.Threading.DispatcherTimer
        $syncHash.UIUpdateTimer.Interval = [System.TimeSpan]::FromMilliseconds(500)
        $syncHash.UIUpdateTimer.Add_Tick({
            try {
                # ===================== DEBUG LOG FLUSH =====================
                if ($syncHash.Debug_TextBox -and $syncHash.Debug_TextBox.IsLoaded) {
                    $logBatch = @()
                    while ($syncHash.DebugOutputQueue.Count -gt 0) {
                        $logBatch += $syncHash.DebugOutputQueue.Dequeue()
                    }

                    if ($logBatch.Count -gt 0) {
                        $textToAdd = $logBatch -join "`r`n"

                        if ($syncHash.Debug_TextBox.Text.Length -gt 0) {
                            $syncHash.Debug_TextBox.AppendText("`r`n$textToAdd")
                        } else {
                            $syncHash.Debug_TextBox.AppendText($textToAdd)
                        }

                        $syncHash.Debug_TextBox.ScrollToEnd()

                        # Trim lines to 1000
                        $lines = $syncHash.Debug_TextBox.Text -split "`r`n"
                        if ($lines.Count -gt 1000) {
                            $syncHash.Debug_TextBox.Text = ($lines[-1000..-1] -join "`r`n")
                        }
                    }
                }

                # ===================== GENERAL UI UPDATE =====================

                # Auto-collect general settings from UI controls
                Save-GeneralSettingsFromInput

                # Only update if there have been changes
                if ($syncHash.DataChanged) {
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message $syncHash.UIConfigs.LocaleDebugOutput.UITimerDebug -Source "UI Timer" -Level "Debug"}
                    Update-AllUIFromData
                    $syncHash.DataChanged = $false
                }

            } catch {
                If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UITimerError -f $_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Level "Error"}
            }
        })


        # Initialize change tracking
        $syncHash.DataChanged = $false

        Write-DebugOutput -Message $syncHash.UIConfigs.LocaleDebugOutput.UIInitializationInfo -Source "UI Initialization" -Level "Info"

        $syncHash.LastUpdateHash = @{
            Exclusions = ""
            Omissions = ""
            Annotations = ""
            GeneralSettings = ""
        }


        # Function to mark data as changed
        Function Set-DataChanged {
            <#
            .SYNOPSIS
            Sets the DataChanged flag to indicate user modifications.
            .DESCRIPTION
            This function sets the DataChanged flag in the syncHash to track when user input has modified the configuration data.
            #>
            $syncHash.DataChanged = $true
        }

        # Recursively find all controls
        function Find-ControlElement {
            <#
            .SYNOPSIS
            Recursively searches for all control elements within a WPF container.
            .DESCRIPTION
            This function traverses the visual tree to find and return all control elements contained within a specified parent container.
            #>
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.DependencyObject]$Parent
            )

            $results = @()

            for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent); $i++) {
                $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
                if ($child -is [System.Windows.Controls.Control]) {
                    $results += $child
                }
                $results += Find-ControlElement -Parent $child
            }

            return $results
        }

        # Function to add event handlers to a specific control (for dynamically created controls)
        Function Add-ControlEventHandler {
            <#
            .SYNOPSIS
            Adds event handlers to dynamically created WPF controls.
            .DESCRIPTION
            This function attaches appropriate event handlers to different types of WPF controls (TextBox, ComboBox, Button, etc.) for user interaction tracking.
            #>
            param(
                [System.Windows.Controls.Control]$Control
            )

            If($syncHash.DebugMode -match 'Verbose|Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventAdd -f $Control.GetType().Name, $Control.Name) -Source $MyInvocation.MyCommand.Name -Level "Verbose"}

            switch($Control.GetType().Name) {
                'TextBox' {
                    # Add LostFocus event
                    $Control.Add_LostFocus({
                        $controlName = if ($this.Name) { $this.Name } else { "Unnamed TextBox" }
                        $controlValue = $this.Text
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventSelect -f $Control.GetType().Name, $controlName, $controlValue) -Source "Control Handler" -Level "Info"
                    }.GetNewClosure())
                }
                'ComboBox' {
                    # Add SelectionChanged event
                    $Control.Add_SelectionChanged({
                        $controlName = if ($this.Name) { $this.Name } else { "Unnamed ComboBox" }
                        $selectedItem = $this.SelectedItem
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventSelect -f $Control.GetType().Name, $controlName, $selectedItem) -Source "Control Handler" -Level "Info"
                    }.GetNewClosure())
                }
                'Button' {
                    # Add Click event
                    $Control.Add_Click({
                        $controlName = if ($this.Name) { $this.Name } else { "Unnamed Button" }
                        $controlContent = if ($this.Content) { " ($($this.Content))" } else { "" }
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventSelect -f $Control.GetType().Name, $controlName, $controlContent) -Source "Control Handler" -Level "Info"
                    }.GetNewClosure())
                }
                'CheckBox' {
                    # Add Checked event
                    $Control.Add_Checked({
                        $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                        $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventSelect -f $Control.GetType().Name, $controlName, $controlTag) -Source "Control Handler" -Level "Info"
                    }.GetNewClosure())

                    # Add Unchecked event
                    $Control.Add_Unchecked({
                        $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                        $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ControlEventSelect -f $Control.GetType().Name, $controlName, $controlTag) -Source "Control Handler" -Level "Info"
                    }.GetNewClosure())
                }
            }
        }

        #===========================================================================
        # UI Helper Functions
        #===========================================================================
        # Add global event handlers to all UI controls after everything is loaded
        #Closes UI objects and exits (within runspace)
        Function Close-UIMainWindow {
            <#
            .SYNOPSIS
            Closes the main UI window and performs cleanup operations.
            .DESCRIPTION
            This function handles the controlled closing of the main application window, including memory cleanup and event handler removal.
            #>
            # Prevent recursive calls
            if ($syncHash.isClosing) {
                return
            }

            try {
                # Mark as closing immediately
                $syncHash.isClosing = $true

                # 1. Stop timers only (essential for preventing crashes)
                if ($syncHash.UIUpdateTimer) {
                    $syncHash.UIUpdateTimer.Stop()
                    $syncHash.UIUpdateTimer = $null
                }

                if ($syncHash.DebugFlushTimer) {
                    $syncHash.DebugFlushTimer.Stop()
                    $syncHash.DebugFlushTimer = $null
                }

                # 2. Disconnect from Microsoft Graph if connected (important)
                if ($syncHash.GraphConnected) {
                    try {
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                    } catch {
                        Write-Error "Failed to disconnect from Microsoft Graph: $_"
                    }
                }

                # 3. Close window immediately - let .NET handle the rest
                if ($syncHash.Window) {
                    try {
                        # Just close the window - skip all the detailed cleanup
                        if ($syncHash.Window.Dispatcher.CheckAccess()) {
                            $syncHash.Window.Close()
                        } else {
                            $syncHash.Window.Dispatcher.Invoke([Action]{ $syncHash.Window.Close() })
                        }
                    } catch {
                        # If window can't close normally, force it
                        $syncHash.Window = $null
                    }
                }

                # 4. Set flags and let .NET garbage collector handle the cleanup
                $syncHash.isClosed = $true

                # Optional: Quick memory cleanup (much faster than the detailed version)
                [System.GC]::Collect()

            } catch {
                # Don't block closing on any errors
                $syncHash.Window = $null
                $syncHash.isClosed = $true
            }
        }

        # Helper function to recursively remove event handlers
        Function Remove-AllEventHandlers {
            <#
            .SYNOPSIS
            Recursively removes event handlers from all controls in the UI.
            .DESCRIPTION
            This function traverses the visual tree and removes event handlers from controls to prevent memory leaks during application cleanup.
            #>
            param(
                [System.Windows.DependencyObject]$Parent
            )

            if ($null -eq $Parent) { return }

            try {
                # Remove common event handlers - use try-catch for each to avoid stopping on errors
                switch ($Parent.GetType().Name) {
                    'Button' {
                        try {
                            # Clear the click event handlers
                            $Parent.ClearValue([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
                        } catch {  Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.RemoveEventHandlerError -f $Parent.GetType().Name) -Source "Event Handler Cleanup" -Level "Warning"}
                    }
                    'TextBox' {
                        try {
                            $Parent.ClearValue([System.Windows.Controls.TextBox]::TextChangedEvent)
                            $Parent.ClearValue([System.Windows.UIElement]::GotFocusEvent)
                            $Parent.ClearValue([System.Windows.UIElement]::LostFocusEvent)
                        } catch {  Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.RemoveEventHandlerError -f $Parent.GetType().Name) -Source "Event Handler Cleanup" -Level "Warning"}
                    }
                    'CheckBox' {
                        try {
                            $Parent.ClearValue([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent)
                            $Parent.ClearValue([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent)
                        } catch {  Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.RemoveEventHandlerError -f $Parent.GetType().Name) -Source "Event Handler Cleanup" -Level "Warning"}
                    }
                    'ComboBox' {
                        try {
                            $Parent.ClearValue([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent)
                        } catch {  Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.RemoveEventHandlerError -f $Parent.GetType().Name) -Source "Event Handler Cleanup" -Level "Warning"}
                    }
                }

                # Recursively process children
                $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
                for ($i = 0; $i -lt $childCount; $i++) {
                    $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
                    Remove-AllEventHandlers -Parent $child
                }

            } catch {
                # Silently continue if we can't remove handlers - but add to debug if possible
                if ($syncHash.DebugOutputQueue) {
                    try {
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                        $logEntry = "[$timestamp] [Warning] [Event Handler Cleanup] Could not remove event handlers from $($Parent.GetType().Name): $($_.Exception.Message)"
                        $syncHash.DebugOutputQueue.Enqueue($logEntry)
                    } catch { Write-Error "Failed to add debug message: $($_.Exception.Message)" }
                }
            }
        }

        # Enhanced memory cleanup function that also uses debug queue
        Function Invoke-MemoryCleanup {
            <#
            .SYNOPSIS
            Performs comprehensive memory cleanup for the application.
            .DESCRIPTION
            This function executes garbage collection and clears various data structures to free up memory resources.
            #>
            param(
                [switch]$Aggressive
            )

            # Helper function for this cleanup
            Function Add-MemoryDebugMessage {
                <#
                .SYNOPSIS
                Adds debug messages specifically for memory cleanup operations.
                .DESCRIPTION
                This function creates debug messages related to memory cleanup activities and adds them to the debug queue.
                #>
                param([string]$Message, [string]$Level = "Info")

                if ($syncHash.DebugOutputQueue) {
                    try {
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
                        $logEntry = "[$timestamp] [$Level] [Memory Cleanup] $Message"
                        $syncHash.DebugOutputQueue.Enqueue($logEntry)
                    } catch { Write-Error "Failed to add debug message: $($_.Exception.Message)" }
                }
            }

            Add-MemoryDebugMessage "Running memory cleanup..." "Info"

            try {
                $memoryBefore = [System.GC]::GetTotalMemory($false)

                # Standard cleanup
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()

                if ($Aggressive) {
                    Add-MemoryDebugMessage "Running aggressive memory cleanup..." "Info"

                    # Force cleanup of all generations
                    for ($i = 0; $i -le [System.GC]::MaxGeneration; $i++) {
                        [System.GC]::Collect($i, [System.GCCollectionMode]::Forced)
                    }

                    # Wait for finalizers again
                    [System.GC]::WaitForPendingFinalizers()

                    # Final collection
                    [System.GC]::Collect()

                    # Compact Large Object Heap if available (.NET 4.5.1+)
                    try {
                        [System.Runtime.GCSettings]::LargeObjectHeapCompactionMode = [System.Runtime.GCLargeObjectHeapCompactionMode]::CompactOnce
                        [System.GC]::Collect()
                        Add-MemoryDebugMessage "Large Object Heap compaction completed" "Info"
                    } catch {
                        Add-MemoryDebugMessage "LOH compaction not available: $($_.Exception.Message)" "Warning"
                    }
                }

                # Report memory usage
                $memoryAfter = [System.GC]::GetTotalMemory($true)
                $freed = $memoryBefore - $memoryAfter

                Add-MemoryDebugMessage "Memory cleanup completed" "Info"
                Add-MemoryDebugMessage "  Before: $([math]::Round($memoryBefore / 1MB, 2)) MB" "Info"
                Add-MemoryDebugMessage "  After:  $([math]::Round($memoryAfter / 1MB, 2)) MB" "Info"
                Add-MemoryDebugMessage "  Freed:  $([math]::Round($freed / 1MB, 2)) MB" "Info"

            } catch {
                Add-MemoryDebugMessage "Error during memory cleanup: $($_.Exception.Message)" "Error"
            }
        }

        # Add cleanup function to syncHash
        $syncHash.InvokeMemoryCleanup = { Invoke-MemoryCleanup }

        # Function to initialize placeholder text behavior for TextBox controls
        Function Initialize-PlaceholderTextBox {
            <#
            .SYNOPSIS
            Configures placeholder text behavior for TextBox controls.
            .DESCRIPTION
            This function sets up placeholder text that appears when TextBox controls are empty and manages the visual styling for placeholder display.
            #>
            param(
                [System.Windows.Controls.TextBox]$TextBox,
                [string]$PlaceholderText,
                [string]$InitialValue = $null
            )

            # Set initial value or placeholder
            if (![string]::IsNullOrWhiteSpace($InitialValue)) {
                $TextBox.Text = $InitialValue
                $TextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $TextBox.FontStyle = [System.Windows.FontStyles]::Normal
            } else {
                $TextBox.Text = $PlaceholderText
                $TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $TextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }

            <#
            # Add PreviewTextInput to handle any character input (including space)
            $TextBox.Add_PreviewTextInput({
                param($sender, $e)
                # If we're in placeholder mode, clear it before allowing input
                if ($this.Text -eq $PlaceholderText -and $this.Foreground -eq [System.Windows.Media.Brushes]::Gray) {
                    $this.Clear()
                    $this.Foreground = [System.Windows.Media.Brushes]::Black
                    $this.FontStyle = [System.Windows.FontStyles]::Normal
                }
                # Don't handle the event - let the character be processed normally
                $e.Handled = $false
            }.GetNewClosure())
            #>

            # Add GotFocus event handler (as backup)
            $TextBox.Add_GotFocus({
                if ($this.Text -eq $placeholderText) {
                    $this.Text = ""
                    $this.Foreground = [System.Windows.Media.Brushes]::Black
                    $this.FontStyle = "Normal"
                }
            }.GetNewClosure())

            # Add LostFocus event handler
            $TextBox.Add_LostFocus({
                #param($sender, $e)
                if ([string]::IsNullOrWhiteSpace($this.Text)) {
                    $this.Text = $placeholderText
                    $this.Foreground = [System.Windows.Media.Brushes]::Gray
                    $this.FontStyle = "Italic"
                }
            }.GetNewClosure())

#param($sender, $e)
        }

        # Helper function to find control by setting name
        Function Find-ControlBySettingName {
            <#
            .SYNOPSIS
            Searches for WPF controls using various naming conventions.
            .DESCRIPTION
            This function attempts to locate controls by trying multiple naming patterns and conventions commonly used in the application.
            #>
            param([string]$SettingName)

            # Define naming patterns to try
            $namingPatterns = @(
                $SettingName,                           # Direct name
                "$SettingName`_TextBox"                # SettingName_TextBox
                "$SettingName`_TextBlock"              # SettingName_TextBlock
                "$SettingName`_CheckBox"               # SettingName_CheckBox
                "$SettingName`_ComboBox"               # SettingName_ComboBox
                "$SettingName`_Label"                  # SettingName_Label
                "$SettingName`TextBox"                 # SettingNameTextBox
                "$SettingName`TextBlock"               # SettingNameTextBlock
                "$SettingName`CheckBox"                # SettingNameCheckBox
                "$SettingName`ComboBox"                # SettingNameComboBox
                "$SettingName`Label"                   # SettingNameLabel
            )

            # Try each pattern
            foreach ($pattern in $namingPatterns) {
                if ($syncHash.$pattern) {
                    return $syncHash.$pattern
                }
            }

            return $null
        }

        # Helper function to update control value based on type
        Function Set-ControlValue {
            <#
            .SYNOPSIS
            Updates control values based on their type and handles focus preservation.
            .DESCRIPTION
            This function sets values on different types of WPF controls while preserving cursor position and preventing timer interference.
            #>
            param(
                [object]$Control,
                [object]$Value,
                [string]$SettingKey
            )

            switch ($Control.GetType().Name) {
                'TextBox' {
                    # Skip updating if user is currently typing in this control
                    if ($Control.IsFocused -and $Control.IsKeyboardFocused) {
                        return
                    }
                    $Control.Text = $Value
                    $Control.Foreground = [System.Windows.Media.Brushes]::Black
                    $Control.FontStyle = [System.Windows.FontStyles]::Normal
                }
                'TextBlock' {
                    # Skip updating if user is currently typing in this control
                    if ($Control.IsFocused -and $Control.IsKeyboardFocused) {
                        return
                    }
                    $Control.Text = $Value
                }
                'CheckBox' {
                    $Control.IsChecked = [bool]$Value
                }
                'ComboBox' {
                    Set-ComboBoxValue -ComboBox $Control -Value $Value -SettingKey $SettingKey
                }
                'Label' {
                    $Control.Content = $Value
                }
                'String' {
                    # Skip updating if user is currently typing in this control
                    if ($Control.IsFocused -and $Control.IsKeyboardFocused) {
                        return
                    }
                    $syncHash.$Control = $Value
                }
                default {
                    Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.SetControlWarning -f $SettingKey,$Control.GetType().Name) -Source $MyInvocation.MyCommand.Name -Level "Warning"
                }
            }
        }

        # Helper function to update ComboBox values
        Function Set-ComboBoxValue {
            <#
            .SYNOPSIS
            Updates ComboBox control selection to match a specified value.
            .DESCRIPTION
            This function sets the selected item in a ComboBox control by matching the provided value against available items.
            #>
            param(
                [System.Windows.Controls.ComboBox]$ComboBox,
                [object]$Value,
                [string]$SettingKey
            )

            # SPECIAL handling for M365Environment ComboBox
            if ($SettingKey -eq "M365Environment" -or $ComboBox.Name -eq "M365Environment_ComboBox") {
                # For M365Environment, we need to check both id and name values
                # First try to match by id (stored in Tag)
                $selectedItem = $ComboBox.Items | Where-Object { $_.Tag -eq $Value }

                # If not found by id, try to find by name in the configuration
                if (-not $selectedItem) {
                    $envConfig = $syncHash.UIConfigs.M365Environment | Where-Object { $_.name -eq $Value }
                    if ($envConfig) {
                        $selectedItem = $ComboBox.Items | Where-Object { $_.Tag -eq $envConfig.id }
                    }
                }

                if ($selectedItem) {
                    $ComboBox.SelectedItem = $selectedItem
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message "Selected M365Environment: $Value" -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                    return
                }
            }

            # Default ComboBox handling for other ComboBoxes
            # Try to find item by Tag first (common for environment selection)
            $selectedItem = $ComboBox.Items | Where-Object { $_.Tag -eq $Value }

            # If not found, try by Content
            if (-not $selectedItem) {
                $selectedItem = $ComboBox.Items | Where-Object { $_.Content -eq $Value }
            }

            # If still not found, try by string representation
            if (-not $selectedItem) {
                $selectedItem = $ComboBox.Items | Where-Object { $_.ToString() -eq $Value }
            }

            if ($selectedItem) {
                $ComboBox.SelectedItem = $selectedItem
                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.SetComboBoxInfo -f $SettingKey,$Value)  -Source $MyInvocation.MyCommand.Name -Level "Info"
            } else {
                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.SetComboBoxWarning -f $SettingKey,$Value) -Source $MyInvocation.MyCommand.Name -Level "Warning"
            }
        }

        # Function to validate UI field based on regex and required status
        Function Confirm-UIField {
            <#
            .SYNOPSIS
            Validates UI field values using regex patterns and required field checks.
            .DESCRIPTION
            This function performs validation on user input fields using regular expressions and enforces required field constraints.
            #>
            param(
                [System.Windows.Controls.Control]$UIElement,
                [string]$RegexPattern,
                [string]$ErrorMessage,
                [string]$PlaceholderText = "",
                [switch]$Required,
                [switch]$ShowMessageBox
            )

            $isValid = $true
            $currentValue = ""

            # Get the current value based on control type
            if ($UIElement -is [System.Windows.Controls.TextBox]) {
                $currentValue = $UIElement.Text
            } elseif ($UIElement -is [System.Windows.Controls.ComboBox]) {
                $currentValue = $UIElement.SelectedItem
            }

            # Check if field is required and empty/placeholder
            if ($Required -and ([string]::IsNullOrWhiteSpace($currentValue) -or $currentValue -eq $PlaceholderText)) {
                $isValid = $false
            }
            # Check regex pattern if provided and field has content
            elseif (![string]::IsNullOrWhiteSpace($RegexPattern) -and
                     ![string]::IsNullOrWhiteSpace($currentValue) -and
                     $currentValue -ne $PlaceholderText -and
                     -not ($currentValue -match $RegexPattern)) {
                $isValid = $false
            }

            # Apply visual feedback
            if ($UIElement -is [System.Windows.Controls.TextBox]) {
                if (-not $isValid) {
                    $UIElement.BorderBrush = [System.Windows.Media.Brushes]::Red
                    $UIElement.BorderThickness = "2"
                } else {
                    $UIElement.BorderBrush = [System.Windows.Media.Brushes]::Gray
                    $UIElement.BorderThickness = "1"
                }
            }

            # Show error message if requested
            if (-not $isValid -and $ShowMessageBox) {
                [System.Windows.MessageBox]::Show($ErrorMessage, "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            }

            return $isValid
        }

        Function Update-ProductNames {
            <#
            .SYNOPSIS
            Updates product name selections from UI checkboxes.
            .DESCRIPTION
            This function collects selected products from checkboxes and updates the GeneralSettings data structure accordingly.
            #>
            # Collect ProductNames from checked checkboxes and update GeneralSettings
            $selectedProducts = @()
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }
            foreach ($checkbox in $allProductCheckboxes) {
                if ($checkbox.IsChecked) {
                    $selectedProducts += $checkbox.Tag
                    Write-DebugOutput -Message "Checked checkbox: $($checkbox.Tag)" -Source $MyInvocation.MyCommand.Name -Level "Info"
                }
            }

            # Update the GeneralSettings
            if ($selectedProducts.Count -gt 0) {
                $syncHash.GeneralSettings["ProductNames"] = $selectedProducts
            } else {
                # Remove ProductNames if no products are selected
                $syncHash.GeneralSettings.Remove("ProductNames")
            }

            Write-DebugOutput -Message "Updated ProductNames: [$($selectedProducts -join ', ')]" -Source $MyInvocation.MyCommand.Name -Level "Info"
        }
        #===========================================================================
        # UPDATE UI FUNCTIONS
        #===========================================================================

        # Function to update all UI elements from data
        Function Update-AllUIFromData {
            <#
            .SYNOPSIS
            Updates all UI elements from the current data structures.
            .DESCRIPTION
            This function refreshes all UI components to reflect the current state of the configuration data.
            #>
            # Update general settings
            Update-GeneralSettingsFromData

            # Handle Product Name CheckBox
            Update-ProductNameCheckboxFromData

            # Update exclusions
            Update-ExclusionsFromData

            # Update annotations
            Update-AnnotationsFromData

            # Update omissions
            Update-OmissionsFromData
        }


        # Function to update general settings UI from data (Dynamic Version)
        Function Update-GeneralSettingsFromData {
            <#
            .SYNOPSIS
            Updates general settings UI controls from data.
            .DESCRIPTION
            This function populates general settings controls with values from the GeneralSettings data structure.
            #>
            if (-not $syncHash.GeneralSettings) { return }

            try {
                foreach ($settingKey in $syncHash.GeneralSettings.Keys) {
                    $settingValue = $syncHash.GeneralSettings[$settingKey]

                    # Skip if value is null or empty
                    if ($null -eq $settingValue) { return }

                    # Find the corresponding XAML control using various naming patterns
                    $control = Find-ControlBySettingName -SettingName $settingKey

                    if ($control) {
                        Set-ControlValue -Control $control -Value $settingValue -SettingKey $settingKey
                    } else {
                        If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UpdateGeneralSettingInfo -f $settingKey) -Source $MyInvocation.MyCommand.Name -Level "Warning"}
                    }
                }
            }
            catch {
                If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UpdateGeneralSettingError -f $_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Level "Error"}
            }
        }

        Function Update-ProductNameCheckboxFromData{
            <#
            .SYNOPSIS
            Updates product name checkbox controls from data.
            .DESCRIPTION
            This function sets the checked state of product name checkboxes based on the current configuration data.
            #>
            <#
            .SYNOPSIS
            Updates UI product checkboxes and ensures tabs/content are properly enabled and created
            #>
            param([string[]]$ProductNames = $null)

            # Get all product checkboxes
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }

            # Get all available product IDs
            $allProductIds = $syncHash.UIConfigs.products | Select-Object -ExpandProperty id

            # Determine which products to select
            $productsToSelect = @()

            if ($ProductNames) {
                if ($ProductNames -contains '*') {
                    $productsToSelect = $allProductIds
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UpdateProductNameDebug) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                } else {
                    $productsToSelect = $ProductNames
                }
            } elseif ($syncHash.GeneralSettings.ProductNames -and $syncHash.GeneralSettings.ProductNames.Count -gt 0) {
                $productsToSelect = $syncHash.GeneralSettings.ProductNames
            }

            # Set bulk update flag to prevent event cascades
            $syncHash.BulkUpdateInProgress = $true

            try {
                # First, uncheck all checkboxes and disable all tabs
                foreach ($checkbox in $allProductCheckboxes) {
                    $checkbox.IsChecked = $false
                    $productId = $checkbox.Tag

                    # Disable tabs for this product
                    $omissionTab = $syncHash.("$($productId)OmissionTab")
                    $annotationTab = $syncHash.("$($productId)AnnotationTab")
                    $exclusionTab = $syncHash.("$($productId)ExclusionTab")

                    if ($omissionTab) { $omissionTab.IsEnabled = $false }
                    if ($annotationTab) { $annotationTab.IsEnabled = $false }
                    if ($exclusionTab) { $exclusionTab.IsEnabled = $false }
                }

                # Disable main tabs if no products selected
                if ($productsToSelect.Count -eq 0) {
                    $syncHash.ExclusionsTab.IsEnabled = $false
                    $syncHash.AnnotatePolicyTab.IsEnabled = $false
                    $syncHash.OmitPolicyTab.IsEnabled = $false
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.DisableProductTabsDebug) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                } else {
                    # Enable main tabs when products are selected
                    $syncHash.ExclusionsTab.IsEnabled = $true
                    $syncHash.AnnotatePolicyTab.IsEnabled = $true
                    $syncHash.OmitPolicyTab.IsEnabled = $true
                }

                # Now check selected products and create their content
                foreach ($productId in $productsToSelect) {
                    $checkbox = $allProductCheckboxes | Where-Object { $_.Tag -eq $productId }
                    if ($checkbox) {
                        $checkbox.IsChecked = $true
                        $product = $syncHash.UIConfigs.products | Where-Object { $_.id -eq $productId }

                        # Enable and ensure content exists for omissions
                        $omissionTab = $syncHash.("$($productId)OmissionTab")
                        if ($omissionTab) {
                            $omissionTab.IsEnabled = $true
                            $container = $syncHash.("$($productId)OmissionContent")
                            if ($container -and $container.Children.Count -eq 0) {
                                New-ProductOmissions -ProductName $productId -Container $container
                                If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.CreatedOmissionContentDebug -f $productId) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                            }
                        }

                        # Enable and ensure content exists for annotations
                        $annotationTab = $syncHash.("$($productId)AnnotationTab")
                        if ($annotationTab) {
                            $annotationTab.IsEnabled = $true
                            $container = $syncHash.("$($productId)AnnotationContent")
                            if ($container -and $container.Children.Count -eq 0) {
                                New-ProductAnnotations -ProductName $productId -Container $container
                                If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.CreatedAnnotationContentDebug -f $productId) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                            }
                        }

                        # Enable and ensure content exists for exclusions (if supported)
                        if ($product -and $product.supportsExclusions) {
                            $exclusionTab = $syncHash.("$($productId)ExclusionTab")
                            if ($exclusionTab) {
                                $exclusionTab.IsEnabled = $true
                                $container = $syncHash.("$($productId)ExclusionContent")
                                if ($container -and $container.Children.Count -eq 0) {
                                    New-ProductExclusions -ProductName $productId -Container $container
                                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.CreatedExclusionContentDebug -f $productId) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                                }
                            }
                        }

                        If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.EnabledTabsAndContentDebug -f $productId) -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                    }
                }

                # Update GeneralSettings
                if ($productsToSelect.Count -gt 0) {
                    $syncHash.GeneralSettings["ProductNames"] = $productsToSelect
                } else {
                    $syncHash.GeneralSettings.Remove("ProductNames")
                }

            } finally {
                $syncHash.BulkUpdateInProgress = $false
            }

            If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UpdatedCheckboxesAndTabsDebug -f $productsToSelect -join ', ') -Source $MyInvocation.MyCommand.Name -Level "Debug"}
        }

        # Updated Update-ExclusionsFromData Function for hashtable structure
        Function Update-ExclusionsFromData {
            if (-not $syncHash.Exclusions) { return }

            # Iterate through products and policies in hashtable structure
            foreach ($productName in $syncHash.Exclusions.Keys) {
                foreach ($policyId in $syncHash.Exclusions[$productName].Keys) {
                    try {
                        # Find the exclusionField from the baseline config
                        $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                        if ($baseline -and $baseline.exclusionField -ne "none") {

                            # Find the existing card checkbox
                            $checkboxName = ($policyId.replace('.', '_') + "_ExclusionCheckbox")
                            $checkbox = $syncHash.$checkboxName

                            if ($checkbox) {
                                # Mark as checked
                                $checkbox.IsChecked = $true

                                # Get exclusion data for this policy
                                $exclusionData = $syncHash.Exclusions[$productName][$policyId]

                                # Iterate through exclusionFields (YAML key names)
                                foreach ($yamlKeyName in $exclusionData.Keys) {
                                    $fieldData = $exclusionData[$yamlKeyName]

                                    # Get the exclusionField configuration
                                    $FieldListConfig = $syncHash.UIConfigs.inputTypes.($baseline.exclusionField)

                                    if ($FieldListConfig) {
                                        # Populate the exclusion data fields based on exclusionField configuration
                                        foreach ($field in $FieldListConfig.fields) {
                                            $fieldName = $field.name
                                            $controlName = ($policyId.replace('.', '_') + "_" + $baseline.exclusionField + "_" + $fieldName)

                                            if ($fieldData.Keys -contains $fieldName) {
                                                $fieldValue = $fieldData[$fieldName]

                                                if ($field.type -eq "array" -and $fieldValue -is [array]) {
                                                    # Handle array fields
                                                    #to pass PSAvoidInvokingEmptyMembers
                                                    $listControl = ($controlName + "_List")

                                                    $listContainer = $syncHash.$listControl
                                                    if ($listContainer) {
                                                        # Clear existing items
                                                        $listContainer.Children.Clear()

                                                        # Add each array item
                                                        foreach ($item in $fieldValue) {
                                                            $itemPanel = New-Object System.Windows.Controls.StackPanel
                                                            $itemPanel.Orientation = "Horizontal"
                                                            $itemPanel.Margin = "0,2,0,2"

                                                            $itemText = New-Object System.Windows.Controls.TextBlock
                                                            $itemText.Text = $item
                                                            $itemText.VerticalAlignment = "Center"
                                                            $itemText.Margin = "0,0,8,0"

                                                            $removeBtn = New-Object System.Windows.Controls.Button
                                                            $removeBtn.Content = "Remove"
                                                            $removeBtn.Background = [System.Windows.Media.Brushes]::Red
                                                            $removeBtn.Foreground = [System.Windows.Media.Brushes]::White
                                                            $removeBtn.Width = 60
                                                            $removeBtn.Height = 20
                                                            $removeBtn.Add_Click({
                                                                $listContainer.Children.Remove($itemPanel)
                                                                Write-DebugOutput -Message "Item removed: $item" -Source $listContainer -Level "Info"
                                                            }.GetNewClosure())

                                                            [void]$itemPanel.Children.Add($itemText)
                                                            [void]$itemPanel.Children.Add($removeBtn)
                                                            [void]$listContainer.Children.Add($itemPanel)
                                                        }
                                                    }
                                                } else {
                                                    #to pass PSAvoidInvokingEmptyMembers
                                                    $TextboxControl = ($controlName + "_TextBox")

                                                    # Handle single value fields
                                                    $control = $syncHash.$TextboxControl
                                                    if ($control) {
                                                        $control.Text = $fieldValue
                                                        $control.Foreground = [System.Windows.Media.Brushes]::Black
                                                        $control.FontStyle = [System.Windows.FontStyles]::Normal
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                # Show remove button and make header bold
                                $removeButtonName = ($policyId.replace('.', '_') + "_RemoveExclusion")
                                $removeButton = $syncHash.$removeButtonName
                                if ($removeButton) {
                                    $removeButton.Visibility = "Visible"
                                }

                                # Make policy header bold
                                $policyHeaderName = ($policyId.replace('.', '_') + "_PolicyHeader")
                                if ($syncHash.$policyHeaderName) {
                                    $syncHash.$policyHeaderName.FontWeight = "Bold"
                                }
                            }
                        }
                    }
                    catch {
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ErrorUpdatingExclusionUI -f $policyId, $_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }
                }
            }
        }

        # Updated Update-OmissionsFromData Function
        Function Update-OmissionsFromData {
            <#
            .SYNOPSIS
            Updates omission controls from the Omissions data structure.
            .DESCRIPTION
            This function populates omission checkboxes, rationale fields, and expiration dates with values from the configuration data.
            #>
            if (-not $syncHash.Omissions) { return }

            # Iterate through products and policies in hashtable structure
            foreach ($productName in $syncHash.Omissions.Keys) {
                foreach ($policyId in $syncHash.Omissions[$productName].Keys) {
                    $omission = $syncHash.Omissions[$productName][$policyId]

                    try {
                        # Find the existing card controls
                        $checkboxName = ($policyId.replace('.', '_') + "_OmissionCheckbox")
                        $rationaleTextBoxName = ($policyId.replace('.', '_') + "_Rationale_TextBox")
                        $expirationTextBoxName = ($policyId.replace('.', '_') + "_Expiration_TextBox")
                        $removeButtonName = ($policyId.replace('.', '_') + "_RemoveOmission")

                        $checkbox = $syncHash.$checkboxName
                        $rationaleTextBox = $syncHash.$rationaleTextBoxName
                        $expirationTextBox = $syncHash.$expirationTextBoxName
                        $removeButton = $syncHash.$removeButtonName

                        if ($checkbox -and $rationaleTextBox) {
                            # Mark as checked (but don't expand details)
                            $checkbox.IsChecked = $true

                            # Populate rationale
                            $rationaleTextBox.Text = $omission.Rationale

                            # Populate expiration if exists
                            if ($expirationTextBox) {
                                if ($omission.Expiration) {
                                    $expirationTextBox.Text = $omission.Expiration
                                    $expirationTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                                    $expirationTextBox.FontStyle = [System.Windows.FontStyles]::Normal
                                }
                            }

                            # Show remove button
                            if ($removeButton) {
                                $removeButton.Visibility = "Visible"
                            }

                            # Make policy header bold
                            $policyHeaderName = ($policyId.replace('.', '_') + "_PolicyHeader")
                            if ($syncHash.$policyHeaderName) {
                                $syncHash.$policyHeaderName.FontWeight = "Bold"
                            }
                        }
                    }
                    catch {
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ErrorUpdatingOmissionUI -f $policyId, $_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }
                }
            }
        }

        # Updated Update-AnnotationsFromData Function
        Function Update-AnnotationsFromData {
            <#
            .SYNOPSIS
            Updates annotation controls from the Annotations data structure.
            .DESCRIPTION
            This function populates annotation checkboxes and comment fields with values from the configuration data.
            #>
            if (-not $syncHash.Annotations) { return }

            # Iterate through products and policies in hashtable structure
            foreach ($productName in $syncHash.Annotations.Keys) {
                foreach ($policyId in $syncHash.Annotations[$productName].Keys) {
                    $annotation = $syncHash.Annotations[$productName][$policyId]

                    try {
                        # Find the existing card controls
                        $checkboxName = ($policyId.replace('.', '_') + "_AnnotationCheckbox")
                        $commentTextBoxName = ($policyId.replace('.', '_') + "_Comment_TextBox")
                        $removeButtonName = ($policyId.replace('.', '_') + "_RemoveAnnotation")

                        $checkbox = $syncHash.$checkboxName
                        $commentTextBox = $syncHash.$commentTextBoxName
                        $removeButton = $syncHash.$removeButtonName

                        if ($checkbox -and $commentTextBox) {
                            # Mark as checked (but don't expand details)
                            $checkbox.IsChecked = $true

                            # Populate comment
                            $commentTextBox.Text = $annotation.Comment

                            # Show remove button
                            if ($removeButton) {
                                $removeButton.Visibility = "Visible"
                            }

                            # Make policy header bold
                            $policyHeaderName = ($policyId.replace('.', '_') + "_PolicyHeader")
                            if ($syncHash.$policyHeaderName) {
                                $syncHash.$policyHeaderName.FontWeight = "Bold"
                            }
                        }
                    }
                    catch {
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.ErrorUpdatingAnnotationUI -f $policyId, $_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }
                }
            }
        }

        # Function to collect ProductNames from UI checkboxes and update GeneralSettings
        Function Update-ProductNames {
            <#
            .SYNOPSIS
            Collects checked product checkboxes from UI and updates $syncHash.GeneralSettings.ProductNames

            .DESCRIPTION
            This function scans the UI for checked product checkboxes and updates the GeneralSettings
            with the actual list of selected products. This is used for UI-to-data synchronization.
            #>

            # Collect ProductNames from checked checkboxes
            $selectedProducts = @()
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }

            foreach ($checkbox in $allProductCheckboxes) {
                if ($checkbox.IsChecked) {
                    $selectedProducts += $checkbox.Tag
                    Write-DebugOutput -Message "Checked checkbox: $($checkbox.Tag)" -Source $MyInvocation.MyCommand.Name -Level "Info"
                }
            }

            # Update the GeneralSettings with actual product list
            if ($selectedProducts.Count -gt 0) {
                $syncHash.GeneralSettings["ProductNames"] = $selectedProducts
            } else {
                # Remove ProductNames if no products are selected
                $syncHash.GeneralSettings.Remove("ProductNames")
            }

            Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.UpdatedProductNamesDebug -f ($selectedProducts -join ', ')) -Source $MyInvocation.MyCommand.Name -Level "Info"
        }



        # Function to get ProductNames formatted for YAML output
        Function Get-ProductNamesForYaml {
            <#
            .SYNOPSIS
            Returns ProductNames formatted appropriately for YAML output

            .DESCRIPTION
            This function determines the correct format for ProductNames in YAML output.
            If all available products are selected, returns ['*'].
            Otherwise, returns the actual list of selected products.
            #>

            # Check if we have any selected products
            if (-not $syncHash.GeneralSettings.ProductNames -or $syncHash.GeneralSettings.ProductNames.Count -eq 0) {
                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.YamlImportNoProductSelectedWarning) -Source $MyInvocation.MyCommand.Name -Level "Warning"
                return @()
            }

            # Get all available product IDs
            $allProductIds = $syncHash.UIConfigs.products | Select-Object -ExpandProperty id

            # Check if all products are selected
            $selectedProducts = $syncHash.GeneralSettings.ProductNames | Sort-Object
            $availableProducts = $allProductIds | Sort-Object
            $isAllProductsSelected = ($selectedProducts.Count -eq $availableProducts.Count) -and
                                   (-not (Compare-Object $selectedProducts $availableProducts))

            if ($isAllProductsSelected) {
                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.YamlAllProductsSelectedInfo) -Source $MyInvocation.MyCommand.Name -Level "Info"
                return "`nProductNames: ['*']"
            } else {
                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleDebugOutput.YamlSpecificProductListInfo -f ($selectedProducts -join ', ')) -Source $MyInvocation.MyCommand.Name -Level "Info"
                Return ("`nProductNames: " + ($selectedProducts | ForEach-Object { "`n  - $_" }) -join '')
            }
        }

        # Function to populate cards for product policies
        Function New-ProductPolicyCards {
            <#
            .SYNOPSIS
            Creates policy cards for exclusions, omissions, or annotations.
            .DESCRIPTION
            This function generates UI cards for each policy baseline of a product, allowing users to configure exclusions, omissions, or annotations.
            #>
            param(
                [string]$ProductName,
                [System.Windows.Controls.StackPanel]$Container,
                [string]$ControlType  # "Exclusions", "Omissions", or "Annotations"
            )

            $Container.Children.Clear()

            # Get the baseline control configuration for this type
            $baselineControl = $syncHash.UIConfigs.baselineControls | Where-Object { $_.dataControlOutput -eq $ControlType }

            if (-not $baselineControl) {
                Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.NoBaselineControlConfigFound -f $ControlType) -Source $MyInvocation.MyCommand.Name -Level "Warning"
                return
            }

            # Get baselines for this product
            $baselines = $syncHash.UIConfigs.baselines.$ProductName

            if ($baselines -and $baselines.Count -gt 0) {
                # Filter baselines based on the control type
                $filteredBaselines = switch ($ControlType) {
                    "Exclusions" {
                        $baselines | Where-Object { $_.$($baselineControl.fieldControlName) -ne 'none' }
                    }
                    default {
                        $baselines  # Omissions and Annotations use all baselines
                    }
                }

                if ($filteredBaselines -and $filteredBaselines.Count -gt 0) {
                    foreach ($baseline in $filteredBaselines) {
                        # Get the field list for this baseline
                        $fieldList = if ($baseline.PSObject.Properties.Name -contains $baselineControl.fieldControlName) {
                            $baseline.$($baselineControl.fieldControlName)
                        } else {
                            $baselineControl.defaultFields
                        }

                        # Get the output data hashtable dynamically
                        $outputData = $syncHash.$($baselineControl.dataControlOutput)

                        $card = New-FieldListCard `
                            -CardName $baselineControl.tabName.Replace(' Policies', '').Replace('s', '') `
                            -PolicyId $baseline.id `
                            -ProductName $ProductName `
                            -PolicyName $baseline.name `
                            -PolicyDescription $baseline.rationale `
                            -FieldList $fieldList `
                            -OutputData $outputData `
                            -ShowFieldType:$baselineControl.showFieldType `
                            -ShowDescription:$baselineControl.showDescription

                        if ($card) {
                            [void]$Container.Children.Add($card)
                        }
                    }
                } else {
                    # No applicable baselines
                    $noDataText = New-Object System.Windows.Controls.TextBlock
                    $noDataText.Text = $syncHash.UIConfigs.LocaleInfoMessages.NoPoliciesAvailable
                    $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                    $noDataText.FontStyle = "Italic"
                    $noDataText.HorizontalAlignment = "Center"
                    $noDataText.Margin = "0,50,0,0"
                    [void]$Container.Children.Add($noDataText)
                }
            } else {
                # No baselines available for this product
                $noDataText = New-Object System.Windows.Controls.TextBlock
                $noDataText.Text = $syncHash.UIConfigs.LocaleInfoMessages.NoPoliciesAvailable
                $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $noDataText.FontStyle = "Italic"
                $noDataText.HorizontalAlignment = "Center"
                $noDataText.Margin = "0,50,0,0"
                [void]$Container.Children.Add($noDataText)
            }
        }
        #===========================================================================
        # Dynamic controls - CARDS
        #===========================================================================
        # Function to validate required fields

        Function Test-RequiredField {
            <#
            .SYNOPSIS
            Validates required fields in a policy card details panel.
            .DESCRIPTION
            This function checks that all required fields in a policy configuration card have valid values before allowing the configuration to be saved.
            #>
            param(
                [System.Windows.Controls.StackPanel]$detailsPanel,
                [array]$validInputFields,
                [string]$policyId,
                [string]$CardName
            )

            $missingRequiredFields = @()

            # Function to search recursively for controls
            function Find-ControlByName {
                <#
                .SYNOPSIS
                Recursively searches for a control by name within a parent container.
                .DESCRIPTION
                This nested function traverses the visual tree to locate a control with a specific name.
                #>
                param($parent, $targetName)

                if ($parent.Name -eq $targetName) {
                    return $parent
                }

                if ($parent.Children) {
                    foreach ($child in $parent.Children) {
                        $result = Find-ControlByName -parent $child -targetName $targetName
                        if ($result) { return $result }
                    }
                }

                if ($parent.Content -and $parent.Content.Children) {
                    foreach ($child in $parent.Content.Children) {
                        $result = Find-ControlByName -parent $child -targetName $targetName
                        if ($result) { return $result }
                    }
                }

                if ($parent.Items) {
                    foreach ($item in $parent.Items) {
                        if ($item.Content -and $item.Content.Children) {
                            foreach ($child in $item.Content.Children) {
                                $result = Find-ControlByName -parent $child -targetName $targetName
                                if ($result) { return $result }
                            }
                        }
                    }
                }

                return $null
            }

            # Build dynamic placeholders list based on the fields we're actually validating
            $dynamicPlaceholders = @("Enter value", "No date selected")  # Keep basic fallbacks

            foreach ($inputData in $validInputFields) {
                $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                if (-not $FieldListDef) { continue }

                foreach ($field in $FieldListDef.fields) {
                    # Add the specific placeholder for this field's valueType
                    if ($syncHash.UIConfigs.valueValidations.($field.valueType)) {
                        $fieldPlaceholder = $syncHash.UIConfigs.valueValidations.($field.valueType).sample
                        if ($dynamicPlaceholders -notcontains $fieldPlaceholder) {
                            $dynamicPlaceholders += $fieldPlaceholder
                        }
                    }
                }
            }

            Write-DebugOutput -Message ($syncHash.UIConfigs.localeVerboseMessages.DynamicPlaceholders -f ($dynamicPlaceholders -join ', ')) -Source "Test-RequiredField" -Level "Verbose"

            foreach ($inputData in $validInputFields) {
                $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                if (-not $FieldListDef) { continue }

                foreach ($field in $FieldListDef.fields) {
                    # Skip if field is not required
                    if (-not $field.required) { continue }

                    $controlFieldName = ($policyId.replace('.', '_') + "_" + $CardName + "_" + $field.value)
                    $hasValue = $false

                    Write-DebugOutput -Message ($syncHash.UIConfigs.localeVerboseMessages.RequiredFieldCheck -f $field.name, $controlFieldName) -Source "Test-RequiredField" -Level "Verbose"

                    if ($field.type -eq "array") {
                        # For arrays, check if list container has any items
                        $listContainerName = ($controlFieldName + "_List")
                        $listContainer = Find-ControlByName -parent $detailsPanel -targetName $listContainerName

                        if ($listContainer -and $listContainer.Children.Count -gt 0) {
                            $hasValue = $true
                        }

                    } elseif ($field.type -eq "boolean") {
                        # Boolean fields always have a value (true or false)
                        $hasValue = $true

                    } elseif ($field.type -eq "dateString" -and $field.valueType -eq "yearmonthday") {
                        # Check DatePicker for date fields
                        $datePickerName = ($controlFieldName + "_DatePicker")
                        $datePicker = Find-ControlByName -parent $detailsPanel -targetName $datePickerName

                        if ($datePicker -and $datePicker.SelectedDate) {
                            $hasValue = $true
                        }

                    } else {
                        # For all other string-based fields, check TextBox
                        $stringFieldName = ($controlFieldName + "_TextBox")
                        $stringTextBox = Find-ControlByName -parent $detailsPanel -targetName $stringFieldName

                        if ($stringTextBox -and ![string]::IsNullOrWhiteSpace($stringTextBox.Text)) {
                            # Check if it's not a placeholder text using our dynamic list
                            $currentText = $stringTextBox.Text.Trim()
                            if ($dynamicPlaceholders -notcontains $currentText) {
                                $hasValue = $true
                            }
                        }
                    }

                    # If required field doesn't have a value, add to missing list
                    if (-not $hasValue) {
                        $missingRequiredFields += $field.name
                        Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.RequiredFieldMissing -f $field.name) -Source "Test-RequiredField" -Level "Warning"
                    } else {
                        Write-DebugOutput -Message ($syncHash.UIConfigs.localeVerboseMessages.RequiredFieldHasValue -f $field.name) -Source "Test-RequiredField" -Level "Verbose"
                    }
                }
            }

            return $missingRequiredFields
        }
        Function New-FieldListControl {
            <#
            .SYNOPSIS
            Creates a dynamic field list control for policy configurations.
            .DESCRIPTION
            This function generates interactive UI controls for managing lists of field values in policy configurations.
            #>
            param(
                [string]$ControlName,
                [string]$PolicyId,
                [object]$Field,
                [System.Windows.Controls.StackPanel]$Container
            )

            $fieldPanel = New-Object System.Windows.Controls.StackPanel
            $fieldPanel.Margin = "0,0,0,12"

            # Field label - USE field.name for UI display
            $fieldLabel = New-Object System.Windows.Controls.TextBlock
            $fieldLabel.Text = $Field.name
            $fieldLabel.FontWeight = "SemiBold"
            $fieldLabel.Margin = "0,0,0,4"
            [void]$fieldPanel.Children.Add($fieldLabel)

            # Field description
            $fieldDesc = New-Object System.Windows.Controls.TextBlock
            $fieldDesc.Text = $Field.description
            $fieldDesc.FontSize = 11
            $fieldDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
            $fieldDesc.Margin = "0,0,0,4"
            $fieldDesc.TextWrapping = "Wrap"
            [void]$fieldPanel.Children.Add($fieldDesc)

            # Use field.value for control naming (for data storage)
            $fieldName = ($PolicyId.replace('.', '_') + "_" + $ControlName + "_" + $Field.value)

            if ($Field.type -eq "array") {
                # Create array input with add/remove functionality
                $arrayContainer = New-Object System.Windows.Controls.StackPanel
                $arrayContainer.Name = $fieldName + "_Container"

                # Input row for new entries
                $inputRow = New-Object System.Windows.Controls.StackPanel
                $inputRow.Orientation = "Horizontal"
                $inputRow.Margin = "0,0,0,8"

                $inputTextBox = New-Object System.Windows.Controls.TextBox
                $inputTextBox.Name = $fieldName + "_TextBox"
                $inputTextBox.Width = 250
                $inputTextBox.Height = 28
                $inputTextBox.VerticalContentAlignment = "Center"
                $inputTextBox.Margin = "0,0,8,0"
                $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $inputTextBox.FontStyle = "Italic"

                # Set placeholder text based on field type using valueValidations
                $placeholderText = if ($syncHash.UIConfigs.valueValidations.($Field.valueType)) {
                    $syncHash.UIConfigs.valueValidations.($Field.valueType).sample
                } else {
                    "Enter value"  # fallback if valueType not found
                }

                $inputTextBox.Text = $placeholderText

                # Add global event handlers to dynamically created inputTextBox
                Add-ControlEventHandler -Control $inputTextBox

                # placeholder functionality - capture placeholder in closure properly
                $inputTextBox.Add_GotFocus({
                    #param($sender, $e)
                    if ($this.Text -eq $placeholderText) {
                        $this.Text = ""
                        $this.Foreground = [System.Windows.Media.Brushes]::Black
                        $this.FontStyle = "Normal"
                    }
                }.GetNewClosure())

                $inputTextBox.Add_LostFocus({
                    #param($sender, $e)
                    if ([string]::IsNullOrWhiteSpace($this.Text)) {
                        $this.Text = $placeholderText
                        $this.Foreground = [System.Windows.Media.Brushes]::Gray
                        $this.FontStyle = "Italic"
                    }
                }.GetNewClosure())

                $addButton = New-Object System.Windows.Controls.Button
                $addButton.Content = "Add"
                $addButton.Name = $fieldName + "_Add"
                $addButton.Style = $syncHash.Window.FindResource("PrimaryButton")
                $addButton.Width = 60
                $addButton.Height = 28

                [void]$inputRow.Children.Add($inputTextBox)
                [void]$inputRow.Children.Add($addButton)
                [void]$arrayContainer.Children.Add($inputRow)

                # List container for added items
                $listContainer = New-Object System.Windows.Controls.StackPanel
                $listContainer.Name = $fieldName + "_List"
                [void]$arrayContainer.Children.Add($listContainer)

                Add-ControlEventHandler -Control $addButton

                # add button functionality - capture placeholder properly
                $addButton.Add_Click({
                    $inputBox = $this.Parent.Children[0]
                    $listPanel = $this.Parent.Parent.Children[1]

                    if (![string]::IsNullOrWhiteSpace($inputBox.Text) -and $inputBox.Text -ne $placeholderText) {
                        # Trim the input value
                        $trimmedValue = $inputBox.Text.Trim()

                        # Check if value already exists
                        if ($listContainer.Children.Children | Where-Object { $_.Text -contains $trimmedValue }) {
                            [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localePopupMessages.DuplicateEntry, "Duplicate Entry", "OK", "Warning")
                            return
                        }

                        # Validate input based on valueType using valueValidations
                        $isValid = $true
                        $errorMessage = ""

                        if ($syncHash.UIConfigs.valueValidations.($Field.valueType)) {
                            $validation = $syncHash.UIConfigs.valueValidations.($Field.valueType)
                            if ($validation.pattern) {
                                $isValid = $trimmedValue -match $validation.pattern
                                $errorMessage = ($syncHash.UIConfigs.localeErrorMessages.InvalidFieldFormat -f $Field.valueType, $validation.sample)
                            }
                        } else {
                            # Fallback for unknown types
                            $isValid = $trimmedValue.Length -gt 0
                            $errorMessage = $syncHash.UIConfigs.localeErrorMessages.InvalidFieldValue
                        }

                        if ($isValid) {
                            # Create item row
                            $itemRow = New-Object System.Windows.Controls.StackPanel
                            $itemRow.Orientation = "Horizontal"
                            $itemRow.Margin = "0,2,0,2"

                            $itemText = New-Object System.Windows.Controls.TextBlock
                            $itemText.Text = $trimmedValue
                            $itemText.VerticalAlignment = "Center"
                            $itemText.Width = 250
                            $itemText.Margin = "0,0,8,0"

                            $removeButton = New-Object System.Windows.Controls.Button
                            $removeButton.Content = "Remove"
                            $removeButton.Width = 60
                            $removeButton.Height = 22
                            $removeButton.Background = [System.Windows.Media.Brushes]::Red
                            $removeButton.Foreground = [System.Windows.Media.Brushes]::White
                            $removeButton.BorderThickness = "0"
                            $removeButton.FontSize = 10

                            $removeButton.Add_Click({
                                $this.Parent.Parent.Children.Remove($this.Parent)
                            }.GetNewClosure())

                            [void]$itemRow.Children.Add($itemText)
                            [void]$itemRow.Children.Add($removeButton)
                            [void]$listPanel.Children.Add($itemRow)

                            # Clear input - use captured placeholder
                            $inputBox.Text = $placeholderText
                            $inputBox.Foreground = [System.Windows.Media.Brushes]::Gray
                            $inputBox.FontStyle = "Italic"
                        } else {
                            [System.Windows.MessageBox]::Show($errorMessage, "Validation Error", "OK", "Warning")
                        }
                    }
                }.GetNewClosure())

                # Add input row and list container to field panel
                [void]$fieldPanel.Children.Add($arrayContainer)

            } elseif ($Field.type -eq "boolean") {
                # Create boolean checkbox control
                $booleanCheckBox = New-Object System.Windows.Controls.CheckBox
                $booleanCheckBox.Name = $fieldName + "_CheckBox"
                $booleanCheckBox.Content = "Enable this setting"
                $booleanCheckBox.Margin = "0,4,0,0"
                $booleanCheckBox.IsChecked = $false

                # Add global event handlers to dynamically created checkbox
                Add-ControlEventHandler -Control $booleanCheckBox

                [void]$fieldPanel.Children.Add($booleanCheckBox)

            } elseif ($Field.type -match "string") {

                switch($Field.type){
                    "string" {
                        $stringTextBox = New-Object System.Windows.Controls.TextBox
                        $stringTextBox.Name = $fieldName + "_TextBox"
                        $stringTextBox.HorizontalAlignment = "Left"
                        $stringTextBox.Width = 400
                        $stringTextBox.Height = 28
                        $stringTextBox.VerticalContentAlignment = "Center"
                    }
                    "longstring" {
                        $stringTextBox = New-Object System.Windows.Controls.TextBox
                        $stringTextBox.Name = $fieldName + "_TextBox"
                        $stringTextBox.HorizontalAlignment = "Left"
                        $stringTextBox.Width = 500
                        $stringTextBox.Height = 80
                        $stringTextBox.AcceptsReturn = $true
                        $stringTextBox.TextWrapping = "Wrap"
                        $stringTextBox.VerticalScrollBarVisibility = "Auto"
                        $stringTextBox.VerticalContentAlignment = "Top"
                        $stringTextBox.Margin = "0,0,0,8"
                    }
                    "dateString" {
                        # Set placeholder text for string fields using valueValidations
                        $stringPlaceholder = if ($syncHash.UIConfigs.valueValidations.($Field.valueType)) {
                            $syncHash.UIConfigs.valueValidations.($Field.valueType).sample
                        } else {
                            "Enter value"  # fallback
                        }

                        # Create a horizontal stack panel for date picker and text box
                        $datePanel = New-Object System.Windows.Controls.StackPanel
                        $datePanel.Orientation = "Horizontal"
                        $datePanel.HorizontalAlignment = "Left"

                        # Create DatePicker
                        $datePicker = New-Object System.Windows.Controls.DatePicker
                        $datePicker.Name = $fieldName + "_DatePicker"
                        $datePicker.Width = 150
                        $datePicker.Height = 28
                        $datePicker.Margin = "0,0,8,0"
                        $datePicker.SelectedDateFormat = "Short"

                        # Add global event handlers to dynamically created DatePicker
                        Add-ControlEventHandler -Control $datePicker

                        # Create a "Clear" button next to the DatePicker
                        $clearButton = New-Object System.Windows.Controls.Button
                        $clearButton.Content = "Clear"
                        $clearButton.Name = $fieldName + "_ClearDate"
                        $clearButton.Width = 50
                        $clearButton.Height = 28
                        $clearButton.Margin = "0,0,8,0"

                        # Add global event handlers to dynamically created clear button
                        Add-ControlEventHandler -Control $clearButton

                        # Clear button functionality
                        $clearButton.Add_Click({
                            # Find the DatePicker (previous sibling)
                            $parentPanel = $this.Parent
                            $datePicker = $parentPanel.Children[0]
                            $datePicker.SelectedDate = $null
                        }.GetNewClosure())

                        # Optional: Add a readonly TextBox to show the formatted date
                        $stringTextBox = New-Object System.Windows.Controls.TextBox
                        $stringTextBox.Name = $fieldName + "_TextBox"
                        $stringTextBox.Width = 120
                        $stringTextBox.Height = 28
                        $stringTextBox.IsReadOnly = $true
                        $stringTextBox.Text = "No date selected"
                        $stringTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                        $stringTextBox.FontStyle = "Italic"

                        # Update text box when date picker changes
                        $datePicker.Add_SelectedDateChanged({
                            $parentPanel = $this.Parent
                            $textBox = $parentPanel.Children | Where-Object { $_.Name -like "*_TextBox" }
                            if ($this.SelectedDate) {
                                # Format as yyyy-MM-dd for YAML output
                                $formattedDate = $this.SelectedDate.ToString($syncHash.UIConfigs.valueValidations.($Field.valueType).format)
                                $textBox.Text = $formattedDate
                                $textBox.Foreground = [System.Windows.Media.Brushes]::Black
                                $textBox.FontStyle = "Normal"
                            } else {
                                $textBox.Text = "No date selected"
                                $textBox.Foreground = [System.Windows.Media.Brushes]::Gray
                                $textBox.FontStyle = "Italic"
                            }
                        }.GetNewClosure())

                        # Add controls to the date panel
                        [void]$datePanel.Children.Add($datePicker)
                        [void]$datePanel.Children.Add($clearButton)
                        [void]$datePanel.Children.Add($stringTextBox)

                        # Add the date panel to the field panel instead of individual textbox
                        [void]$fieldPanel.Children.Add($datePanel)

                        # Skip the normal textbox creation for dateString
                        continue
                    }
                }

                # Only add the regular textbox if we didn't create a date panel
                if ($Field.type -ne "dateString") {
                    # placeholder text for string fields using valueValidations
                    $stringPlaceholder = if ($syncHash.UIConfigs.valueValidations.($Field.valueType)) {
                        $syncHash.UIConfigs.valueValidations.($Field.valueType).sample
                    } else {
                        "Enter value"  # fallback if valueType not found in config
                    }

                    $stringTextBox.Text = $stringPlaceholder
                    $stringTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                    $stringTextBox.FontStyle = "Italic"

                    # string field placeholder functionality
                    $stringTextBox.Add_GotFocus({
                        #param($sender, $e)
                        if ($this.Text -eq $stringPlaceholder) {
                            $this.Text = ""
                            $this.Foreground = [System.Windows.Media.Brushes]::Black
                            $this.FontStyle = "Normal"
                        }
                    }.GetNewClosure())

                    $stringTextBox.Add_LostFocus({
                        #param($sender, $e)
                        if ([string]::IsNullOrWhiteSpace($this.Text)) {
                            $this.Text = $stringPlaceholder
                            $this.Foreground = [System.Windows.Media.Brushes]::Gray
                            $this.FontStyle = "Italic"
                        }
                    }.GetNewClosure())

                    # Add global event handlers to dynamically created stringTextBox
                    Add-ControlEventHandler -Control $stringTextBox
                    [void]$fieldPanel.Children.Add($stringTextBox)
                }
            }

            # Add Graph connectivity buttons for those that match the field value
            if ($syncHash.GraphConnected -and ($Field.value -in $syncHash.UIConfigs.graphQueries.Psobject.properties.name) ) {
                $GraphQueryData = ($syncHash.UIConfigs.graphQueries.PSObject.Properties | Where-Object { $_.Name -eq $Field.value }).Value

                $graphGetButton = New-Object System.Windows.Controls.Button
                $graphGetButton.Content = "Get $($GraphQueryData.Name)"
                $graphGetButton.Width = 100
                $graphGetButton.Height = 28
                $graphGetButton.Margin = "8,0,0,0"

                # Add global event handlers to dynamically created graphGetButton
                Add-ControlEventHandler -Control $graphGetButton

                $graphGetButton.Add_Click({
                try {
                    # Get search term from input box
                    $searchTerm = if ($inputTextBox.Text -ne $placeholderText -and ![string]::IsNullOrWhiteSpace($inputTextBox.Text)) { $inputTextBox.Text } else { "" }

                    $selectedItems = Show-GraphSelector -GraphEntityType $GraphQueryData.Name -SearchTerm $searchTerm

                    # More robust check for valid results
                    if ($selectedItems -and $selectedItems.Count -gt 0 -and $null -ne $selectedItems[0]) {
                        # Add selected users to the list
                        foreach ($item in $selectedItems) {
                            # Skip if item is null or empty
                            if (-not $item -or [string]::IsNullOrWhiteSpace($item.($GraphQueryData.outProperty))) {
                                continue
                            }

                            # Check if user already exists in the list
                            if ($listContainer.Children.Children | Where-Object { $_.Text -contains $item.Id }) {
                                continue
                            }

                            # Create item panel
                            $itemPanel = New-Object System.Windows.Controls.StackPanel
                            $itemPanel.Orientation = "Horizontal"
                            $itemPanel.Margin = "0,2,0,2"

                            # Create item text block
                            $itemText = New-Object System.Windows.Controls.TextBlock
                            $itemText.Text = "$($item.($GraphQueryData.outProperty))"
                            $itemText.Width = 250
                            $itemText.VerticalAlignment = "Center"
                            $itemText.ToolTip = "$($item.($GraphQueryData.tipProperty))"

                            # Create remove button
                            $removeUserButton = New-Object System.Windows.Controls.Button
                            $removeUserButton.Content = "Remove"
                            $removeUserButton.Width = 60
                            $removeUserButton.Height = 20
                            $removeUserButton.Margin = "8,0,0,0"
                            $removeUserButton.FontSize = 10
                            $removeUserButton.Background = [System.Windows.Media.Brushes]::Red
                            $removeUserButton.Foreground = [System.Windows.Media.Brushes]::White
                            $removeUserButton.BorderThickness = "0"

                            $removeUserButton.Add_Click({
                                $parentItem = $this.Parent
                                $parentContainer = $parentItem.Parent
                                $parentContainer.Children.Remove($parentItem)
                            }.GetNewClosure())

                            [void]$itemPanel.Children.Add($itemText)
                            [void]$itemPanel.Children.Add($removeUserButton)
                            [void]$listContainer.Children.Add($itemPanel)
                        }

                        # Clear the input box
                        $inputTextBox.Text = $placeholderText
                        $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                        $inputTextBox.FontStyle = "Italic"
                    } else {
                        Write-DebugOutput -Message "No items selected or found from Graph query for $($GraphQueryData.Name)" -Source "Graph Button Click" -Level "Info"
                    }
                }
                catch {
                    Write-DebugOutput -Message "Error in Graph button click: $($_.Exception.Message)" -Source "Graph Button Click" -Level "Error"
                    [System.Windows.MessageBox]::Show("Error retrieving data from Graph: $($_.Exception.Message)", "Graph Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }.GetNewClosure())

                [void]$inputRow.Children.Add($graphGetButton)
            }
            [void]$Container.Children.Add($fieldPanel)
        }

        # Updated Function to create a field card UI element that handles multiple fields
        Function New-FieldListCard {
            <#
            .SYNOPSIS
            Creates a comprehensive field card UI element for policy configuration.
            .DESCRIPTION
            This function generates a complete card interface with checkboxes, input fields, tabs, and buttons for configuring multiple field types within policy settings including exclusions, omissions, and annotations.
            #>
            param(
                [string]$CardName,
                [string]$PolicyId,
                [string]$ProductName,
                [string]$PolicyName,
                [string]$PolicyDescription,
                [string[]]$FieldList,  # Can be string or array
                $OutputData,
                [switch]$ShowFieldType,
                [switch]$ShowDescription
            )

            # Handle both string and array inputs for FieldLists
            $inputFields = @()
            if ($FieldList -is [array]) {
                $inputFields = $FieldList
            } else {
                $inputFields = @($FieldList)
            }

            # Skip if inputField is "none" or empty
            if ($inputFields -contains "none" -or $inputFields.Count -eq 0) {
                return $null
            }

            # Validate all inputFields exist
            $validInputFields = @()
            foreach ($inputData in $inputFields) {
                $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                if ($FieldListDef) {
                    $validInputFields += $inputData
                } else {
                    Write-DebugOutput -Message ($syncHash.UIConfigs.localeErrorMessages.InputDataNotFound -f $inputData) -Source "New-FieldListCard" -Level "Verbose"
                }
            }

            # Return null if no valid inputFields found
            if ($validInputFields.Count -eq 0) {
                return $null
            }

            # Create the main card border
            $card = New-Object System.Windows.Controls.Border
            $card.Style = $syncHash.Window.FindResource("Card")
            $card.Margin = "0,0,0,12"

            # Create main grid for the card
            $cardGrid = New-Object System.Windows.Controls.Grid
            [void]$cardGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))
            [void]$cardGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))

            # Create header with checkbox and policy info
            $headerGrid = New-Object System.Windows.Controls.Grid
            [void]$headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))
            [void]$headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))
            [System.Windows.Controls.Grid]::SetRow($headerGrid, 0)

            # Create checkbox
            $checkbox = New-Object System.Windows.Controls.CheckBox
            $checkbox.Name = ($PolicyId.replace('.', '_') + "_" + $CardName + "_FieldListCheckbox")
            $checkbox.VerticalAlignment = "Top"
            $checkbox.Margin = "0,0,12,0"
            [System.Windows.Controls.Grid]::SetColumn($checkbox, 0)

            # Add global event handlers to dynamically created checkbox
            Add-ControlEventHandler -Control $checkbox

            # Create policy info stack panel
            $policyInfoStack = New-Object System.Windows.Controls.StackPanel
            [System.Windows.Controls.Grid]::SetColumn($policyInfoStack, 1)

            # Policy ID and name
            $policyHeader = New-Object System.Windows.Controls.TextBlock
            $policyHeader.Text = "$PolicyId`: $PolicyName"
            $policyHeader.FontWeight = "SemiBold"
            $policyHeader.Foreground = $syncHash.Window.FindResource("PrimaryBrush")
            $policyHeader.TextWrapping = "Wrap"
            $policyHeader.Margin = "0,0,0,4"
            [void]$policyInfoStack.Children.Add($policyHeader)

            # Add cursor and click handler to policy header
            $policyHeader.Cursor = [System.Windows.Input.Cursors]::Hand
            $policyHeader.Add_MouseLeftButtonDown({
                # Navigate to checkbox: this -> policyInfoStack -> headerGrid -> checkbox (first child)
                $headerGrid = $this.Parent.Parent
                $checkbox = $headerGrid.Children[0]
                $checkbox.IsChecked = -not $checkbox.IsChecked
            }.GetNewClosure())

            If($ShowDescription){
                # Policy description
                $policyDesc = New-Object System.Windows.Controls.TextBlock
                $policyDesc.Text = $PolicyDescription
                $policyDesc.FontSize = 11
                $policyDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $policyDesc.TextWrapping = "Wrap"
                [void]$policyInfoStack.Children.Add($policyDesc)
            }

            If($ShowFieldType){
                 # Field types info (show all available fields)
                $FieldLists = @()
                foreach ($inputData in $validInputFields) {
                    $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                    if ($FieldListDef) {
                        $FieldLists += $FieldListDef.name
                    }
                }
                $FieldListHeader = New-Object System.Windows.Controls.TextBlock
                $FieldListHeader.Text = "${CardName}: $($FieldLists -join ', ')"
                $FieldListHeader.FontSize = 12
                $FieldListHeader.Foreground = $syncHash.Window.FindResource("AccentBrush")
                $FieldListHeader.Margin = "0,0,0,4"
                [void]$policyInfoStack.Children.Add($FieldListHeader)
            }

            # Add elements to header grid
            [void]$headerGrid.Children.Add($checkbox)
            [void]$headerGrid.Children.Add($policyInfoStack)

            # Create details panel (initially collapsed)
            $detailsPanel = New-Object System.Windows.Controls.StackPanel
            $detailsPanel.Visibility = "Collapsed"
            $detailsPanel.Margin = "24,12,0,0"
            [System.Windows.Controls.Grid]::SetRow($detailsPanel, 1)

            # Create tab control for multiple inputFields if more than one
            if ($validInputFields.Count -gt 1) {
                $tabControl = New-Object System.Windows.Controls.TabControl
                $tabControl.Margin = "0,0,0,16"

                foreach ($inputData in $validInputFields) {
                    $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                    if ($FieldListDef) {
                        # Create tab item
                        $tabItem = New-Object System.Windows.Controls.TabItem
                        $tabItem.Header = $FieldListDef.name

                        # Create tab content
                        $tabContent = New-Object System.Windows.Controls.StackPanel
                        $tabContent.Margin = "8"

                        # Add description
                        $fieldDesc = New-Object System.Windows.Controls.TextBlock
                        $fieldDesc.Text = $FieldListDef.description
                        $fieldDesc.FontStyle = "Italic"
                        $fieldDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                        $fieldDesc.TextWrapping = "Wrap"
                        $fieldDesc.Margin = "0,0,0,16"
                        [void]$tabContent.Children.Add($fieldDesc)

                        # Generate fields for this field
                        foreach ($field in $FieldListDef.fields) {
                            New-FieldListControl -ControlName $CardName -PolicyId $PolicyId -Field $field -Container $tabContent
                        }

                        $tabItem.Content = $tabContent
                        [void]$tabControl.Items.Add($tabItem)
                    }
                }
                [void]$detailsPanel.Children.Add($tabControl)
            } else {
                # Single Field - no tabs needed
                $FieldListDef = $syncHash.UIConfigs.inputTypes.$validInputFields[0]

                If($ShowDescription){
                    # Card description
                    $fieldDesc = New-Object System.Windows.Controls.TextBlock
                    $fieldDesc.Text = $FieldListDef.description
                    $fieldDesc.FontStyle = "Italic"
                    $fieldDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                    $fieldDesc.TextWrapping = "Wrap"
                    $fieldDesc.Margin = "0,0,0,16"
                    [void]$detailsPanel.Children.Add($fieldDesc)
                }
                # Generate fields based on Field
                foreach ($field in $FieldListDef.fields) {
                    New-FieldListControl -ControlName $CardName -PolicyId $PolicyId -Field $field -Container $detailsPanel
                }
            }

            # Button panel
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.Margin = "0,16,0,0"

            # Save button
            $saveButton = New-Object System.Windows.Controls.Button
            $saveButton.Content = "Save $CardName"
            $saveButton.Name = ($PolicyId.replace('.', '_') + "_" + $CardName + "_SaveButton")
            $saveButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $saveButton.Width = 120
            $saveButton.Height = 28
            $saveButton.Margin = "0,0,10,0"

            # Remove button (initially hidden)
            $removeButton = New-Object System.Windows.Controls.Button
            $removeButton.Content = "Remove $CardName"
            $removeButton.Name = ($PolicyId.replace('.', '_') + "_" + $CardName + "_RemoveButton")
            $removeButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $removeButton.Width = 120
            $removeButton.Height = 28
            $removeButton.Background = [System.Windows.Media.Brushes]::Red
            $removeButton.Foreground = [System.Windows.Media.Brushes]::White
            $removeButton.BorderThickness = "0"
            $removeButton.FontWeight = "SemiBold"
            $removeButton.Cursor = [System.Windows.Input.Cursors]::Hand
            $removeButton.Visibility = "Collapsed"

            [void]$buttonPanel.Children.Add($saveButton)
            [void]$buttonPanel.Children.Add($removeButton)
            [void]$detailsPanel.Children.Add($buttonPanel)

            # Add elements to main grid
            [void]$cardGrid.Children.Add($headerGrid)
            [void]$cardGrid.Children.Add($detailsPanel)
            $card.Child = $cardGrid

            # Add checkbox event handler
            $checkbox.Add_Checked({
                $detailsPanel = $this.Parent.Parent.Children | Where-Object { $_.GetType().Name -eq "StackPanel" }
                $detailsPanel.Visibility = "Visible"
            }.GetNewClosure())

            $checkbox.Add_Unchecked({
                $detailsPanel = $this.Parent.Parent.Children | Where-Object { $_.GetType().Name -eq "StackPanel" }
                $detailsPanel.Visibility = "Collapsed"
            }.GetNewClosure())

            Add-ControlEventHandler -Control $saveButton
            Add-ControlEventHandler -Control $removeButton

            # Create click event for save button
            $saveButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace( ("_" + $CardName + "_SaveButton"), "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.PolicySaving -f $CardName.ToLower(), $policyId) -Source $this.Name -Level "Info"

                # Get the details panel (parent of button panel)
                $detailsPanel = $this.Parent.Parent

                # Validate required fields BEFORE processing data
                $missingRequiredFields = Test-RequiredField -detailsPanel $detailsPanel -validInputFields $validInputFields -policyId $policyId -CardName $CardName

                if ($missingRequiredFields.Count -gt 0) {
                    $errorMessage = if ($missingRequiredFields.Count -eq 1) {
                        $syncHash.UIConfigs.localeErrorMessages.RequiredFieldValidation -f $missingRequiredFields[0]
                    } else {
                        $syncHash.UIConfigs.localeErrorMessages.RequiredFieldsValidation -f ($missingRequiredFields -join ", ")
                    }

                    [System.Windows.MessageBox]::Show($errorMessage, "Required Fields Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return # Exit save operation if validation fails
                }

                # Initialize policy level if not exists
                if (-not $OutputData[$ProductName]) {
                    $OutputData[$ProductName] = @{}
                }
                if (-not $OutputData[$ProductName][$policyId]) {
                    $OutputData[$ProductName][$policyId] = @{}
                }

                $hasOutputData = $false
                $savedinputTypes = @()

                # Get the details panel (parent of button panel)
                $detailsPanel = $this.Parent.Parent

                # Function to search recursively for the list container
                function Find-ListContainer {
                    <#
                    .SYNOPSIS
                    Recursively searches for a list container control by name.
                    .DESCRIPTION
                    This nested function traverses the WPF control hierarchy to locate a list container used for array field values.
                    #>

                    param($parent, $targetName)

                    if ($parent.Name -eq $targetName) {
                        return $parent
                    }

                    if ($parent.Children) {
                        foreach ($child in $parent.Children) {
                            $result = Find-ListContainer -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Content -and $parent.Content.Children) {
                        foreach ($child in $parent.Content.Children) {
                            $result = Find-ListContainer -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Items) {
                        foreach ($item in $parent.Items) {
                            if ($item.Content -and $item.Content.Children) {
                                foreach ($child in $item.Content.Children) {
                                    $result = Find-ListContainer -parent $child -targetName $targetName
                                    if ($result) { return $result }
                                }
                            }
                        }
                    }

                    return $null
                }

                # Function to search recursively for the checkbox
                function Find-CheckBox {
                    <#
                    .SYNOPSIS
                    Recursively searches for a CheckBox control by name.
                    .DESCRIPTION
                    This nested function traverses the WPF control hierarchy to locate a specific CheckBox control used for boolean field values.
                    #>
                    param($parent, $targetName)

                    if ($parent.Name -eq $targetName -and $parent -is [System.Windows.Controls.CheckBox]) {
                        return $parent
                    }

                    if ($parent.Children) {
                        foreach ($child in $parent.Children) {
                            $result = Find-CheckBox -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Content -and $parent.Content.Children) {
                        foreach ($child in $parent.Content.Children) {
                            $result = Find-CheckBox -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Items) {
                        foreach ($item in $parent.Items) {
                            if ($item.Content -and $item.Content.Children) {
                                foreach ($child in $item.Content.Children) {
                                    $result = Find-CheckBox -parent $child -targetName $targetName
                                    if ($result) { return $result }
                                }
                            }
                        }
                    }

                    return $null
                }

                # Function to search recursively for the textbox
                function Find-TextBox {
                    <#
                    .SYNOPSIS
                    Recursively searches for a TextBox control by name.
                    .DESCRIPTION
                    This nested function traverses the WPF control hierarchy to locate a specific TextBox control used for string field values.
                    #>
                    param($parent, $targetName)

                    if ($parent.Name -eq $targetName -and $parent -is [System.Windows.Controls.TextBox]) {
                        return $parent
                    }

                    if ($parent.Children) {
                        foreach ($child in $parent.Children) {
                            $result = Find-TextBox -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Content -and $parent.Content.Children) {
                        foreach ($child in $parent.Content.Children) {
                            $result = Find-TextBox -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Items) {
                        foreach ($item in $parent.Items) {
                            if ($item.Content -and $item.Content.Children) {
                                foreach ($child in $item.Content.Children) {
                                    $result = Find-TextBox -parent $child -targetName $targetName
                                    if ($result) { return $result }
                                }
                            }
                        }
                    }

                    return $null
                }

                function Find-DatePicker {
                    <#
                    .SYNOPSIS
                    Recursively searches for a DatePicker control by name.
                    .DESCRIPTION
                    This nested function traverses the WPF control hierarchy to locate a specific DatePicker control used for date field values.
                    #>
                    param($parent, $targetName)

                    if ($parent.Name -eq $targetName -and $parent -is [System.Windows.Controls.DatePicker]) {
                        return $parent
                    }

                    if ($parent.Children) {
                        foreach ($child in $parent.Children) {
                            $result = Find-DatePicker -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Content -and $parent.Content.Children) {
                        foreach ($child in $parent.Content.Children) {
                            $result = Find-DatePicker -parent $child -targetName $targetName
                            if ($result) { return $result }
                        }
                    }

                    if ($parent.Items) {
                        foreach ($item in $parent.Items) {
                            if ($item.Content -and $item.Content.Children) {
                                foreach ($child in $item.Content.Children) {
                                    $result = Find-DatePicker -parent $child -targetName $targetName
                                    if ($result) { return $result }
                                }
                            }
                        }
                    }

                    return $null
                }

                # Process each Field and MERGE them into a single policy entry
                foreach ($inputData in $validInputFields) {
                    $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                    if (-not $FieldListDef) { continue }

                    # Get the YAML key name
                    $yamlKeyName = $FieldListDef.name

                    Write-DebugOutput -Message (($syncHash.UIConfigs.LocaleVerboseMessages.ControlFieldProcessing -f $CardName, $inputData)) -Source $this.Name -Level "Verbose"

                    # Collect field values for this Field
                    $fieldCardData = @{}
                    foreach ($field in $FieldListDef.fields) {
                        # Use $field.value for control names (matching creation)
                        $controlFieldName = ($policyId.replace('.', '_') + "_" + $CardName + "_" + $field.value)

                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleVerboseMessages.ControlFieldCardSearch -f $controlFieldName, $field.type) -Source $this.Name -Level "Verbose"

                        if ($field.type -eq "array") {
                            # For arrays, look for the list container
                            $listContainerName = ($controlFieldName + "_List")
                            $listContainer = Find-ListContainer -parent $detailsPanel -targetName $listContainerName

                            if ($listContainer -and $listContainer.Children.Count -gt 0) {
                                $items = @()
                                foreach ($childPanel in $listContainer.Children) {
                                    if ($childPanel -is [System.Windows.Controls.StackPanel]) {
                                        foreach ($element in $childPanel.Children) {
                                            if ($element -is [System.Windows.Controls.TextBlock]) {
                                                $text = $element.Text.Trim()
                                                if (-not [string]::IsNullOrWhiteSpace($text)) {
                                                    $items += $text
                                                }
                                            }
                                        }
                                    }
                                }

                                if ($items.Count -gt 0) {
                                    # Use $field.value for data storage key (YAML output)
                                    $fieldCardData[$field.value] = $items
                                    Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.CollectedArrayField -f $inputData, $field.value, ($items -join ', ')) -Source $this.Name -Level "Info"
                                }
                            } else {
                                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleErrorMessages.ListContainerNotFound -f $listContainerName) -Source $this.Name -Level "Warning"
                            }

                        } elseif ($field.type -eq "boolean") {
                            # For boolean, look for the CheckBox
                            $booleanFieldName = ($controlFieldName + "_CheckBox")
                            $booleanCheckBox = Find-CheckBox -parent $detailsPanel -targetName $booleanFieldName

                            if ($booleanCheckBox) {
                                $value = [bool]$booleanCheckBox.IsChecked
                                # Use $field.value for data storage key (YAML output)
                                $fieldCardData[$field.value] = $value
                                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.CollectedBooleanField -f $inputData, $field.value, $value) -Source $this.Name -Level "Info"
                            } else {
                                Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleErrorMessages.ListContainerNotFound -f $booleanFieldName) -Source $this.Name -Level "Warning"
                            }

                        } elseif ($field.type -match "string") {
                            # Check if this is a dateString with yearmonthday - look for DatePicker
                            if ($field.type -eq "dateString") {
                                $datePickerName = ($controlFieldName + "_DatePicker")
                                $datePicker = Find-DatePicker -parent $detailsPanel -targetName $datePickerName

                                if ($datePicker -and $datePicker.SelectedDate) {
                                    $value = $datePicker.SelectedDate.ToString($syncHash.UIConfigs.valueValidations.($Field.valueType).format)
                                    $fieldCardData[$field.value] = $value
                                    Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.CollectedStringField -f $inputData, $field.value, $value) -Source $this.Name -Level "Info"
                                }
                            } else {
                                $stringFieldName = ($controlFieldName + "_TextBox")
                                $stringTextBox = Find-TextBox -parent $detailsPanel -targetName $stringFieldName

                                if ($stringTextBox -and ![string]::IsNullOrWhiteSpace($stringTextBox.Text)) {
                                    # Check if it's not just the placeholder text using config-driven approach
                                    $placeholderText = if ($syncHash.UIConfigs.valueValidations.($field.valueType)) {
                                        $syncHash.UIConfigs.valueValidations.($field.valueType).sample
                                    } else {
                                        "Enter value"  # fallback
                                    }

                                    if ($stringTextBox.Text.Trim() -ne $placeholderText) {
                                        $value = $stringTextBox.Text.Trim()
                                        # Use $field.value for data storage key (YAML output)
                                        $fieldCardData[$field.value] = $value
                                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleVerboseMessages.CollectedStringField -f $inputData, $field.value, $value) -Source $this.Name -Level "Info"
                                    }
                                } else {
                                    Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleErrorMessages.StringFieldNotFound -f $stringFieldName) -Source $this.Name -Level "Warning"
                                }
                            }
                        }
                    }

                    # Store the data with proper nesting for YAML output
                    if ($fieldCardData.Count -gt 0) {
                        # Get the YAML section name (exclusion type value like "CapExclusions", "RoleExclusions")
                        $FieldListValue = $FieldListDef.value

                        # Handle empty values - store fields directly under policy if no group value
                        if ([string]::IsNullOrWhiteSpace($FieldListValue)) {
                            # If the inputType has no value, store fields directly under the policy
                            foreach ($fieldKey in $fieldCardData.Keys) {
                                $OutputData[$ProductName][$policyId][$fieldKey] = $fieldCardData[$fieldKey]
                            }
                            Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.MergedCardField -f $CardName, $ProductName, $policyId, "Direct", ($fieldCardData | ConvertTo-Json -Compress)) -Source $this.Name -Level "Info"
                        } else {
                            # If the inputType has a value, create nested structure
                            # Initialize the exclusion type container if it doesn't exist
                            if (-not $OutputData[$ProductName][$policyId][$FieldListValue]) {
                                $OutputData[$ProductName][$policyId][$FieldListValue] = @{}
                            }

                            # Store field data under the exclusion type
                            foreach ($fieldKey in $fieldCardData.Keys) {
                                $OutputData[$ProductName][$policyId][$FieldListValue][$fieldKey] = $fieldCardData[$fieldKey]
                            }
                            Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleInfoMessages.MergedCardField -f $CardName, $ProductName, $policyId, $FieldListValue, ($fieldCardData | ConvertTo-Json -Compress)) -Source $this.Name -Level "Info"
                        }

                        $hasOutputData = $true
                        $savedinputTypes += $yamlKeyName
                    } else {
                        Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleErrorMessages.NoEntriesCollected -f $CardName.ToLower(), $inputData) -Source $this.Name -Level "Warning"
                    }
                }

                if ($hasOutputData) {
                    # Log the final merged structure
                    [System.Windows.MessageBox]::Show(($syncHash.UIConfigs.LocalePopupMessages.CardSavedSuccess -f $CardName, $ProductName, $policyId, ($savedinputTypes -join ', ')), "Success", "OK", "Information")

                    # Make remove button visible and header bold
                    $removeButton.Visibility = "Visible"
                    $policyHeader.FontWeight = "Bold"

                    # Collapse details panel and uncheck checkbox
                    $detailsPanel.Visibility = "Collapsed"
                    $checkbox.IsChecked = $false
                } else {
                    Write-DebugOutput -Message ($syncHash.UIConfigs.LocaleErrorMessages.NoEntriesFound -f $CardName.ToLower(), $inputData) -Source $this.Name -Level "Error"
                    [System.Windows.MessageBox]::Show(($syncHash.UIConfigs.LocalePopupMessages.NoEntriesFound -f $CardName.ToLower()), "Validation Error", "OK", "Warning")
                }
            }.GetNewClosure())


            # Enhanced remove button click handler for multiple exclusionFields
            $removeButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace(("_" + $CardName + "_RemoveButton"), "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show(($syncHash.UIConfigs.LocalePopupMessages.RemoveCardPolicyConfirmation -f $CardName.ToLower(), $policyId), "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {

                    # Remove the policy from the nested structure
                    if ($OutputData[$ProductName] -and $OutputData[$ProductName][$policyId]) {
                        $OutputData[$ProductName].Remove($policyId)

                        # If no more policies for this product, remove the product
                        if ($OutputData[$ProductName].Count -eq 0) {
                            $OutputData.Remove($ProductName)
                        }
                    }

                    # Clear all field values for all exclusionFields
                    foreach ($inputData in $validInputFields) {
                        $FieldListDef = $syncHash.UIConfigs.inputTypes.$inputData
                        if ($FieldListDef) {
                            foreach ($field in $FieldListDef.fields) {
                                $fieldName = ($policyId.replace('.', '_') + "_" + $inputData + "_" + $field.name)

                                if ($field.type -eq "array") {
                                    # Clear list containers
                                    $listContainerName = ($fieldName + "_List")
                                    $listContainer = $detailsPanel.Children | ForEach-Object {
                                        if ($_ -is [System.Windows.Controls.StackPanel]) {
                                            $arrayContainer = $_.Children | Where-Object { $_.Name -eq ($fieldName + "_" + $CardName + "_Container") }
                                            if ($arrayContainer) {
                                                return $arrayContainer.Children | Where-Object { $_.Name -eq $listContainerName }
                                            }
                                        } elseif ($_ -is [System.Windows.Controls.TabControl]) {
                                            # Search within tab control
                                            foreach ($tabItem in $_.Items) {
                                                if ($tabItem.Header -eq $FieldListDef.name) {
                                                    $tabContent = $tabItem.Content
                                                    $arrayContainer = $tabContent.Children | Where-Object { $_.Name -eq ($fieldName + "_" + $CardName + "_Container") }
                                                    if ($arrayContainer) {
                                                        return $arrayContainer.Children | Where-Object { $_.Name -eq $listContainerName }
                                                    }
                                                }
                                            }
                                        }
                                    } | Select-Object -First 1

                                    if ($listContainer) {
                                        $listContainer.Children.Clear()
                                    }
                                } elseif ($field.type -eq "boolean") {
                                    # Reset checkbox
                                    $booleanFieldName = ($fieldName + "_" + $CardName + "_CheckBox")
                                    $booleanCheckBox = $detailsPanel.Children | ForEach-Object {
                                        if ($_ -is [System.Windows.Controls.StackPanel]) {
                                            return $_.Children | Where-Object { $_.Name -eq $booleanFieldName -and $_ -is [System.Windows.Controls.CheckBox] }
                                        } elseif ($_ -is [System.Windows.Controls.TabControl]) {
                                            # Search within tab control
                                            foreach ($tabItem in $_.Items) {
                                                if ($tabItem.Header -eq $FieldListDef.name) {
                                                    $tabContent = $tabItem.Content
                                                    return $tabContent.Children | Where-Object { $_.Name -eq $booleanFieldName -and $_ -is [System.Windows.Controls.CheckBox] }
                                                }
                                            }
                                        }
                                    } | Select-Object -First 1

                                    if ($booleanCheckBox) {
                                        $booleanCheckBox.IsChecked = $false
                                    }
                                } elseif ($field.type -eq "string") {
                                    # Clear text boxes
                                    $stringFieldName = ($fieldName + "_" + $CardName + "_TextBox")
                                    $stringTextBox = $detailsPanel.Children | ForEach-Object {
                                        if ($_ -is [System.Windows.Controls.StackPanel]) {
                                            return $_.Children | Where-Object { $_.Name -eq $stringFieldName -and $_ -is [System.Windows.Controls.TextBox] }
                                        } elseif ($_ -is [System.Windows.Controls.TabControl]) {
                                            # Search within tab control
                                            foreach ($tabItem in $_.Items) {
                                                if ($tabItem.Header -eq $FieldListDef.name) {
                                                    $tabContent = $tabItem.Content
                                                    return $tabContent.Children | Where-Object { $_.Name -eq $stringFieldName -and $_ -is [System.Windows.Controls.TextBox] }
                                                }
                                            }
                                        }
                                    } | Select-Object -First 1

                                    if ($stringTextBox) {
                                        $stringTextBox.Text = ""
                                    }
                                }
                            }
                        }
                    }

                    [System.Windows.MessageBox]::Show(($syncHash.UIConfigs.LocalePopupMessages.RemoveCardEntrySuccess -f $CardName, $policyId), "Success", "OK", "Information")

                    # Hide remove button and unbold header
                    $this.Visibility = "Collapsed"
                    $policyHeader.FontWeight = "SemiBold"
                    $checkbox.IsChecked = $false
                }
            }.GetNewClosure())

            return $card
        }#end New-FieldListsCard

        #===========================================================================
        #
        # GRAPH HELPER
        #
        #===========================================================================

        # Enhanced Graph Query Function with Filter Support
        function Invoke-GraphQueryWithFilter {
            <#
            .SYNOPSIS
            Executes Microsoft Graph API queries with filtering support in a background thread.
            .DESCRIPTION
            This function performs asynchronous Microsoft Graph API queries with optional filtering, returning data for users, groups, or other Graph entities.
            #>
            param(
                [string]$QueryType,
                $GraphConfig,
                [string]$FilterString,
                [int]$Top = 999
            )

            # Create runspace
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()

            # Create PowerShell instance
            $powershell = [powershell]::Create()
            $powershell.Runspace = $runspace

            # Add script block
            $scriptBlock = {
                param($QueryType, $GraphConfig, $FilterString, $Top)

                try {
                    # Get query configuration
                    $queryConfig = $GraphConfig.$QueryType
                    if (-not $queryConfig) {
                        #Write-DebugOutput -Message "Query configuration not found for: $QueryType" -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }

                    # Build query parameters
                    $queryParams = @{
                        Uri = $queryConfig.endpoint
                        Method = "Get"
                    }

                    # Build query string
                    $queryStringParts = @()

                    # Add existing query parameters from config
                    if ($queryConfig.queryParameters) {
                        foreach ($param in $queryConfig.queryParameters.psobject.properties.name) {
                            $queryStringParts += "$param=$($queryConfig.queryParameters.$param)"
                        }
                    }

                    # Add filter if provided
                    if (![string]::IsNullOrWhiteSpace($FilterString)) {
                        $queryStringParts += "`$filter=$FilterString"
                    }

                    # Add top parameter
                    $queryStringParts += "`$top=$Top"

                    # Combine query string
                    if ($queryStringParts.Count -gt 0) {
                        $queryParams.Uri += $syncHash.GraphEndpoint + "?" + ($queryStringParts -join "&")
                    }

                    #Write-DebugOutput -Message "Graph Query URI: $($queryParams.Uri)" -Source $MyInvocation.MyCommand.Name -Level "Information"

                    # Execute the Graph request
                    $result = Invoke-MgGraphRequest @queryParams

                    # Return the result
                    return @{
                        Success = $true
                        Data = $result
                        QueryConfig = $queryConfig
                        Message = "Successfully retrieved $($result.value.Count) items"
                        FilterApplied = ![string]::IsNullOrWhiteSpace($FilterString)
                    }
                }
                catch {
                    #Write-DebugOutput -Message "Error executing Graph query: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                    return @{
                        Success = $false
                        Error = $_.Exception.Message
                        Message = "Failed to retrieve data from uri [{0}]: {1}" -f $queryParams.Uri, $($_.Exception.Message)
                        FilterApplied = ![string]::IsNullOrWhiteSpace($FilterString)
                    }
                }
            }

            # Add parameters and start execution
            $powershell.AddScript($scriptBlock).AddParameter("QueryType", $QueryType).AddParameter("GraphConfig", $GraphConfig).AddParameter("FilterString", $FilterString).AddParameter("Top", $Top)
            $asyncResult = $powershell.BeginInvoke()

            return @{
                PowerShell = $powershell
                AsyncResult = $asyncResult
                Runspace = $runspace
            }
        }


        #===========================================================================
        # Graph Selection Function
        #===========================================================================
        function Show-GraphProgressWindow {
            <#
            .SYNOPSIS
            Displays a progress window while executing Graph queries and shows results in a selection interface.
            .DESCRIPTION
            This function shows a progress dialog during Graph API operations and presents the results in a searchable, selectable data grid for users and groups.
            #>
            param(
                [string]$GraphEntityType,
                [string]$SearchTerm = "",
                [int]$Top = 100
            )

            try {
                # Define entity-specific configurations
                $entityConfigs = @{
                    users = @{
                        Title = "Select Users"
                        SearchPlaceholder = "Search by display name..."
                        LoadingMessage = "Loading users..."
                        NoResultsMessage = "No users found matching the search criteria."
                        NoResultsTitle = "No Users Found"
                        FilterProperty = "userPrincipalName"
                        SearchProperty = "DisplayName"
                        DisplayOrder = @("DisplayName", "UserPrincipalName", "AccountEnabled", "Id")
                        QueryType = "users"
                        DataTransform = {
                            param($item)
                            [PSCustomObject]@{
                                DisplayName = $item.DisplayName
                                UserPrincipalName = $item.UserPrincipalName
                                AccountEnabled = $item.AccountEnabled
                                Id = $item.Id
                                OriginalObject = $item
                            }
                        }
                        ColumnConfig = [ordered]@{
                            DisplayName = @{ Header = "Display Name"; Width = 200 }
                            UserPrincipalName = @{ Header = "User Principal Name"; Width = 250 }
                            AccountEnabled = @{ Header = "Enabled"; Width = 80 }
                            Id = @{ Header = "ID"; Width = 150 }
                        }
                    }
                    groups = @{
                        Title = "Select Groups"
                        SearchPlaceholder = "Search by group name..."
                        LoadingMessage = "Loading groups..."
                        NoResultsMessage = "No groups found matching the search criteria."
                        NoResultsTitle = "No Groups Found"
                        FilterProperty = "displayName"
                        SearchProperty = "DisplayName"
                        DisplayOrder = @("DisplayName", "Description", "GroupType", "Id")
                        QueryType = "groups"
                        DataTransform = {
                            param($item)
                            $groupType = "Distribution"
                            if ($item.SecurityEnabled) { $groupType = "Security" }
                            if ($item.GroupTypes -contains "Unified") { $groupType = "Microsoft 365" }

                            [PSCustomObject]@{
                                DisplayName = $item.DisplayName
                                Description = $item.Description
                                GroupType = $groupType
                                Id = $item.Id
                                OriginalObject = $item
                            }
                        }
                        ColumnConfig = [ordered]@{
                            DisplayName = @{ Header = "Group Name"; Width = 200 }
                            Description = @{ Header = "Description"; Width = 250 }
                            GroupType = @{ Header = "Type"; Width = 100 }
                            Id = @{ Header = "ID"; Width = 150 }
                        }
                    }
                }

                # Get configuration for the specified entity type
                $config = $entityConfigs[$GraphEntityType]
                if ($config) {
                    Write-DebugOutput -Message "Using configuration for graph entity type: $GraphEntityType" -Source $MyInvocation.MyCommand.Name -Level "Info"
                }else{
                    Write-DebugOutput -Message "Unsupported graph entity type: $GraphEntityType" -Source $MyInvocation.MyCommand.Name -Level "Error"
                }

                # Build filter string
                $filterString = $null
                if (![string]::IsNullOrWhiteSpace($SearchTerm)) {
                    $filterString = "startswith($($config.FilterProperty),'$SearchTerm')"
                }

                # Show progress window
                $progressWindow = New-Object System.Windows.Window
                $progressWindow.Title = $config.Title
                $progressWindow.Width = 300
                $progressWindow.Height = 120
                $progressWindow.WindowStartupLocation = "CenterOwner"
                $progressWindow.Owner = $syncHash.Window
                $progressWindow.Background = [System.Windows.Media.Brushes]::White

                $progressPanel = New-Object System.Windows.Controls.StackPanel
                $progressPanel.Margin = "20"
                $progressPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
                $progressPanel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

                $progressLabel = New-Object System.Windows.Controls.Label
                $progressLabel.Content = $config.LoadingMessage
                $progressLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center

                $progressBar = New-Object System.Windows.Controls.ProgressBar
                $progressBar.Width = 200
                $progressBar.Height = 20
                $progressBar.IsIndeterminate = $true

                [void]$progressPanel.Children.Add($progressLabel)
                [void]$progressPanel.Children.Add($progressBar)
                $progressWindow.Content = $progressPanel

                # Start async operation
                Write-DebugOutput -Message "Starting async operation for graph query type: $($config.QueryType) with filter: $filterString" -Source $MyInvocation.MyCommand.Name -Level "Info"
                $asyncOp = Invoke-GraphQueryWithFilter `
                                -QueryType $config.QueryType `
                                -GraphConfig $syncHash.UIConfigs.graphQueries `
                                -FilterString $filterString -Top $Top

                # Show progress window
                $progressWindow.Show()

                # Wait for completion
                while (-not $asyncOp.AsyncResult.IsCompleted) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 100
                }

                # Close progress window
                $progressWindow.Close()

                # Get results
                $result = $asyncOp.PowerShell.EndInvoke($asyncOp.AsyncResult)
                $asyncOp.PowerShell.Dispose()
                $asyncOp.Runspace.Close()
                $asyncOp.Runspace.Dispose()

                if ($result.Success) {
                    Write-DebugOutput -Message "Graph query successful for entity type: $GraphEntityType, items found: $($result.Data.value.Count)" -Source $MyInvocation.MyCommand.Name -Level "Info"
                    $items = $result.Data.value
                    if (-not $items -or $items.Count -eq 0) {
                        [System.Windows.MessageBox]::Show($config.NoResultsMessage, $config.NoResultsTitle,
                                                        [System.Windows.MessageBoxButton]::OK,
                                                        [System.Windows.MessageBoxImage]::Information)
                        return $null
                    }

                    # Transform data using entity-specific transformer
                    $displayItems = $items | ForEach-Object {
                        & $config.DataTransform $_
                    } | Sort-Object DisplayName

                    # Show selector using the universal selection window
                    $selectedItems = Show-UISelectionWindow `
                                        -Title $config.Title `
                                        -SearchPlaceholder $config.SearchPlaceholder `
                                        -Items $displayItems `
                                        -ColumnConfig $config.ColumnConfig `
                                        -SearchProperty $config.SearchProperty `
                                        -DisplayOrder $config.ColumnConfig.Keys `
                                        -AllowMultiple

                    return $selectedItems
                }
                else {
                    Write-DebugOutput -Message "Graph query failed for entity type: $GraphEntityType, error: $($result.Error)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                    [System.Windows.MessageBox]::Show($result.Message, "Error",
                                                    [System.Windows.MessageBoxButton]::OK,
                                                    [System.Windows.MessageBoxImage]::Error)
                    return $null
                }
            }
            catch {
                [System.Windows.MessageBox]::Show(("{0} {1}: {2}" -f $syncHash.UIConfigs.localeErrorMessages.WindowError,$GraphEntityType, $_.Exception.Message),
                                                "Error",
                                                [System.Windows.MessageBoxButton]::OK,
                                                [System.Windows.MessageBoxImage]::Error)
                return $null
            }
        }

        function Show-GraphSelector {
            <#
            .SYNOPSIS
            Shows a Graph entity selector with optional search functionality.
            .DESCRIPTION
            This function displays a selector interface for Microsoft Graph entities (users, groups) with optional search term filtering and result limiting.
            #>
            param(
                [string]$GraphEntityType,
                [string]$SearchTerm = "",
                [int]$Top = 100
            )
            If([string]::IsNullOrWhiteSpace($SearchTerm)) {
                Write-DebugOutput -Message "Showing $($GraphEntityType.ToLower()) selector with top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }Else {
                Write-DebugOutput -Message "Showing $($GraphEntityType.ToLower()) selector with search term: $SearchTerm, top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }
            return Show-GraphProgressWindow -GraphEntityType $GraphEntityType -SearchTerm $SearchTerm -Top $Top
        }

        #build UI selection window
        function Show-UISelectionWindow {
            <#
            .SYNOPSIS
            Creates a universal selection window with search and filtering capabilities.
            .DESCRIPTION
            This function generates a reusable selection dialog with a searchable data grid, supporting single or multiple selection modes for various data types.
            #>
            param(
                [Parameter(Mandatory)]
                [string]$Title,

                [Parameter(Mandatory)]
                [string]$SearchPlaceholder,

                [Parameter(Mandatory)]
                [array]$Items,

                [Parameter(Mandatory)]
                [hashtable]$ColumnConfig,

                [Parameter()]
                [string[]]$DisplayOrder,

                [Parameter()]
                [string]$SearchProperty = "DisplayName",

                [Parameter()]
                [int]$WindowWidth = 1000,

                [Parameter()]
                [string]$ReturnProperty,

                [Parameter()]
                [switch]$AllowMultiple
            )

            #[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | out-null
            #[System.Reflection.Assembly]::LoadWithPartialName('System.Security') | out-null

            try {
                # Create selection window
                $selectionWindow = New-Object System.Windows.Window
                $selectionWindow.Title = $Title
                $selectionWindow.Width = $WindowWidth
                $selectionWindow.Height = 500
                $selectionWindow.WindowStartupLocation = "CenterOwner"
                $selectionWindow.Owner = $syncHash.Window
                $selectionWindow.Background = [System.Windows.Media.Brushes]::White

                # Create main grid
                $mainGrid = New-Object System.Windows.Controls.Grid
                $rowDef1 = New-Object System.Windows.Controls.RowDefinition
                $rowDef1.Height = [System.Windows.GridLength]::Auto
                $rowDef2 = New-Object System.Windows.Controls.RowDefinition
                $rowDef2.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                $rowDef3 = New-Object System.Windows.Controls.RowDefinition
                $rowDef3.Height = [System.Windows.GridLength]::Auto
                [void]$mainGrid.RowDefinitions.Add($rowDef1)
                [void]$mainGrid.RowDefinitions.Add($rowDef2)
                [void]$mainGrid.RowDefinitions.Add($rowDef3)

                # Search panel
                $searchPanel = New-Object System.Windows.Controls.StackPanel
                $searchPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                $searchPanel.Margin = "10"

                $searchLabel = New-Object System.Windows.Controls.Label
                $searchLabel.Content = "Search:"
                $searchLabel.Width = 60
                $searchLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

                $searchBox = New-Object System.Windows.Controls.TextBox
                $searchBox.Width = 300
                $searchBox.Height = 25
                $searchBox.Text = $SearchPlaceholder
                $searchBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $searchBox.FontStyle = [System.Windows.FontStyles]::Italic
                $searchBox.Margin = "5,0"

                # Search box placeholder functionality
                $searchBox.Add_GotFocus({
                    if ($searchBox.Text -eq $SearchPlaceholder) {
                        $searchBox.Text = ""
                        $searchBox.Foreground = [System.Windows.Media.Brushes]::Black
                        $searchBox.FontStyle = [System.Windows.FontStyles]::Normal
                    }
                })

                $searchBox.Add_LostFocus({
                    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
                        $searchBox.Text = $SearchPlaceholder
                        $searchBox.Foreground = [System.Windows.Media.Brushes]::Gray
                        $searchBox.FontStyle = [System.Windows.FontStyles]::Italic
                    }
                })

                [void]$searchPanel.Children.Add($searchLabel)
                [void]$searchPanel.Children.Add($searchBox)

                [System.Windows.Controls.Grid]::SetRow($searchPanel, 0)
                [void]$mainGrid.Children.Add($searchPanel)

                # Create DataGrid
                $dataGrid = New-Object System.Windows.Controls.DataGrid
                $dataGrid.AutoGenerateColumns = $false
                $dataGrid.CanUserAddRows = $false
                $dataGrid.CanUserDeleteRows = $false
                $dataGrid.IsReadOnly = $true
                $dataGrid.SelectionMode = if ($AllowMultiple) { [System.Windows.Controls.DataGridSelectionMode]::Extended } else { [System.Windows.Controls.DataGridSelectionMode]::Single }
                $dataGrid.GridLinesVisibility = [System.Windows.Controls.DataGridGridLinesVisibility]::Horizontal
                $dataGrid.HeadersVisibility = [System.Windows.Controls.DataGridHeadersVisibility]::Column
                $dataGrid.Margin = "10"

                # Display order handling if specified
                if ($DisplayOrder -and $DisplayOrder.Count -gt 0) {
                    $keyOrder = $DisplayOrder
                } else {
                    # Use keys from ColumnConfig if no display order specified
                    $keyOrder = $ColumnConfig.Keys | Sort-Object
                }

                # Create columns based on ColumnConfig
                foreach ($columnKey in $keyOrder) {
                    if ($ColumnConfig.ContainsKey($columnKey)) {
                        $column = New-Object System.Windows.Controls.DataGridTextColumn
                        $column.Header = $ColumnConfig[$columnKey].Header
                        $column.Binding = New-Object System.Windows.Data.Binding($columnKey)
                        $column.Width = $ColumnConfig[$columnKey].Width
                        $dataGrid.Columns.Add($column)
                    } else {
                        Write-DebugOutput -Message "Column configuration for '$columnKey' not found in ColumnConfig." -Source $MyInvocation.MyCommand.Name -Level "Warning"
                    }
                }

                # Store original items for filtering
                $originalItems = $Items

                # Filter function
                $FilterItems = {
                    $searchText = $searchBox.Text.ToLower()
                    if ([string]::IsNullOrWhiteSpace($searchText) -or $searchText -eq $SearchPlaceholder.ToLower()) {
                        $dataGrid.ItemsSource = $originalItems
                    } else {
                        $filteredItems = @($originalItems | Where-Object {
                            $_.$SearchProperty.ToLower().Contains($searchText)
                        })
                        $dataGrid.ItemsSource = $filteredItems
                    }
                }

                # Initial load
                $dataGrid.ItemsSource = $originalItems

                # Search on text change
                $searchBox.Add_TextChanged($FilterItems)

                [System.Windows.Controls.Grid]::SetRow($dataGrid, 1)
                [void]$mainGrid.Children.Add($dataGrid)

                # Create button panel
                $buttonPanel = New-Object System.Windows.Controls.StackPanel
                $buttonPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
                $buttonPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
                $buttonPanel.Margin = "10"

                $selectButton = New-Object System.Windows.Controls.Button
                $selectButton.Content = "Select"
                $selectButton.Width = 80
                $selectButton.Height = 30
                $selectButton.Margin = "0,0,10,0"
                $selectButton.IsDefault = $true

                $cancelButton = New-Object System.Windows.Controls.Button
                $cancelButton.Content = "Cancel"
                $cancelButton.Width = 80
                $cancelButton.Height = 30
                $cancelButton.IsCancel = $true

                [void]$buttonPanel.Children.Add($selectButton)
                [void]$buttonPanel.Children.Add($cancelButton)

                [System.Windows.Controls.Grid]::SetRow($buttonPanel, 2)
                [void]$mainGrid.Children.Add($buttonPanel)

                $selectionWindow.Content = $mainGrid

                # Event handlers
                $selectButton.Add_Click({
                    if ($dataGrid.SelectedItems.Count -gt 0) {
                        $selectedResults = @()
                        foreach ($selectedItem in $dataGrid.SelectedItems) {
                            $selectedResults += $selectedItem
                        }
                        $selectionWindow.Tag = $selectedResults
                        $selectionWindow.DialogResult = $true
                        $selectionWindow.Close()
                    } else {
                        [System.Windows.MessageBox]::Show("Please select an item.", "No Selection",
                                                        [System.Windows.MessageBoxButton]::OK,
                                                        [System.Windows.MessageBoxImage]::Warning)
                    }
                })

                $cancelButton.Add_Click({
                    $selectionWindow.DialogResult = $false
                    $selectionWindow.Close()
                })

                $dataGrid.Add_MouseDoubleClick({
                    if ($dataGrid.SelectedItem) {
                        $selectionWindow.Tag = @($dataGrid.SelectedItem)
                        $selectionWindow.DialogResult = $true
                        $selectionWindow.Close()
                    }
                })

                # Show dialog
                $result = $selectionWindow.ShowDialog()

                if ($result -eq $true) {
                    If($ReturnProperty) {
                        # Return the specified property from the selected items
                        $returnValues = @()
                        foreach ($item in $selectionWindow.Tag) {
                            if ($item -is [PSCustomObject] -and $item.PSObject.Properties[$ReturnProperty]) {
                                $returnValues += $item.$ReturnProperty
                            } else {
                                Write-DebugOutput -Message "Selected item does not have property '$ReturnProperty': $($item | ConvertTo-Json -Compress)" -Source $MyInvocation.MyCommand.Name -Level "Warning"
                            }
                        }
                        return $returnValues
                    }Else{
                        return $selectionWindow.Tag
                    }
                }
                return $null
            }
            catch {
                [System.Windows.MessageBox]::Show( ("{0} {1}: {2}" -f $syncHash.UIConfigs.localeErrorMessages.WindowError, $Title, $_.Exception.Message),
                                                "Error",
                                                [System.Windows.MessageBoxButton]::OK,
                                                [System.Windows.MessageBoxImage]::Error)
                return $null
            }
        }


        #======================================================
        # YAML CONTROL
        #======================================================
        # Function to import YAML data into core data structures (without UI updates)
        # Updated Import-YamlToDataStructures Function
        Function Import-YamlToDataStructures {
            <#
            .SYNOPSIS
            Imports YAML configuration data into internal data structures.
            .DESCRIPTION
            This function parses YAML configuration data and populates the application's internal data structures without updating the UI.
            #>
            param($Config)

            try {
                # Get top-level keys (now always hashtable)
                $topLevelKeys = $Config.Keys

                #get all products from UIConfigs
                $productIds = $syncHash.UIConfigs.products | Select-Object -ExpandProperty id

                # Import General Settings that are not product
                $generalSettingsFields = $topLevelKeys | Where-Object {$_ -notin $productIds}

                foreach ($field in $generalSettingsFields) {
                    # Special handling for ProductNames to expand '*' wildcard
                    if ($field -eq "ProductNames" -and $Config[$field] -contains "*") {
                        # Expand '*' to all available products
                        $syncHash.GeneralSettings[$field] = $productIds
                    } else {
                        $syncHash.GeneralSettings[$field] = $Config[$field]
                    }
                }

                # Import Exclusions
                $hasProductSections = $productIds | Where-Object { $topLevelKeys -contains $_ }

                if ($hasProductSections) {
                    foreach ($productName in $productIds) {
                        if ($topLevelKeys -contains $productName) {
                            $productData = $Config[$productName]
                            $productKeys = $productData.Keys

                            foreach ($policyId in $productKeys) {
                                $policyData = $productData[$policyId]

                                # Find the exclusionField from baseline config
                                $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                                if ($baseline -and $baseline.exclusionField -ne "none") {
                                    # Get the exclusionField configuration to get the actual YAML key name
                                    $FieldListConfig = $syncHash.UIConfigs.inputTypes.($baseline.exclusionField)
                                    $yamlKeyName = if ($FieldListConfig) { $FieldListConfig.name } else { $baseline.exclusionField }

                                    # Check if the policy data contains the expected exclusionField key
                                    if ($policyData.ContainsKey($yamlKeyName)) {
                                        # Initialize product level if not exists
                                        if (-not $syncHash.Exclusions[$productName]) {
                                            $syncHash.Exclusions[$productName] = @{}
                                        }

                                        # Initialize policy level if not exists
                                        if (-not $syncHash.Exclusions[$productName][$policyId]) {
                                            $syncHash.Exclusions[$productName][$policyId] = @{}
                                        }

                                        # CORRECTED: Extract the field data from inside the exclusionField
                                        # Instead of: $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $policyData
                                        # We want: $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $policyData[$yamlKeyName]
                                        $FieldListData = $policyData[$yamlKeyName]
                                        $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $FieldListData

                                        Write-DebugOutput -Message "Imported exclusion for [$productName][$policyId][$yamlKeyName]: $($FieldListData | ConvertTo-Json -Compress)" -Source $MyInvocation.MyCommand.Name -Level "Info"
                                    }
                                }
                            }
                        }
                    }
                }

                # Import Omissions - Now using hashtable structure
                if ($topLevelKeys -contains "OmitPolicy") {
                    $omissionKeys = $Config["OmitPolicy"].Keys

                    foreach ($policyId in $omissionKeys) {
                        $omissionData = $Config["OmitPolicy"][$policyId]

                        # Find which product this policy belongs to
                        $productName = $null
                        foreach ($product in $syncHash.UIConfigs.products) {
                            $baseline = $syncHash.UIConfigs.baselines.($product.id) | Where-Object { $_.id -eq $policyId }
                            if ($baseline) {
                                $productName = $product.id
                                break
                            }
                        }

                        if ($productName) {
                            # Store as hashtable: $syncHash.Omissions[$productName][$policyId] = @{data}
                            if (-not $syncHash.Omissions[$productName]) {
                                $syncHash.Omissions[$productName] = @{}
                            }

                            $syncHash.Omissions[$productName][$policyId] = @{
                                Id = $policyId
                                Product = $productName
                                Rationale = $omissionData["Rationale"]
                                Expiration = $omissionData["Expiration"]
                            }
                        }
                    }
                }

                # Import Annotations - Now using hashtable structure
                if ($topLevelKeys -contains "AnnotatePolicy") {
                    $annotationKeys = $Config["AnnotatePolicy"].Keys

                    foreach ($policyId in $annotationKeys) {
                        $annotationData = $Config["AnnotatePolicy"][$policyId]

                        # Find which product this policy belongs to
                        $productName = $null
                        foreach ($product in $syncHash.UIConfigs.products) {
                            $baseline = $syncHash.UIConfigs.baselines.($product.id) | Where-Object { $_.id -eq $policyId }
                            if ($baseline) {
                                $productName = $product.id
                                break
                            }
                        }

                        if ($productName) {
                            # Store as hashtable: $syncHash.Annotations[$productName][$policyId] = @{data}
                            if (-not $syncHash.Annotations[$productName]) {
                                $syncHash.Annotations[$productName] = @{}
                            }

                            $syncHash.Annotations[$productName][$policyId] = @{
                                Id = $policyId
                                Product = $productName
                                Comment = $annotationData["Comment"]
                            }
                        }
                    }
                }

            }
            catch {
                Write-DebugOutput -Message "Error importing data: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
            }
        }

        Function New-YamlPreview {
            <#
            .SYNOPSIS
            Generates YAML configuration preview from current UI settings.
            .DESCRIPTION
            This function creates a YAML preview string by collecting values from all UI controls and formatting them according to ScubaGear configuration standards.
            #>
            $yamlPreview = @()
            $yamlPreview += '# ScubaGear Configuration File'
            $yamlPreview += "`n# Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $yamlPreview += "`n`n# Organization Configuration"

            # Process main settings in order of localePlaceholder keys
            if ($syncHash.UIConfigs.localePlaceholder) {
                foreach ($placeholderKey in $syncHash.UIConfigs.localePlaceholder.PSObject.Properties.Name) {
                    $control = $syncHash.$placeholderKey
                    if ($control -is [System.Windows.Controls.TextBox]) {
                        $currentValue = $control.Text
                        $placeholderValue = $syncHash.UIConfigs.localePlaceholder.$placeholderKey

                        # Only include if it's not empty and not a placeholder
                        if (![string]::IsNullOrWhiteSpace($currentValue) -and $currentValue -ne $placeholderValue) {
                            # Convert control name to YAML field name (remove _TextBox suffix)
                            $yamlFieldName = $placeholderKey -replace '_TextBox$', ''
                            $trimmedValue = $currentValue.Trim()

                            # Handle special formatting for description
                            if ($yamlFieldName -eq 'Description') {
                                $escapedDescription = $trimmedValue.Replace('"', '""')
                                $yamlPreview += "`n$yamlFieldName`: `"$escapedDescription`""
                            } else {
                                $yamlPreview += "`n$yamlFieldName`: $trimmedValue"
                            }
                        }
                    }
                }
            }

            $yamlPreview += "`n`n# Configuration Details"

            # Handle ProductNames using the enhanced function
            $yamlPreview += Get-ProductNamesForYaml

            # Handle M365Environment
            $selectedEnv = $syncHash.UIConfigs.M365Environment | Where-Object { $_.id -eq $syncHash.M365Environment_ComboBox.SelectedItem.Tag } | Select-Object -ExpandProperty name
            $yamlPreview += "`n`nM365Environment: $selectedEnv"
            $yamlPreview += "`n"

            # Process advanced settings in order of defaultAdvancedSettings keys
            if ($syncHash.UIConfigs.defaultAdvancedSettings) {
                $hasAdvancedSettings = $false
                $advancedSettingsContent = @()

                foreach ($advancedKey in $syncHash.UIConfigs.defaultAdvancedSettings.PSObject.Properties.Name) {
                    $control = $syncHash.$advancedKey
                    $defaultValue = $syncHash.UIConfigs.defaultAdvancedSettings.$advancedKey

                    # Only include settings that are explicitly enabled by the user
                    # Check if there's an associated checkbox to enable/disable this setting
                    $enabledCheckBoxName = $advancedKey -replace '_TextBox$', '_EnabledCheckBox'
                    $enabledCheckBox = $syncHash.$enabledCheckBoxName

                    # If there's no enabled checkbox, check if setting has been modified from default
                    $shouldInclude = $false

                    if ($enabledCheckBox -is [System.Windows.Controls.CheckBox]) {
                        # If there's an enabled checkbox, only include if checked
                        $shouldInclude = $enabledCheckBox.IsChecked -eq $true
                    } elseif ($control -is [System.Windows.Controls.TextBox]) {
                        # If no enabled checkbox, only include if user has changed from default
                        $currentValue = $control.Text
                        $shouldInclude = (![string]::IsNullOrWhiteSpace($currentValue) -and
                                        ![string]::IsNullOrWhiteSpace($defaultValue))
                    } elseif ($control -is [System.Windows.Controls.CheckBox]) {
                        # For checkbox controls, only include if different from default
                        $shouldInclude = $control.IsChecked -ne $defaultValue
                    }

                    # Only process if the setting should be included
                    if ($shouldInclude) {
                        if ($control -is [System.Windows.Controls.TextBox]) {
                            $currentValue = $control.Text

                            # Use current value or default if empty
                            $valueToUse = if (![string]::IsNullOrWhiteSpace($currentValue)) { $currentValue } else { $defaultValue }

                            if (![string]::IsNullOrWhiteSpace($valueToUse)) {
                                # Convert control name to YAML field name (remove _TextBox suffix)
                                $yamlFieldName = $advancedKey -replace '_TextBox$', ''

                                # Handle path escaping
                                if ($valueToUse -match '\\') {
                                    $advancedSettingsContent += "`n$yamlFieldName`: `"$($valueToUse.Replace('\', '\\'))`""
                                } else {
                                    $advancedSettingsContent += "`n$yamlFieldName`: $valueToUse"
                                }
                                $hasAdvancedSettings = $true
                            }
                        }
                        elseif ($control -is [System.Windows.Controls.CheckBox]) {
                            # Convert control name to YAML field name (remove _CheckBox suffix)
                            $yamlFieldName = $advancedKey -replace '_CheckBox$', ''
                            $boolValue = if ($control.IsChecked) { "true" } else { "false" }
                            $advancedSettingsContent += "`n$yamlFieldName`: $boolValue"
                            $hasAdvancedSettings = $true
                        }
                    }
                }

                # Only add advanced settings section if there are settings to show
                if ($hasAdvancedSettings) {
                    $yamlPreview += "`n`n# Advanced Configuration"
                    $yamlPreview += $advancedSettingsContent
                    $yamlPreview += "`n"
                }
            }

            <#
            Function New-DynamicYamlSections {
                foreach ($baselineControl in $syncHash.UIConfigs.baselineControls) {
                    $dataHashtable = $syncHash.($baselineControl.dataControlOutput)

                    if ($dataHashtable -and $dataHashtable.Keys.Count -gt 0) {
                        $yamlPreview += "`n`n$($baselineControl.yamlValue):"

                        # Generate content based on the data structure
                        foreach ($productName in $dataHashtable.Keys) {
                            foreach ($policyId in $dataHashtable[$productName].Keys) {
                                # Dynamic YAML generation logic here
                            }
                        }
                    }
                }
            }
            #>


            # Handle Exclusions (unchanged - already hashtable)
            foreach ($productName in $syncHash.Exclusions.Keys) {
                #$productHasExclusions = $false
                $productExclusions = @()

                foreach ($policyId in $syncHash.Exclusions[$productName].Keys) {
                    $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                    $policyBlock = @()

                    if ($baseline) {
                        $policyBlock += "`n  # $($baseline.name)"
                    }

                    $policyBlock += "`n  $policyId`:"  # Only once per policy

                    $exclusions = $syncHash.Exclusions[$productName][$policyId]

                    foreach ($exclusionKey in $exclusions.Keys) {
                        $exclusionData = $exclusions[$exclusionKey]

                        if ($null -ne $exclusionData -and ($exclusionData -isnot [System.Collections.ICollection] -or $exclusionData.Count -gt 0)) {

                            # Handle types: boolean / string
                            if ($exclusionData -is [bool] -or $exclusionData -is [string] -or $exclusionData -is [int]) {
                                $policyBlock += "`n    $exclusionKey`: $exclusionData"
                            }

                            # Array
                            elseif ($exclusionData -is [array]) {
                                $policyBlock += "`n    $exclusionKey`:"
                                foreach ($item in $exclusionData) {
                                    $policyBlock += "`n      - $item"
                                }
                            }

                            # Hashtable
                            elseif ($exclusionData -is [hashtable]) {
                                $policyBlock += "`n    $exclusionKey`:"
                                foreach ($fieldName in $exclusionData.Keys) {
                                    $fieldValue = $exclusionData[$fieldName]

                                    if ($null -eq $fieldValue) { continue }

                                    $policyBlock += "`n      $fieldName`:"

                                    if ($fieldValue -is [array]) {
                                        foreach ($item in $fieldValue) {
                                            $policyBlock += "`n        - $item"
                                        }
                                    } else {
                                        $policyBlock += "`n        - $fieldValue"
                                    }
                                }
                            }
                        }
                    }

                    # Append the entire block only ONCE per policy
                    if ($policyBlock.Count -gt 0) {
                        $productExclusions += $policyBlock
                    }
                }

                if ($productExclusions.Count -gt 0) {
                    $yamlPreview += "`n$productName`:"
                    $yamlPreview += $productExclusions
                }
            }

            # Handle Annotations - Dynamic version using UIConfig fields
            $annotationCount = 0
            foreach ($productName in $syncHash.Annotations.Keys) {
                $annotationCount += $syncHash.Annotations[$productName].Count
            }

            if ($annotationCount -gt 0) {
                # Get the YAML section name from UIConfig
                $annotationSectionName = $syncHash.UIConfigs.inputTypes.annotation.value
                $yamlPreview += "`n`n$annotationSectionName`:"

                # Group annotations by product
                foreach ($productName in $syncHash.Annotations.Keys) {
                    #$yamlPreview += "`n  # $productName Annotations:"

                    # Sort policies by ID
                    $sortedPolicies = $syncHash.Annotations[$productName].Keys | Sort-Object
                    foreach ($policyId in $sortedPolicies) {
                        $annotationData = $syncHash.Annotations[$productName][$policyId]

                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $policyId`:"

                        # Access the nested annotation data correctly
                        # The structure is: $syncHash.Annotations[Product][PolicyId][AnnotationType][FieldName]
                        # We need to get the annotation type (like "Annotate") first
                        foreach ($annotationType in $annotationData.Keys) {
                            $annotation = $annotationData[$annotationType]

                            # Process fields in the order defined in UIConfig
                            $annotationFields = $syncHash.UIConfigs.inputTypes.annotation.fields
                            foreach ($field in $annotationFields) {
                                $fieldValue = $annotation[$field.value]

                                if ($null -ne $fieldValue -and ![string]::IsNullOrEmpty($fieldValue)) {
                                    # Handle different field types
                                    if ($field.type -eq "boolean") {
                                        $yamlPreview += "`n    $($field.value): $($fieldValue.ToString().ToLower())"
                                    }
                                    elseif ($field.type -eq "longstring" -or $field.type -eq "string") {
                                        if ($fieldValue -match "`n") {
                                            # Use quoted string format with \n for line breaks
                                            $escapedValue = $fieldValue.Replace('"', '""').Replace("`n", "\n")
                                            $yamlPreview += "`n    $($field.value): `"$escapedValue`""
                                        } else {
                                            $yamlPreview += "`n    $($field.value): `"$fieldValue`""
                                        }
                                    }
                                    elseif ($field.type -eq "dateString") {
                                        $yamlPreview += "`n    $($field.value): $fieldValue"
                                    }
                                    else {
                                        # Default handling
                                        $yamlPreview += "`n    $($field.value): $fieldValue"
                                    }
                                }
                            }
                        }
                    }
                }
            }

            # Handle Omissions - Dynamic version using UIConfig fields
            $omissionCount = 0
            foreach ($productName in $syncHash.Omissions.Keys) {
                $omissionCount += $syncHash.Omissions[$productName].Count
            }

            if ($omissionCount -gt 0) {
                # Get the YAML section name from UIConfig
                $omissionSectionName = $syncHash.UIConfigs.inputTypes.omissions.value
                $yamlPreview += "`n`n$omissionSectionName`:"

                # Group omissions by product
                foreach ($productName in $syncHash.Omissions.Keys) {
                    #$yamlPreview += "`n  # $productName Omissions:"

                    # Sort policies by ID
                    $sortedPolicies = $syncHash.Omissions[$productName].Keys | Sort-Object
                    foreach ($policyId in $sortedPolicies) {
                        $omissionData = $syncHash.Omissions[$productName][$policyId]

                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $policyId`:"

                        # Access the nested omission data correctly
                        # The structure is: $syncHash.Omissions[Product][PolicyId][OmissionType][FieldName]
                        # We need to get the omission type (like "Omit") first
                        foreach ($omissionType in $omissionData.Keys) {
                            $omission = $omissionData[$omissionType]

                            # Process fields in the order defined in UIConfig
                            $omissionFields = $syncHash.UIConfigs.inputTypes.omissions.fields
                            foreach ($field in $omissionFields) {
                                $fieldValue = $omission[$field.value]

                                if ($null -ne $fieldValue -and ![string]::IsNullOrEmpty($fieldValue)) {
                                    # Handle different field types
                                    if ($field.type -eq "boolean") {
                                        $yamlPreview += "`n    $($field.value): $($fieldValue.ToString().ToLower())"
                                    }
                                    elseif ($field.type -eq "longstring" -or $field.type -eq "string") {
                                        if ($fieldValue -match "`n") {
                                            # Use quoted string format with \n for line breaks
                                            $escapedValue = $fieldValue.Replace('"', '""').Replace("`n", "\n")
                                            $yamlPreview += "`n    $($field.value): `"$escapedValue`""
                                        } else {
                                            $yamlPreview += "`n    $($field.value): `"$fieldValue`""
                                        }
                                    }
                                    elseif ($field.type -eq "dateString") {
                                        $yamlPreview += "`n    $($field.value): $fieldValue"
                                    }
                                    else {
                                        # Default handling
                                        $yamlPreview += "`n    $($field.value): $fieldValue"
                                    }
                                }
                            }
                        }
                    }
                }
            }

            # Display in preview tab
            $syncHash.YamlPreview_TextBox.Text = $yamlPreview

            foreach ($tab in $syncHash.MainTabControl.Items) {
                if ($tab -is [System.Windows.Controls.TabItem] -and $tab.Header -eq "Preview") {
                    $syncHash.MainTabControl.SelectedItem = $syncHash.PreviewTab
                    break
                }
            }
        }#end function : New-YamlPreview

        #>
        #===========================================================================
        # RESET: NEW SESSION
        #===========================================================================
        Function Clear-FieldValue {
            <#
            .SYNOPSIS
            Clears all field values and resets the UI to default state.
            .DESCRIPTION
            This function resets all configuration data structures and UI controls to their initial empty state for starting a new session.
            #>

            # Clear core data structures
            $syncHash.Exclusions = @{}
            $syncHash.Omissions = @{}
            $syncHash.Annotations = @{}
            $syncHash.GeneralSettings = @{}

            # Dynamically reset all controls using configuration
            $syncHash.GetEnumerator() | ForEach-Object {
                $controlName = $_.Key
                $control = $_.Value

                if ($control -is [System.Windows.Controls.TextBox]) {
                    # First check if there's a placeholder value
                    if ($syncHash.UIConfigs.localePlaceholder.$controlName) {
                        # Reset to placeholder value with placeholder styling
                        $control.Text = $syncHash.UIConfigs.localePlaceholder.$controlName
                        $control.Foreground = [System.Windows.Media.Brushes]::Gray
                        $control.FontStyle = [System.Windows.FontStyles]::Italic
                        $control.BorderBrush = [System.Windows.Media.Brushes]::Gray
                        $control.BorderThickness = "1"
                    }
                    # Then check if there's a default value in defaultSettings
                    elseif ($syncHash.UIConfigs.defaultAdvancedSettings.$controlName) {
                        $control.Text = $syncHash.UIConfigs.defaultAdvancedSettings.$controlName
                        $control.Foreground = [System.Windows.Media.Brushes]::Black
                        $control.FontStyle = [System.Windows.FontStyles]::Normal
                        $control.BorderBrush = [System.Windows.Media.Brushes]::Gray
                        $control.BorderThickness = "1"
                    }
                    # Fallback for special cases not in config
                    else {
                        $control.Text = ""
                        $control.Foreground = [System.Windows.Media.Brushes]::Black
                        $control.FontStyle = [System.Windows.FontStyles]::Normal
                        $control.BorderBrush = [System.Windows.Media.Brushes]::Gray
                        $control.BorderThickness = "1"
                    }
                }
                elseif ($control -is [System.Windows.Controls.CheckBox]) {
                    # Check if there's a default value in defaultSettings
                    if ($syncHash.UIConfigs.defaultAdvancedSettings.PSObject.Properties.Name -contains $controlName) {
                        $control.IsChecked = $syncHash.UIConfigs.defaultAdvancedSettings.$controlName
                    }
                    # Fallback for controls not in config
                    else {
                        # Don't reset product checkboxes here - handle them separately
                        if (-not $controlName.EndsWith('ProductCheckBox')) {
                            $control.IsChecked = $false
                        }
                    }
                }
                Write-DebugOutput -Message "Cleared value for: $controlName" -Source $MyInvocation.MyCommand.Name -Level "Verbose"
            }

            # Reset specific UI elements that need special handling

            # Uncheck all product checkboxes
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }
            foreach ($checkbox in $allProductCheckboxes) {
                $checkbox.IsChecked = $false
            }

            # Reset M365 Environment to default
            $syncHash.M365Environment_ComboBox.SelectedIndex = 0

            # Reset Advanced Tab toggles (these control visibility, not data)
            $toggleControls = $syncHash.GetEnumerator() | Where-Object { $_.Name -like '*_Toggle' }
            foreach ($toggleName in $toggleControls) {
                if ($toggleName.Value -is [System.Windows.Controls.CheckBox]) {
                    $syncHash.$toggleName.IsChecked = $false
                    $contentName = $toggleName.Replace('_Toggle', '_Content')
                    if ($syncHash.$contentName) {
                        $syncHash.$contentName.Visibility = [System.Windows.Visibility]::Collapsed
                    }
                }
            }

            # Mark data as changed to trigger UI update
            Set-DataChanged

        }#end function : Clear-FieldValue

        # Function to collect general settings from UI controls
        Function Save-GeneralSettingsFromInput {
            <#
            .SYNOPSIS
            Saves general settings from UI controls to data structures.
            .DESCRIPTION
            This function collects values from UI controls and stores them in the GeneralSettings data structure for YAML export.
            #>

            # Collect ProductNames from checked checkboxes - use helper function
            #Update-ProductNames

            # Collect from localePlaceholder TextBox controls
            if ($syncHash.UIConfigs.localePlaceholder)
            {
                foreach ($placeholderKey in $syncHash.UIConfigs.localePlaceholder.PSObject.Properties.Name) {
                    $control = $syncHash.$placeholderKey
                    if ($control -is [System.Windows.Controls.TextBox]) {
                        $currentValue = $control.Text
                        $placeholderValue = $syncHash.UIConfigs.localePlaceholder.$placeholderKey

                        # Only include if it's not empty and not a placeholder
                        if (![string]::IsNullOrWhiteSpace($currentValue) -and $currentValue -ne $placeholderValue) {
                            # Convert control name to setting name (remove _TextBox suffix)
                            $settingName = $placeholderKey -replace '_TextBox$', ''
                            $syncHash.GeneralSettings[$settingName] = $currentValue.Trim()
                        }
                    }
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message "Collected Main setting: $placeholderKey = $($syncHash.GeneralSettings[$settingName])" -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                }
            }

            # Collect M365Environment
            if ($syncHash.M365Environment_ComboBox.SelectedItem) {
                $selectedEnv = $syncHash.UIConfigs.M365Environment | Where-Object { $_.id -eq $syncHash.M365Environment_ComboBox.SelectedItem.Tag } | Select-Object -ExpandProperty name
                if ($selectedEnv) {
                    $syncHash.GeneralSettings["M365Environment"] = $selectedEnv
                }
                If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message "Collected M365Environment: $selectedEnv" -Source $MyInvocation.MyCommand.Name -Level "Debug"}
            }

            # Collect from advanced settings
            if ($syncHash.UIConfigs.defaultAdvancedSettings) {
                foreach ($advancedKey in $syncHash.UIConfigs.defaultAdvancedSettings.PSObject.Properties.Name) {
                    $control = $syncHash.$advancedKey

                    if ($control -is [System.Windows.Controls.TextBox]) {
                        $currentValue = $control.Text
                        if (![string]::IsNullOrWhiteSpace($currentValue)) {
                            # Convert control name to setting name (remove _TextBox suffix)
                            $settingName = $advancedKey -replace '_TextBox$', ''
                            $syncHash.GeneralSettings[$settingName] = $currentValue.Trim()
                        }
                    }
                    elseif ($control -is [System.Windows.Controls.CheckBox]) {
                        # Convert control name to setting name (remove _CheckBox suffix)
                        $settingName = $advancedKey -replace '_CheckBox$', ''
                        $syncHash.GeneralSettings[$settingName] = $control.IsChecked
                    }
                    If($syncHash.DebugMode -match 'Debug'){Write-DebugOutput -Message "Collected Advanced setting: $settingName = $($syncHash.GeneralSettings[$settingName])" -Source $MyInvocation.MyCommand.Name -Level "Debug"}
                }
            }

            # Mark data as changed
            Set-DataChanged
        } #end function : Save-GeneralSettingsFromInput
        #===========================================================================
        # Make UI functional
        #===========================================================================
        Function Initialize-DynamicTabs {
            <#
            .SYNOPSIS
            Initializes dynamic tab content for product-specific policy configurations.
            .DESCRIPTION
            This function sets up dynamic tabs for different baseline controls, creating appropriate containers and UI elements for policy management.
            #>
            .SYNOPSIS
            Initializes dynamic tab content for product-specific policy configurations.
            .DESCRIPTION
            This function creates and configures dynamic tabs for annotations, omissions, and exclusions based on the product configurations.
            #>
            foreach ($baselineControl in $syncHash.UIConfigs.baselineControls) {
                # Skip if this is the special exclusions case (handled differently due to product support)
                if ($baselineControl.dataControlOutput -eq "Exclusions") {
                    continue  # Handle exclusions with existing logic due to supportsExclusions
                }

                # Create tab
                $tab = New-Object System.Windows.Controls.TabItem
                $tab.Name = $baselineControl.dataControlOutput + "Tab"
                $tab.Header = $baselineControl.tabName
                $tab.IsEnabled = $false

                # Create tab content with nested product tabs
                $tabContent = New-Object System.Windows.Controls.Grid
                $tabContent.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))
                $tabContent.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))

                # Header information
                $headerBorder = New-Object System.Windows.Controls.Border
                $headerBorder.Style = $syncHash.Window.FindResource("Card")
                $headerBorder.Margin = "0,0,0,16"
                [System.Windows.Controls.Grid]::SetRow($headerBorder, 0)

                $headerText = New-Object System.Windows.Controls.TextBlock
                $headerText.Name = $baselineControl.dataControlOutput + "TabInfo_TextBlock"
                $headerText.TextWrapping = "Wrap"
                $headerBorder.Child = $headerText
                [void]$tabContent.Children.Add($headerBorder)

                # Product tab control
                $productTabControl = New-Object System.Windows.Controls.TabControl
                $productTabControl.Name = $baselineControl.dataControlOutput + "ProductTabControl"
                [System.Windows.Controls.Grid]::SetRow($productTabControl, 1)
                [void]$tabContent.Children.Add($productTabControl)

                $tab.Content = $tabContent

                # Add to main tab control
                [void]$syncHash.MainTabControl.Items.Add($tab)
                $syncHash.($tab.Name) = $tab
            }
        }

        #update version
        $syncHash.Version_TextBlock.Text = "v$($syncHash.UIConfigs.Version)"

        # Show/Hide Debug tab based on DebugUI parameter
        if ($syncHash.DebugMode -ne 'None') {
            $syncHash.DebugTab.Visibility = "Visible"
            $syncHash.DebugTabInfo_TextBlock.Text = "Debug output is enabled. Real-time debugging information will appear below."
            Write-DebugOutput -Message "Debug is enabled in mode: $($syncHash.DebugMode)" -Source "UI Launch" -Level "Info"
        } else {
            $syncHash.DebugTab.Visibility = "Collapsed"
        }

        #override locale context
        foreach ($localeElement in $syncHash.UIConfigs.localeContext.PSObject.Properties) {
            $LocaleControl = $syncHash.($localeElement.Name)
            if ($LocaleControl){
                #get type of control
                switch($LocaleControl.GetType().Name) {
                    'TextBlock' {
                        $LocaleControl.Text = $localeElement.Value
                    }
                    'Button' {
                        $LocaleControl.Content = $localeElement.Value
                    }
                    'ComboBox' {
                        $LocaleControl.ToolTip = $localeElement.Value
                    }
                    'CheckBox' {
                        $LocaleControl.Content = $localeElement.Value
                    }
                    'Label' {
                        $LocaleControl.Content = $localeElement.Value
                    }
                }
            }
            Write-DebugOutput -Message "$($localeElement.Name): $($localeElement.Value)" -Source "UI Launch" -Level "Info"

        }

        $syncHash.PreviewTab.IsEnabled = $false

        #$syncHash.ExclusionsTab.IsEnabled = $false
        #$syncHash.AnnotatePolicyTab.IsEnabled = $false
        #$syncHash.OmitPolicyTab.IsEnabled = $false

        foreach ($env in $syncHash.UIConfigs.M365Environment) {
            $comboItem = New-Object System.Windows.Controls.ComboBoxItem
            $comboItem.Content = "$($env.displayName) ($($env.name))"
            $comboItem.Tag = $env.id

            $syncHash.M365Environment_ComboBox.Items.Add($comboItem)
            Write-DebugOutput -Message "M365Environment_ComboBox added: $($env.displayName) ($($env.name))" -Source "UI Launch" -Level "Info"
        }

        Add-ControlEventHandler -Control $syncHash.M365Environment_ComboBox

        # Set selection based on parameter or defa          ult to first item
        if ($syncHash.M365Environment) {
            $selectedEnv = $syncHash.M365Environment_ComboBox.Items | Where-Object { $_.Tag -eq $syncHash.M365Environment }
            if ($selectedEnv) {
                $syncHash.M365Environment_ComboBox.SelectedItem = $selectedEnv
            } else {
                # If the specified environment isn't found, default to first item
                $syncHash.M365Environment_ComboBox.SelectedIndex = 0
            }
        } else {
            # Set default selection to first item if no parameter specified
            $syncHash.M365Environment_ComboBox.SelectedIndex = 0
        }
        Write-DebugOutput -Message "M365Environment_ComboBox set: $($syncHash.M365Environment_ComboBox.SelectedItem.Content)" -Source "UI Update" -Level "Info"

        # Populate Products Checkbox dynamically within the ProductsGrid
        #only list three rows then use next column
        # Assume 3 rows, then wrap to next column
        $maxRows = 3
        for ($i = 0; $i -lt $syncHash.UIConfigs.products.Count; $i++) {
            $product = $syncHash.UIConfigs.products[$i]

            $checkBox = New-Object System.Windows.Controls.CheckBox
            $checkBox.Content = $product.displayName
            $checkBox.Name = $product.id + "ProductCheckBox"
            $checkBox.Tag = $product.id
            $checkBox.Margin = "0,5"

            $row = $i % $maxRows
            $column = [math]::Floor($i / $maxRows)

            [System.Windows.Controls.Grid]::SetRow($checkBox, $row)
            [System.Windows.Controls.Grid]::SetColumn($checkBox, $column)

            [void]$syncHash.ProductsGrid.Children.Add($checkBox)

            # Add event handlers for checked/unchecked
            $checkBox.Add_Checked({
                # Skip if this is a UI update from timer (not user action)
                if ($syncHash.BulkUpdateInProgress) {
                    return
                }

                $productId = $this.Tag

                # Only update the data - let the timer handle UI updates
                if (-not $syncHash.GeneralSettings.ProductNames) {
                    $syncHash.GeneralSettings.ProductNames = @()
                }

                # Add to GeneralSettings if not already present
                if ($syncHash.GeneralSettings.ProductNames -notcontains $productId) {
                    $syncHash.GeneralSettings.ProductNames += $productId
                    Write-DebugOutput -Message "Added [$productId] to ProductNames data" -Source "User Action" -Level "Info"
                }

                #enable the main tabs
                $syncHash.ExclusionsTab.IsEnabled = $true
                $syncHash.AnnotatePolicyTab.IsEnabled = $true
                $syncHash.OmitPolicyTab.IsEnabled = $true

                #omissions tab
                $omissionTab = $syncHash.("$($productId)OmissionTab")
                $omissionTab.IsEnabled = $true
                Write-DebugOutput -Message "Omit Policy sub tab enabled: $($productId)" -Source "UI Update" -Level "Info"

                $container = $syncHash.("$($productId)OmissionContent")
                <#
                if ($container -and $container.Children.Count -eq 0) {
                    New-ProductOmissions -ProductName $productId -Container $container
                }
                #>
                New-ProductPolicyCards -ProductName $productId -Container $container -ControlType "Omissions"


                #annotations tab
                $AnnotationTab = $syncHash.("$($productId)AnnotationTab")
                $AnnotationTab.IsEnabled = $true
                Write-DebugOutput -Message "Annotation sub tab enabled: $($productId)" -Source "UI Update" -Level "Info"

                $container = $syncHash.("$($productId)AnnotationContent")
                <#
                if ($container -and $container.Children.Count -eq 0) {
                    New-ProductAnnotations -ProductName $productId -Container $container
                }
                #>
                New-ProductPolicyCards -ProductName $productId -Container $container -ControlType "Annotations"

                #exclusions tab
                if ($product.supportsExclusions)
                {
                    $ExclusionsTab = $syncHash.("$($productId)ExclusionTab")
                    $ExclusionsTab.IsEnabled = $true
                    Write-DebugOutput -Message "Exclusion sub tab enabled: $($productId)" -Source "UI Update" -Level "Info"

                    $container = $syncHash.("$($productId)ExclusionContent")
                    <#
                    if ($container -and $container.Children.Count -eq 0) {
                        New-ProductExclusions -ProductName $productId -Container $container
                    }
                    #>
                    New-ProductPolicyCards -ProductName $productId -Container $container -ControlType "Exclusions"
                }

                # Mark data as changed for timer to pick up
                Set-DataChanged
            }.GetNewClosure())

            $checkBox.Add_Unchecked({
                # Skip if this is a UI update from timer (not user action)
                if ($syncHash.BulkUpdateInProgress) {
                    return
                }

                $productId = $this.Tag

                # Check minimum selection requirement
                if ($syncHash.GeneralSettings.ProductNames -and $syncHash.GeneralSettings.ProductNames.Count -eq 1 -and $syncHash.GeneralSettings.ProductNames -contains $productId) {
                    # This is the last selected product - prevent unchecking
                    Write-DebugOutput -Message "Prevented unchecking last product: $productId" -Source "User Action" -Level "Warning"
                    [System.Windows.MessageBox]::Show("At least one product must be selected for the configuration to be valid.", "Minimum Selection Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

                    # Set the checkbox back to checked (this won't trigger the event because of the check above)
                    $syncHash.BulkUpdateInProgress = $true
                    $this.IsChecked = $true
                    $syncHash.BulkUpdateInProgress = $false
                    return
                }

                # Remove from GeneralSettings
                if ($syncHash.GeneralSettings.ProductNames -contains $productId) {
                    $syncHash.GeneralSettings.ProductNames = $syncHash.GeneralSettings.ProductNames | Where-Object { $_ -ne $productId }
                    Write-DebugOutput -Message "Removed [$productId] from ProductNames data" -Source "User Action" -Level "Info"
                }

                # Clear data for this product
                if ($syncHash.Exclusions.ContainsKey($productId)) {
                    $syncHash.Exclusions.Remove($productId)
                    Write-DebugOutput -Message "Cleared exclusions data for: $productId" -Source "User Action" -Level "Info"
                }

                if ($syncHash.Omissions.ContainsKey($productId)) {
                    $syncHash.Omissions.Remove($productId)
                    Write-DebugOutput -Message "Cleared omissions data for: $productId" -Source "User Action" -Level "Info"
                }

                if ($syncHash.Annotations.ContainsKey($productId)) {
                    $syncHash.Annotations.Remove($productId)
                    Write-DebugOutput -Message "Cleared annotations data for: $productId" -Source "User Action" -Level "Info"
                }

                # Clear and disable tabs for this product
                $omissionTab = $syncHash.("$($productId)OmissionTab")
                $annotationTab = $syncHash.("$($productId)AnnotationTab")
                $exclusionTab = $syncHash.("$($productId)ExclusionTab")

                if ($omissionTab) {
                    $omissionTab.IsEnabled = $false
                    # Clear the container
                    $container = $syncHash.("$($productId)OmissionContent")
                    if ($container) {
                        $container.Children.Clear()
                        Write-DebugOutput -Message "Cleared omission container for: $productId" -Source "User Action" -Level "Info"
                    }
                }

                if ($annotationTab) {
                    $annotationTab.IsEnabled = $false
                    # Clear the container
                    $container = $syncHash.("$($productId)AnnotationContent")
                    if ($container) {
                        $container.Children.Clear()
                        Write-DebugOutput -Message "Cleared annotation container for: $productId" -Source "User Action" -Level "Info"
                    }
                }

                if ($exclusionTab) {
                    $exclusionTab.IsEnabled = $false
                    # Clear the container
                    $container = $syncHash.("$($productId)ExclusionContent")
                    if ($container) {
                        $container.Children.Clear()
                        Write-DebugOutput -Message "Cleared exclusion container for: $productId" -Source "User Action" -Level "Info"
                    }
                }

                # Check if any products are still selected - if not, disable main tabs
                if (-not $syncHash.GeneralSettings.ProductNames -or $syncHash.GeneralSettings.ProductNames.Count -eq 0) {
                    $syncHash.ExclusionsTab.IsEnabled = $false
                    $syncHash.AnnotatePolicyTab.IsEnabled = $false
                    $syncHash.OmitPolicyTab.IsEnabled = $false
                    Write-DebugOutput -Message "Disabled main tabs - no products selected" -Source "User Action" -Level "Info"
                }

                # Mark data as changed for timer to pick up
                Set-DataChanged
            }.GetNewClosure())

        }
        $ExclusionSupport = $syncHash.UIConfigs.products | Where-Object { $_.supportsExclusions -eq $true } | Select-Object -ExpandProperty id
        $syncHash.ExclusionsInfo_TextBlock.Text = ($syncHash.UIConfigs.localeContext.ExclusionsInfo_TextBlock -f ($ExclusionSupport -join ', ').ToUpper())

        Foreach($product in $syncHash.UIConfigs.products) {
            # Initialize the OmissionTab and ExclusionTab for each product
            $exclusionTab = $syncHash.("$($product.id)ExclusionTab")

            if ($product.supportsExclusions) {
                $exclusionTab.Visibility = "Visible"
                Write-DebugOutput -Message "Enabled Exclusion sub tab for: $($product.id)" -Source "UI" -Level "Info"
            }else{
                # Disable the Exclusions tab if the product does not support exclusions
                $exclusionTab.Visibility = "Collapsed"
                Write-DebugOutput -Message "Disabled Exclusion sub tab for: $($product.id)" -Source "UI" -Level "Info"
            }
        }

        # added events to all tab toggles
        $toggleControls = $syncHash.GetEnumerator() | Where-Object { $_.Name -like '*_Toggle' }
        foreach ($toggleName in $toggleControls) {
            $contentName = $toggleName.Name.Replace('_Toggle', '_Content')
            $contentControl = $syncHash[$contentName]

            # Add Checked event handler
            $syncHash[$toggleName.Name].Add_Checked({
                $contentControl.Visibility = "Visible"
            }.GetNewClosure())

            # Add Unchecked event handler
            $syncHash[$toggleName.Name].Add_Unchecked({
                $contentControl.Visibility = "Collapsed"
            }.GetNewClosure())
        }

        # Add event to placeholder TextBoxes
        foreach ($placeholderKey in $syncHash.UIConfigs.localePlaceholder.PSObject.Properties.Name) {
            $control = $syncHash.$placeholderKey
            if ($control -is [System.Windows.Controls.TextBox]) {
                $placeholderText = $syncHash.UIConfigs.localePlaceholder.$placeholderKey
                Initialize-PlaceholderTextBox -TextBox $control -PlaceholderText $placeholderText
            }
        }

        # Handle Organization TextBox with special Graph Connected logic
        if ($syncHash.GraphConnected) {
            try {
                $tenantDetails = (Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/organization").Value
                $tenantName = ($tenantDetails.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
                $syncHash.Organization_TextBox.Text = $tenantName
                $syncHash.Organization_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.Organization_TextBox.FontStyle = [System.Windows.FontStyles]::Normal
                $syncHash.Organization_TextBox.BorderBrush = [System.Windows.Media.Brushes]::Gray
                $syncHash.Organization_TextBox.BorderThickness = "1"
                $syncHash.Organization_TextBox.isEnabled = $false # Disable editing if Graph is connected
            } catch {
                # Fallback to placeholder if Graph request fails
                Initialize-PlaceholderTextBox -TextBox $syncHash.Organization_TextBox -PlaceholderText $syncHash.UIConfigs.localePlaceholder.Organization_TextBox
            }
        }

        #===========================================================================
        # Button Event Handlers
        #===========================================================================
        # add event handlers to all buttons
        # Add global event handlers to dynamically created remove button
        <#
        foreach($button in $syncHash.GetEnumerator() | Where-Object { $_.Value -is 'System.Windows.Controls.Button' }) {
            Add-ControlEventHandler -Control $syncHash.$($button.Name)
        }
        #>

        # New Session Button
        $syncHash.NewSessionButton.Add_Click({
            $result = [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localePopupMessages.NewSessionConfirmation, "New Session", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                # Reset all form fields
                Clear-FieldValue
            }
        })

        # Import Button
        $syncHash.ImportButton.Add_Click({
            $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
            $openFileDialog.Filter = "YAML Files (*.yaml;*.yml)|*.yaml;*.yml|All Files (*.*)|*.*"
            $openFileDialog.Title = "Import ScubaGear Configuration"

            if ($openFileDialog.ShowDialog() -eq $true) {
                try {
                    # Load and parse the YAML file
                    $yamlContent = Get-Content -Path $openFileDialog.FileName -Raw
                    $yamlHash = $yamlContent | ConvertFrom-Yaml

                    # Clear existing data
                    $syncHash.Exclusions = @{}
                    $syncHash.Omissions = @{}
                    $syncHash.Annotations = @{}
                    $syncHash.GeneralSettings = @{}

                    # Import data into the core data structures
                    Import-YamlToDataStructures -Config $yamlHash

                    # Trigger UI update through reactive system
                    Set-DataChanged

                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localePopupMessages.ImportSuccess, "Import Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
                catch {
                    [System.Windows.MessageBox]::Show("Error importing configuration: $($_.Exception.Message)", "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        })

        # Preview & Generate Button
        $syncHash.PreviewButton.Add_Click({

            $syncHash.Window.Dispatcher.Invoke([action]{
                $overallValid = $true
                $errorMessages = @()

                # Organization validation (required)
                $orgValid = Confirm-UIField -UIElement $syncHash.Organization_TextBox `
                                        -RegexPattern $syncHash.UIConfigs.valueValidations.tenantDomain.pattern `
                                        -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.OrganizationValidation `
                                        -PlaceholderText $syncHash.UIConfigs.localePlaceholder.Organization_TextBox `
                                        -Required `
                                        -ShowMessageBox:$false

                if (-not $orgValid) {
                    $overallValid = $false
                    $errorMessages += $syncHash.UIConfigs.localeErrorMessages.OrganizationValidation
                    #navigate to General tab
                    $syncHash.MainTabControl.SelectedItem = $syncHash.MainTab
                }

                # Advanced Tab Validations (only if sections are toggled on)

                # Application Section Validations
                if ($syncHash.ApplicationSection_Toggle.IsChecked) {

                    # AppID validation (GUID format)
                    $appIdValid = Confirm-UIField -UIElement $syncHash.AppId_TextBox `
                                                    -RegexPattern $syncHash.UIConfigs.valueValidations.guid.pattern `
                                                    -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.AppIdValidation `
                                                    -PlaceholderText $syncHash.UIConfigs.localePlaceholder.AppId_TextBox `
                                                    -ShowMessageBox:$false

                    if (-not $appIdValid) {
                        $overallValid = $false
                        $errorMessages += $syncHash.UIConfigs.localeErrorMessages.AppIdValidation
                        $syncHash.MainTabControl.SelectedItem = $syncHash.AdvancedTab
                    }

                    # Certificate Thumbprint validation (40 character hex)
                    $certValid = Confirm-UIField -UIElement $syncHash.CertificateThumbprint_TextBox `
                                                -RegexPattern $syncHash.UIConfigs.valueValidations.thumbprint.pattern `
                                                -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.CertificateValidation `
                                                -PlaceholderText $syncHash.UIConfigs.localePlaceholder.CertificateThumbprint_TextBox `
                                                -ShowMessageBox:$false

                    if (-not $certValid) {
                        $overallValid = $false
                        $errorMessages += $syncHash.UIConfigs.localeErrorMessages.CertificateValidation
                        $syncHash.MainTabControl.SelectedItem = $syncHash.AdvancedTab
                    }
                }

                # Show consolidated error message if there are validation errors
                if (-not $overallValid) {
                    $syncHash.PreviewTab.IsEnabled = $false
                    $consolidatedMessage = $syncHash.UIConfigs.localeErrorMessages.PreviewValidation + "`n`n" + ($errorMessages -join "`n")
                    [System.Windows.MessageBox]::Show($consolidatedMessage, "Validation Errors", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                }else {
                    $syncHash.PreviewTab.IsEnabled = $true
                }

                if ($overallValid) {
                    New-YamlPreview
                }
            }) #end Dispatcher.Invoke
        })

        # Copy to Clipboard Button
        $syncHash.CopyYamlButton.Add_Click({
            try {
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    if (![string]::IsNullOrWhiteSpace($syncHash.YamlPreview_TextBox.Text)) {
                        [System.Windows.Clipboard]::SetText($syncHash.YamlPreview_TextBox.Text)
                        [System.Windows.MessageBox]::Show("YAML preview copied to clipboard successfully.", "Copy Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    } else {
                        [System.Windows.MessageBox]::Show("No YAML preview to copy.", "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    }
                })
            }
            catch {
                # Even this must go in Dispatcher
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    [System.Windows.MessageBox]::Show("Error copying YAML preview to clipboard: $($_.Exception.Message)", "Copy Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                })
            }
        })

        # Download YAML Button
        $syncHash.DownloadYamlButton.Add_Click({
            try {
                if ([string]::IsNullOrWhiteSpace($syncHash.YamlPreview_TextBox.Text)) {
                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeErrorMessages.DownloadNullError, "Nothing to Download", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }

                # Generate filename based on organization name
                $orgName = $syncHash.Organization_TextBox.Text
                if ([string]::IsNullOrWhiteSpace($orgName) -or $orgName -eq $syncHash.UIConfigs.localePlaceholder.Organization_TextBox) {
                    $filename = "ScubaGear-Config.yaml"
                } else {
                    # Remove any invalid filename characters and use organization name
                    $cleanOrgName = $orgName -replace '[\\/:*?"<>|]', '_'
                    $filename = "$cleanOrgName.yaml"
                }

                # Create SaveFileDialog
                $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveFileDialog.Filter = "YAML Files (*.yaml)|*.yaml|All Files (*.*)|*.*"
                $saveFileDialog.Title = "Save ScubaGear Configuration"
                $saveFileDialog.FileName = $filename
                $saveFileDialog.DefaultExt = ".yaml"

                if ($saveFileDialog.ShowDialog() -eq $true) {
                    # Save the YAML content to file
                    $yamlContent = $syncHash.YamlPreview_TextBox.Text
                    [System.IO.File]::WriteAllText($saveFileDialog.FileName, $yamlContent, [System.Text.Encoding]::UTF8)
                    #$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
                    #[System.IO.File]::WriteAllText($saveFileDialog.FileName, $yamlContent, $utf8NoBom)
                    #$yamlContent | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8NoBOM

                    [System.Windows.MessageBox]::Show("Configuration saved successfully to: $($saveFileDialog.FileName)", "Save Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Error saving file: $($_.Exception.Message)", "Save Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })

        # Browse Output Path Button
        $syncHash.BrowseOutPathButton.Add_Click({
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select Output Path"
            $folderDialog.ShowNewFolderButton = $true

            if ($syncHash.OutPath_TextBox.Text -ne "." -and (Test-Path $syncHash.OutPath_TextBox.Text)) {
                $folderDialog.SelectedPath = $syncHash.OutPath_TextBox.Text
            }

            $result = $folderDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $syncHash.OutPath_TextBox.Text = $folderDialog.SelectedPath
                #New-YamlPreview
            }
        })

        # Browse OPA Path Button
        $syncHash.BrowseOpaPathButton.Add_Click({
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select OPA Path"
            $folderDialog.ShowNewFolderButton = $true

            if ($syncHash.OpaPath_TextBox.Text -ne "." -and (Test-Path $syncHash.OpaPath_TextBox.Text)) {
                $folderDialog.SelectedPath = $syncHash.OpaPath_TextBox.Text
            }

            $result = $folderDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $syncHash.OpaPath_TextBox.Text = $folderDialog.SelectedPath
                #New-YamlPreview
            }
        })

        # Select Certificate Button
        $syncHash.SelectCertificateButton.Add_Click({
            try {
                # Get user certificates with better error handling
                $userCerts = @()
                try {
                    $userCerts = Get-ChildItem -Path "Cert:\CurrentUser\My" -ErrorAction Stop | Where-Object {
                        $_.HasPrivateKey -and
                        $_.NotAfter -gt (Get-Date) -and
                        $_.Subject -notlike "*Microsoft*"
                    } | Sort-Object Subject
                }
                catch {
                    Write-DebugOutput -Message "Error accessing certificate store: $($_.Exception.Message)" -Source "Certificate Selection" -Level "Error"
                    [System.Windows.MessageBox]::Show("Error accessing certificate store: $($_.Exception.Message)", "Certificate Store Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }

                Write-DebugOutput -Message "Found $($userCerts.Count) certificates" -Source "Certificate Selection" -Level "Info"

                if ($userCerts.Count -eq 0) {
                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeErrorMessages.CertificateNotFound,
                                                    "No Certificates",
                                                    [System.Windows.MessageBoxButton]::OK,
                                                    [System.Windows.MessageBoxImage]::Information)
                    return
                }

                # Prepare data for display
                $displayCerts = $userCerts | ForEach-Object {
                    [PSCustomObject]@{
                        Subject = $_.Subject
                        Issuer = $_.Issuer
                        NotAfter = $_.NotAfter.ToString("yyyy-MM-dd")
                        Thumbprint = $_.Thumbprint
                        Certificate = $_
                    }
                } | Sort-Object Subject

                # Column configuration
                $columnConfig = [ordered]@{
                    Thumbprint = @{ Header = "Thumbprint"; Width = 120 }
                    Subject = @{ Header = "Subject"; Width = 250 }
                    Issuer = @{ Header = "Issued By"; Width = 200 }
                    NotAfter = @{ Header = "Expires"; Width = 100 }
                }

                # Show selector (single selection only for certificates)
                $selectedThumbprint = Show-UISelectionWindow `
                                    -WindowWidth 740 `
                                    -Title "Select Certificate" `
                                    -SearchPlaceholder "Search by subject..." `
                                    -Items $displayCerts `
                                    -ColumnConfig $columnConfig `
                                    -DisplayOrder $columnConfig.Keys `
                                    -SearchProperty "Subject" `
                                    -ReturnProperty "Thumbprint"

                $syncHash.CertificateThumbprint_TextBox.Text = $selectedThumbprint
                Write-DebugOutput -Message "Selected certificate thumbprint: $selectedThumbprint" -Source "Certificate Selection" -Level "Info"
            }
            catch {
                Write-DebugOutput -Message "Certificate selection error: $($_.Exception.Message)" -Source "Certificate Selection" -Level "Error"
                [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeErrorMessages.WindowError,
                                                "Error",
                                                [System.Windows.MessageBoxButton]::OK,
                                                [System.Windows.MessageBoxImage]::Error)

            }
        })

        # Copy Debug Logs Button
        $syncHash.CopyDebugLogsButton.Add_Click({
            try {
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    if (![string]::IsNullOrWhiteSpace($syncHash.Debug_TextBox.Text)) {
                        [System.Windows.Clipboard]::SetText($syncHash.Debug_TextBox.Text)
                        [System.Windows.MessageBox]::Show("Debug logs copied to clipboard successfully.", "Copy Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    } else {
                        [System.Windows.MessageBox]::Show("No debug logs to copy.", "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                    }
                })
            }
            catch {
                # Even this must go in Dispatcher
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    [System.Windows.MessageBox]::Show("Error copying debug logs to clipboard: $($_.Exception.Message)", "Copy Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                })
            }
        })
        #=======================================
        # CLOSE UI
        #=======================================

        #Add smooth closing for Window
        $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
    	$syncHash.Window.Add_Loaded({
            $syncHash.isLoaded = $true
            Write-Host "Window loaded successfully" -ForegroundColor Green
        })

        # Closing event - called when user closes window or Close() is called
        $syncHash.Window.Add_Closing({
            #param($sender, $e)

            # Only run cleanup if not already closing
            if (-not $syncHash.isClosing) {
                $syncHash.isClosing = $true

                # Run cleanup but DON'T call window.Close() again
                try {
                    # Run memory cleanup
                    if ($syncHash.InvokeMemoryCleanup) {
                        $syncHash.InvokeMemoryCleanup.Invoke($true)
                    }

                    # Run cleanup but skip the window closing part
                    Close-UIMainWindow

                } catch {
                    Write-Error "Error during closing cleanup: $($_.Exception.Message)"
                }
            }
        })

        # Closed event - called after window is fully closed
        $syncHash.Window.Add_Closed({

            # Final cleanup
            try {
                if ($syncHash.UIUpdateTimer) {
                    $syncHash.UIUpdateTimer.Stop()
                    $syncHash.UIUpdateTimer = $null
                }
                if ($syncHash.DebugFlushTimer) {
                    $syncHash.DebugFlushTimer.Stop()
                    $syncHash.DebugFlushTimer.Dispose()
                    $syncHash.DebugFlushTimer = $null
                }

                $syncHash.isClosed = $true

                # Final aggressive memory cleanup
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                [System.GC]::Collect()

            } catch {
                Write-Error "Error in final cleanup: $($_.Exception.Message)"
            }
        })

        #always force windows on bottom
        $syncHash.Window.Topmost = $True



        $syncHash.UIUpdateTimer.Start()

        #Allow UI to be dragged around screen
        <#
        $syncHash.Window.Add_MouseLeftButtonDown( {
            $syncHash.Window.DragMove()
        })
        #>

        #action for exit button
        <#
        $syncHash.btnExit.Add_Click({
            Close-UIMainWindow
        })
        #>

        $syncHash.Window.ShowDialog()
        #$Runspace.Close()
        #$Runspace.Dispose()
        $syncHash.Error = $Error
    }) # end scriptblock

    #collect data from runspace
    $Data = $syncHash
    #invoke scriptblock in runspace
    $PowerShellCommand.Runspace = $Runspace
    $AsyncHandle = $PowerShellCommand.BeginInvoke()


    #wait until runspace is completed before ending
    #do {
    #    Start-sleep -m 100 }
    #while (!$AsyncHandle.IsCompleted)
    #end invoked process
    $null = $PowerShellCommand.EndInvoke($AsyncHandle)

    If($Passthru){
        return $Data
    }

}


Export-ModuleMember -Function @(
    'Invoke-SCuBAConfigAppUI'
)
