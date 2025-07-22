class ScubaConfig {
    <#
    .SYNOPSIS
    This singleton class stores Scuba config data loaded from a file.
    .DESCRIPTION
    This class is designed to function as a singleton. The singleton instance
    is cached on the ScubaConfig type itself. In the context of tests, it may be
    important to call `.ResetInstance` before and after tests as needed to
    ensure any preexisting configs are not inadvertantly used for the test,
    or left in place after the test is finished. The singleton will persist
    for the life of the powershell session unless the ScubaConfig module is
    removed. Note that `.LoadConfig` internally calls `.ResetInstance` to avoid
    issues.
    .EXAMPLE
    $Config = [ScubaConfig]::GetInstance()
    [ScubaConfig]::LoadConfig($SomePath)
    #>
    hidden static [ScubaConfig]$_Instance = [ScubaConfig]::new()
    hidden static [Boolean]$_IsLoaded = $false
    hidden static [hashtable]$ScubaDefaults = @{
        DefaultOPAPath = try {Join-Path -Path $env:USERPROFILE -ChildPath ".scubagear\Tools"} catch {"."};
        DefaultProductNames = @("aad", "defender", "exo", "sharepoint", "teams")
        AllProductNames = @("aad", "defender", "exo", "powerplatform", "sharepoint", "teams")
        DefaultM365Environment = "commercial"
        DefaultLogIn = $true
        DefaultOutPath = Get-Location | Select-Object -ExpandProperty ProviderPath
        DefaultOutFolderName = "M365BaselineConformance"
        DefaultOutProviderFileName = "ProviderSettingsExport"
        DefaultOutRegoFileName = "TestResults"
        DefaultOutReportName = "BaselineReports"
        DefaultOutJsonFileName = "ScubaResults"
        DefaultOutCsvFileName = "ScubaResults"
        DefaultOutActionPlanFileName = "ActionPlan"
        DefaultNumberOfUUIDCharactersToTruncate = 18
        DefaultPrivilegedRoles = @(
            "Global Administrator",
            "Privileged Role Administrator",
            "User Administrator",
            "SharePoint Administrator",
            "Exchange Administrator",
            "Hybrid Identity Administrator",
            "Application Administrator",
            "Cloud Application Administrator")
        DefaultOPAVersion = '1.6.0'
    }

    static [object]ScubaDefault ([string]$Name){
        return [ScubaConfig]::ScubaDefaults[$Name]
    }

    static [string]GetOpaVersion() {
        return [ScubaConfig]::ScubaDefault('DefaultOPAVersion')
    }

    [Boolean]LoadConfig([System.IO.FileInfo]$Path){
        if (-Not (Test-Path -PathType Leaf $Path)){
            throw [System.IO.FileNotFoundException]"Failed to load: $Path"
        }
        [ScubaConfig]::ResetInstance()
        $Content = Get-Content -Raw -Path $Path
        try {
            $this.Configuration = $Content | ConvertFrom-Yaml
        }
        catch {
            $ParseError = $($_.Exception.Message) -Replace '^Exception calling "Load" with "1" argument\(s\): ', ''
            throw "Error loading config file: $ParseError"
        }

        $this.SetParameterDefaults()
        [ScubaConfig]::_IsLoaded = $true

        # If OmitPolicy was included in the config file, validate the policy IDs included there.
        if ($this.Configuration.ContainsKey("OmitPolicy")) {
            foreach ($Policy in $this.Configuration.OmitPolicy.Keys) {
                if (-not ($Policy -match "^ms\.[a-z]+\.[0-9]+\.[0-9]+v[0-9]+$")) {
                    # Note that -match is a case insensitive match
                    # Note that the regex does not validate the product name, this will be done later
                    $Warning = "Config file indicates omitting $Policy, but $Policy is not a valid control ID. "
                    $Warning += "Expected format is 'MS.[PRODUCT].[GROUP].[NUMBER]v[VERSION]', "
                    $Warning += "e.g., 'MS.DEFENDER.1.1v1'. Control will not be omitted."
                    Write-Warning $Warning
                    Continue
                }
                $Product = ($Policy -Split "\.")[1]
                # Here's where the product name is validated
                if (-not ($this.Configuration.ProductNames -Contains $Product)) {
                    $Warning = "Config file indicates omitting $Policy, but $Product is not one of the products "
                    $Warning += "specified in the ProductNames parameter. Control will not be omitted."
                    Write-Warning $Warning
                    Continue
                }
            }
        }

        # If AnnotatePolicy was included in the config file, validate the policy IDs included there.
        if ($this.Configuration.ContainsKey("AnnotatePolicy")) {
            foreach ($Policy in $this.Configuration.AnnotatePolicy.Keys) {
                if (-not ($Policy -match "^ms\.[a-z]+\.[0-9]+\.[0-9]+v[0-9]+$")) {
                    # Note that -match is a case insensitive match
                    # Note that the regex does not validate the product name, this will be done later
                    $Warning = "Config file adds annotation for $Policy, "
                    $Warning += "but $Policy is not a valid control ID. "
                    $Warning += "Expected format is 'MS.[PRODUCT].[GROUP].[NUMBER]v[VERSION]', "
                    $Warning += "e.g., 'MS.DEFENDER.1.1v1'."
                    Write-Warning $Warning
                    Continue
                }
                $Product = ($Policy -Split "\.")[1]
                # Here's where the product name is validated
                if (-not ($this.Configuration.ProductNames -Contains $Product)) {
                    $Warning = "Config file adds annotation for $Policy, "
                    $Warning += "but $Product is not one of the products "
                    $Warning += "specified in the ProductNames parameter."
                    Write-Warning $Warning
                    Continue
                }
            }
        }

        return [ScubaConfig]::_IsLoaded
    }

    hidden [void]ClearConfiguration(){
        $this.Configuration = $null
    }

    hidden [Guid]$Uuid = [Guid]::NewGuid()
    hidden [hashtable]$Configuration

    hidden [void]SetParameterDefaults(){
        Write-Debug "Setting ScubaConfig default values."
        if (-Not $this.Configuration.ProductNames){
            $this.Configuration.ProductNames = [ScubaConfig]::ScubaDefault('DefaultProductNames')
        }
        else{
            # Transform ProductNames into list of all products if it contains wildcard
            if ($this.Configuration.ProductNames.Contains('*')){
                $this.Configuration.ProductNames = [ScubaConfig]::ScubaDefault('AllProductNames')
                Write-Debug "Setting ProductNames to all products because of wildcard"
            }
            else{
                Write-Debug "ProductNames provided - using as is."
                $this.Configuration.ProductNames = $this.Configuration.ProductNames | Sort-Object -Unique
            }
        }

        if (-Not $this.Configuration.M365Environment){
            $this.Configuration.M365Environment = [ScubaConfig]::ScubaDefault('DefaultM365Environment')
        }

        if (-Not $this.Configuration.OPAPath){
            $this.Configuration.OPAPath = [ScubaConfig]::ScubaDefault('DefaultOPAPath')
        }

        if (-Not $this.Configuration.LogIn){
            $this.Configuration.LogIn = [ScubaConfig]::ScubaDefault('DefaultLogIn')
        }

        if (-Not $this.Configuration.DisconnectOnExit){
            $this.Configuration.DisconnectOnExit = $false
        }

        if (-Not $this.Configuration.OutPath){
            $this.Configuration.OutPath = [ScubaConfig]::ScubaDefault('DefaultOutPath')
        }

        if (-Not $this.Configuration.OutFolderName){
            $this.Configuration.OutFolderName = [ScubaConfig]::ScubaDefault('DefaultOutFolderName')
        }

        if (-Not $this.Configuration.OutProviderFileName){
            $this.Configuration.OutProviderFileName = [ScubaConfig]::ScubaDefault('DefaultOutProviderFileName')
        }

        if (-Not $this.Configuration.OutRegoFileName){
            $this.Configuration.OutRegoFileName = [ScubaConfig]::ScubaDefault('DefaultOutRegoFileName')
        }

        if (-Not $this.Configuration.OutReportName){
            $this.Configuration.OutReportName = [ScubaConfig]::ScubaDefault('DefaultOutReportName')
        }

        if (-Not $this.Configuration.OutJsonFileName){
            $this.Configuration.OutJsonFileName = [ScubaConfig]::ScubaDefault('DefaultOutJsonFileName')
        }

        if (-Not $this.Configuration.OutCsvFileName){
            $this.Configuration.OutCsvFileName = [ScubaConfig]::ScubaDefault('DefaultOutCsvFileName')
        }

        if (-Not $this.Configuration.OutActionPlanFileName){
            $this.Configuration.OutActionPlanFileName = [ScubaConfig]::ScubaDefault('DefaultOutActionPlanFileName')
        }

        if (-Not $this.Configuration.NumberOfUUIDCharactersToTruncate){
            $this.Configuration.NumberOfUUIDCharactersToTruncate = [ScubaConfig]::ScubaDefault('DefaultNumberOfUUIDCharactersToTruncate')
        }
        return
    }

    hidden ScubaConfig(){
    }

    static [void]ResetInstance(){
        [ScubaConfig]::_Instance.ClearConfiguration()
        [ScubaConfig]::_IsLoaded = $false
        return
    }

    static [ScubaConfig]GetInstance(){
        return [ScubaConfig]::_Instance
    }
}

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
            Write-Warning "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
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

        # Store Form Objects In PowerShell
        $UIXML.SelectNodes("//*[@Name]") | ForEach-Object{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

        #Import UI configuration file
        $syncHash.UIConfigs = Get-Content -Path $syncHash.UIConfigPath -Raw | ConvertFrom-Json
        Write-DebugOutput -Message "UIConfigs loaded: $($syncHash.UIConfigPath)" -Source "UI Launch" -Level "Info"

        $syncHash.DebugMode = $syncHash.UIConfigs.DebugMode

        # If YAMLImport is specified, load the YAML configuration
        If($syncHash.YAMLImport){
            $syncHash.YAMLConfig = Get-Content -Path $syncHash.YAMLImport -Raw | ConvertFrom-Yaml
            Write-DebugOutput -Message "YAMLConfig loaded: $($syncHash.YAMLImport)" -Source "UI Launch" -Level "Info"
        }

        function Write-DebugOutput {
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

        $syncHash.Window.Dispatcher.Invoke([Action]{
            try {
                $syncHash.Debug_TextBox.AppendText("UI START`r`n")
                $syncHash.Debug_TextBox.ScrollToEnd()
            } catch {
                Write-Warning "Dispatcher error: $($_.Exception.Message)"
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
                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Data changes detected - updating UI displays" -Source "UI Timer" -Level "Verbose"}
                    Update-AllUIFromData
                    $syncHash.DataChanged = $false
                }

            } catch {
                If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Error in UI update timer: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"}
            }
        })


        # Initialize change tracking
        $syncHash.DataChanged = $false

        Write-DebugOutput -Message "UI initialization started - creating timer and data structures" -Source "UI Initialization" -Level "Info"

        $syncHash.LastUpdateHash = @{
            Exclusions = ""
            Omissions = ""
            Annotations = ""
            GeneralSettings = ""
        }


        # Function to mark data as changed
        Function Set-DataChanged {
            $syncHash.DataChanged = $true
        }

        # Recursively find all controls
        function Find-AllControls {
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
                $results += Find-AllControls -Parent $child
            }

            return $results
        }

        function Add-ControlEventHandlers {
            param(
                [System.Windows.Controls.Control]$Control
            )

            switch ($Control.GetType().Name) {
                'CheckBox' {
                    $Control.Add_Checked({
                        $name = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                        Write-DebugOutput "User checked: $name" -Source "Global Event Handler" -Level "Info"
                    }.GetNewClosure())

                    $Control.Add_Unchecked({
                        $name = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                        Write-DebugOutput "User unchecked: $name" -Source "Global Event Handler" -Level "Info"
                    }.GetNewClosure())
                }
                'Button' {
                    $Control.Add_Click({
                        $name = if ($this.Name) { $this.Name } else { "Unnamed Button" }
                        Write-DebugOutput "User clicked: $name" -Source "Global Event Handler" -Level "Info"
                    }.GetNewClosure())
                }
                'TextBox' {
                    $Control.Add_LostFocus({
                        $name = if ($this.Name) { $this.Name } else { "Unnamed TextBox" }
                        $value = $this.Text
                        Write-DebugOutput "User changed text in: $name => $value" -Source "Global Event Handler" -Level "Info"
                    }.GetNewClosure())
                }
                default {
                    # Do nothing for other types
                }
            }
        }

        <#
        Find-AllControls -Parent $syncHash.Window

        Function Add-GlobalEventHandlers {
            # Add event handlers to all checkboxes
            foreach ($checkbox in $allCheckboxes) {
                # Add Checked event
                $checkbox.Add_Checked({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                    $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                    Write-DebugOutput -Message "User checked checkbox: $controlName$controlTag" -Source "Global Event Handler" -Level "Info"
                }.GetNewClosure())

                # Add Unchecked event
                $checkbox.Add_Unchecked({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                    $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                    Write-DebugOutput -Message "User unchecked checkbox: $controlName$controlTag" -Source "Global Event Handler" -Level "Info"
                }.GetNewClosure())
            }

            # Add event handlers to all buttons
            foreach ($button in $allButtons) {
                # Add Click event
                $button.Add_Click({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed Button" }
                    $controlContent = if ($this.Content) { " ($($this.Content))" } else { "" }
                    Write-DebugOutput -Message "User clicked button: $controlName$controlContent" -Source "Global Event Handler" -Level "Info"
                }.GetNewClosure())
            }

            Write-DebugOutput -Message "Global event handlers added - CheckBoxes: $($allCheckboxes.Count), Buttons: $($allButtons.Count)" -Source "UI Initialization" -Level "Info"
        }

        # Function to add event handlers to a specific control (for dynamically created controls)
        Function Add-ControlEventHandlers {
            param(
                [System.Windows.Controls.Control]$Control
            )

            if ($Control -is [System.Windows.Controls.CheckBox]) {
                # Add Checked event
                $Control.Add_Checked({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                    $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                    Write-DebugOutput -Message "User checked checkbox: $controlName$controlTag" -Source "Checkbox Action" -Level "Info"
                }.GetNewClosure())

                # Add Unchecked event
                $Control.Add_Unchecked({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed CheckBox" }
                    $controlTag = if ($this.Tag) { " (Tag: $($this.Tag))" } else { "" }
                    Write-DebugOutput -Message "User unchecked checkbox: $controlName$controlTag" -Source "Checkbox Action" -Level "Info"
                }.GetNewClosure())
            }
            elseif ($Control -is [System.Windows.Controls.Button]) {
                # Add Click event
                $Control.Add_Click({
                    $controlName = if ($this.Name) { $this.Name } else { "Unnamed Button" }
                    $controlContent = if ($this.Content) { " ($($this.Content))" } else { "" }
                    Write-DebugOutput -Message "User clicked button: $controlName$controlContent" -Source "Button Action" -Level "Info"
                }.GetNewClosure())
            }
        }

        #>
        #===========================================================================
        # UI Helper Functions
        #===========================================================================
        #Closes UI objects and exits (within runspace)

        Function Close-UIMainWindow
        {
            if ($syncHash.hadCritError) { Write-Error -Message ("Critical error occurred, closing UI: {0}" -f $syncHash.Error) }
            #if runspace has not errored Dispose the UI
            if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
        }



        # Function to initialize placeholder text behavior for TextBox controls
        Function Initialize-PlaceholderTextBox {
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

            # Add GotFocus event handler
            $TextBox.Add_GotFocus({
                if ($this.Text -eq $PlaceholderText) {
                    $this.Text = ""
                    $this.Foreground = [System.Windows.Media.Brushes]::Black
                    $this.FontStyle = [System.Windows.FontStyles]::Normal
                }
            }.GetNewClosure())

            # Add LostFocus event handler
            $TextBox.Add_LostFocus({
                if ([string]::IsNullOrWhiteSpace($this.Text)) {
                    $this.Text = $PlaceholderText
                    $this.Foreground = [System.Windows.Media.Brushes]::Gray
                    $this.FontStyle = [System.Windows.FontStyles]::Italic
                }
            }.GetNewClosure())

        }

        # Helper function to find control by setting name
        Function Find-ControlBySettingName {
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
            param(
                [object]$Control,
                [object]$Value,
                [string]$SettingKey
            )

            switch ($Control.GetType().Name) {
                'TextBox' {
                    $Control.Text = $Value
                    $Control.Foreground = [System.Windows.Media.Brushes]::Black
                    $Control.FontStyle = [System.Windows.FontStyles]::Normal
                }
                'TextBlock' {
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
                    #this would update values in the syncHash directly
                    $syncHash.$Control = $Value
                }
                default {
                    Write-DebugOutput -Message "Unknown control type for $SettingKey`: $($Control.GetType().Name)" -Source $MyInvocation.MyCommand.Name -Level "Warning"
                }
            }
        }

        # Helper function to update ComboBox values
        Function Set-ComboBoxValue {
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
                    Write-DebugOutput -Message "Selected M365Environment: $Value" -Source $MyInvocation.MyCommand.Name -Level "Info"
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
                Write-DebugOutput -Message "Set ComboBox [$SettingKey] to value: $Value" -Source $MyInvocation.MyCommand.Name -Level "Info"
            } else {
                Write-DebugOutput -Message "Could not find ComboBox [$SettingKey] with value: $Value" -Source $MyInvocation.MyCommand.Name -Level "Warning"
            }
        }

        # Function to validate UI field based on regex and required status
        Function Confirm-UIField {
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
                        If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "No UI control found for setting: $settingKey" -Source $MyInvocation.MyCommand.Name -Level "Warning"}
                    }
                }
            }
            catch {
                If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Error updating general settings UI: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"}
            }
        }

        Function Update-ProductNameCheckboxFromData{
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
                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Selecting all products due to '*' value" -Source $MyInvocation.MyCommand.Name -Level "Info"}
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
                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Disabled main tabs - no products selected" -Source $MyInvocation.MyCommand.Name -Level "Info"}
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
                                If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Created omission content for: $productId" -Source $MyInvocation.MyCommand.Name -Level "Info"}
                            }
                        }

                        # Enable and ensure content exists for annotations
                        $annotationTab = $syncHash.("$($productId)AnnotationTab")
                        if ($annotationTab) {
                            $annotationTab.IsEnabled = $true
                            $container = $syncHash.("$($productId)AnnotationContent")
                            if ($container -and $container.Children.Count -eq 0) {
                                New-ProductAnnotations -ProductName $productId -Container $container
                                If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Created annotation content for: $productId" -Source $MyInvocation.MyCommand.Name -Level "Info"}
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
                                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Created exclusion content for: $productId" -Source $MyInvocation.MyCommand.Name -Level "Info"}
                                }
                            }
                        }

                        If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Enabled tabs and ensured content for: $productId" -Source $MyInvocation.MyCommand.Name -Level "Info"}
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

            If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Updated checkboxes and tabs for products: [$($productsToSelect -join ', ')]" -Source $MyInvocation.MyCommand.Name -Level "Info"}
        }

        # Updated Update-ExclusionsFromData Function for hashtable structure
        Function Update-ExclusionsFromData {
            if (-not $syncHash.Exclusions) { return }

            # Iterate through products and policies in hashtable structure
            foreach ($productName in $syncHash.Exclusions.Keys) {
                foreach ($policyId in $syncHash.Exclusions[$productName].Keys) {
                    try {
                        # Find the exclusion type from the baseline config
                        $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                        if ($baseline -and $baseline.exclusionType -ne "none") {

                            # Find the existing card checkbox
                            $checkboxName = ($policyId.replace('.', '_') + "_ExclusionCheckbox")
                            $checkbox = $syncHash.$checkboxName

                            if ($checkbox) {
                                # Mark as checked
                                $checkbox.IsChecked = $true

                                # Get exclusion data for this policy
                                $exclusionData = $syncHash.Exclusions[$productName][$policyId]

                                # Iterate through exclusion types (YAML key names)
                                foreach ($yamlKeyName in $exclusionData.Keys) {
                                    $fieldData = $exclusionData[$yamlKeyName]

                                    # Get the exclusion type configuration
                                    $exclusionTypeConfig = $syncHash.UIConfigs.exclusionTypes.($baseline.exclusionType)

                                    if ($exclusionTypeConfig) {
                                        # Populate the exclusion data fields based on exclusion type configuration
                                        foreach ($field in $exclusionTypeConfig.fields) {
                                            $fieldName = $field.name
                                            $controlName = ($policyId.replace('.', '_') + "_" + $baseline.exclusionType + "_" + $fieldName)

                                            if ($fieldData.Keys -contains $fieldName) {
                                                $fieldValue = $fieldData[$fieldName]

                                                if ($field.type -eq "array" -and $fieldValue -is [array]) {
                                                    # Handle array fields
                                                    $listContainer = $syncHash.($controlName + "_List")
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
                                                                Write-DebugOutput -Message "User removed item: $item" -Source $listContainer -Level "Info"
                                                            }.GetNewClosure())

                                                            [void]$itemPanel.Children.Add($itemText)
                                                            [void]$itemPanel.Children.Add($removeBtn)
                                                            [void]$listContainer.Children.Add($itemPanel)
                                                        }
                                                    }
                                                } else {
                                                    # Handle single value fields
                                                    $control = $syncHash.($controlName + "_TextBox")
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
                        Write-DebugOutput -Message "Error updating exclusion UI for $policyId`: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }
                }
            }
        }

        # Updated Update-OmissionsFromData Function
        Function Update-OmissionsFromData {
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
                        Write-DebugOutput -Message "Error updating omission UI for $policyId`: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                    }
                }
            }
        }

        # Updated Update-AnnotationsFromData Function
        Function Update-AnnotationsFromData {
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
                        Write-DebugOutput -Message "Error updating annotation UI for $policyId`: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
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

            Write-DebugOutput -Message "Updated ProductNames in GeneralSettings: [$($selectedProducts -join ', ')]" -Source $MyInvocation.MyCommand.Name -Level "Info"
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
                Write-DebugOutput -Message "No ProductNames selected, returning empty array for YAML" -Source $MyInvocation.MyCommand.Name -Level "Info"
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
                Write-DebugOutput -Message "All products selected, returning '*' for YAML output" -Source $MyInvocation.MyCommand.Name -Level "Info"
                return "`nProductNames: ['*']"
            } else {
                Write-DebugOutput -Message "Returning specific product list for YAML: [$($selectedProducts -join ', ')]" -Source $MyInvocation.MyCommand.Name -Level "Info"
                Return ("`nProductNames: " + ($selectedProducts | ForEach-Object { "`n  - $_" }) -join '')
            }
        }

        #===========================================================================
        #
        # DYNAMIC ELEMENT FUNCTIONS
        #
        #===========================================================================

        #===========================================================================
        # ANNOTATION dynamic controls
        #===========================================================================
        # Function to create an annotation card UI element
        Function New-AnnotationCard {
            param(
                [string]$PolicyId,
                [string]$ProductName,
                [string]$PolicyName,
                [string]$PolicyDescription
            )

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
            $checkbox.Name = ($PolicyId.replace('.', '_') + "_AnnotationCheckbox")
            $checkbox.VerticalAlignment = "Top"
            $checkbox.Margin = "0,0,12,0"
            [System.Windows.Controls.Grid]::SetColumn($checkbox, 0)

            # Add global event handlers to dynamically created checkbox
            Add-ControlEventHandlers -Control $checkbox

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

            # Policy description
            $policyDesc = New-Object System.Windows.Controls.TextBlock
            $policyDesc.Text = $PolicyDescription
            $policyDesc.FontSize = 11
            $policyDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
            $policyDesc.TextWrapping = "Wrap"
            $policyDesc.Margin = "0,0,0,4"
            [void]$policyInfoStack.Children.Add($policyDesc)

            # Add elements to header grid
            [void]$headerGrid.Children.Add($checkbox)
            [void]$headerGrid.Children.Add($policyInfoStack)

            # Create details panel (initially collapsed)
            $detailsPanel = New-Object System.Windows.Controls.StackPanel
            $detailsPanel.Visibility = "Collapsed"
            $detailsPanel.Margin = "24,12,0,0"
            [System.Windows.Controls.Grid]::SetRow($detailsPanel, 1)

            # Comment label
            $commentLabel = New-Object System.Windows.Controls.TextBlock
            $commentLabel.Text = "Comment:"
            $commentLabel.FontWeight = "SemiBold"
            $commentLabel.Margin = "0,0,0,4"
            [void]$detailsPanel.Children.Add($commentLabel)

            # Comment TextBox
            $commentTextBox = New-Object System.Windows.Controls.TextBox
            $commentTextBox.Name = ($PolicyId.replace('.', '_') + "_Comment_TextBox")
            $commentTextBox.Height = 80
            $commentTextBox.TextWrapping = "Wrap"
            $commentTextBox.AcceptsReturn = $true
            $commentTextBox.VerticalScrollBarVisibility = "Auto"
            $commentTextBox.Margin = "0,0,0,16"
            [void]$detailsPanel.Children.Add($commentTextBox)

            # Button panel
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.Margin = "0,16,0,0"

            # Save button
            $saveButton = New-Object System.Windows.Controls.Button
            $saveButton.Content = "Save Annotation"
            $saveButton.Name = ($PolicyId.replace('.', '_') + "_SaveAnnotation")
            $saveButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $saveButton.HorizontalAlignment = "Left"
            $saveButton.Width = 120
            $saveButton.Height = 26
            $saveButton.Margin = "0,0,10,0"

            # Add global event handlers to dynamically created save button
            Add-ControlEventHandlers -Control $saveButton

            # Remove button (initially hidden)
            $removeButton = New-Object System.Windows.Controls.Button
            $removeButton.Content = "Remove Annotation"
            $removeButton.Name = ($PolicyId.replace('.', '_') + "_RemoveAnnotation")
            $removeButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $removeButton.HorizontalAlignment = "Left"
            $removeButton.Width = 120
            $removeButton.Height = 26
            $removeButton.Background = [System.Windows.Media.Brushes]::Red
            $removeButton.Foreground = [System.Windows.Media.Brushes]::White
            $removeButton.Cursor = [System.Windows.Input.Cursors]::Hand
            $removeButton.Visibility = "Collapsed"

            # Add global event handlers to dynamically created remove button
            Add-ControlEventHandlers -Control $removeButton

            [void]$buttonPanel.Children.Add($saveButton)
            [void]$buttonPanel.Children.Add($removeButton)
            [void]$detailsPanel.Children.Add($buttonPanel)

            # Add elements to main grid
            [void]$cardGrid.Children.Add($headerGrid)
            [void]$cardGrid.Children.Add($detailsPanel)
            $card.Child = $cardGrid

            # Add checkbox event handler
            $checkbox.Add_Checked({
                $detailsPanel.Visibility = "Visible"
            }.GetNewClosure())

            $checkbox.Add_Unchecked({
                $detailsPanel.Visibility = "Collapsed"
            }.GetNewClosure())

            # Updated Annotation Save Button Handler
            $saveButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_SaveAnnotation", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                # Get the comment text
                $commentTextBox = $this.Parent.Parent.Children | Where-Object { $_.Name -eq ($policyIdWithUnderscores + "_Comment_TextBox") }
                $comment = $commentTextBox.Text

                # Initialize product level if not exists
                if (-not $syncHash.Annotations[$ProductName]) {
                    $syncHash.Annotations[$ProductName] = @{}
                }

                # Only create annotation if comment is not empty
                if (![string]::IsNullOrWhiteSpace($comment)) {
                    $syncHash.Annotations[$ProductName][$policyId] = @{
                        Id = $policyId
                        Product = $ProductName
                        Comment = $comment.Trim()
                    }
                } else {

                    if ($syncHash.Annotations[$ProductName]) {
                        $syncHash.Annotations[$ProductName].Remove($policyId)
                        # Remove product level if empty
                        if ($syncHash.Annotations[$ProductName].Count -eq 0) {
                            $syncHash.Annotations.Remove($ProductName)
                            Write-DebugOutput -Message "Removed [$policyId] from product [$ProductName]" -Source $policyIdWithUnderscores -Level "Info"
                        }
                    }
                }

                [System.Windows.MessageBox]::Show("[$policyId] annotation saved successfully.", "Success", "OK", "Information")

                # Make remove button visible and header bold
                $removeButton.Visibility = "Visible"
                $policyHeader.FontWeight = "Bold"

                # collapse details panel
                $detailsPanel.Visibility = "Collapsed"

                #uncheck checkbox
                $checkbox.IsChecked = $false
            }.GetNewClosure())


            # Updated Annotation Remove Button Handler
            $removeButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_RemoveAnnotation", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove the annotation for [$policyId]?", "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Write-DebugOutput -Message "User confirmed removal of annotation for policy: $policyId" -Source "User Action" -Level "Info"

                    # Remove annotation from hashtable
                    if ($syncHash.Annotations[$ProductName]) {
                        $syncHash.Annotations[$ProductName].Remove($policyId)
                        # Remove product level if empty
                        if ($syncHash.Annotations[$ProductName].Count -eq 0) {
                            $syncHash.Annotations.Remove($ProductName)
                        }
                    }

                    # Clear comment textbox
                    $commentTextBox = $this.Parent.Parent.Children | Where-Object { $_.Name -eq ($policyIdWithUnderscores + "_Comment_TextBox") }
                    $commentTextBox.Text = ""

                    [System.Windows.MessageBox]::Show("[$policyId] annotation removed successfully.", "Success", "OK", "Information")

                    # Hide remove button and unbold header
                    $this.Visibility = "Collapsed"

                    #change weight of policy header
                    $policyHeader.FontWeight = "SemiBold"

                    #uncheck checkbox
                    $checkbox.IsChecked = $false
                }
            }.GetNewClosure())

            return $card
        }#end Function : New-AnnotationCard

        Function New-ProductAnnotations {
            param(
                [string]$ProductName,
                [System.Windows.Controls.StackPanel]$Container
            )

            $Container.Children.Clear()

            # Get baselines for this product
            $baselines = $syncHash.UIConfigs.baselines.$ProductName | Select-Object id, name, rationale

            if ($baselines -and $baselines.Count -gt 0) {
                #TEST $baseline = $baselines[0]
                foreach ($baseline in $baselines) {
                    $card = New-AnnotationCard -PolicyId $baseline.id -ProductName $ProductName -PolicyName $baseline.name -PolicyDescription $baseline.rationale
                    [void]$Container.Children.Add($card)
                }
            } else {
                # No baselines available
                $noDataText = New-Object System.Windows.Controls.TextBlock
                $noDataText.Text = "No policies available for annotation in this product."
                $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $noDataText.FontStyle = "Italic"
                $noDataText.HorizontalAlignment = "Center"
                $noDataText.Margin = "0,50,0,0"
                [void]$Container.Children.Add($noDataText)
            }
        }

        #===========================================================================
        # OMISSIONS dynamic controls
        #===========================================================================

        # Function to create an omission card UI element
        Function New-OmissionCard {
            param(
                [string]$PolicyId,
                [string]$ProductName,
                [string]$PolicyName,
                [string]$PolicyDescription
            )

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
            $checkbox.Name = ($PolicyId.replace('.', '_') + "_OmissionCheckbox")
            $checkbox.VerticalAlignment = "Top"
            $checkbox.Margin = "0,0,12,0"
            [System.Windows.Controls.Grid]::SetColumn($checkbox, 0)

            # Add global event handlers to dynamically created checkbox
            Add-ControlEventHandlers -Control $checkbox

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

            # Policy description
            $policyDesc = New-Object System.Windows.Controls.TextBlock
            $policyDesc.Text = $PolicyDescription
            $policyDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
            $policyDesc.TextWrapping = "Wrap"
            [void]$policyInfoStack.Children.Add($policyDesc)

            # Add elements to header grid
            [void]$headerGrid.Children.Add($checkbox)
            [void]$headerGrid.Children.Add($policyInfoStack)

            # Create details panel (initially collapsed)
            $detailsPanel = New-Object System.Windows.Controls.StackPanel
            $detailsPanel.Visibility = "Collapsed"
            $detailsPanel.Margin = "24,12,0,0"
            [System.Windows.Controls.Grid]::SetRow($detailsPanel, 1)

            # Rationale section
            $rationaleLabel = New-Object System.Windows.Controls.TextBlock
            $rationaleLabel.Text = "Rationale (Required)"
            $rationaleLabel.FontWeight = "SemiBold"
            $rationaleLabel.Margin = "0,0,0,4"
            [void]$detailsPanel.Children.Add($rationaleLabel)

            $rationaleTextBox = New-Object System.Windows.Controls.TextBox
            $rationaleTextBox.Name = ($PolicyId.replace('.', '_') + "_Rationale_TextBox")
            $rationaleTextBox.Height = 60
            $rationaleTextBox.AcceptsReturn = $true
            $rationaleTextBox.TextWrapping = "Wrap"
            $rationaleTextBox.VerticalScrollBarVisibility = "Auto"
            $rationaleTextBox.Margin = "0,0,0,12"
            [void]$detailsPanel.Children.Add($rationaleTextBox)

            # Add global event handlers to dynamically created checkbox
            Add-ControlEventHandlers -Control $rationaleTextBox

            # Expiration date section
            $expirationLabel = New-Object System.Windows.Controls.TextBlock
            $expirationLabel.Text = "Expiration Date (Optional)"
            $expirationLabel.FontWeight = "SemiBold"
            $expirationLabel.Margin = "0,0,0,4"
            [void]$detailsPanel.Children.Add($expirationLabel)

            $expirationTextBox = New-Object System.Windows.Controls.TextBox
            $expirationTextBox.Name = ($PolicyId.replace('.', '_') + "_Expiration_TextBox")
            $expirationTextBox.Text = "mm/dd/yyyy"
            $expirationTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $expirationTextBox.FontStyle = "Italic"
            $expirationTextBox.Margin = "0,0,0,12"
            $expirationTextBox.Width = 200
            $expirationTextBox.VerticalContentAlignment = "Center"
            $expirationTextBox.HorizontalAlignment = "Left"
            [void]$detailsPanel.Children.Add($expirationTextBox)

            # Add global event handlers to dynamically created expirationTextBox
            Add-ControlEventHandlers -Control $expirationTextBox

            # Add placeholder functionality for expiration date
            $expirationTextBox.Add_GotFocus({
                if ($this.Text -eq "mm/dd/yyyy") {
                    $this.Text = ""
                    $this.Foreground = [System.Windows.Media.Brushes]::Black
                    $this.FontStyle = "Normal"
                }
            }.GetNewClosure())

            $expirationTextBox.Add_LostFocus({
                if ([string]::IsNullOrWhiteSpace($this.Text)) {
                    $this.Text = "mm/dd/yyyy"
                    $this.Foreground = [System.Windows.Media.Brushes]::Gray
                    $this.FontStyle = "Italic"
                }
            }.GetNewClosure())

             # Button panel
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.Margin = "0,16,0,0"

            # Save button
            $saveButton = New-Object System.Windows.Controls.Button
            $saveButton.Content = "Save Omission"
            $saveButton.Name = ($PolicyId.replace('.', '_') + "_SaveOmission")
            $saveButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $saveButton.HorizontalAlignment = "Left"
            $saveButton.Width = 120
            $saveButton.Height = 26
            $saveButton.Margin = "0,0,10,0"

            # Remove button (initially hidden)
            $removeButton = New-Object System.Windows.Controls.Button
            $removeButton.Content = "Remove Omission"
            $removeButton.Name = ($PolicyId.replace('.', '_') + "_RemoveOmission")
            $removeButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $removeButton.HorizontalAlignment = "Left"
            $removeButton.Width = 120
            $removeButton.Height = 26
            $removeButton.Background = [System.Windows.Media.Brushes]::Red
            $removeButton.Foreground = [System.Windows.Media.Brushes]::White
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

            # Updated Omission Save Button Handler
            $saveButton.Add_Click({

                # Get the correct policy ID from the button name
                $policyIdWithUnderscores = $this.Name.Replace("_SaveOmission", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                # Since button is now in buttonPanel, we need to go up to detailsPanel
                $detailsPanel = $this.Parent.Parent
                $rationaleTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Rationale_TextBox") }
                $expirationTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Expiration_TextBox") }

                if ([string]::IsNullOrWhiteSpace($rationaleTextBox.Text)) {
                    Write-DebugOutput -Message "User attempted to save omission for $policyId without rationale" -Source "User Action" -Level "Warning"
                    [System.Windows.MessageBox]::Show("Rationale is required for policy omissions.", "Validation Error", "OK", "Warning")
                    return
                }

                $expirationDate = $null
                if ($expirationTextBox.Text -ne "mm/dd/yyyy" -and -not [string]::IsNullOrWhiteSpace($expirationTextBox.Text)) {
                    try {
                        $expirationDate = [DateTime]::Parse($expirationTextBox.Text).ToString("yyyy-MM-dd")
                    }
                    catch {
                        Write-DebugOutput -Message "User entered invalid date format for $policyId expiration: $($expirationTextBox.Text)" -Source "User Action" -Level "Warning"
                        Write-DebugOutput -Message "Error parsing expiration date for $policyId`: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                        [System.Windows.MessageBox]::Show("Invalid date format. Please use mm/dd/yyyy format.", "Validation Error", "OK", "Warning")
                        return
                    }
                }

                # Initialize product level if not exists
                if (-not $syncHash.Omissions[$ProductName]) {
                    $syncHash.Omissions[$ProductName] = @{}
                }

                Write-DebugOutput -Message "Saving omission for policy: $policyId with rationale: $($rationaleTextBox.Text)" -Source "User Action" -Level "Info"
                # Add omission to hashtable
                $syncHash.Omissions[$ProductName][$policyId] = @{
                    Id = $policyId
                    Product = $ProductName
                    Rationale = $rationaleTextBox.Text
                    Expiration = $expirationDate
                }

                [System.Windows.MessageBox]::Show("[$policyId] omission saved successfully.", "Success", "OK", "Information")

                #make remove button visible
                $removeButton.Visibility = "Visible"

                #make policy bold
                $policyHeader.FontWeight = "Bold"

                # collapse details panel
                $detailsPanel.Visibility = "Collapsed"

                #uncheck checkbox
                $checkbox.IsChecked = $false
            }.GetNewClosure())

            # Updated Omission Remove Button Handler
            $removeButton.Add_Click({
                # Get the correct policy ID from the button name
                $policyIdWithUnderscores = $this.Name.Replace("_RemoveOmission", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove [$policyId] from omission?", "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {

                    # Since button is now in buttonPanel, we need to go up to detailsPanel
                    $detailsPanel = $this.Parent.Parent
                    $rationaleTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Rationale_TextBox") }
                    $expirationTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Expiration_TextBox") }

                    if ($rationaleTextBox) {
                        $rationaleTextBox.Text = ""
                    }
                    if ($expirationTextBox) {
                        $expirationTextBox.Text = "mm/dd/yyyy"
                        $expirationTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                        $expirationTextBox.FontStyle = "Italic"
                    }

                    # Remove omission from hashtable
                    if ($syncHash.Omissions[$ProductName]) {
                        $syncHash.Omissions[$ProductName].Remove($policyId)
                        # Remove product level if empty
                        if ($syncHash.Omissions[$ProductName].Count -eq 0) {
                            $syncHash.Omissions.Remove($ProductName)
                        }
                    }

                    [System.Windows.MessageBox]::Show("[$policyId] omission removed successfully.", "Success", "OK", "Information")

                    # Hide the remove button
                    $this.Visibility = "Collapsed"

                    # Make policy header normal weight
                    $policyHeader.FontWeight = "SemiBold"

                    # collapse details panel
                    $detailsPanel.Visibility = "Collapsed"

                    #uncheck checkbox
                    $checkbox.IsChecked = $false
                }
            }.GetNewClosure())

            return $card
        }#end Function : New-OmissionCard

        # Function to populate omissions for a product
        Function New-ProductOmissions {
            param(
                [string]$ProductName,
                [System.Windows.Controls.StackPanel]$Container
            )

            $Container.Children.Clear()

            # Get baselines for this product
            $baselines = $syncHash.UIConfigs.baselines.$ProductName | Select-Object id, name, rationale

            if ($baselines -and $baselines.Count -gt 0) {
                #TEST $baseline = $baselines[0]
                foreach ($baseline in $baselines) {
                    $card = New-OmissionCard -PolicyId $baseline.id -ProductName $ProductName -PolicyName $baseline.name -PolicyDescription $baseline.rationale
                    [void]$Container.Children.Add($card)
                }
            } else {
                # No baselines available
                $noDataText = New-Object System.Windows.Controls.TextBlock
                $noDataText.Text = "No policies available for omission in this product."
                $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $noDataText.FontStyle = "Italic"
                $noDataText.HorizontalAlignment = "Center"
                $noDataText.Margin = "0,50,0,0"
                [void]$Container.Children.Add($noDataText)
            }
        }

        #===========================================================================
        # EXCLUSIONS dynamic controls
        #===========================================================================

        # Function to create exclusion field UI based on field type and valueType
        Function New-ExclusionFieldControl {
            param(
                [string]$PolicyId,
                [string]$ExclusionTypeName,
                [object]$Field,
                [System.Windows.Controls.StackPanel]$Container
            )

            $fieldPanel = New-Object System.Windows.Controls.StackPanel
            $fieldPanel.Margin = "0,0,0,12"

            # Field label
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

            $fieldName = ($PolicyId.replace('.', '_') + "_" + $ExclusionTypeName + "_" + $Field.name)

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
                $inputTextBox.Height = 26
                $inputTextBox.VerticalContentAlignment = "Center"
                $inputTextBox.Margin = "0,0,8,0"

                # Set placeholder based on valueType
                switch ($Field.valueType) {
                    "email" { $inputTextBox.Text = "user@domain.com"; $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray; $inputTextBox.FontStyle = "Italic" }
                    "guid" { $inputTextBox.Text = "12345678-1234-1234-1234-123456789abc"; $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray; $inputTextBox.FontStyle = "Italic" }
                    "domain" { $inputTextBox.Text = "example.com"; $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray; $inputTextBox.FontStyle = "Italic" }
                    "string" { $inputTextBox.Text = "Enter value"; $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray; $inputTextBox.FontStyle = "Italic" }
                    default { $inputTextBox.Text = "Enter value"; $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray; $inputTextBox.FontStyle = "Italic" }
                }

                # Add global event handlers to dynamically created inputTextBox
                Add-ControlEventHandlers -Control $inputTextBox

                # Placeholder functionality
                $placeholder = $inputTextBox.Text
                $inputTextBox.Add_GotFocus({
                    if ($this.Text -eq $placeholder) {
                        $this.Text = ""
                        $this.Foreground = [System.Windows.Media.Brushes]::Black
                        $this.FontStyle = "Normal"
                    }
                }.GetNewClosure())

                $inputTextBox.Add_LostFocus({
                    if ([string]::IsNullOrWhiteSpace($this.Text)) {
                        $this.Text = $placeholder
                        $this.Foreground = [System.Windows.Media.Brushes]::Gray
                        $this.FontStyle = "Italic"
                    }
                }.GetNewClosure())

                $addButton = New-Object System.Windows.Controls.Button
                $addButton.Content = "Add"
                $addButton.Name = $fieldName + "_Add"
                $addButton.Style = $syncHash.Window.FindResource("PrimaryButton")
                $addButton.Width = 60
                $addButton.Height = 26

                [void]$inputRow.Children.Add($inputTextBox)
                [void]$inputRow.Children.Add($addButton)
                [void]$arrayContainer.Children.Add($inputRow)

                # List container for added items
                $listContainer = New-Object System.Windows.Controls.StackPanel
                $listContainer.Name = $fieldName + "_List"
                [void]$arrayContainer.Children.Add($listContainer)

                Add-ControlEventHandlers -Control $addButton

                # Register the listContainer in syncHash for data collection
                #syncHash.$fieldName = $listContainer

                # Add button functionality
                $addButton.Add_Click({
                    $inputBox = $this.Parent.Children[0]
                    $listPanel = $this.Parent.Parent.Children[1]

                    if (![string]::IsNullOrWhiteSpace($inputBox.Text) -and $inputBox.Text -ne $placeholder) {
                        # Trim the input value
                        $trimmedValue = $inputBox.Text.Trim()

                        If($listContainer.Children.Children | Where-Object { $_.Text -contains $trimmedValue }) {
                            # User already exists, skip
                            return
                        }

                        # Validate input based on valueType
                        $isValid = $true
                        switch ($Field.valueType) {
                            "email" { $isValid = $trimmedValue -match "^[^\s@]+@[^\s@]+\.[^\s@]+$" }
                            "guid" { $isValid = $trimmedValue -match "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$" }
                            "domain" { $isValid = $trimmedValue -match "^[a-zA-Z0-9][a-zA-Z0-9-]*[a-zA-Z0-9]*\.([a-zA-Z]{2,})+$" }
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

                            # Clear input
                            $inputBox.Text = $placeholder
                            $inputBox.Foreground = [System.Windows.Media.Brushes]::Gray
                            $inputBox.FontStyle = "Italic"
                        } else {
                            [System.Windows.MessageBox]::Show("Invalid $($Field.valueType) format.", "Validation Error", "OK", "Warning")
                        }
                    }
                }.GetNewClosure())

                [void]$fieldPanel.Children.Add($arrayContainer)

            } elseif ($Field.type -eq "string") {
                # Create single string input
                $stringTextBox = New-Object System.Windows.Controls.TextBox
                $stringTextBox.Name = $fieldName + "_TextBox"
                $stringTextBox.Width = 400
                $stringTextBox.Height = 26
                $stringTextBox.VerticalContentAlignment = "Center"

                if ($Field.valueType -eq "semicolonList") {
                    $stringTextBox.Text = "user1@domain.com;user2@domain.com"
                    $stringTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                    $stringTextBox.FontStyle = "Italic"

                    $stringTextBox.Add_GotFocus({
                        if ($this.Text -eq "user1@domain.com;user2@domain.com") {
                            $this.Text = ""
                            $this.Foreground = [System.Windows.Media.Brushes]::Black
                            $this.FontStyle = "Normal"
                        }
                    })

                    $stringTextBox.Add_LostFocus({
                        if ([string]::IsNullOrWhiteSpace($this.Text)) {
                            $this.Text = "user1@domain.com;user2@domain.com"
                            $this.Foreground = [System.Windows.Media.Brushes]::Gray
                            $this.FontStyle = "Italic"
                        }
                    })
                }
                # Add global event handlers to dynamically created stringTextBox
                Add-ControlEventHandlers -Control $stringTextBox
                [void]$fieldPanel.Children.Add($stringTextBox)
            }

            If($syncHash.GraphConnected)
            {
                if ($Field.name -eq "Users") {
                    $getUsersButton = New-Object System.Windows.Controls.Button
                    $getUsersButton.Content = "Get Users"
                    $getUsersButton.Width = 80
                    $getUsersButton.Height = 26
                    $getUsersButton.Margin = "8,0,0,0"

                    $getUsersButton.Add_Click({
                        try {
                            # Get search term from input box
                            $searchTerm = if ($inputTextBox.Text -ne $placeholder -and ![string]::IsNullOrWhiteSpace($inputTextBox.Text)) { $inputTextBox.Text } else { "" }

                            # Show user selector
                            $selectedUsers = Show-UserSelector -SearchTerm $searchTerm

                            if ($selectedUsers -and $selectedUsers.Count -gt 0) {
                                # Clear existing items
                                #$listContainer.Children.Clear()


                                # Add selected users to the list
                                foreach ($user in $selectedUsers) {
                                    #check if group already exists in the list
                                    If ($listContainer.Children.Children | Where-Object { $_.Text -contains $user.Id }) {
                                        # User already exists, skip
                                        return
                                    }
                                    $userItem = New-Object System.Windows.Controls.StackPanel
                                    $userItem.Orientation = "Horizontal"
                                    $userItem.Margin = "0,2,0,2"

                                    $userText = New-Object System.Windows.Controls.TextBlock
                                    $userText.Text = "$($user.id)"
                                    $userText.Width = 250
                                    $userText.VerticalAlignment = "Center"
                                    $userText.ToolTip = "$($user.DisplayName) ($($user.UserPrincipalName))"

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

                                    [void]$userItem.Children.Add($userText)
                                    [void]$userItem.Children.Add($removeUserButton)
                                    [void]$listContainer.Children.Add($userItem)

                                }

                                # Clear the input box
                                $inputTextBox.Text = $placeholder
                                $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                                $inputTextBox.FontStyle = "Italic"

                                Write-DebugOutput -Message "Added $($selectedUsers.Count) users to exclusion list"
                            }
                        }
                        catch {
                            Write-DebugOutput -Message "Error selecting users: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                            [System.Windows.MessageBox]::Show("Error selecting users: $($_.Exception.Message)", "Error",
                                                            [System.Windows.MessageBoxButton]::OK,
                                                            [System.Windows.MessageBoxImage]::Error)
                        }
                    }.GetNewClosure())

                    [void]$inputRow.Children.Add($getUsersButton)
                }

                # Check if field is for groups
                elseif ($Field.name -eq "Groups") {
                    $getGroupsButton = New-Object System.Windows.Controls.Button
                    $getGroupsButton.Content = "Get Groups"
                    $getGroupsButton.Width = 80
                    $getGroupsButton.Height = 26
                    $getGroupsButton.Margin = "8,0,0,0"

                    $getGroupsButton.Add_Click({
                        try {
                            # Get search term from input box
                            $searchTerm = if ($inputTextBox.Text -ne $placeholder -and ![string]::IsNullOrWhiteSpace($inputTextBox.Text)) { $inputTextBox.Text } else { "" }

                            # Show group selector
                            $selectedGroups = Show-GroupSelector -SearchTerm $searchTerm

                            if ($selectedGroups -and $selectedGroups.Count -gt 0) {
                                # Clear existing items
                                #$listContainer.Children.Clear()

                                # Add selected groups to the list
                                foreach ($group in $selectedGroups)
                                {
                                    #check if group already exists in the list
                                    If ($listContainer.Children.Children | Where-Object { $_.Text -contains $group.Id }) {
                                        # User already exists, skip
                                        Return

                                    }
                                    $groupItem = New-Object System.Windows.Controls.StackPanel
                                    $groupItem.Orientation = "Horizontal"
                                    $groupItem.Margin = "0,2,0,2"

                                    $groupText = New-Object System.Windows.Controls.TextBlock
                                    $groupText.Text = "$($group.Id)"
                                    $groupText.Width = 250
                                    $groupText.VerticalAlignment = "Center"
                                    $groupText.ToolTip = "$($group.DisplayName)"

                                    $removeGroupButton = New-Object System.Windows.Controls.Button
                                    $removeGroupButton.Content = "Remove"
                                    $removeGroupButton.Width = 60
                                    $removeGroupButton.Height = 20
                                    $removeGroupButton.Margin = "8,0,0,0"
                                    $removeGroupButton.FontSize = 10
                                    $removeGroupButton.Background = [System.Windows.Media.Brushes]::Red
                                    $removeGroupButton.Foreground = [System.Windows.Media.Brushes]::White
                                    $removeGroupButton.BorderThickness = "0"

                                    $removeGroupButton.Add_Click({
                                        $parentItem = $this.Parent
                                        $parentContainer = $parentItem.Parent
                                        $parentContainer.Children.Remove($parentItem)
                                    }.GetNewClosure())

                                    [void]$groupItem.Children.Add($groupText)
                                    [void]$groupItem.Children.Add($removeGroupButton)
                                    [void]$listContainer.Children.Add($groupItem)


                                }

                                # Clear the input box
                                $inputTextBox.Text = $placeholder
                                $inputTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                                $inputTextBox.FontStyle = "Italic"

                                Write-DebugOutput -Message "Added $($selectedGroups.Count) groups to exclusion list" -Source $MyInvocation.MyCommand.Name -Level "Info"
                            }
                        }
                        catch {
                            Write-DebugOutput -Message "Error selecting groups: $($_.Exception.Message)" -Source $MyInvocation.MyCommand.Name -Level "Error"
                            [System.Windows.MessageBox]::Show("Error selecting groups: $($_.Exception.Message)", "Error",
                                                            [System.Windows.MessageBoxButton]::OK,
                                                            [System.Windows.MessageBoxImage]::Error)
                        }
                    }.GetNewClosure())

                    [void]$inputRow.Children.Add($getGroupsButton)
                }
            }
            [void]$Container.Children.Add($fieldPanel)
        }

        # Function to create an exclusion card UI element
        Function New-ExclusionCard {
            param(
                [string]$PolicyId,
                [string]$ProductName,
                [string]$PolicyName,
                [string]$PolicyDescription,
                [string]$ExclusionType
            )

            # Skip if exclusion type is "none"
            if ($ExclusionType -eq "none") {
                return $null
            }

            # Get exclusion type definition from config
            $exclusionTypeDef = $syncHash.UIConfigs.exclusionTypes.$ExclusionType
            if (-not $exclusionTypeDef) {
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
            $checkbox.Name = ($PolicyId.replace('.', '_') + "_ExclusionCheckbox")
            $checkbox.VerticalAlignment = "Top"
            $checkbox.Margin = "0,0,12,0"
            [System.Windows.Controls.Grid]::SetColumn($checkbox, 0)

            # Add global event handlers to dynamically created checkbox
            Add-ControlEventHandlers -Control $checkbox

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

            # Exclusion type info
            $exclusionTypeHeader = New-Object System.Windows.Controls.TextBlock
            $exclusionTypeHeader.Text = "Exclusion Type: $($exclusionTypeDef.name)"
            $exclusionTypeHeader.FontSize = 12
            $exclusionTypeHeader.Foreground = $syncHash.Window.FindResource("AccentBrush")
            $exclusionTypeHeader.Margin = "0,0,0,4"
            [void]$policyInfoStack.Children.Add($exclusionTypeHeader)

            # Add elements to header grid
            [void]$headerGrid.Children.Add($checkbox)
            [void]$headerGrid.Children.Add($policyInfoStack)

            # Create details panel (initially collapsed)
            $detailsPanel = New-Object System.Windows.Controls.StackPanel
            $detailsPanel.Visibility = "Collapsed"
            $detailsPanel.Margin = "24,12,0,0"
            [System.Windows.Controls.Grid]::SetRow($detailsPanel, 1)

            # Exclusion type description
            $exclusionDesc = New-Object System.Windows.Controls.TextBlock
            $exclusionDesc.Text = $exclusionTypeDef.description
            $exclusionDesc.FontStyle = "Italic"
            $exclusionDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
            $exclusionDesc.TextWrapping = "Wrap"
            $exclusionDesc.Margin = "0,0,0,16"
            [void]$detailsPanel.Children.Add($exclusionDesc)

            # Generate fields based on exclusion type
            foreach ($field in $exclusionTypeDef.fields) {
                New-ExclusionFieldControl -PolicyId $PolicyId -ExclusionTypeName $ExclusionType -Field $field -Container $detailsPanel
            }

            # Button panel
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.Margin = "0,16,0,0"

            # Save button
            $saveButton = New-Object System.Windows.Controls.Button
            $saveButton.Content = "Save Exclusion"
            $saveButton.Name = ($PolicyId.replace('.', '_') + "_SaveExclusion")
            $saveButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $saveButton.Width = 120
            $saveButton.Height = 26
            $saveButton.Margin = "0,0,10,0"

            # Remove button (initially hidden)
            $removeButton = New-Object System.Windows.Controls.Button
            $removeButton.Content = "Remove Exclusion"
            $removeButton.Name = ($PolicyId.replace('.', '_') + "_RemoveExclusion")
            $removeButton.Style = $syncHash.Window.FindResource("PrimaryButton")
            $removeButton.Width = 120
            $removeButton.Height = 26
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

            # Add save button click handler
            $saveButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_SaveExclusion", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                Write-DebugOutput -Message "Saving exclusion for policy: $policyId" -Source $this.Name -Level "Info"

                # Get the exclusion type configuration to get the actual YAML key name
                $exclusionTypeConfig = $syncHash.UIConfigs.exclusionTypes.$ExclusionType
                $yamlKeyName = if ($exclusionTypeConfig) { $exclusionTypeConfig.name } else { $ExclusionType }

                # INITIALIZE exclusionData hashtable
                $exclusionData = @{}

                # Collect field values by traversing the UI tree
                $detailsPanel = $this.Parent.Parent

                foreach ($field in $exclusionTypeDef.fields) {
                    $fieldName = ($policyId.replace('.', '_') + "_" + $ExclusionType + "_" + $field.name)

                    if ($field.type -eq "array") {
                        # For arrays, look for the list container
                        $listContainerName = ($fieldName + "_List")
                        $listContainer = $detailsPanel.Children | ForEach-Object {
                            if ($_ -is [System.Windows.Controls.StackPanel]) {
                                $arrayContainer = $_.Children | Where-Object { $_.Name -eq ($fieldName + "_Container") }
                                if ($arrayContainer) {
                                    Write-DebugOutput -Message "Found array container for field [$field.name]: $($arrayContainer.Name)" -Source $this.Name -Level "Info"
                                    return $arrayContainer.Children | Where-Object { $_.Name -eq $listContainerName }

                                }
                            }
                        } | Select-Object -First 1

                        if ($listContainer -and $listContainer.Children.Count -gt 0) {
                            $items = @()
                            foreach ($childPanel in $listContainer.Children) {
                                # Each child is a StackPanel containing TextBlock and Button
                                if ($childPanel -is [System.Windows.Controls.StackPanel] -and $childPanel.Children.Count -gt 0) {
                                    $textBlock = $childPanel.Children[0]
                                    if ($textBlock -is [System.Windows.Controls.TextBlock] -and ![string]::IsNullOrWhiteSpace($textBlock.Text)) {
                                        $items += $textBlock.Text.Trim()
                                    }
                                }
                            }
                            if ($items.Count -gt 0) {
                                $exclusionData[$field.name] = $items
                                Write-DebugOutput -Message "Collected array field [$field.name]: $($items -join ', ')" -Source $this.Name -Level "Info"
                            }
                        }
                    } elseif ($field.type -eq "string") {
                        # For strings, look for the TextBox
                        $stringFieldName = ($fieldName + "_TextBox")
                        $stringTextBox = $detailsPanel.Children | ForEach-Object {
                            if ($_ -is [System.Windows.Controls.StackPanel]) {
                                return $_.Children | Where-Object { $_.Name -eq $stringFieldName -and $_ -is [System.Windows.Controls.TextBox] }
                            }
                        } | Select-Object -First 1

                        if ($stringTextBox -and ![string]::IsNullOrWhiteSpace($stringTextBox.Text)) {
                            $value = $stringTextBox.Text.Trim()
                            $exclusionData[$field.name] = $value
                            Write-DebugOutput -Message "Collected string field [$field.name]: $value" -Source $this.Name -Level "Info"
                        }
                    }
                }

                # Only create exclusion if we have data
                if ($exclusionData.Count -gt 0) {
                    # Initialize product level if not exists
                    if (-not $syncHash.Exclusions[$ProductName]) {
                        $syncHash.Exclusions[$ProductName] = @{}
                    }

                    # Initialize policy level if not exists
                    if (-not $syncHash.Exclusions[$ProductName][$policyId]) {
                        $syncHash.Exclusions[$ProductName][$policyId] = @{}
                    }

                    # Set the exclusion data using the YAML key name
                    $syncHash.Exclusions[$ProductName][$policyId][$yamlKeyName] = $exclusionData

                    Write-DebugOutput -Message "Exclusion saved for [$ProductName][$policyId][$yamlKeyName]: $($exclusionData | ConvertTo-Json -Compress)" -Source $this.Name -Level "Info"

                    [System.Windows.MessageBox]::Show("[$policyId] exclusion saved successfully.", "Success", "OK", "Information")

                    # Make remove button visible and header bold
                    $removeButton.Visibility = "Visible"
                    $policyHeader.FontWeight = "Bold"
                } else {
                    [System.Windows.MessageBox]::Show("No exclusion data entered. Please fill in at least one field.", "Validation Error", "OK", "Warning")
                    return
                }

                # collapse details panel
                $detailsPanel.Visibility = "Collapsed"

                #uncheck checkbox
                $checkbox.IsChecked = $false
            }.GetNewClosure())

            # Add remove button click handler
            $removeButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_RemoveExclusion", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove exclusions for [$policyId]?", "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {

                    # Remove the policy from the nested structure
                    if ($syncHash.Exclusions[$ProductName] -and $syncHash.Exclusions[$ProductName][$policyId]) {
                        $syncHash.Exclusions[$ProductName].Remove($policyId)

                        # If no more policies for this product, remove the product
                        if ($syncHash.Exclusions[$ProductName].Count -eq 0) {
                            $syncHash.Exclusions.Remove($ProductName)
                        }
                    }

                    # Clear all field values
                    foreach ($field in $exclusionTypeDef.fields) {
                        $fieldName = ($policyId.replace('.', '_') + "_" + $ExclusionType + "_" + $field.name)
                        $fieldControl = $syncHash.$fieldName

                        if ($fieldControl) {
                            if ($field.type -eq "array") {
                                $fieldControl.Items.Clear()
                            } elseif ($field.type -eq "string") {
                                $fieldControl.Text = ""
                            }
                        }
                    }

                    [System.Windows.MessageBox]::Show("[$policyId] exclusions removed successfully.", "Success", "OK", "Information")

                    # Hide remove button and unbold header
                    $this.Visibility = "Collapsed"
                    $policyHeader.FontWeight = "SemiBold"
                    $checkbox.IsChecked = $false
                }
            }.GetNewClosure())

            return $card
        }

        # Function to populate exclusions for a product
        Function New-ProductExclusions {
            param(
                [string]$ProductName,
                [System.Windows.Controls.StackPanel]$Container
            )

            $Container.Children.Clear()

            # Get baselines for this product that support exclusions
            $baselines = $syncHash.UIConfigs.baselines.$ProductName | Where-Object { $_.exclusionType -ne 'none' } | Select-Object id, name, rationale, exclusionType

            if ($baselines -and $baselines.Count -gt 0) {
                foreach ($baseline in $baselines) {
                    $card = New-ExclusionCard -PolicyId $baseline.id -ProductName $ProductName -PolicyName $baseline.name -PolicyDescription $baseline.rationale -ExclusionType $baseline.exclusionType
                    if ($card) {
                        $Container.Children.Add($card)
                    }
                }
            } else {
                # No baselines available for exclusions
                $noDataText = New-Object System.Windows.Controls.TextBlock
                $noDataText.Text = "No policies support exclusions in this product."
                $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $noDataText.FontStyle = "Italic"
                $noDataText.HorizontalAlignment = "Center"
                $noDataText.Margin = "0,50,0,0"
                $Container.Children.Add($noDataText)
            }
        }

        #===========================================================================
        #
        # GRAPH HELPER
        #
        #===========================================================================

        # Enhanced Graph Query Function with Filter Support
        function Invoke-GraphQueryWithFilter {
            param(
                [string]$QueryType,
                $GraphConfig,
                [string]$FilterString,
                [int]$Top = 999,
                [string]$ProgressMessage = "Retrieving data..."
            )

            # Create runspace
            $runspace = [runspacefactory]::CreateRunspace()
            $runspace.Open()

            # Create PowerShell instance
            $powershell = [powershell]::Create()
            $powershell.Runspace = $runspace

            # Add script block
            $scriptBlock = {
                param($QueryType, $GraphConfig, $FilterString, $Top, $ProgressMessage)

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
            $powershell.AddScript($scriptBlock).AddParameter("QueryType", $QueryType).AddParameter("GraphConfig", $GraphConfig).AddParameter("FilterString", $FilterString).AddParameter("Top", $Top).AddParameter("ProgressMessage", $ProgressMessage)
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
            param(
                [Parameter(Mandatory)]
                [ValidateSet("users", "groups")]
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
                $asyncOp = Invoke-GraphQueryWithFilter -QueryType $config.QueryType -GraphConfig $syncHash.UIConfigs.graphQueries -FilterString $filterString -Top $Top

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

        function Show-UserSelector {
            param(
                [string]$SearchTerm = "",
                [int]$Top = 100
            )
            If([string]::IsNullOrWhiteSpace($SearchTerm)) {
                Write-DebugOutput -Message "Showing user selector with top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }Else {
                Write-DebugOutput -Message "Showing user selector with search term: $SearchTerm, top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }
            return Show-GraphProgressWindow -GraphEntityType "users" -SearchTerm $SearchTerm -Top $Top
        }

        function Show-GroupSelector {
            param(
                [string]$SearchTerm = "",
                [int]$Top = 100
            )
            If([string]::IsNullOrWhiteSpace($SearchTerm)) {
                Write-DebugOutput -Message "Showing group selector with top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }Else {
                Write-DebugOutput -Message "Showing group selector with search term: $SearchTerm, top: $Top" -Source $MyInvocation.MyCommand.Name -Level "Info"
            }
            return Show-GraphProgressWindow -GraphEntityType "groups" -SearchTerm $SearchTerm -Top $Top
        }

        #build UI selection window
        function Show-UISelectionWindow {
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
                        $filteredItems = $originalItems | Where-Object {
                            $_.$SearchProperty.ToLower().Contains($searchText)
                        }
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

                                # Find the exclusion type from baseline config
                                $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                                if ($baseline -and $baseline.exclusionType -ne "none") {
                                    # Get the exclusion type configuration to get the actual YAML key name
                                    $exclusionTypeConfig = $syncHash.UIConfigs.exclusionTypes.($baseline.exclusionType)
                                    $yamlKeyName = if ($exclusionTypeConfig) { $exclusionTypeConfig.name } else { $baseline.exclusionType }

                                    # Check if the policy data contains the expected exclusion type key
                                    if ($policyData.ContainsKey($yamlKeyName)) {
                                        # Initialize product level if not exists
                                        if (-not $syncHash.Exclusions[$productName]) {
                                            $syncHash.Exclusions[$productName] = @{}
                                        }

                                        # Initialize policy level if not exists
                                        if (-not $syncHash.Exclusions[$productName][$policyId]) {
                                            $syncHash.Exclusions[$productName][$policyId] = @{}
                                        }

                                        # CORRECTED: Extract the field data from inside the exclusion type
                                        # Instead of: $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $policyData
                                        # We want: $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $policyData[$yamlKeyName]
                                        $exclusionFieldData = $policyData[$yamlKeyName]
                                        $syncHash.Exclusions[$productName][$policyId][$yamlKeyName] = $exclusionFieldData

                                        Write-DebugOutput -Message "Imported exclusion for [$productName][$policyId][$yamlKeyName]: $($exclusionFieldData | ConvertTo-Json -Compress)" -Source $MyInvocation.MyCommand.Name -Level "Info"
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
            # Generate YAML preview
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
            #$productNamesForYaml = Update-ProductName -AsYamlOutput
            $yamlPreview += Get-ProductNamesForYaml

            <#if ($productNamesForYaml.Count -gt 0) {
                if ($productNamesForYaml -contains '*') {
                    $yamlPreview += "`nProductNames: ['*']"
                } else {
                    $yamlPreview += "`nProductNames:"
                    foreach ($product in $productNamesForYaml) {
                        $yamlPreview += "`n  - $product"
                    }
                }
            }
            #>

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

                    if ($control -is [System.Windows.Controls.TextBox]) {
                        $currentValue = $control.Text

                        # Include if it has a value (use current value or default if empty)
                        $valueToUse = if (![string]::IsNullOrWhiteSpace($currentValue)) { $currentValue } else { $defaultValue }

                        if (![string]::IsNullOrWhiteSpace($valueToUse)) {
                            # Convert control name to YAML field name (remove _TextBox suffix)
                            $yamlFieldName = $advancedKey -replace '_TextBox$', ''

                            # Handle path escaping
                            if ($valueToUse -match '\\') {
                                $advancedSettingsContent += "`n$yamlFieldName`: `"$($valueToUse.Replace('\', '\\'))`""
                            } else {
                                $advancedSettingsContent += "`n$yamlFieldName`: `"$valueToUse`""
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

                # Only add advanced settings section if there are settings to show
                if ($hasAdvancedSettings) {
                    $yamlPreview += "`n`n# Advanced Configuration"
                    $yamlPreview += $advancedSettingsContent
                    $yamlPreview += "`n"
                }
            }

            # Handle Exclusions (unchanged - already hashtable)
            #TEST $productName = $syncHash.Exclusions.Keys[0]
            foreach ($productName in $syncHash.Exclusions.Keys) {
                $productHasExclusions = $false
                $productExclusions = @()

                #TEST $policyId = $syncHash.Exclusions[$productName].Keys[0]
                foreach ($policyId in $syncHash.Exclusions[$productName].Keys) {
                    #get baseline from config
                    $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                    #TEST $exclusionKey = $syncHash.Exclusions[$productName][$policyId].Keys[0]
                    foreach ($exclusionKey in $syncHash.Exclusions[$productName][$policyId].Keys) {
                        $exclusionData = $syncHash.Exclusions[$productName][$policyId][$exclusionKey]

                        if ($exclusionData.Count -gt 0) {
                            if (-not $productHasExclusions) {
                                $productExclusions += "`n$productName`:"
                                $productHasExclusions = $true
                            }

                            #add comment for the policy
                            if ($baseline) {
                                $productExclusions += "`n  # $($baseline.name)"
                            }
                            $productExclusions += "`n  $policyId`:"
                            $productExclusions += "`n    $exclusionKey`:"

                            # $exclusionData contains the field data directly (e.g., @{Groups = @("guid1", "guid2")})
                            if ($exclusionData.Count -gt 0) {
                                foreach ($fieldName in $exclusionData.Keys) {
                                    $productExclusions += "`n      $fieldName`:"
                                    $fieldValue = $exclusionData[$fieldName]

                                    # Handle both array and single values
                                    if ($fieldValue -is [Array]) {
                                        foreach ($item in $fieldValue) {
                                            $productExclusions += "`n        - $item"
                                        }
                                    } else {
                                        $productExclusions += "`n        - $fieldValue"
                                    }
                                }
                            }
                        }
                    }
                }

                if ($productHasExclusions) {
                    $yamlPreview += $productExclusions
                }
            }

            # Handle Annotations - Updated for hashtable structure
            $annotationCount = 0
            foreach ($productName in $syncHash.Annotations.Keys) {
                $annotationCount += $syncHash.Annotations[$productName].Count
            }

            if ($annotationCount -gt 0) {
                $yamlPreview += "`n`nAnnotatePolicy:"

                # Group annotations by product
                foreach ($productName in $syncHash.Annotations.Keys) {
                    $yamlPreview += "`n  # $productName Annotations:"

                    # Sort policies by ID
                    $sortedPolicies = $syncHash.Annotations[$productName].Keys | Sort-Object
                    foreach ($policyId in $sortedPolicies) {
                        $annotation = $syncHash.Annotations[$productName][$policyId]

                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $policyId`:"

                        if ($annotation.Comment -match "`n") {
                            # Use quoted string format with \n for line breaks
                            $escapedComment = $annotation.Comment.Replace('"', '""').Replace("`n", "\n")
                            $yamlPreview += "`n    Comment: `"$escapedComment`""
                        } else {
                            $yamlPreview += "`n    Comment: `"$($annotation.Comment)`""
                        }
                    }
                }
            }

            # Handle Omissions - Updated for hashtable structure
            $omissionCount = 0
            foreach ($productName in $syncHash.Omissions.Keys) {
                $omissionCount += $syncHash.Omissions[$productName].Count
            }

            if ($omissionCount -gt 0) {
                $yamlPreview += "`n`nOmitPolicy:"

                # Group omissions by product
                foreach ($productName in $syncHash.Omissions.Keys) {
                    $yamlPreview += "`n  # $productName Omissions:"

                    # Sort policies by ID
                    $sortedPolicies = $syncHash.Omissions[$productName].Keys | Sort-Object
                    foreach ($policyId in $sortedPolicies) {
                        $omission = $syncHash.Omissions[$productName][$policyId]

                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $policyId`:"
                        $yamlPreview += "`n    Rationale: $($omission.Rationale)"
                        if ($omission.Expiration) {
                            $yamlPreview += "`n    Expiration: $($omission.Expiration)"
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

        #===========================================================================
        # RESET: NEW SESSION
        #===========================================================================
        Function Reset-UIFields {

            # Reset core data structures
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
                Write-DebugOutput -Message "Reset control: $controlName" -Source $MyInvocation.MyCommand.Name -Level "Info"
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

        }#end function : Reset-UIFields

        # Function to collect general settings from UI controls
        Function Save-GeneralSettingsFromInput {

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
                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Collected Main setting: $placeholderKey = $($syncHash.GeneralSettings[$settingName])" -Source $MyInvocation.MyCommand.Name -Level "Info"}
                }
            }

            # Collect M365Environment
            if ($syncHash.M365Environment_ComboBox.SelectedItem) {
                $selectedEnv = $syncHash.UIConfigs.M365Environment | Where-Object { $_.id -eq $syncHash.M365Environment_ComboBox.SelectedItem.Tag } | Select-Object -ExpandProperty name
                if ($selectedEnv) {
                    $syncHash.GeneralSettings["M365Environment"] = $selectedEnv
                }
                If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Collected M365Environment: $selectedEnv" -Source $MyInvocation.MyCommand.Name -Level "Info"}
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
                    If($syncHash.DebugMode -match 'Timer|All'){Write-DebugOutput -Message "Collected Advanced setting: $settingName = $($syncHash.GeneralSettings[$settingName])" -Source $MyInvocation.MyCommand.Name -Level "Info"}
                }
            }

            # Mark data as changed
            Set-DataChanged
        } #end function : Save-GeneralSettingsFromInput
        #===========================================================================
        # Make UI functional
        #===========================================================================
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
            Write-DebugOutput -Message "M365Environment ComboBox added: $($env.displayName) ($($env.name))" -Source "UI Launch" -Level "Info"
        }

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
        Write-DebugOutput -Message "M365 Environment ComboBox set: $($syncHash.M365Environment_ComboBox.SelectedItem.Content)" -Source "UI Update" -Level "Info"

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
                    Write-DebugOutput -Message "Added $productId to ProductNames data" -Source "User Action" -Level "Info"
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
                if ($container -and $container.Children.Count -eq 0) {
                    New-ProductOmissions -ProductName $productId -Container $container
                }

                #annotations tab
                $AnnotationTab = $syncHash.("$($productId)AnnotationTab")
                $AnnotationTab.IsEnabled = $true
                Write-DebugOutput -Message "Annotation sub tab enabled: $($productId)" -Source "UI Update" -Level "Info"

                $container = $syncHash.("$($productId)AnnotationContent")
                if ($container -and $container.Children.Count -eq 0) {
                    New-ProductAnnotations -ProductName $productId -Container $container
                }

                #exclusions tab
                if ($product.supportsExclusions)
                {
                    $ExclusionsTab = $syncHash.("$($productId)ExclusionTab")
                    $ExclusionsTab.IsEnabled = $true
                    Write-DebugOutput -Message "Exclusion sub tab enabled: $($productId)" -Source "UI Update" -Level "Info"

                    $container = $syncHash.("$($productId)ExclusionContent")
                    if ($container -and $container.Children.Count -eq 0) {
                        New-ProductExclusions -ProductName $productId -Container $container
                    }
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
                    Write-DebugOutput -Message "Removed $productId from ProductNames data" -Source "User Action" -Level "Info"
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
        $PlaceholderTextBoxes = $syncHash.UIConfigs.localePlaceholder
        Foreach($placeholderKey in $PlaceholderTextBoxes.PSObject.Properties.Name) {
            $control = $syncHash.$placeholderKey
            if ($control -is [System.Windows.Controls.TextBox]) {
                Initialize-PlaceholderTextBox -TextBox $control -PlaceholderText $PlaceholderTextBoxes.$placeholderKey
            }
        }

        # Handle Organization TextBox with special Graph Connected logic
        if ($syncHash.GraphConnected) {
            try {
                $tenantDetails = (Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/organization").Value
                $tenantName = ($tenantDetails.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
                Initialize-PlaceholderTextBox -TextBox $syncHash.Organization_TextBox -PlaceholderText $syncHash.UIConfigs.localePlaceholder.Organization_TextBox -InitialValue $tenantName
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
            Add-ControlEventHandlers -Control $syncHash.$($button.Name)
        }
        #>

        # New Session Button
        $syncHash.NewSessionButton.Add_Click({
            $result = [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeInfoMessages.NewSessionConfirmation, "New Session", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                # Reset all form fields
                Reset-UIFields
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

                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeInfoMessages.ImportSuccess, "Import Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
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
                                        -RegexPattern "^(.*\.)?(onmicrosoft\.com|onmicrosoft\.us)$" `
                                        -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.OrganizationValidation `
                                        -PlaceholderText $syncHash.UIConfigs.localePlaceholder.Organization_TextBox `
                                        -Required `
                                        -ShowMessageBox:$false

                if (-not $orgValid) {
                    $overallValid = $false
                    $errorMessages += $syncHash.UIConfigs.localeErrorMessages.OrganizationValidation
                }

                # Advanced Tab Validations (only if sections are toggled on)

                # Application Section Validations
                if ($syncHash.ApplicationSection_Toggle.IsChecked) {

                    # AppID validation (GUID format)
                    $appIdValid = Confirm-UIField -UIElement $syncHash.AppId_TextBox `
                                                    -RegexPattern "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$" `
                                                    -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.AppIdValidation `
                                                    -PlaceholderText $syncHash.UIConfigs.localePlaceholder.AppId_TextBox `
                                                    -ShowMessageBox:$false

                    if (-not $appIdValid) {
                        $overallValid = $false
                        $errorMessages += $syncHash.UIConfigs.localeErrorMessages.AppIdValidation
                    }

                    # Certificate Thumbprint validation (40 character hex)
                    $certValid = Confirm-UIField -UIElement $syncHash.CertificateThumbprint_TextBox `
                                                -RegexPattern "^[0-9a-fA-F]{40}$" `
                                                -ErrorMessage $syncHash.UIConfigs.localeErrorMessages.CertificateValidation `
                                                -PlaceholderText $syncHash.UIConfigs.localePlaceholder.CertificateThumbprint_TextBox `
                                                -ShowMessageBox:$false

                    if (-not $certValid) {
                        $overallValid = $false
                        $errorMessages += $syncHash.UIConfigs.localeErrorMessages.CertificateValidation
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
    	$syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-UIMainWindow})
    	$syncHash.Window.Add_Closed({
            if ($syncHash.UIUpdateTimer) {
                $syncHash.UIUpdateTimer.Stop()
                $syncHash.UIUpdateTimer = $null
            }
            if ($syncHash.DebugFlushTimer) {
                $syncHash.DebugFlushTimer.Stop()
                $syncHash.DebugFlushTimer.Dispose()
            }
            $syncHash.isClosed = $True
        })

        #always force windows on bottom
        $syncHash.Window.Topmost = $True

        # Add global event handlers to all UI controls after everything is loaded
         Write-DebugOutput -Message "Adding event handlers to all controls" -Source "UI Initialization" -Level "Info"
        $allControls = Find-AllControls -Parent $syncHash.Window
        foreach ($control in $allControls) {
            try {
                Write-DebugOutput -Message "Adding event handlers to control: $($control.Name)" -Source "UI Initialization" -Level "Info"
                Add-ControlEventHandlers -Control $control
            } catch {
                Write-DebugOutput -Message "Failed to add event handler to control: $($_.Exception.Message)" -Source "UI Initialization" -Level "Warning"
            }
        }

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