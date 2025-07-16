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
    #>

    [CmdletBinding()]
    Param(
        $YAMLConfigFile,

        [ValidateSet('en-US')]
        $Language = 'en-US',

        [switch]$Online,

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
    $Runspace =[runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $Runspace
    $syncHash.GraphConnected = $Online
    $syncHash.XamlPath = "$PSScriptRoot\ScubaConfigAppUI.xaml"
    $syncHash.UIConfigPath = "$PSScriptRoot\ScubaConfig_$Language.json"
    $syncHash.YAMLImport = $YAMLConfigFile
    $syncHash.GraphEndpoint = $GraphEndpoint
    $syncHash.M365Environment = $M365Environment
    $syncHash.Exclusions = @()
    $syncHash.Omissions = @()
    $syncHash.Annotations = @()
    $syncHash.GeneralSettings = @{}
    $syncHash.Placeholder = @{}
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
        $UIXML.SelectNodes("//*[@Name]") | %{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

        #===========================================================================
        # UPDATE UI FUNCTIONS
        #===========================================================================

        # Function to update all UI elements from data
        Function Update-AllUIFromData {
            # Update exclusions
            Update-ExclusionsFromData

            # Update omissions
            Update-OmissionsFromData

            # Update annotations
            Update-AnnotationsFromData

            # Update general settings
            Update-GeneralSettingsFromData
        }

        # Function to update exclusions UI from data
        Function Update-ExclusionsFromData {
            if (-not $syncHash.Exclusions) { return }

            foreach ($exclusion in $syncHash.Exclusions) {
                try {
                    $policyId = $exclusion.Id
                    $productName = $exclusion.Product

                    # Find the exclusion type from the baseline config
                    $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                    if ($baseline -and $baseline.exclusionType -ne "none") {
                        Update-ExclusionCardUI -PolicyId $policyId -ProductName $productName -ExclusionData $exclusion.Data -ExclusionType $baseline.exclusionType
                    }
                }
                catch {
                    Write-Output "Error updating exclusion UI for $($exclusion.Id): $($_.Exception.Message)"
                }
            }
        }

        # Function to update omissions UI from data
        Function Update-OmissionsFromData {
            if (-not $syncHash.Omissions) { return }

            foreach ($omission in $syncHash.Omissions) {
                try {
                    Update-OmissionCardUI -PolicyId $omission.Id -ProductName $omission.Product -Rationale $omission.Rationale -Expiration $omission.Expiration
                }
                catch {
                    Write-Output "Error updating omission UI for $($omission.Id): $($_.Exception.Message)"
                }
            }
        }

        # Function to update annotations UI from data
        Function Update-AnnotationsFromData {
            if (-not $syncHash.Annotations) { return }

            foreach ($annotation in $syncHash.Annotations) {
                try {
                    Update-AnnotationCardUI -PolicyId $annotation.Id -ProductName $annotation.Product -Comment $annotation.Comment
                }
                catch {
                    Write-Output "Error updating annotation UI for $($annotation.Id): $($_.Exception.Message)"
                }
            }
        }

        # Function to update general settings UI from data (Dynamic Version)
        Function Update-GeneralSettingsFromData {
            if (-not $syncHash.GeneralSettings) { return }

            try {
                foreach ($settingKey in $syncHash.GeneralSettings.Keys) {
                    $settingValue = $syncHash.GeneralSettings[$settingKey]

                    # Skip if value is null or empty
                    if ($null -eq $settingValue) { continue }

                    # Find the corresponding XAML control using various naming patterns
                    $control = Find-ControlBySettingName -SettingName $settingKey

                    if ($control) {
                        Update-ControlValue -Control $control -Value $settingValue -SettingKey $settingKey
                    } else {
                        Write-Verbose "No UI control found for setting: $settingKey"
                    }
                }

                # Handle special cases that need custom logic
                Handle-SpecialSettingsCases

            }
            catch {
                Write-Output "Error updating general settings UI: $($_.Exception.Message)"
            }
        }

        # Helper function to find control by setting name
        Function Find-ControlBySettingName {
            param([string]$SettingName)

            # Define naming patterns to try
            $namingPatterns = @(
                $SettingName,                           # Direct name
                "$SettingName`_TextBox",                # SettingName_TextBox
                "$SettingName`_TextBlock",              # SettingName_TextBlock
                "$SettingName`_CheckBox",               # SettingName_CheckBox
                "$SettingName`_ComboBox",               # SettingName_ComboBox
                "$SettingName`_Label",                  # SettingName_Label
                "$SettingName`TextBox",                 # SettingNameTextBox
                "$SettingName`TextBlock",               # SettingNameTextBlock
                "$SettingName`CheckBox",                # SettingNameCheckBox
                "$SettingName`ComboBox",                # SettingNameComboBox
                "$SettingName`Label",                   # SettingNameLabel
                "txt$SettingName",                      # txtSettingName
                "chk$SettingName",                      # chkSettingName
                "cbo$SettingName",                      # cboSettingName
                "lbl$SettingName"                       # lblSettingName
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
        Function Update-ControlValue {
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
                    Update-ComboBoxValue -ComboBox $Control -Value $Value -SettingKey $SettingKey
                }
                'Label' {
                    $Control.Content = $Value
                }
                default {
                    Write-Verbose "Unknown control type for $SettingKey`: $($Control.GetType().Name)"
                }
            }
        }

        # Helper function to update ComboBox values
        Function Update-ComboBoxValue {
            param(
                [System.Windows.Controls.ComboBox]$ComboBox,
                [object]$Value,
                [string]$SettingKey
            )

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
            } else {
                Write-Verbose "Could not find ComboBox item for value: $Value in $SettingKey"
            }
        }

        # Function to handle special cases that need custom logic
        Function Handle-SpecialSettingsCases {
            # Handle ProductNames (checkbox selection)
            if ($syncHash.GeneralSettings.ProductNames) {
                $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                    $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
                }

                # First uncheck all products
                foreach ($checkbox in $allProductCheckboxes) {
                    $checkbox.IsChecked = $false
                }

                # Then check the selected products
                if ($syncHash.GeneralSettings.ProductNames -contains "*") {
                    # Select all products
                    foreach ($checkbox in $allProductCheckboxes) {
                        $checkbox.IsChecked = $true
                    }
                } else {
                    # Select specific products
                    foreach ($productName in $syncHash.GeneralSettings.ProductNames) {
                        $checkbox = $allProductCheckboxes | Where-Object { $_.Tag -eq $productName }
                        if ($checkbox) {
                            $checkbox.IsChecked = $true
                        }
                    }
                }
            }

            # Handle any other special cases here
            # For example, if you have complex controls that need special handling
        }

        # Create a DispatcherTimer for periodic UI updates
        $syncHash.UIUpdateTimer = New-Object System.Windows.Threading.DispatcherTimer
        $syncHash.UIUpdateTimer.Interval = [System.TimeSpan]::FromMilliseconds(500) # Check every 500ms
        $syncHash.UIUpdateTimer.Add_Tick({
            try {
                # Only update if there have been changes
                if ($syncHash.DataChanged) {
                    Update-AllUIFromData
                    $syncHash.DataChanged = $false
                }
            }
            catch {
                Write-Output "Error in UI update timer: $($_.Exception.Message)"
            }
        })

        # Initialize change tracking
        $syncHash.DataChanged = $false
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
        #===========================================================================
        #
        # Load UI Configuration
        #
        #===========================================================================
        #Import configuration file
        $syncHash.UIConfigs = Get-Content -Path $syncHash.UIConfigPath -Raw | ConvertFrom-Json

        If($syncHash.YAMLImport){
            $syncHash.YAMLConfig = Get-Content -Path $syncHash.YAMLImport -Raw | ConvertFrom-Yaml
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
        }

        $syncHash.PreviewTab.IsEnabled = $false

        $syncHash.ExclusionsTab.IsEnabled = $false
        $syncHash.AnnotatePolicyTab.IsEnabled = $false
        $syncHash.OmitPolicyTab.IsEnabled = $false

        foreach ($env in $syncHash.UIConfigs.SupportedM365Environment) {
            $comboItem = New-Object System.Windows.Controls.ComboBoxItem
            $comboItem.Content = "$($env.displayName) ($($env.name))"
            $comboItem.Tag = $env.id

            $syncHash.M365Environment_ComboBox.Items.Add($comboItem)
        }

        # Set selection based on parameter or default to first item
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

            #omissions tab
            $OmissionTab = $syncHash.("$($product.id)OmissionTab")

            $checkBox.Add_Checked({
                $syncHash.Window.Dispatcher.Invoke([action]{
                    #enable the main tabs
                    $syncHash.ExclusionsTab.IsEnabled = $true
                    $syncHash.AnnotatePolicyTab.IsEnabled = $true
                    $syncHash.OmitPolicyTab.IsEnabled = $true

                    #omissions tab
                    $omissionTab = $syncHash.("$($product.id)OmissionTab")
                    $omissionTab.IsEnabled = $true

                    $container = $syncHash.("$($product.id)OmissionContent")
                    if ($container -and $container.Children.Count -eq 0) {
                        New-ProductOmissions -ProductName $product.id -Container $container
                    }

                    #annotations tab
                    $AnnotationTab = $syncHash.("$($product.id)AnnotationTab")
                    $AnnotationTab.IsEnabled = $true
                    $container = $syncHash.("$($product.id)AnnotationContent")
                    if ($container -and $container.Children.Count -eq 0) {
                        New-ProductAnnotations -ProductName $product.id -Container $container
                    }

                    #exclusions tab
                    if ($product.supportsExclusions)
                    {
                        $ExclusionsTab = $syncHash.("$($product.id)ExclusionTab")
                        $ExclusionsTab.IsEnabled = $true
                        $container = $syncHash.("$($product.id)ExclusionContent")
                        if ($container -and $container.Children.Count -eq 0) {
                            New-ProductExclusions -ProductName $product.id -Container $container
                        }
                    }
                })

            }.GetNewClosure())

            $checkBox.Add_Unchecked({
                $OmissionTab.IsEnabled = $false
                $ExclusionsTab.IsEnabled = $false
            }.GetNewClosure())

            <#
            #enable individual product tabs
            if ($product.supportsExclusions) {

                $productTab = $syncHash.("$($product.id)Tab")

                $checkBox.Add_Checked({
                    $productTab.IsEnabled = $true
                }.GetNewClosure())

                $checkBox.Add_Unchecked({
                    $productTab.IsEnabled = $false

                    # Optional: Disable Preview tab if no products are checked
                    $anyChecked = $syncHash.ProductsGrid.Children | Where-Object {
                        $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked -eq $true
                    }

                    if (-not $anyChecked) {
                        $syncHash.PreviewTab.IsEnabled = $false
                    }
                }.GetNewClosure())
            }
            #>
        }
        $ExclusionSupport = $syncHash.UIConfigs.products | Where-Object { $_.supportsExclusions -eq $true } | select -ExpandProperty id
        $syncHash.ExclusionsInfo_TextBlock.Text = ($syncHash.UIConfigs.localeContext.ExclusionsInfo_TextBlock -f ($ExclusionSupport -join ', ').ToUpper())

        Foreach($product in $syncHash.UIConfigs.products) {
            # Initialize the OmissionTab and ExclusionTab for each product
            $exclusionTab = $syncHash.("$($product.id)ExclusionTab")

            if ($product.supportsExclusions) {
                $exclusionTab.Visibility = "Visible"
            }else{
                # Disable the Exclusions tab if the product does not support exclusions
                $exclusionTab.Visibility = "Collapsed"
            }
        }

        # Select All/Deselect All Button
        $syncHash.SelectAllButton.Add_Click({
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }

            # Check current state - if all are checked, we'll deselect all
            $allChecked = $allProductCheckboxes | ForEach-Object { $_.IsChecked } | Where-Object { $_ -eq $true }

            if ($allChecked.Count -eq $allProductCheckboxes.Count) {
                # All are checked, so deselect all
                foreach ($checkbox in $allProductCheckboxes) {
                    $checkbox.IsChecked = $false
                }
                $syncHash.SelectAllButton.Content = "Select All"
            } else {
                # Not all are checked, so select all
                foreach ($checkbox in $allProductCheckboxes) {
                    $checkbox.IsChecked = $true
                }
                $syncHash.SelectAllButton.Content = "Select None"
            }
        })

        #===========================================================================
        #
        # ANNOTATION dynamic controls
        #
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

            # Add save button event handler
            $saveButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_SaveAnnotation", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                # Get the comment text
                $commentTextBox = $this.Parent.Parent.Children | Where-Object { $_.Name -eq ($policyIdWithUnderscores + "_Comment_TextBox") }
                $comment = $commentTextBox.Text

                # Initialize if not exists
                if (-not $syncHash.Annotations) {
                    $syncHash.Annotations = @()
                }

                # Remove existing annotation for this policy
                $syncHash.Annotations = @($syncHash.Annotations | Where-Object { $_.Id -ne $policyId })

                # Only create annotation if comment is not empty
                if (![string]::IsNullOrWhiteSpace($comment)) {
                    $annotation = [PSCustomObject]@{
                        Id = $policyId
                        Product = $ProductName
                        Comment = $comment.Trim()
                    }

                    $syncHash.Annotations += $annotation

                    Write-Output "Annotation added for $policyId. Total annotations: $($syncHash.Annotations.Count)"
                    Write-Output "Comment: $comment"
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

            # Add remove button event handler
            $removeButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_RemoveAnnotation", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove the annotation for [$policyId]?", "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {

                    # Remove annotation for this policy
                    $syncHash.Annotations = @($syncHash.Annotations | Where-Object { $_.Id -ne $policyId })

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
        }

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
        #
        # OMISSIONS dynamic controls
        #
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

            # Add save button event handler
            # Add save button event handler
            $saveButton.Add_Click({
                # Get the correct policy ID from the button name
                $policyIdWithUnderscores = $this.Name.Replace("_SaveOmission", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                # Since button is now in buttonPanel, we need to go up to detailsPanel
                $detailsPanel = $this.Parent.Parent
                $rationaleTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Rationale_TextBox") }
                $expirationTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_Expiration_TextBox") }

                if ([string]::IsNullOrWhiteSpace($rationaleTextBox.Text)) {
                    [System.Windows.MessageBox]::Show("Rationale is required for policy omissions.", "Validation Error", "OK", "Warning")
                    return
                }

                $expirationDate = $null
                if ($expirationTextBox.Text -ne "mm/dd/yyyy" -and -not [string]::IsNullOrWhiteSpace($expirationTextBox.Text)) {
                    try {
                        $expirationDate = [DateTime]::Parse($expirationTextBox.Text).ToString("yyyy-MM-dd")
                    }
                    catch {
                        [System.Windows.MessageBox]::Show("Invalid date format. Please use mm/dd/yyyy format.", "Validation Error", "OK", "Warning")
                        return
                    }
                }

                # Remove existing omission for this policy
                $policyId = $this.Name.Replace("_SaveOmission", "").Replace("_", ".")
                $syncHash.Omissions = @($syncHash.Omissions | Where-Object { $_.Id -ne $policyId })

                # Add new omission
                $omission = [PSCustomObject]@{
                    Id = $policyId
                    Product = $ProductName
                    Rationale = $rationaleTextBox.Text
                    Expiration = $expirationDate
                }

                $syncHash.Omissions += $omission

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

                    # Remove the omission from $syncHash.Omissions psobject
                    $syncHash.Omissions = @($syncHash.Omissions | Where-Object { $_.Id -ne $policyId })

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
        }

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
        #
        # EXCLUSIONS dynamic controls
        #
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

                # Add button functionality
                $addButton.Add_Click({
                    $inputBox = $this.Parent.Children[0]
                    $listPanel = $this.Parent.Parent.Children[1]

                    if (![string]::IsNullOrWhiteSpace($inputBox.Text) -and $inputBox.Text -ne $placeholder) {

                        If($listContainer.Children.Children | Where-Object { $_.Text -contains $inputBox.Text }) {
                            # User already exists, skip
                            return
                        }

                        # Validate input based on valueType
                        $isValid = $true
                        switch ($Field.valueType) {
                            "email" { $isValid = $inputBox.Text -match "^[^\s@]+@[^\s@]+\.[^\s@]+$" }
                            "guid" { $isValid = $inputBox.Text -match "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$" }
                            "domain" { $isValid = $inputBox.Text -match "^[a-zA-Z0-9][a-zA-Z0-9-]*[a-zA-Z0-9]*\.([a-zA-Z]{2,})+$" }
                        }

                        if ($isValid) {
                            # Create item row
                            $itemRow = New-Object System.Windows.Controls.StackPanel
                            $itemRow.Orientation = "Horizontal"
                            $itemRow.Margin = "0,2,0,2"

                            $itemText = New-Object System.Windows.Controls.TextBlock
                            $itemText.Text = $inputBox.Text
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

                                Write-Output "Added $($selectedUsers.Count) users to exclusion list"
                            }
                        }
                        catch {
                            Write-Error "Error selecting users: $($_.Exception.Message)"
                            [System.Windows.MessageBox]::Show("Error selecting users: $($_.Exception.Message)", "Error",
                                                            [System.Windows.MessageBoxButton]::OK,
                                                            [System.Windows.MessageBoxImage]::Error)
                        }
                    }.GetNewClosure())

                    [void]$inputRow.Children.Add($getUsersButton)
                }

                # Replace the existing Get Groups button handler with:
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

                                Write-Output "Added $($selectedGroups.Count) groups to exclusion list"
                            }
                        }
                        catch {
                            Write-Error "Error selecting groups: $($_.Exception.Message)"
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

            # Replace the save button click handler in New-ExclusionCard with this:
            $saveButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_SaveExclusion", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                # Initialize if not exists
                if (-not $syncHash.Exclusions) {
                    $syncHash.Exclusions = @()
                }

                # Remove existing exclusions for this policy
                $syncHash.Exclusions = @($syncHash.Exclusions | Where-Object { $_.Id -ne $policyId })

                # Collect field values by traversing the UI tree
                $detailsPanel = $this.Parent.Parent
                $exclusionData = @{}

                foreach ($field in $exclusionTypeDef.fields) {
                    $fieldName = ($policyIdWithUnderscores + "_" + $ExclusionType + "_" + $field.name)

                    if ($field.type -eq "array") {
                        # Find the list container by traversing the panel children
                        $listContainer = $null
                        foreach ($child in $detailsPanel.Children) {
                            if ($child -is [System.Windows.Controls.StackPanel]) {
                                foreach ($subChild in $child.Children) {
                                    if ($subChild -is [System.Windows.Controls.StackPanel] -and $subChild.Name -eq ($fieldName + "_Container")) {
                                        if ($subChild.Children.Count -gt 1) {
                                            $listContainer = $subChild.Children[1]
                                            break
                                        }
                                    }
                                }
                                if ($listContainer) { break }
                            }
                        }

                        if ($listContainer) {
                            $values = @()
                            foreach ($item in $listContainer.Children) {
                                if ($item.Children.Count -gt 0) {
                                    $values += $item.Children[0].Text
                                }
                            }

                            if ($values.Count -gt 0) {
                                $exclusionData[$field.name] = $values
                            }
                        }
                    } elseif ($field.type -eq "string") {
                        # Find the string control by traversing the panel children
                        $stringControl = $null
                        foreach ($child in $detailsPanel.Children) {
                            if ($child -is [System.Windows.Controls.StackPanel]) {
                                foreach ($subChild in $child.Children) {
                                    if ($subChild -is [System.Windows.Controls.TextBox] -and $subChild.Name -eq $fieldName) {
                                        $stringControl = $subChild
                                        break
                                    }
                                }
                                if ($stringControl) { break }
                            }
                        }

                        if ($stringControl -and ![string]::IsNullOrWhiteSpace($stringControl.Text) -and
                            $stringControl.Text -ne "user1@domain.com;user2@domain.com") {
                            $exclusionData[$field.name] = $stringControl.Text
                        }
                    }
                }

                # Only create exclusion if we have data
                if ($exclusionData.Count -gt 0) {
                    $exclusion = [PSCustomObject]@{
                        Id = $policyId
                        Product = $ProductName
                        TypeName = $ExclusionType
                        Data = $exclusionData
                    }

                    $syncHash.Exclusions += $exclusion

                    Write-Output "Exclusion added for $policyId. Total exclusions: $($syncHash.Exclusions.Count)"
                    Write-Output "Exclusion data: $($exclusionData | ConvertTo-Json -Depth 3)"
                }

                [System.Windows.MessageBox]::Show("[$policyId] exclusion saved successfully.", "Success", "OK", "Information")

                # Make remove button visible and header bold
                $removeButton.Visibility = "Visible"
                $policyHeader.FontWeight = "Bold"

                # collapse details panel
                $detailsPanel.Visibility = "Collapsed"

                #uncheck checkbox
                $checkbox.IsChecked = $false
            }.GetNewClosure())

            # Add remove button event handler
            $removeButton.Add_Click({
                $policyIdWithUnderscores = $this.Name.Replace("_RemoveExclusion", "")
                $policyId = $policyIdWithUnderscores.Replace("_", ".")

                $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove exclusions for [$policyId]?", "Confirm Remove", "YesNo", "Question")
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {

                    # Remove all exclusions for this policy
                    $syncHash.Exclusions = @($syncHash.Exclusions | Where-Object { $_.Id -ne $policyId })

                    # Clear all field values
                    foreach ($field in $exclusionTypeDef.fields) {
                        $fieldName = ($policyIdWithUnderscores + "_" + $ExclusionType + "_" + $field.name)

                        if ($field.type -eq "array") {
                            $listContainer = $syncHash.Window.FindName($fieldName + "_List")
                            if ($listContainer) {
                                $listContainer.Children.Clear()
                            }
                        } elseif ($field.type -eq "string") {
                            $stringControl = $syncHash.Window.FindName($fieldName)
                            if ($stringControl) {
                                if ($field.valueType -eq "semicolonList") {
                                    $stringControl.Text = "user1@domain.com;user2@domain.com"
                                    $stringControl.Foreground = [System.Windows.Media.Brushes]::Gray
                                    $stringControl.FontStyle = "Italic"
                                }
                            }
                        }
                    }

                    [System.Windows.MessageBox]::Show("[$policyId] exclusions removed successfully.", "Success", "OK", "Information")

                    # Hide remove button and unbold header
                    $this.Visibility = "Collapsed"

                    #change weight of policy header
                    $policyHeader.FontWeight = "SemiBold"

                    #uncheck checkbox
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
        # GRAPH controls
        #
        #===========================================================================

        # Enhanced Graph Query Function with Filter Support
        function Invoke-GraphQueryWithFilter {
            param(
                [string]$QueryType,
                $GraphConfig,
                [string]$FilterString,
                [int]$Top = 500,
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
                        throw "Query configuration not found for: $QueryType"
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
                            if ($param -ne '$top') {  # We'll handle $top separately
                                $queryStringParts += "$param=$($queryConfig.queryParameters.$param)"
                            }
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

                    Write-Output "Graph Query URI: $($queryParams.Uri)"

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
        # Placeholder Text Functionality
        #===========================================================================
        If($syncHash.GraphConnected) {
            # Placeholder for Organization Name TextBox
            $tenantDetails = (Invoke-MgGraphRequest -Method GET -Uri "$GraphEndpoint/v1.0/organization").Value
            $tenantName = ($tenantDetails.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
            $syncHash.Organization_TextBox.Text = $tenantName
            $syncHash.Organization_TextBox.Foreground = [System.Windows.Media.Brushes]::Black
        }Else{
            # Organization Name TextBox with placeholder
            $syncHash.Organization_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.Organization_TextBox
            $syncHash.Organization_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $syncHash.Organization_TextBox.FontStyle = [System.Windows.FontStyles]::Italic
        }


        $syncHash.Organization_TextBox.Add_GotFocus({
            if ($syncHash.Organization_TextBox.Text -eq $syncHash.UIConfigs.localePlaceholder.Organization_TextBox) {
                $syncHash.Organization_TextBox.Text = ""
                $syncHash.Organization_TextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.Organization_TextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.Organization_TextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.Organization_TextBox.Text)) {
                $syncHash.Organization_TextBox.Text = $OrganizationPlaceholder
                $syncHash.Organization_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.Organization_TextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        # Organization Name TextBox with placeholder
        $syncHash.OrgName_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.OrgName_TextBox
        $syncHash.OrgName_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.OrgName_TextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.OrgName_TextBox.Add_GotFocus({
            if ($syncHash.OrgName_TextBox.Text -eq $syncHash.UIConfigs.localePlaceholder.OrgName_TextBox) {
                $syncHash.OrgName_TextBox.Text = ""
                $syncHash.OrgName_TextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.OrgName_TextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.OrgName_TextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.OrgName_TextBox.Text)) {
                $syncHash.OrgName_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.OrgName_TextBox
                $syncHash.OrgName_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.OrgName_TextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        # Organization Unit TextBox with placeholder
        $syncHash.OrgUnit_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.OrgUnit_TextBox
        $syncHash.OrgUnit_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.OrgUnit_TextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.OrgUnit_TextBox.Add_GotFocus({
            if ($syncHash.OrgUnit_TextBox.Text -eq $syncHash.UIConfigs.localePlaceholder.OrgUnit_TextBox) {
                $syncHash.OrgUnit_TextBox.Text = ""
                $syncHash.OrgUnit_TextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.OrgUnit_TextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.OrgUnit_TextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.OrgUnit_TextBox.Text)) {
                $syncHash.OrgUnit_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.OrgUnit_TextBox
                $syncHash.OrgUnit_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.OrgUnit_TextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })


        # Description TextBox with placeholder
        $syncHash.Description_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.Description_TextBox
        $syncHash.Description_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.Description_TextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.Description_TextBox.Add_GotFocus({
            if ($syncHash.Description_TextBox.Text -eq $syncHash.UIConfigs.localePlaceholder.Description_TextBox) {
                $syncHash.Description_TextBox.Text = ""
                $syncHash.Description_TextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.Description_TextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.Description_TextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.Description_TextBox.Text)) {
                $syncHash.Description_TextBox.Text = $syncHash.UIConfigs.localePlaceholder.Description_TextBox
                $syncHash.Description_TextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.Description_TextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        #===========================================================================
        # Advanced Tab Toggle Functionality
        #===========================================================================

        # Application Section Toggle
        $syncHash.ApplicationSection_Toggle.Add_Checked({
            $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.ApplicationSection_Toggle.Add_Unchecked({
            $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # Output Section Toggle
        $syncHash.OutputSection_Toggle.Add_Checked({
            $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.OutputSection_Toggle.Add_Unchecked({
            $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # OPA Section Toggle
        $syncHash.OpaSection_Toggle.Add_Checked({
            $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.OpaSection_Toggle.Add_Unchecked({
            $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # General Section Toggle
        $syncHash.GeneralSection_Toggle.Add_Checked({
            $syncHash.GeneralSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.GeneralSection_Toggle.Add_Unchecked({
            $syncHash.GeneralSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        #===========================================================================
        # Browse and Select Button Functionality
        #===========================================================================

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

                # Get user certificates
                $userCerts = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object {
                    $_.HasPrivateKey -and
                    $_.NotAfter -gt (Get-Date) -and
                    $_.Subject -notlike "*Microsoft*"
                } | Sort-Object Subject

                Write-Output "Found $($userCerts.Count) certificates"

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
                $selectedCerts = Show-UISelectionWindow -Title "Select Certificate" -SearchPlaceholder "Search by subject..." -Items $displayCerts -ColumnConfig $columnConfig -SearchProperty "Subject"

                if ($selectedCerts -and $selectedCerts.Count -gt 0) {
                    # Get the first (and only) selected certificate
                    $selectedCert = $selectedCerts[0]

                    $syncHash.CertificateThumbprint_TextBox.Text = $selectedCert.Thumbprint

                }
            }
            catch {

                [System.Windows.MessageBox]::Show(("{0} certificate store: {1}" -f $syncHash.UIConfigs.localeErrorMessages.WindowError, $_.Exception.Message),
                                                "Error",
                                                [System.Windows.MessageBoxButton]::OK,
                                                [System.Windows.MessageBoxImage]::Error)
            }
        })


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
                if (-not $config) {
                    throw "Unsupported graph entity type: $GraphEntityType"
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
                    $selectedItems = Show-UISelectionWindow -Title $config.Title -SearchPlaceholder $config.SearchPlaceholder -Items $displayItems -ColumnConfig $config.ColumnConfig -SearchProperty $config.SearchProperty -AllowMultiple

                    return $selectedItems
                }
                else {
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

         #===========================================================================
        # Wrapper Functions for Backward Compatibility
        #===========================================================================
        function Show-UserSelector {
            param(
                [string]$SearchTerm = "",
                [int]$Top = 100
            )
            return Show-GraphProgressWindow -GraphEntityType "users" -SearchTerm $SearchTerm -Top $Top
        }

        function Show-GroupSelector {
            param(
                [string]$SearchTerm = "",
                [int]$Top = 100
            )
            return Show-GraphProgressWindow -GraphEntityType "groups" -SearchTerm $SearchTerm -Top $Top
        }

        #===========================================================================
        # Universal Selection Function
        #===========================================================================
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
                [string]$SearchProperty = "DisplayName",

                [Parameter()]
                [string]$ReturnProperty = "Id",

                [Parameter()]
                [switch]$AllowMultiple
            )

            try {
                # Create selection window
                $selectionWindow = New-Object System.Windows.Window
                $selectionWindow.Title = $Title
                $selectionWindow.Width = 700
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

                # Create columns based on config
                foreach ($columnKey in $ColumnConfig.Keys) {
                    $column = New-Object System.Windows.Controls.DataGridTextColumn
                    $column.Header = $ColumnConfig[$columnKey].Header
                    $column.Binding = New-Object System.Windows.Data.Binding($columnKey)
                    $column.Width = $ColumnConfig[$columnKey].Width
                    $dataGrid.Columns.Add($column)
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
                    return $selectionWindow.Tag
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

        #===========================================================================
        # Validation Functions
        #===========================================================================
        Function Confirm-UIField {
            param(
                [System.Windows.Controls.Control]$UIElement,
                [string]$RegexPattern,
                [string]$ErrorMessage,
                [string]$PlaceholderText = "",
                [switch]$Required,
                [switch]$ShowMessageBox = $true
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

        #===========================================================================
        # Button Event Handlers
        #===========================================================================

        # Preview & Generate Button
        $syncHash.PreviewButton.Add_Click({
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
        })


        # New Session Button
        $syncHash.NewSessionButton.Add_Click({
            $result = [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeInfoMessages.NewSessionConfirmation, "New Session", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                # Reset all form fields
                Reset-FormFields
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
                    $yamlObject = $yamlContent | ConvertFrom-Yaml

                    # Clear existing data
                    $syncHash.Exclusions = @()
                    $syncHash.Omissions = @()
                    $syncHash.Annotations = @()
                    $syncHash.GeneralSettings = @{}

                    # Import data into the core data structures
                    Import-YamlToDataStructures -Config $yamlObject

                    # Trigger UI update through reactive system
                    Set-DataChanged

                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeInfoMessages.ImportSuccess, "Import Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
                catch {
                    [System.Windows.MessageBox]::Show("Error importing configuration: $($_.Exception.Message)", "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        })

        # Copy to Clipboard Button
        $syncHash.CopyYamlButton.Add_Click({
            try {
                if (![string]::IsNullOrWhiteSpace($syncHash.YamlPreview_TextBox.Text)) {
                    [System.Windows.Clipboard]::SetText($syncHash.YamlPreview_TextBox.Text)
                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeInfoMessages.CopySuccess, "Copy Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    [System.Windows.MessageBox]::Show($syncHash.UIConfigs.localeErrorMessages.CopyError, "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Error copying to clipboard: $($_.Exception.Message)", "Copy Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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

                    [System.Windows.MessageBox]::Show("Configuration saved successfully to: $($saveFileDialog.FileName)", "Save Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Error saving file: $($_.Exception.Message)", "Save Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })



        #===========================================================================
        # Helper Functions
        #===========================================================================
        # Function to import YAML configuration and populate UI fields
        # Function to import YAML data into core data structures (without UI updates)
        Function Import-YamlToDataStructures {
            param($Config)

            try {
                # Get top-level keys regardless of object type
                $topLevelKeys = if ($Config -is [hashtable]) { $Config.Keys } else { $Config.PSObject.Properties.Name }

                #get all products from UIConfigs
                $productIds = $syncHash.UIConfigs.products | Select-Object -ExpandProperty id

                # Import General Settings that are not product
                $generalSettingsFields = $topLevelKeys | Where-Object {$_ -notin $productIds}

                foreach ($field in $generalSettingsFields) {
                    $syncHash.GeneralSettings[$field] = $Config.$field
                }

                # Import Exclusions
                $hasProductSections = $productIds | Where-Object { $topLevelKeys -contains $_ }

                if ($hasProductSections) {
                    foreach ($productName in $productIds) {
                        if ($topLevelKeys -contains $productName) {
                            $productData = $Config.$productName

                            # Get product exclusion keys
                            $productKeys = if ($productData -is [hashtable]) { $productData.Keys } else { $productData.PSObject.Properties.Name }

                            foreach ($policyId in $productKeys) {
                                $policyData = $productData.$policyId

                                # Find the exclusion type from baseline config
                                $baseline = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $policyId }
                                if ($baseline -and $baseline.exclusionType -ne "none") {
                                    $exclusion = [PSCustomObject]@{
                                        Id = $policyId
                                        Product = $productName
                                        Data = $policyData
                                        ExclusionType = $baseline.exclusionType
                                    }
                                    $syncHash.Exclusions += $exclusion
                                }
                            }
                        }
                    }
                }

                # Import Omissions
                if ($topLevelKeys -contains "OmitPolicy") {
                    $omissionKeys = if ($Config.OmitPolicy -is [hashtable]) { $Config.OmitPolicy.Keys } else { $Config.OmitPolicy.PSObject.Properties.Name }

                    foreach ($policyId in $omissionKeys) {
                        $omissionData = $Config.OmitPolicy.$policyId

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
                            $omission = [PSCustomObject]@{
                                Id = $policyId
                                Product = $productName
                                Rationale = $omissionData.Rationale
                                Expiration = $omissionData.Expiration
                            }
                            $syncHash.Omissions += $omission
                        }
                    }
                }

                # Import Annotations
                if ($topLevelKeys -contains "AnnotatePolicy") {
                    $annotationKeys = if ($Config.AnnotatePolicy -is [hashtable]) { $Config.AnnotatePolicy.Keys } else { $Config.AnnotatePolicy.PSObject.Properties.Name }

                    foreach ($policyId in $annotationKeys) {
                        $annotationData = $Config.AnnotatePolicy.$policyId

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
                            $annotation = [PSCustomObject]@{
                                Id = $policyId
                                Product = $productName
                                Comment = $annotationData.Comment
                            }
                            $syncHash.Annotations += $annotation
                        }
                    }
                }

            }
            catch {
                Write-Output "Error importing data: $($_.Exception.Message)"
                throw
            }
        }

        Function Reset-FormFields {
            # Reset core data structures
            $syncHash.Exclusions = @()
            $syncHash.Omissions = @()
            $syncHash.Annotations = @()
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
            $toggleControls = @('ApplicationSection_Toggle', 'OutputSection_Toggle', 'OpaSection_Toggle', 'GeneralSection_Toggle')
            foreach ($toggleName in $toggleControls) {
                if ($syncHash.$toggleName) {
                    $syncHash.$toggleName.IsChecked = $false
                    $contentName = $toggleName.Replace('_Toggle', 'Content')
                    if ($syncHash.$contentName) {
                        $syncHash.$contentName.Visibility = [System.Windows.Visibility]::Collapsed
                    }
                }
            }

            # Mark data as changed to trigger UI update
            Set-DataChanged
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

                            # Handle special formatting for description
                            if ($yamlFieldName -eq 'Description') {
                                $escapedDescription = $currentValue.Replace('"', '""')
                                $yamlPreview += "`n$yamlFieldName`: `"$escapedDescription`""
                            } else {
                                $yamlPreview += "`n$yamlFieldName`: $currentValue"
                            }
                        }
                    }
                }
            }

            $yamlPreview += "`n`n# Configuration Details"

            # Handle ProductNames (existing logic)
            $selectedProducts = @()
            foreach ($item in $syncHash.ProductsGrid.Children) {
                if ($item -is [System.Windows.Controls.CheckBox] -and
                    $item.Name -like "*ProductCheckBox" -and
                    $item.IsChecked) {
                    $selectedProducts += $item.Tag
                }
            }

            # Check if all products are selected
            $allProductCheckboxes = $syncHash.ProductsGrid.Children | Where-Object {
                $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "*ProductCheckBox"
            }

            if ($selectedProducts.Count -eq $allProductCheckboxes.Count -and $selectedProducts.Count -gt 0) {
                $yamlPreview += "`nProductNames: ['*']"
            } else {
                $yamlPreview += "`nProductNames:"
                foreach ($product in $selectedProducts) {
                    $yamlPreview += "`n  - $product"
                }
            }

            # Handle M365Environment
            #$selectedEnv = $syncHash.M365Environment_ComboBox.SelectedItem.Tag
            $selectedEnv = $syncHash.UIConfigs.SupportedM365Environment | Where-Object { $_.id -eq $syncHash.M365Environment_ComboBox.SelectedItem.Tag } | Select-Object -ExpandProperty name
            $yamlPreview += "`n`nM365Environment: $selectedEnv"

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
                }
            }

            # Handle Exclusions (existing logic)
            if ($syncHash.Exclusions.Count -gt 0) {
                $yamlPreview += "`n`n# Policy Exclusions"

                # Group exclusions by product and sort by ID
                $syncHash.Exclusions | Group-Object Product | ForEach-Object {
                    $productName = $_.Name
                    $productExclusions = $_.Group | Sort-Object Id

                    $yamlPreview += "`n$($productName):"

                    foreach ($exclusion in $productExclusions) {
                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$productName | Where-Object { $_.id -eq $exclusion.Id }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                            $yamlPreview += "`n  $($exclusion.Id):"

                            # Get exclusion type definition
                            $exclusionTypeDef = $syncHash.UIConfigs.exclusionTypes.$($exclusion.TypeName)
                            if ($exclusionTypeDef) {
                                # Check if this exclusion type has only one field and if that field name matches the type name
                                $fields = $exclusionTypeDef.fields
                                $isSingleFieldSameName = ($fields.Count -eq 1) -and ($fields[0].name -eq $exclusionTypeDef.name)

                                if ($isSingleFieldSameName) {
                                    # Skip the extra nesting level - go directly to the field values
                                    $fieldName = $fields[0].name
                                    $fieldValue = $exclusion.Data[$fieldName]
                                    $yamlPreview += "`n    $($fieldName):"

                                    if ($fieldValue -is [array]) {
                                        # Handle array values
                                        foreach ($value in $fieldValue) {
                                            $yamlPreview += "`n      - $value"
                                        }
                                    } else {
                                        # Handle single values
                                        $yamlPreview += "`n      - $fieldValue"
                                    }
                                } else {
                                    # Multiple fields or field name doesn't match type name - use normal nesting
                                    $yamlPreview += "`n    $($exclusionTypeDef.name):"

                                    # Add each field from the exclusion data
                                    foreach ($fieldName in $exclusion.Data.Keys) {
                                        $fieldValue = $exclusion.Data[$fieldName]
                                        $yamlPreview += "`n      $($fieldName):"

                                        if ($fieldValue -is [array]) {
                                            # Handle array values
                                            foreach ($value in $fieldValue) {
                                                $yamlPreview += "`n        - $value"
                                            }
                                        } else {
                                            # Handle single values
                                            $yamlPreview += "`n        - $fieldValue"
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            # Handle Annotations
            if ($syncHash.Annotations.Count -gt 0) {
                $yamlPreview += "`n`nAnnotatePolicy:"

                # Group annotations by product
                $syncHash.Annotations | Group-Object Product | ForEach-Object {
                    $product = $_.Name
                    $yamlPreview += "`n  # $product Annotations:"

                    foreach ($item in $_.Group | Sort-Object Id) {
                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$product | Where-Object { $_.id -eq $item.Id }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $($item.Id):"

                        if ($item.Comment -match "`n") {
                            # Use quoted string format with \n for line breaks
                            $escapedComment = $item.Comment.Replace('"', '""').Replace("`n", "\n")
                            $yamlPreview += "`n    Comment: `"$escapedComment`""
                        } else {
                            $yamlPreview += "`n    Comment: `"$($item.Comment)`""
                        }
                    }
                }
            }

            # Handle Omissions
            if ($syncHash.Omissions.Count -gt 0) {
                $yamlPreview += "`n`nOmitPolicy:"

                # Group omissions by product
                $syncHash.Omissions | Group-Object Product | ForEach-Object {
                    $product = $_.Name
                    $yamlPreview += "`n  # $product Omissions:"

                    foreach ($item in $_.Group | Sort-Object Id) {
                        # Get policy details from baselines
                        $policyInfo = $syncHash.UIConfigs.baselines.$product | Where-Object { $_.id -eq $item.Id }

                        if ($policyInfo) {
                            # Add policy comment with description
                            $yamlPreview += "`n  # $($policyInfo.name)"
                        }

                        $yamlPreview += "`n  $($item.Id):"
                        $yamlPreview += "`n    Rationale: $($item.Rationale)"
                        if ($item.Expiration) {
                            $yamlPreview += "`n    Expiration: $($item.Expiration)"
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
        }

        # CLOSE and LAUNCH UI
        #Closes UI objects and exits (within runspace)
        Function Close-UIMainWindow
        {
            if ($syncHash.hadCritError) { Write-UILogEntry -Message ("Critical error occurred, closing UI: {0}" -f $syncHash.Error) -Source 'Close-UIMainWindow' -Severity 3 }
            #if runspace has not errored Dispose the UI
            if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
        }

        #Add smooth closing for Window
        $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
    	$syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-UIMainWindow})
    	$syncHash.Window.Add_Closed({
            if ($syncHash.UIUpdateTimer) {
                $syncHash.UIUpdateTimer.Stop()
                $syncHash.UIUpdateTimer = $null
            }
            $syncHash.isClosed = $True
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
<#

#run UI
Invoke-SCuBAConfigAppUI

#UI functions used within runspace

Set-DataChanged
Update-AllUIFromData
Update-ExclusionsFromData
Update-OmissionsFromData
Update-AnnotationsFromData
Update-GeneralSettingsFromData

Show-UISelectionWindow
Show-GraphProgressWindow

Invoke-GraphQueryWithFilter
New-ProductExclusions
New-ExclusionCard
New-ExclusionFieldControl

New-ProductOmissions
New-OmissionCard

New-ProductAnnotations
New-AnnotationCard



Show-GroupSelector
Show-UserSelector
Confirm-UIField

Reset-FormFields
New-YamlPreview

Close-UIMainWindow

#>