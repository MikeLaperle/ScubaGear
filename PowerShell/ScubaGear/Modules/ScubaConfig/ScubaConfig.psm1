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
        $YAMLConfig,

        [switch]$Online,

        [switch]$Passthru
    )

    [string]${CmdletName} = $MyInvocation.MyCommand
    Write-Verbose ("{0}: Sequencer started" -f ${CmdletName})

    # build a hash table with locale data to pass to runspace
    $syncHash = [hashtable]::Synchronized(@{})
    $Runspace =[runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $Runspace
    $syncHash.Online = $Online.IsPresent
    $syncHash.XamlPath = "$PSScriptRoot\ScubaConfigAppUI.xaml"
    $syncHash.UIConfigPath = "$PSScriptRoot\ScubaConfig.json"
    $syncHash.YAMLImport = $YAMLConfig
    $syncHash.Exclusions = @()
    $syncHash.Omissions = @()
    $syncHash.ExportSettings = @{}
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

        #===========================================================================
        # Store Form Objects In PowerShell
        #===========================================================================
        $UIXML.SelectNodes("//*[@Name]") | %{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

        # INNER  FUNCTIONS
        #Closes UI objects and exits (within runspace)
        Function Close-UIMainWindow
        {
            if ($syncHash.hadCritError) { Write-UILogEntry -Message ("Critical error occurred, closing UI: {0}" -f $syncHash.Error) -Source 'Close-UIMainWindow' -Severity 3 }
            #if runspace has not errored Dispose the UI
            if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
        }
        #Import configuration file
        $syncHash.UIConfigs = Get-Content -Path $syncHash.UIConfigPath -Raw | ConvertFrom-Json

        If($syncHash.YAMLImport){
            $syncHash.YAMLConfig = Get-Content -Path $syncHash.YAMLImport -Raw | ConvertFrom-Yaml
        }

        #$syncHash.AadTab.IsEnabled = $false
        #$syncHash.DefenderTab.IsEnabled = $false
        #$syncHash.ExoTab.IsEnabled = $false
        $syncHash.PreviewTab.IsEnabled = $false
        #===========================================================================
        # Populate Dynamic Controls from Configuration
        #===========================================================================
        # Populate M365 Environment ComboBox
        foreach ($env in $syncHash.UIConfigs.SupportedM365Environment) {
            $syncHash.EnvironmentComboBox.Items.Add($env.displayName)
        }
        # Set default selection
        $syncHash.EnvironmentComboBox.SelectedIndex = 0

        <#populate each baseline items
        foreach ($baseline in $syncHash.UIConfigs.baselines.aad | Where-Object {$_.exclusionType -ne 'none'}) {
            $ListName = ($baseline.id + ": " + $baseline.name)
            $syncHash.AadBaselineComboBox.Items.Add($ListName)
        }

        foreach ($baseline in $syncHash.UIConfigs.baselines.defender | Where-Object {$_.exclusionType -ne 'none'}) {
            $ListName = ($baseline.id + ": " + $baseline.name)
            $syncHash.DefenderBaselineComboBox.Items.Add($ListName)
        }

        foreach ($baseline in $syncHash.UIConfigs.baselines.exo | Where-Object {$_.exclusionType -ne 'none'}) {
            $ListName = ($baseline.id + ": " + $baseline.name)
            $syncHash.ExoBaselineComboBox.Items.Add($ListName)
        }
        #>

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

            $syncHash.ProductsGrid.Children.Add($checkBox)

            #omissions tab
            $OmissionTab = $syncHash.("$($product.id)OmissionTab")

            $checkBox.Add_Checked({
                $syncHash.Window.Dispatcher.Invoke([action]{
                    $omissionTab = $syncHash.("$($product.id)OmissionTab")
                    $omissionTab.IsEnabled = $true

                    $container = $syncHash.("$($product.id)OmissionContent")
                    if ($container -and $container.Children.Count -eq 0) {
                        New-ProductOmissions -ProductName $product.id -Container $container
                    }
                })
            }.GetNewClosure())

            $checkBox.Add_Checked({
                if ($product.supportsExclusions)
                {
                    $syncHash.Window.Dispatcher.Invoke([action]{
                        $ExclusionTab = $syncHash.("$($product.id)ExclusionTab")
                        $ExclusionTab.IsEnabled = $true

                        $container = $syncHash.("$($product.id)ExclusionContent")
                        if ($container -and $container.Children.Count -eq 0) {
                            New-ProductExclusions -ProductName $product.id -Container $container
                        }
                    })
                }
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
        $syncHash.ExclusionsInfoTextBlock.Text = ("Exclusions are supported for the following products: {0}" -f ($ExclusionSupport -join ', ').ToUpper())

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
        #===========================================================================
        # Omissions dynamic controls
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

            <# Policy description
            $policyDesc = New-Object System.Windows.Controls.TextBlock
            $policyDesc.Text = $PolicyDescription
            $policyDesc.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
            $policyDesc.TextWrapping = "Wrap"
            [void]$policyInfoStack.Children.Add($policyDesc)
            #>

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
            $rationaleTextBox.Name = ($PolicyId.replace('.', '_') + "_RationaleTextBox")
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
            $expirationTextBox.Name = ($PolicyId.replace('.', '_') + "_ExpirationTextBox")
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

            # Style the remove button
            #[void]$detailsPanel.Children.Add($saveButton)
            #[void]$detailsPanel.Children.Add($removeButton)

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
                $rationaleTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_RationaleTextBox") }
                $expirationTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_ExpirationTextBox") }

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
                    $rationaleTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_RationaleTextBox") }
                    $expirationTextBox = $detailsPanel.Children | Where-Object { $_.Name -eq ($policyId.replace('.', '_') + "_ExpirationTextBox") }

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
            $baselines = $syncHash.UIConfigs.baselines.$ProductName

            if ($baselines -and $baselines.Count -gt 0) {
                #TEST $baseline = $baselines[0]
                foreach ($baseline in $baselines) {
                    $card = New-OmissionCard -PolicyId $baseline.id -ProductName $ProductName -PolicyName $baseline.name -PolicyDescription $baseline.description
                    $Container.Children.Add($card)
                }
            } else {
                # No baselines available
                $noDataText = New-Object System.Windows.Controls.TextBlock
                $noDataText.Text = "No policies available for omission in this product."
                $noDataText.Foreground = $syncHash.Window.FindResource("MutedTextBrush")
                $noDataText.FontStyle = "Italic"
                $noDataText.HorizontalAlignment = "Center"
                $noDataText.Margin = "0,50,0,0"
                $Container.Children.Add($noDataText)
            }
        }

        #===========================================================================
        # Exclusions dynamic controls
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
                $inputTextBox.Name = $fieldName + "_Input"
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
                $stringTextBox.Name = $fieldName
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

                    Write-Host "Exclusion added for $policyId. Total exclusions: $($syncHash.Exclusions.Count)" -ForegroundColor Green
                    Write-Host "Exclusion data: $($exclusionData | ConvertTo-Json -Depth 3)" -ForegroundColor Cyan
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
            $baselines = $syncHash.UIConfigs.baselines.$ProductName | Where-Object { $_.exclusionType -and $_.exclusionType -ne 'none' }

            if ($baselines -and $baselines.Count -gt 0) {
                foreach ($baseline in $baselines) {
                    $card = New-ExclusionCard -PolicyId $baseline.id -ProductName $ProductName -PolicyName $baseline.name -PolicyDescription $baseline.description -ExclusionType $baseline.exclusionType
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
        # Placeholder Text Functionality
        #===========================================================================
        # Organization Name TextBox with placeholder
        $syncHash.OrganizationPlaceholder = "Enter tenant name (e.g., example.onmicrosoft.com)"
        $syncHash.OrganizationTextBox.Text = $syncHash.OrganizationPlaceholder
        $syncHash.OrganizationTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.OrganizationTextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.OrganizationTextBox.Add_GotFocus({
            if ($syncHash.OrganizationTextBox.Text -eq $syncHash.OrganizationPlaceholder) {
                $syncHash.OrganizationTextBox.Text = ""
                $syncHash.OrganizationTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.OrganizationTextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.OrganizationTextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.OrganizationTextBox.Text)) {
                $syncHash.OrganizationTextBox.Text = $OrganizationPlaceholder
                $syncHash.OrganizationTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.OrganizationTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        # Organization Name TextBox with placeholder
        $syncHash.OrgNamePlaceholder = "Enter organization name"
        $syncHash.OrgNameTextBox.Text = $syncHash.OrgNamePlaceholder
        $syncHash.OrgNameTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.OrgNameTextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.OrgNameTextBox.Add_GotFocus({
            if ($syncHash.OrgNameTextBox.Text -eq $syncHash.OrgNamePlaceholder) {
                $syncHash.OrgNameTextBox.Text = ""
                $syncHash.OrgNameTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.OrgNameTextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.OrgNameTextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.OrgNameTextBox.Text)) {
                $syncHash.OrgNameTextBox.Text = $syncHash.OrgNamePlaceholder
                $syncHash.OrgNameTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.OrgNameTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        # Organization Unit TextBox with placeholder
        $syncHash.OrgUnitPlaceholder = "Enter organizational unit (optional)"
        $syncHash.OrgUnitTextBox.Text = $syncHash.OrgUnitPlaceholder
        $syncHash.OrgUnitTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.OrgUnitTextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.OrgUnitTextBox.Add_GotFocus({
            if ($syncHash.OrgUnitTextBox.Text -eq $syncHash.OrgUnitPlaceholder) {
                $syncHash.OrgUnitTextBox.Text = ""
                $syncHash.OrgUnitTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.OrgUnitTextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.OrgUnitTextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.OrgUnitTextBox.Text)) {
                $syncHash.OrgUnitTextBox.Text = $syncHash.OrgUnitPlaceholder
                $syncHash.OrgUnitTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.OrgUnitTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })


        # Description TextBox with placeholder
        $syncHash.DescriptionPlaceholder = "Enter a description for this configuration (optional)"
        $syncHash.DescriptionTextBox.Text = $syncHash.DescriptionPlaceholder
        $syncHash.DescriptionTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
        $syncHash.DescriptionTextBox.FontStyle = [System.Windows.FontStyles]::Italic

        $syncHash.DescriptionTextBox.Add_GotFocus({
            if ($syncHash.DescriptionTextBox.Text -eq $syncHash.DescriptionPlaceholder) {
                $syncHash.DescriptionTextBox.Text = ""
                $syncHash.DescriptionTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                $syncHash.DescriptionTextBox.FontStyle = [System.Windows.FontStyles]::Normal
            }
        })

        $syncHash.DescriptionTextBox.Add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($syncHash.DescriptionTextBox.Text)) {
                $syncHash.DescriptionTextBox.Text = $syncHash.DescriptionPlaceholder
                $syncHash.DescriptionTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
                $syncHash.DescriptionTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            }
        })

        #===========================================================================
        # Advanced Tab Toggle Functionality
        #===========================================================================

        # Application Section Toggle
        $syncHash.ApplicationSectionToggle.Add_Checked({
            $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.ApplicationSectionToggle.Add_Unchecked({
            $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # Output Section Toggle
        $syncHash.OutputSectionToggle.Add_Checked({
            $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.OutputSectionToggle.Add_Unchecked({
            $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # OPA Section Toggle
        $syncHash.OpaSectionToggle.Add_Checked({
            $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.OpaSectionToggle.Add_Unchecked({
            $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
        })

        # General Section Toggle
        $syncHash.GeneralSectionToggle.Add_Checked({
            $syncHash.GeneralSectionContent.Visibility = [System.Windows.Visibility]::Visible
        })

        $syncHash.GeneralSectionToggle.Add_Unchecked({
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

            if ($syncHash.OutPathTextBox.Text -ne "." -and (Test-Path $syncHash.OutPathTextBox.Text)) {
                $folderDialog.SelectedPath = $syncHash.OutPathTextBox.Text
            }

            $result = $folderDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $syncHash.OutPathTextBox.Text = $folderDialog.SelectedPath
                New-YamlPreview
            }
        })

        # Browse OPA Path Button
        $syncHash.BrowseOpaPathButton.Add_Click({
            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderDialog.Description = "Select OPA Path"
            $folderDialog.ShowNewFolderButton = $true

            if ($syncHash.OpaPathTextBox.Text -ne "." -and (Test-Path $syncHash.OpaPathTextBox.Text)) {
                $folderDialog.SelectedPath = $syncHash.OpaPathTextBox.Text
            }

            $result = $folderDialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $syncHash.OpaPathTextBox.Text = $folderDialog.SelectedPath
                New-YamlPreview
            }
        })

        # Select Certificate Button
        $syncHash.SelectCertificateButton.Add_Click({
            Show-CertificateSelector
        })

        Function Show-CertificateSelector {
            try {
                # Get user certificates
                $userCerts = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object {
                    $_.HasPrivateKey -and
                    $_.NotAfter -gt (Get-Date) -and
                    $_.Subject -notlike "*Microsoft*"
                } | Sort-Object Subject

                if ($userCerts.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No suitable certificates found in the current user's personal certificate store.",
                                                    "No Certificates",
                                                    [System.Windows.MessageBoxButton]::OK,
                                                    [System.Windows.MessageBoxImage]::Information)
                    return
                }

                # Create certificate selection window
                $certWindow = New-Object System.Windows.Window
                $certWindow.Title = "Select Certificate"
                $certWindow.Width = 600
                $certWindow.Height = 400
                $certWindow.WindowStartupLocation = "CenterOwner"
                $certWindow.Owner = $syncHash.Window
                $certWindow.Background = [System.Windows.Media.Brushes]::White

                # Create main grid
                $mainGrid = New-Object System.Windows.Controls.Grid
                $rowDef1 = New-Object System.Windows.Controls.RowDefinition
                $rowDef1.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
                $rowDef2 = New-Object System.Windows.Controls.RowDefinition
                $rowDef2.Height = [System.Windows.GridLength]::Auto
                $mainGrid.RowDefinitions.Add($rowDef1)
                $mainGrid.RowDefinitions.Add($rowDef2)

                # Create DataGrid for certificates
                $dataGrid = New-Object System.Windows.Controls.DataGrid
                $dataGrid.AutoGenerateColumns = $false
                $dataGrid.CanUserAddRows = $false
                $dataGrid.CanUserDeleteRows = $false
                $dataGrid.IsReadOnly = $true
                $dataGrid.SelectionMode = [System.Windows.Controls.DataGridSelectionMode]::Single
                $dataGrid.GridLinesVisibility = [System.Windows.Controls.DataGridGridLinesVisibility]::Horizontal
                $dataGrid.HeadersVisibility = [System.Windows.Controls.DataGridHeadersVisibility]::Column
                $dataGrid.Margin = "10"

                # Create columns
                $subjectColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $subjectColumn.Header = "Subject"
                $subjectColumn.Binding = New-Object System.Windows.Data.Binding("Subject")
                $subjectColumn.Width = 250
                $dataGrid.Columns.Add($subjectColumn)

                $issuerColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $issuerColumn.Header = "Issued By"
                $issuerColumn.Binding = New-Object System.Windows.Data.Binding("Issuer")
                $issuerColumn.Width = 200
                $dataGrid.Columns.Add($issuerColumn)

                $expiryColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $expiryColumn.Header = "Expires"
                $expiryColumn.Binding = New-Object System.Windows.Data.Binding("NotAfter")
                $expiryColumn.Width = 100
                $dataGrid.Columns.Add($expiryColumn)

                $thumbprintColumn = New-Object System.Windows.Controls.DataGridTextColumn
                $thumbprintColumn.Header = "Thumbprint"
                $thumbprintColumn.Binding = New-Object System.Windows.Data.Binding("Thumbprint")
                $thumbprintColumn.Width = 120
                $dataGrid.Columns.Add($thumbprintColumn)

                # Bind data
                $certData = $userCerts | ForEach-Object {
                    [PSCustomObject]@{
                        Subject = $_.Subject
                        Issuer = $_.Issuer
                        NotAfter = $_.NotAfter.ToString("yyyy-MM-dd")
                        Thumbprint = $_.Thumbprint
                        Certificate = $_
                    }
                }
                $dataGrid.ItemsSource = $certData

                [System.Windows.Controls.Grid]::SetRow($dataGrid, 0)
                $mainGrid.Children.Add($dataGrid)

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

                $buttonPanel.Children.Add($selectButton)
                $buttonPanel.Children.Add($cancelButton)

                [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
                $mainGrid.Children.Add($buttonPanel)

                $certWindow.Content = $mainGrid

                # Event handlers
                $selectButton.Add_Click({
                    if ($dataGrid.SelectedItem) {
                        $selectedCert = $dataGrid.SelectedItem.Certificate
                        $syncHash.CertificateTextBox.Text = $selectedCert.Thumbprint
                        $certWindow.DialogResult = $true
                        $certWindow.Close()
                        New-YamlPreview
                    } else {
                        [System.Windows.MessageBox]::Show("Please select a certificate.", "No Selection",
                                                        [System.Windows.MessageBoxButton]::OK,
                                                        [System.Windows.MessageBoxImage]::Warning)
                    }
                })

                $cancelButton.Add_Click({
                    $certWindow.DialogResult = $false
                    $certWindow.Close()
                })

                $dataGrid.Add_MouseDoubleClick({
                    if ($dataGrid.SelectedItem) {
                        $selectedCert = $dataGrid.SelectedItem.Certificate
                        $syncHash.CertificateTextBox.Text = $selectedCert.Thumbprint
                        $certWindow.DialogResult = $true
                        $certWindow.Close()
                        New-YamlPreview
                    }
                })

                # Show dialog
                $certWindow.ShowDialog()
            }
            catch {
                [System.Windows.MessageBox]::Show("Error loading certificates: $($_.Exception.Message)",
                                                "Error",
                                                [System.Windows.MessageBoxButton]::OK,
                                                [System.Windows.MessageBoxImage]::Error)
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
            $orgValid = Confirm-UIField -UIElement $syncHash.OrganizationTextBox `
                                       -RegexPattern "^(.*\.)?(onmicrosoft\.com|onmicrosoft\.us)$" `
                                       -ErrorMessage "Organization Name is required and must be in format: name.onmicrosoft.com or name.onmicrosoft.us" `
                                       -PlaceholderText $syncHash.OrganizationPlaceholder `
                                       -Required `
                                       -ShowMessageBox:$false

            if (-not $orgValid) {
                $overallValid = $false
                $errorMessages += "Organization Name is required and must be in format: name.onmicrosoft.com or name.onmicrosoft.us"
            }

            # Advanced Tab Validations (only if sections are toggled on)

            # Application Section Validations
            if ($syncHash.ApplicationSectionToggle.IsChecked) {

                # AppID validation (GUID format)


                    $appIdValid = Confirm-UIField -UIElement $syncHash.AppIdTextBox `
                                                 -RegexPattern "^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$" `
                                                 -ErrorMessage "Application ID must be in GUID format (e.g., 12345678-1234-1234-1234-123456789abc)" `
                                                 -PlaceholderText "Your Application ID" `
                                                 -ShowMessageBox:$false

                    if (-not $appIdValid) {
                        $overallValid = $false
                        $errorMessages += "Application ID must be in GUID format"
                    }


                # Certificate Thumbprint validation (40 character hex)


                    $certValid = Confirm-UIField -UIElement $syncHash.CertificateTextBox `
                                                -RegexPattern "^[0-9a-fA-F]{40}$" `
                                                -ErrorMessage "Certificate Thumbprint must be 40 hexadecimal characters" `
                                                -PlaceholderText "Certificate Thumbprint" `
                                                -ShowMessageBox:$false

                    if (-not $certValid) {
                        $overallValid = $false
                        $errorMessages += "Certificate Thumbprint must be 40 hexadecimal characters"
                    }

            }

            # Show consolidated error message if there are validation errors
            if (-not $overallValid) {
                $syncHash.PreviewTab.IsEnabled = $false
                $consolidatedMessage = "Please fix the following validation errors:`n`n" + ($errorMessages -join "`n")
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
            $result = [System.Windows.MessageBox]::Show("Are you sure you want to start a new session? All current data will be lost.", "New Session", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
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

                    # Store the imported configuration for potential timer-based updates
                    $syncHash.ImportedConfig = $yamlObject

                    # Import and populate all form fields
                    Import-YamlConfiguration -Config $yamlObject

                    [System.Windows.MessageBox]::Show("Configuration imported successfully.", "Import Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
                catch {
                    [System.Windows.MessageBox]::Show("Error importing configuration: $($_.Exception.Message)", "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        })

        # Copy to Clipboard Button
        $syncHash.CopyYamlButton.Add_Click({
            try {
                if (![string]::IsNullOrWhiteSpace($syncHash.YamlPreviewTextBox.Text)) {
                    [System.Windows.Clipboard]::SetText($syncHash.YamlPreviewTextBox.Text)
                    [System.Windows.MessageBox]::Show("YAML configuration copied to clipboard successfully.", "Copy Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    [System.Windows.MessageBox]::Show("No YAML content to copy. Please generate the configuration first.", "Nothing to Copy", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Error copying to clipboard: $($_.Exception.Message)", "Copy Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })

        # Download YAML Button
        $syncHash.DownloadYamlButton.Add_Click({
            try {
                if ([string]::IsNullOrWhiteSpace($syncHash.YamlPreviewTextBox.Text)) {
                    [System.Windows.MessageBox]::Show("No YAML content to download. Please generate the configuration first.", "Nothing to Download", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }

                # Generate filename based on organization name
                $orgName = $syncHash.OrganizationTextBox.Text
                if ([string]::IsNullOrWhiteSpace($orgName) -or $orgName -eq $syncHash.OrganizationPlaceholder) {
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
                    $yamlContent = $syncHash.YamlPreviewTextBox.Text
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
        Function Import-YamlConfiguration {
            param($Config)

            try {
                # Reset form first
                Reset-FormFields

                # Main tab fields
                if ($Config.Organization) {
                    $syncHash.OrganizationTextBox.Text = $Config.Organization
                    $syncHash.OrganizationTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                    $syncHash.OrganizationTextBox.FontStyle = [System.Windows.FontStyles]::Normal
                }

                if ($Config.OrgName) {
                    $syncHash.OrgNameTextBox.Text = $Config.OrgName
                    $syncHash.OrgNameTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                    $syncHash.OrgNameTextBox.FontStyle = [System.Windows.FontStyles]::Normal
                }

                if ($Config.OrgUnit) {
                    $syncHash.OrgUnitTextBox.Text = $Config.OrgUnit
                    $syncHash.OrgUnitTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                    $syncHash.OrgUnitTextBox.FontStyle = [System.Windows.FontStyles]::Normal
                }

                if ($Config.Description) {
                    $syncHash.DescriptionTextBox.Text = $Config.Description
                    $syncHash.DescriptionTextBox.Foreground = [System.Windows.Media.Brushes]::Black
                    $syncHash.DescriptionTextBox.FontStyle = [System.Windows.FontStyles]::Normal
                }

                # M365 Environment
                if ($Config.M365Environment) {
                    $envItem = $syncHash.UIConfigs.SupportedM365Environment | Where-Object { $_.id -eq $Config.M365Environment }
                    if ($envItem) {
                        $syncHash.EnvironmentComboBox.SelectedItem = $envItem.displayName
                    }
                }

                # Products
                if ($Config.ProductNames) {
                    foreach ($item in $syncHash.ProductsGrid.Children) {
                        if ($item -is [System.Windows.Controls.CheckBox] -and
                            $item.Name -like "*ProductCheckBox") {

                            # Check if this product is in the imported list
                            if ($Config.ProductNames -contains $item.Tag) {
                                $item.IsChecked = $true
                            }
                        }
                    }
                }

                # Advanced Tab - Application Section
                if ($Config.AppId -or $Config.CertificateThumbprint) {
                    $syncHash.ApplicationSectionToggle.IsChecked = $true
                    $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Visible

                    if ($Config.AppId) {
                        $syncHash.AppIdTextBox.Text = $Config.AppId
                    }

                    if ($Config.CertificateThumbprint) {
                        $syncHash.CertificateTextBox.Text = $Config.CertificateThumbprint
                    }
                }

                # Advanced Tab - Output Section
                if ($Config.OutPath -or $Config.OutFolderName -or $Config.OutJsonFileName -or $Config.OutCsvFileName) {
                    $syncHash.OutputSectionToggle.IsChecked = $true
                    $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Visible

                    if ($Config.OutPath) {
                        $syncHash.OutPathTextBox.Text = $Config.OutPath
                    }

                    if ($Config.OutFolderName) {
                        $syncHash.OutFolderNameTextBox.Text = $Config.OutFolderName
                    }

                    if ($Config.OutJsonFileName) {
                        $syncHash.OutJsonFileNameTextBox.Text = $Config.OutJsonFileName
                    }

                    if ($Config.OutCsvFileName) {
                        $syncHash.OutCsvFileNameTextBox.Text = $Config.OutCsvFileName
                    }
                }

                # Advanced Tab - OPA Section
                if ($Config.OPAPath) {
                    $syncHash.OpaSectionToggle.IsChecked = $true
                    $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Visible
                    $syncHash.OpaPathTextBox.Text = $Config.OPAPath
                }

                # Advanced Tab - General Section
                if ($Config.PSObject.Properties.Name -contains "LogIn" -or $Config.PSObject.Properties.Name -contains "DisconnectOnExit") {
                    $syncHash.GeneralSectionToggle.IsChecked = $true
                    $syncHash.GeneralSectionContent.Visibility = [System.Windows.Visibility]::Visible

                    if ($Config.PSObject.Properties.Name -contains "LogIn") {
                        $syncHash.LogInCheckBox.IsChecked = [System.Convert]::ToBoolean($Config.LogIn)
                    }

                    if ($Config.PSObject.Properties.Name -contains "DisconnectOnExit") {
                        $syncHash.DisconnectOnExitCheckBox.IsChecked = [System.Convert]::ToBoolean($Config.DisconnectOnExit)
                    }
                }

                # Validate and enable preview if organization is valid
                if (Confirm-UIFields) {
                    $syncHash.PreviewTab.IsEnabled = $true
                }

                # Force update the preview
                New-YamlPreview

            }
            catch {
                [System.Windows.MessageBox]::Show("Error populating form fields: $($_.Exception.Message)", "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }

        Function Reset-FormFields {
            # Reset Tenane Name
            $syncHash.OrganizationTextBox.Text = $syncHash.OrganizationPlaceholder
            $syncHash.OrganizationTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $syncHash.OrganizationTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            $syncHash.OrganizationTextBox.BorderBrush = [System.Windows.Media.Brushes]::Gray
            $syncHash.OrganizationTextBox.BorderThickness = "1"

            # Reset Organization Name TextBox
            $syncHash.OrgNameTextBox.Text = $syncHash.OrgNamePlaceholder
            $syncHash.OrgNameTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $syncHash.OrgNameTextBox.FontStyle = [System.Windows.FontStyles]::Italic
            $syncHash.OrgNameTextBox.BorderBrush = [System.Windows.Media.Brushes]::Gray
            $syncHash.OrgNameTextBox.BorderThickness = "1"

            # Reset Organization Unit
            $syncHash.OrgUnitTextBox.Text = $syncHash.OrgUnitPlaceholder
            $syncHash.OrgUnitTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $syncHash.OrgUnitTextBox.FontStyle = [System.Windows.FontStyles]::Italic

            # Reset Description
            $syncHash.DescriptionTextBox.Text = $syncHash.DescriptionPlaceholder
            $syncHash.DescriptionTextBox.Foreground = [System.Windows.Media.Brushes]::Gray
            $syncHash.DescriptionTextBox.FontStyle = [System.Windows.FontStyles]::Italic

            # Reset M365 Environment
            $syncHash.EnvironmentComboBox.SelectedIndex = 0

            # Uncheck all products
            foreach ($item in $syncHash | Where-Object { $_.GetType().Name -eq 'CheckBox' }) {
                $item.IsChecked = $false
            }

            # Reset Advanced Tab fields
            if ($syncHash.ApplicationSectionToggle) {
                $syncHash.ApplicationSectionToggle.IsChecked = $false
                $syncHash.ApplicationSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
                #$syncHash.AppIdTextBox.Text = "Your Application ID"
                #$syncHash.CertificateTextBox.Text = "Certificate Thumbprint"
            }

            if ($syncHash.OutputSectionToggle) {
                $syncHash.OutputSectionToggle.IsChecked = $false
                $syncHash.OutputSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
                $syncHash.OutPathTextBox.Text = "."
                $syncHash.OutFolderNameTextBox.Text = "M365BaselineConformance"
                $syncHash.OutJsonFileNameTextBox.Text = "ScubaResults"
                $syncHash.OutCsvFileNameTextBox.Text = "ScubaResults"
            }

            if ($syncHash.OpaSectionToggle) {
                $syncHash.OpaSectionToggle.IsChecked = $false
                $syncHash.OpaSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
                $syncHash.OpaPathTextBox.Text = "."
            }

            if ($syncHash.GeneralSectionToggle) {
                $syncHash.GeneralSectionToggle.IsChecked = $false
                $syncHash.GeneralSectionContent.Visibility = [System.Windows.Visibility]::Collapsed
                $syncHash.LogInCheckBox.IsChecked = $true
                $syncHash.DisconnectOnExitCheckBox.IsChecked = $false
            }
        }

        Function New-YamlPreview {
            # Get selected values
            $orgName = if ($syncHash.OrgNameTextBox.Text -ne $syncHash.OrgNamePlaceholder) { $syncHash.OrgNameTextBox.Text } else { "" }
            $orgUnit = if ($syncHash.OrgUnitTextBox.Text -ne $syncHash.OrgUnitPlaceholder) { $syncHash.OrgUnitTextBox.Text } else { "" }
            $description = if ($syncHash.DescriptionTextBox.Text -ne $syncHash.DescriptionPlaceholder) { $syncHash.DescriptionTextBox.Text } else { "" }

            $selectedEnv = $syncHash.UIConfigs.SupportedM365Environment | Where-Object { $_.displayName -eq $syncHash.EnvironmentComboBox.SelectedItem } | Select-Object -ExpandProperty id

            $selectedProducts = @()
            foreach ($item in $syncHash.ProductsGrid.Children) {
                if ($item -is [System.Windows.Controls.CheckBox] -and
                    $item.Name -like "*ProductCheckBox" -and
                    $item.IsChecked) {
                    $selectedProducts += $item.Tag
                }
            }

            # Generate YAML preview
            $yamlPreview = @()
            $yamlPreview += '# ScubaGear Configuration File'
            $yamlPreview += "`n# Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $yamlPreview += "`n`n# Organization Configuration"
            $yamlPreview += "`nOrganization: $($syncHash.OrganizationTextBox.Text)"
            If($orgName){$yamlPreview += "`nOrgName: $orgName"}
            If($orgUnit){$yamlPreview += "`nOrgUnit: $orgUnit"}
            If($description){$yamlPreview += @"
`nDescription: |
$description
"@
            }
            $yamlPreview += "`n`n# Configuration Details"
            $yamlPreview += @"
`nProductNames:
"@
            foreach ($product in $selectedProducts) {
                $yamlPreview += "`n  - $product"
            }

            $yamlPreview += "`n`nM365Environment: $selectedEnv"

            # Check Advanced Tab sections and add them if toggled on

            # Application Section
            if ($syncHash.ApplicationSectionToggle -and $syncHash.ApplicationSectionToggle.IsChecked) {
                $yamlPreview += "`n`n# Application Configuration"

                # Add AppId if it's not empty and not placeholder
                if (![string]::IsNullOrWhiteSpace($syncHash.AppIdTextBox.Text) -and
                    $syncHash.AppIdTextBox.Text -ne "Your Application ID") {
                    $yamlPreview += "`nAppId: $($syncHash.AppIdTextBox.Text)"
                }

                # Add Certificate if it's not empty and not placeholder
                if (![string]::IsNullOrWhiteSpace($syncHash.CertificateTextBox.Text) -and
                    $syncHash.CertificateTextBox.Text -ne "Certificate Thumbprint") {
                    $yamlPreview += "`nCertificateThumbprint: $($syncHash.CertificateTextBox.Text)"
                }
            }

            # Output Section
            if ($syncHash.OutputSectionToggle -and $syncHash.OutputSectionToggle.IsChecked) {
                $yamlPreview += "`n`n# Output Configuration"

                # Add OutPath (always include if section is toggled)
                $outPath = if (![string]::IsNullOrWhiteSpace($syncHash.OutPathTextBox.Text)) { $syncHash.OutPathTextBox.Text } else { "." }
                $yamlPreview += "`nOutPath: `"$outPath`""

                # Add OutFolderName (always include if section is toggled)
                $outFolderName = if (![string]::IsNullOrWhiteSpace($syncHash.OutFolderNameTextBox.Text)) { $syncHash.OutFolderNameTextBox.Text } else { "M365BaselineConformance" }
                $yamlPreview += "`nOutFolderName: `"$outFolderName`""

                # Add OutJsonFileName (always include if section is toggled)
                $outJsonFileName = if (![string]::IsNullOrWhiteSpace($syncHash.OutJsonFileNameTextBox.Text)) { $syncHash.OutJsonFileNameTextBox.Text } else { "ScubaResults" }
                $yamlPreview += "`nOutJsonFileName: `"$outJsonFileName`""

                # Add OutCsvFileName (always include if section is toggled)
                $outCsvFileName = if (![string]::IsNullOrWhiteSpace($syncHash.OutCsvFileNameTextBox.Text)) { $syncHash.OutCsvFileNameTextBox.Text } else { "ScubaResults" }
                $yamlPreview += "`nOutCsvFileName: `"$outCsvFileName`""
            }

            # OPA Section
            if ($syncHash.OpaSectionToggle -and $syncHash.OpaSectionToggle.IsChecked) {
                $yamlPreview += "`n`n# OPA Configuration"

                # Add OpaPath (always include if section is toggled)
                $opaPath = if (![string]::IsNullOrWhiteSpace($syncHash.OpaPathTextBox.Text)) { $syncHash.OpaPathTextBox.Text } else { "." }
                $yamlPreview += "`nOPAPath: `"$opaPath`""
            }

            # General Section
            if ($syncHash.GeneralSectionToggle -and $syncHash.GeneralSectionToggle.IsChecked) {
                $yamlPreview += "`n`n# General Configuration"

                # Add LogIn setting (always include if section is toggled)
                $logIn = if ($syncHash.LogInCheckBox.IsChecked) { "true" } else { "false" }
                $yamlPreview += "`nLogIn: $logIn"

                # Add DisconnectOnExit setting (always include if section is toggled)
                $disconnectOnExit = if ($syncHash.DisconnectOnExitCheckBox.IsChecked) { "true" } else { "false" }
                $yamlPreview += "`nDisconnectOnExit: $disconnectOnExit"
            }
            else {
                # If General section is not toggled, add default values
                $yamlPreview += "`nLogIn: true"
                $yamlPreview += "`nDisconnectOnExit: false"
            }

            # list exclusions psobject
            # members: Id, Product, TypeName, FieldName, FieldValue
            <#
            Aad: <-- Product
            # Legacy authentication SHALL be blocked.: <--Description
            MS.AAD.1.1v1: <-- Id
                CapExclusions: <-- TypeName
                Groups: <-- FieldName
                    - a90c9846-7c79-4d90-a38a-6d4781106481 <-- FieldValue
                Users: <-- FieldName
                    - a90c9846-7c79-4d90-a38a-6d4781106481 <-- FieldValue
            #>
            # Add this where you have the exclusions comment in New-YamlPreview
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

            # list omisisons psobject
            # members: Id,Product,Rationale,Expiration
            <#
            OmitPolicy: <-- OmitPolicy
            # Conditional access policies SHALL be enforced for all users.: <-- Description
            MS.AAD.2.1v1: <-- Id
                Rationale: I need this <-- Rationale
                Expiration: 2025-07-15 <-- Expiration
            #>
            If($syncHash.Omissions.Count -gt 0) {
                #group all omissions by product so they are in order
                $yamlPreview += "`n`nOmitPolicy:"
                #group all exclusions by product so they are in order
                $syncHash.Omissions | Group-Object Product  | ForEach-Object {
                    $product = $_.Name
                    $yamlPreview += "`n  #$product Omissions:"
                    foreach ($item in $_.Group | Sort-Object Id) {
                        $yamlPreview += "`n  $($item.Id):"
                        $yamlPreview += "`n    Rationale: $($item.Rationale)"
                        if ($item.Expiration) {
                            $yamlPreview += "`n    Expiration: $($item.Expiration)"
                        }
                    }
                }

            }

            # Display in preview tab
            $syncHash.YamlPreviewTextBox.Text = $yamlPreview

            foreach ($tab in $syncHash.MainTabControl.Items) {
                if ($tab -is [System.Windows.Controls.TabItem] -and $tab.Header -eq "Preview") {
                    $syncHash.MainTabControl.SelectedItem = $syncHash.PreviewTab
                    break
                }
            }
        }


        if ($syncHash.YAMLImport) {
            try {
                Write-Verbose "Loading YAML configuration from: $($syncHash.YAMLImport)"

                # Check if file exists
                if (Test-Path -Path $syncHash.YAMLImport -PathType Leaf) {
                    $yamlContent = Get-Content -Path $syncHash.YAMLImport -Raw
                    $yamlObject = $yamlContent | ConvertFrom-Yaml

                    # Import and populate all form fields
                    Import-YamlConfiguration -Config $yamlObject

                    Write-Verbose "YAML configuration loaded successfully"
                } else {
                    Write-Warning "YAML configuration file not found: $($syncHash.YAMLImport)"
                }
            }
            catch {
                Write-Warning "Error loading YAML configuration file: $($_.Exception.Message)"
                # Don't show message box on startup errors - just log the warning
                # The UI will still open with default values
            }
        }

        #Add smooth closing for Window
        $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
    	$syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-UIMainWindow})
    	$syncHash.Window.Add_Closed({ $syncHash.isClosed = $True })

        #always force windows on bottom
        $syncHash.Window.Topmost = $True

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

