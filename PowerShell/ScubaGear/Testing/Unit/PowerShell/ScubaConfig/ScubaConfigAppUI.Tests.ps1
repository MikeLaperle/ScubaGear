using module '..\..\..\..\Modules\ScubaConfig\ScubaConfigAppUI.psm1'

InModuleScope ScubaConfigAppUI {

    Describe -tag "Config" -name 'ScubaConfig JSON Configuration Validation' {
        BeforeAll {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'configPath')]
            $configPath = "$PSScriptRoot\..\..\..\..\Modules\ScubaConfig\ScubaConfig_en-US.json"
        }

        Context 'JSON File Structure Validation' {
            It 'Should have a valid JSON configuration file' {
                Test-Path $configPath | Should -BeTrue -Because "Configuration file should exist at expected location"

                { $script:configContent = Get-Content $configPath -Raw | ConvertFrom-Json } | Should -Not -Throw -Because "JSON should be valid and parseable"
                $script:configContent | Should -Not -BeNullOrEmpty
            }

            It 'Should contain all required root keys' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json

                # Define expected root keys based on your configuration structure
                $expectedRootKeys = @(
                    'DebugMode',
                    'Version',
                    'localeContext',
                    'localePlaceholder',
                    'localeInfoMessages',
                    'localeVerboseMessages',
                    'localeErrorMessages',
                    'localeDebugOutput',
                    'localePopupMessages',
                    'defaultAdvancedSettings',
                    'products',
                    'M365Environment',
                    'baselineControls',
                    'baselines',
                    'inputTypes',
                    'valueValidations',
                    'graphQueries'
                )

                foreach ($key in $expectedRootKeys) {
                    $configContent.PSObject.Properties.Name | Should -Contain $key -Because "Root key '$key' should be present in configuration"
                }
            }

            It 'Should have valid version format' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.Version | Should -Match '^\d+\.\d+\.\d+$' -Because "Version should follow semantic versioning format (x.y.z)"
            }

            It 'Should have valid DebugMode values' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $validDebugModes = @('None', 'Info', 'Verbose', 'Debug')
                $configContent.DebugMode | Should -BeIn $validDebugModes -Because "DebugMode should be one of the valid options"
            }
        }

        Context 'Products Configuration Validation' {
            It 'Should have products array with required properties' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.products | Should -Not -BeNullOrEmpty

                # Ensure products is treated as an array (handle single item case)
                $productsArray = @($configContent.products)
                $productsArray.Count | Should -BeGreaterThan 0 -Because "Should have at least one product"

                foreach ($product in $productsArray) {
                    $product.PSObject.Properties.Name | Should -Contain 'id' -Because "Each product should have an 'id' property"
                    $product.PSObject.Properties.Name | Should -Contain 'name' -Because "Each product should have a 'name' property"
                    $product.PSObject.Properties.Name | Should -Contain 'displayName' -Because "Each product should have a 'displayName' property"
                    $product.PSObject.Properties.Name | Should -Contain 'supportsExclusions' -Because "Each product should have a 'supportsExclusions' property"

                    $product.id | Should -Not -BeNullOrEmpty
                    $product.supportsExclusions | Should -BeOfType [System.Boolean]
                }
            }
        }

        Context 'M365Environment Configuration Validation' {
            It 'Should have M365Environment array with required properties' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.M365Environment | Should -Not -BeNullOrEmpty

                # Ensure M365Environment is treated as an array (handle single item case)
                $environmentsArray = @($configContent.M365Environment)
                $environmentsArray.Count | Should -BeGreaterThan 0 -Because "Should have at least one environment"

                foreach ($env in $environmentsArray) {
                    $env.PSObject.Properties.Name | Should -Contain 'id' -Because "Each environment should have an 'id' property"
                    $env.PSObject.Properties.Name | Should -Contain 'name' -Because "Each environment should have a 'name' property"
                    $env.PSObject.Properties.Name | Should -Contain 'displayName' -Because "Each environment should have a 'displayName' property"

                    $env.id | Should -Not -BeNullOrEmpty
                    $env.name | Should -Not -BeNullOrEmpty
                }
            }
        }

        Context 'BaselineControls Configuration Validation' {
            It 'Should have baselineControls array with required properties' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.baselineControls | Should -Not -BeNullOrEmpty

                # Ensure baselineControls is treated as an array (handle single item case)
                $controlsArray = @($configContent.baselineControls)
                $controlsArray.Count | Should -BeGreaterThan 0 -Because "Should have at least one baseline control"

                $requiredProperties = @('tabName', 'yamlValue', 'dataControlOutput', 'fieldControlName', 'defaultFields', 'cardName', 'showFieldType', 'showDescription', 'supportsAllProducts')

                foreach ($control in $controlsArray) {
                    foreach ($property in $requiredProperties) {
                        $control.PSObject.Properties.Name | Should -Contain $property -Because "Each baseline control should have a '$property' property"
                    }
                }
            }
        }

        Context 'Locale Messages Validation' {
            It 'Should have non-empty locale message sections' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json

                $localeMessageSections = @('localeContext', 'localePlaceholder', 'localeInfoMessages', 'localeVerboseMessages', 'localeErrorMessages', 'localeDebugOutput', 'localePopupMessages')

                foreach ($section in $localeMessageSections) {
                    $configContent.$section | Should -Not -BeNullOrEmpty -Because "Locale section '$section' should not be empty"
                    $configContent.$section.PSObject.Properties.Count | Should -BeGreaterThan 0 -Because "Locale section '$section' should contain message definitions"
                }
            }
        }

        Context 'ValueValidations Configuration Validation' {
            It 'Should have valueValidations with pattern and sample properties' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.valueValidations | Should -Not -BeNullOrEmpty

                foreach ($validation in $configContent.valueValidations.PSObject.Properties) {
                    $validationObj = $validation.Value
                    $validationObj.PSObject.Properties.Name | Should -Contain 'pattern' -Because "Validation '$($validation.Name)' should have a 'pattern' property"
                    $validationObj.PSObject.Properties.Name | Should -Contain 'sample' -Because "Validation '$($validation.Name)' should have a 'sample' property"
                    $validationObj.PSObject.Properties.Name | Should -Contain 'format' -Because "Validation '$($validation.Name)' should have a 'format' property"

                    $validationObj.pattern | Should -Not -BeNullOrEmpty
                    $validationObj.sample | Should -Not -BeNullOrEmpty
                }
            }
        }

        Context 'GraphQueries Configuration Validation' {
            It 'Should have graphQueries with required endpoint properties' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.graphQueries | Should -Not -BeNullOrEmpty

                foreach ($query in $configContent.graphQueries.PSObject.Properties) {
                    $queryObj = $query.Value
                    $queryObj.PSObject.Properties.Name | Should -Contain 'name' -Because "Graph query '$($query.Name)' should have a 'name' property"
                    $queryObj.PSObject.Properties.Name | Should -Contain 'endpoint' -Because "Graph query '$($query.Name)' should have an 'endpoint' property"
                    $queryObj.PSObject.Properties.Name | Should -Contain 'outProperty' -Because "Graph query '$($query.Name)' should have an 'outProperty' property"
                    $queryObj.PSObject.Properties.Name | Should -Contain 'tipProperty' -Because "Graph query '$($query.Name)' should have a 'tipProperty' property"

                    $queryObj.endpoint | Should -Match '^/v1\.0/' -Because "Graph endpoint should start with '/v1.0/'"
                }
            }
        }

        Context 'Baselines Configuration Validation' {
            It 'Should have baselines for each product' {
                $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
                $configContent.baselines | Should -Not -BeNullOrEmpty

                # Verify baselines exist for each product
                $productsArray = @($configContent.products)
                foreach ($product in $productsArray) {
                    $configContent.baselines.PSObject.Properties.Name | Should -Contain $product.id -Because "Baselines should exist for product '$($product.id)'"

                    $productBaselines = $configContent.baselines.($product.id)
                    $productBaselines | Should -Not -BeNullOrEmpty -Because "Product '$($product.id)' should have baseline policies"

                    # Ensure baselines is treated as an array (handle single item case)
                    $baselinesArray = @($productBaselines)
                    $baselinesArray.Count | Should -BeGreaterThan 0 -Because "Product '$($product.id)' should have at least one baseline policy"
                }
            }
        }
    }

    Describe -tag "UI" -name 'ScubaConfigAppUI XAML Validation' {
        BeforeAll {
            # Mock the UI launch function to prevent actual UI from showing
            Mock -CommandName Invoke-SCuBAConfigAppUI { return $true }

            # Helper function to test XAML parsing without UI launch
            function Test-XamlValidity {
                param([string]$XamlPath)

                try {
                    # Load assemblies needed for XAML parsing
                    [System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework') | Out-Null
                    [System.Reflection.Assembly]::LoadWithPartialName('PresentationCore') | Out-Null

                    # Read and process XAML the same way as the main function
                    [string]$XAML = (Get-Content $XamlPath -ReadCount 0) -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window' -replace 'Click=".*','/>'
                    [xml]$UIXML = $XAML
                    $reader = New-Object System.Xml.XmlNodeReader ([xml]$UIXML)

                    # Try to load the XAML - this will throw if invalid
                    $window = [Windows.Markup.XamlReader]::Load($reader)

                    return @{
                        IsValid = $true
                        Window = $window
                        Error = $null
                    }
                }
                catch {
                    return @{
                        IsValid = $false
                        Window = $null
                        Error = $_.Exception.Message
                    }
                }
            }
        }

        Context 'XAML File Validation' {
            It 'Should have a valid XAML file' {
                $xamlPath = "$PSScriptRoot\..\..\..\..\Modules\ScubaConfig\ScubaConfigAppUI.xaml"
                Test-Path $xamlPath | Should -BeTrue

                $result = Test-XamlValidity -XamlPath $xamlPath
                $result.IsValid | Should -BeTrue -Because "XAML should be valid: $($result.Error)"
            }

            It 'Should contain required UI elements' {
                $xamlPath = "$PSScriptRoot\..\..\..\..\Modules\ScubaConfig\ScubaConfigAppUI.xaml"
                $result = Test-XamlValidity -XamlPath $xamlPath

                $result.IsValid | Should -BeTrue
                $result.Window | Should -Not -BeNullOrEmpty

                # Test for specific named elements
                $result.Window.FindName("M365Environment_ComboBox") | Should -Not -BeNullOrEmpty
                $result.Window.FindName("ProductsGrid") | Should -Not -BeNullOrEmpty
                $result.Window.FindName("Organization_TextBox") | Should -Not -BeNullOrEmpty
            }
        }

        Context 'Mocked UI Function' {
            It 'Should not launch actual UI when mocked' {
                # This should return true without launching UI
                Invoke-SCuBAConfigAppUI | Should -BeTrue

                # Verify the mock was called
                Should -Invoke -CommandName Invoke-SCuBAConfigAppUI -Exactly -Times 1
            }
        }
    }
}