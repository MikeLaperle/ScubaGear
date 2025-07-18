using module '..\..\..\..\Modules\ScubaConfig\ScubaConfig.psm1'

InModuleScope ScubaConfig {
    Describe -tag "Utils" -name 'ScubaConfigLoadConfig' {
        BeforeAll {
            Mock -CommandName Write-Warning {}
            function Get-ScubaDefault {throw 'this will be mocked'}
            Mock -ModuleName ScubaConfig Get-ScubaDefault {"."}
	    Remove-Item function:\ConvertFrom-Yaml
        }
        context 'Handling repeated keys in YAML file' {
            It 'Load config with dupliacte keys'{
                # Load the first file and check the ProductNames value.

                {[ScubaConfig]::GetInstance().LoadConfig((Join-Path -Path $PSScriptRoot -ChildPath "./MockLoadConfig.yaml"))} | Should -Throw
            }
            AfterAll {
                [ScubaConfig]::ResetInstance()
            }
        }
        context 'Handling repeated LoadConfig invocations' {
            It 'Load valid config file followed by another'{
                $cfg = [ScubaConfig]::GetInstance()
                # Load the first file and check the ProductNames value.
                function global:ConvertFrom-Yaml {
                    @{
                        ProductNames=@('teams')
                    }
                }
                [ScubaConfig]::GetInstance().LoadConfig($PSCommandPath) | Should -BeTrue
                $cfg.Configuration.ProductNames | Should -Be 'teams'
                # Load the second file and verify that ProductNames has changed.
                function global:ConvertFrom-Yaml {
                    @{
                        ProductNames=@('exo')
                    }
                }
                [ScubaConfig]::GetInstance().LoadConfig($PSCommandPath) | Should -BeTrue
                $cfg.Configuration.ProductNames | Should -Be 'exo'
                Should -Invoke -CommandName Write-Warning -Exactly -Times 0
            }
            AfterAll {
                [ScubaConfig]::ResetInstance()
            }
        }
        context "Handling policy omissions" {
            It 'Does not warn for proper control IDs' {
                function global:ConvertFrom-Yaml {
                    @{
                        ProductNames=@('exo');
                        OmitPolicy=@{"MS.EXO.1.1v2"=@{"Rationale"="Example rationale"}}
                    }
                }
                [ScubaConfig]::GetInstance().LoadConfig($PSCommandPath) | Should -BeTrue
                Should -Invoke -CommandName Write-Warning -Exactly -Times 0
            }

            It 'Warns for malformed control IDs' {
                function global:ConvertFrom-Yaml {
                    @{
                        ProductNames=@('exo');
                        OmitPolicy=@{"MSEXO.1.1v2"=@{"Rationale"="Example rationale"}}
                    }
                }
                [ScubaConfig]::GetInstance().LoadConfig($PSCommandPath) | Should -BeTrue
                Should -Invoke -CommandName Write-Warning -Exactly -Times 1
            }

            It 'Warns for control IDs not encompassed by ProductNames' {
                function global:ConvertFrom-Yaml {
                    @{
                        ProductNames=@('exo');
                        OmitPolicy=@{"MS.Gmail.1.1v1"=@{"Rationale"="Example rationale"}}
                    }
                }
                [ScubaConfig]::GetInstance().LoadConfig($PSCommandPath) | Should -BeTrue
                Should -Invoke -CommandName Write-Warning -Exactly -Times 1
            }
	    AfterAll {
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
