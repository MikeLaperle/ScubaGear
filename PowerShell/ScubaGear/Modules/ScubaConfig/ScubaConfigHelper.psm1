function Update-BaselineConfig {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigFilePath,

        [Parameter(Mandatory=$false)]
        [string]$BaselineDirectory,

        [Parameter(Mandatory=$false)]
        [string]$GitHubDirectoryUrl,

        [Parameter(Mandatory=$false)]
        [hashtable]$ExclusionTypeMapping = @{
            # Default mappings - can be overridden
            "cap" = "cap"
            "role" = "role"
            "sensitiveAccounts" = "sensitiveAccounts"
            "sensitiveUsers" = "sensitiveUsers"
            "partnerDomains" = "partnerDomains"
            "forwardingDomains" = "forwardingDomains"
            "none" = "none"
        }
    )

    # Function to determine exclusion type based on policy content and ID
    function Get-ExclusionTypeFromPolicy {
        param(
            [string]$PolicyId,
            [string]$Title,
            [string]$Implementation,
            [string]$Rationale
        )

        # Define mapping rules based on policy patterns
        $mappingRules = @{
            # Conditional Access Policy related
            "legacy authentication|conditional access|high-risk|phishing-resistant|managed devices|device code" = "cap"
            
            # Role-based exclusions
            "highly privileged roles" = "role"
            
            # Sensitive account configurations
            "sensitive account|preset security policy" = "sensitiveAccounts"
            
            # Sensitive user configurations  
            "user impersonation|sensitive user" = "sensitiveUsers"
            
            # Partner domain configurations
            "partner|domain impersonation" = "partnerDomains"
            
            # Forwarding domain configurations
            "forwarding|spf policy" = "forwardingDomains"
        }

        $contentToCheck = "$Title $Implementation $Rationale".ToLower()
        
        foreach ($pattern in $mappingRules.Keys) {
            if ($contentToCheck -match $pattern) {
                return $mappingRules[$pattern]
            }
        }

        # Default to "none" if no pattern matches
        return "none"
    }

    # Load existing configuration
    if (-not (Test-Path $ConfigFilePath)) {
        throw "Configuration file not found at: $ConfigFilePath"
    }

    $configContent = Get-Content $ConfigFilePath -Raw | ConvertFrom-Json

    # Get baseline policies using the provided function
    $baselinePolicies = Get-BaselinePolicies -BaselineDirectory $BaselineDirectory -GitHubDirectoryUrl $GitHubDirectoryUrl

    # Create new baselines structure
    $newBaselines = @{}

    foreach ($product in $baselinePolicies.Keys) {
        $policies = $baselinePolicies[$product]
        $productBaseline = @()

        foreach ($policy in $policies) {
            # Determine exclusion type
            $exclusionType = Get-ExclusionTypeFromPolicy -PolicyId $policy.PolicyId -Title $policy.Title -Implementation $policy.Implementation -Rationale $policy.Rationale
            
            # Map to known exclusion types
            if ($ExclusionTypeMapping.ContainsKey($exclusionType)) {
                $exclusionType = $ExclusionTypeMapping[$exclusionType]
            }

            # Create policy object
            $policyObj = [PSCustomObject]@{
                id = $policy.PolicyId
                name = $policy.Title
                exclusionType = $exclusionType
                rationale = $policy.Rationale
            }

            # Add optional fields if they exist
            if ($policy.Criticality) {
                $policyObj | Add-Member -MemberType NoteProperty -Name "criticality" -Value $policy.Criticality
            }
            if ($policy.LastModified) {
                $policyObj | Add-Member -MemberType NoteProperty -Name "lastModified" -Value $policy.LastModified
            }
            if ($policy.Implementation) {
                $policyObj | Add-Member -MemberType NoteProperty -Name "implementation" -Value $policy.Implementation
            }
            if ($policy.MITRE_Mapping -and $policy.MITRE_Mapping.Count -gt 0) {
                $policyObj | Add-Member -MemberType NoteProperty -Name "mitreMapping" -Value $policy.MITRE_Mapping
            }
            if ($policy.Resources -and $policy.Resources.Count -gt 0) {
                $policyObj | Add-Member -MemberType NoteProperty -Name "resources" -Value $policy.Resources
            }

            $productBaseline += $policyObj
        }

        if ($productBaseline.Count -gt 0) {
            $newBaselines[$product] = $productBaseline
        }
    }

    # Update the configuration
    $configContent.baselines = $newBaselines

    # Save the updated configuration
    $configContent | ConvertTo-Json -Depth 10 | Set-Content $ConfigFilePath -Encoding UTF8

    Write-Host "Successfully updated baselines in configuration file: $ConfigFilePath"
    Write-Host "Updated products: $($newBaselines.Keys -join ', ')"
    
    # Summary statistics
    foreach ($product in $newBaselines.Keys) {
        $policyCount = $newBaselines[$product].Count
        $exclusionCounts = $newBaselines[$product] | Group-Object exclusionType | ForEach-Object { "$($_.Name): $($_.Count)" }
        Write-Host "  $product`: $policyCount policies ($($exclusionCounts -join ', '))"
    }
}

# Enhanced version with better exclusion type detection
function Update-BaselineConfigAdvanced {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigFilePath,

        [Parameter(Mandatory=$false)]
        [string]$BaselineDirectory,

        [Parameter(Mandatory=$false)]
        [string]$GitHubDirectoryUrl,

        [Parameter(Mandatory=$false)]
        [switch]$UseIntelligentMapping,

        [Parameter(Mandatory=$false)]
        [hashtable]$CustomExclusionRules = @{}
    )

    # Enhanced exclusion type detection
    function Get-SmartExclusionType {
        param(
            [string]$PolicyId,
            [string]$Title,
            [string]$Implementation,
            [string]$Rationale
        )

        # Combine all text for analysis
        $fullText = "$Title $Implementation $Rationale".ToLower()
        
        # Enhanced mapping rules with more specific patterns
        $smartRules = @{
            # Conditional Access Policies
            "cap" = @(
                "conditional access",
                "legacy authentication",
                "high-risk.*user",
                "high-risk.*sign-in",
                "phishing-resistant.*mfa",
                "managed device",
                "device code.*authentication",
                "mfa.*enforced",
                "authentication.*method"
            )
            
            # Role-based exclusions
            "role" = @(
                "role assignment",
                "privileged role",
                "global administrator",
                "pam system",
                "privileged access management",
                "just.*time",
                "activation.*role"
            )
            
            # Sensitive accounts
            "sensitiveAccounts" = @(
                "sensitive account",
                "preset security policy",
                "defender.*office.*365",
                "exchange online protection",
                "strict.*preset"
            )
            
            # Sensitive users
            "sensitiveUsers" = @(
                "user impersonation",
                "sensitive user",
                "impersonation protection"
            )
            
            # Partner domains
            "partnerDomains" = @(
                "partner",
                "domain impersonation",
                "important partner"
            )
            
            # Forwarding domains
            "forwardingDomains" = @(
                "forwarding.*domain",
                "spf.*policy",
                "automatic.*forwarding"
            )
        }

        # Check custom rules first
        foreach ($customType in $CustomExclusionRules.Keys) {
            $customPatterns = $CustomExclusionRules[$customType]
            foreach ($pattern in $customPatterns) {
                if ($fullText -match $pattern) {
                    return $customType
                }
            }
        }

        # Check smart rules
        foreach ($exclusionType in $smartRules.Keys) {
            $patterns = $smartRules[$exclusionType]
            foreach ($pattern in $patterns) {
                if ($fullText -match $pattern) {
                    return $exclusionType
                }
            }
        }

        # Policy ID based detection
        switch -Regex ($PolicyId) {
            "AAD\.3\." { return "cap" }      # MFA policies
            "AAD\.2\." { return "cap" }      # Risk-based policies  
            "AAD\.1\." { return "cap" }      # Legacy auth policies
            "AAD\.7\." { return "role" }     # Privileged role policies
            "DEFENDER\.1\." { return "sensitiveAccounts" }  # Defender preset policies
            "DEFENDER\.2\." { return "sensitiveUsers" }     # Impersonation policies
            "EXO\.2\." { return "forwardingDomains" }       # SPF policies
            default { return "none" }
        }
    }

    # Load existing configuration
    if (-not (Test-Path $ConfigFilePath)) {
        throw "Configuration file not found at: $ConfigFilePath"
    }

    $configContent = Get-Content $ConfigFilePath -Raw | ConvertFrom-Json

    # Get baseline policies
    $baselinePolicies = Get-BaselinePolicies -BaselineDirectory $BaselineDirectory -GitHubDirectoryUrl $GitHubDirectoryUrl

    # Create new baselines structure
    $newBaselines = @{}

    foreach ($product in $baselinePolicies.Keys) {
        $policies = $baselinePolicies[$product]
        $productBaseline = @()

        foreach ($policy in $policies) {
            # Determine exclusion type
            if ($UseIntelligentMapping) {
                $exclusionType = Get-SmartExclusionType -PolicyId $policy.PolicyId -Title $policy.Title -Implementation $policy.Implementation -Rationale $policy.Rationale
            } else {
                $exclusionType = "none"
            }

            # Create policy object with all available fields
            $policyObj = [ordered]@{
                id = $policy.PolicyId
                name = $policy.Title
                exclusionType = $exclusionType
                rationale = $policy.Rationale
            }

            # Add optional fields conditionally
            if ($policy.Criticality) {
                $policyObj.criticality = $policy.Criticality
            }
            if ($policy.LastModified) {
                $policyObj.lastModified = $policy.LastModified
            }
            if ($policy.Implementation) {
                $policyObj.implementation = $policy.Implementation
            }
            if ($policy.MITRE_Mapping -and $policy.MITRE_Mapping.Count -gt 0) {
                $policyObj.mitreMapping = $policy.MITRE_Mapping
            }
            if ($policy.Resources -and $policy.Resources.Count -gt 0) {
                $policyObj.resources = $policy.Resources
            }

            $productBaseline += $policyObj
        }

        if ($productBaseline.Count -gt 0) {
            $newBaselines[$product] = $productBaseline
        }
    }

    # Update the configuration
    $configContent.baselines = $newBaselines

    # Save the updated configuration with proper formatting
    $jsonOutput = $configContent | ConvertTo-Json -Depth 10
    
    # Format JSON nicely (optional - for better readability)
    $jsonOutput | Set-Content $ConfigFilePath -Encoding UTF8

    Write-Host "Successfully updated baselines in configuration file: $ConfigFilePath" -ForegroundColor Green
    Write-Host "Updated products: $($newBaselines.Keys -join ', ')" -ForegroundColor Yellow
    
    # Detailed summary
    foreach ($product in $newBaselines.Keys) {
        $policies = $newBaselines[$product]
        $policyCount = $policies.Count
        $exclusionCounts = $policies | Group-Object exclusionType | ForEach-Object { "$($_.Name): $($_.Count)" }
        Write-Host "  $product`: $policyCount policies ($($exclusionCounts -join ', '))" -ForegroundColor Cyan
    }

    return $newBaselines
}


function Get-BaselinePolicies {
    param(
        [Parameter(Mandatory=$false)]
        [string]$BaselineDirectory,

        [Parameter(Mandatory=$false)]
        [string]$GitHubDirectoryUrl
    )

    function Parse-PolicyContent {
        param([string]$Content)
        $result = @{
            Criticality = $null
            LastModified = $null
            Rationale = $null
            MITRE_Mapping = @()
            Resources = @()
        }
        if ($Content -match '<!--Policy:\s*[^;]+;\s*Criticality:\s*([A-Z]+)\s*-->') {
            $result.Criticality = $matches[1]
        }
        if ($Content -match '- _Last modified:_\s*(.+)') {
            $result.LastModified = $matches[1].Trim()
        }
        if ($Content -match '- _Rationale:_\s*(.+)') {
            $result.Rationale = $matches[1].Trim()
        }
        if ($Content -match '(_MITRE ATT&CK TTP Mapping:_[\s\S]+?)(\n\s*\n|###|$)') {
            $mitreBlock = $matches[1]
            $mitreList = @()
            foreach ($line in $mitreBlock -split "`n") {
                if ($line -match '\[([^\]]+)\]\(([^)]+)\)') {
                    $mitreList += $line.Trim()
                }
            }
            $result.MITRE_Mapping = $mitreList
        }
        if ($Content -match '(?ms)^### Resources\s*(.+?)(^###|\z)') {
            $resourcesBlock = $matches[1]
            $resources = @()
            foreach ($line in $resourcesBlock -split "`n") {
                if ($line -match '^\s*-\s*\[([^\]]+)\]\(([^)]+)\)') {
                    $resources += $line.Trim()
                }
            }
            $result.Resources = $resources
        }
        return $result
    }

    $policyHeaderPattern = '^####\s+([A-Z0-9\.]+v\d+)\s*$'
    $policiesByProduct = @{}

    $files = @()

    if ($GitHubDirectoryUrl) {
        # Convert GitHub URL to API URL
        if ($GitHubDirectoryUrl -match '^https://github.com/([^/]+)/([^/]+)/tree/([^/]+)(?:/(.*))?$') {
            $owner = $matches[1]
            $repo = $matches[2]
            $branch = $matches[3]
            $path = $matches[4]
            if ($null -ne $path -and $path -ne "") {
                $apiUrl = "https://api.github.com/repos/$owner/$repo/contents/$path`?ref=$branch"
            } else {
                $apiUrl = "https://api.github.com/repos/$owner/$repo/contents`?ref=$branch"
            }
        } else {
            throw "Invalid GitHub directory URL."
        }

        # Get list of markdown files in the directory
        $response = Invoke-RestMethod -Uri $apiUrl
        $files = $response | Where-Object { $_.name -like "*.md" }
    } elseif ($BaselineDirectory) {
        $files = Get-ChildItem -Path $BaselineDirectory -Filter *.md
    } else {
        throw "You must provide either -BaselineDirectory or -GitHubDirectoryUrl."
    }

    foreach ($file in $files) {
        if ($GitHubDirectoryUrl) {
            $rawUrl = $file.download_url
            $content = Invoke-WebRequest -Uri $rawUrl -UseBasicParsing | Select-Object -ExpandProperty Content
            $product = [System.IO.Path]::GetFileNameWithoutExtension($file.name)
            $lines = $content -split "`n"
        } else {
            $product = $file.BaseName
            $lines = Get-Content $file.FullName
        }

        $inPoliciesSection = $false
        $currentPolicy = $null
        $currentContent = @()
        $policies = @()
        $expectTitle = $false
        $implementationInstructions = @{}

        # First, find all Implementation instruction blocks for each policy
        $inImplementationSection = $false
        $currentImplementationPolicy = $null
        $currentImplementationContent = @()

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]

            # Check if we're in Implementation section
            if ($line -match '^### Implementation\s*$') {
                $inImplementationSection = $true
                continue
            }

            # Stop looking when we hit the next main section
            if ($inImplementationSection -and $line -match '^## ') {
                $inImplementationSection = $false
                if ($currentImplementationPolicy) {
                    $implementationInstructions[$currentImplementationPolicy] = ($currentImplementationContent -join "`n").Trim()
                    $currentImplementationPolicy = $null
                    $currentImplementationContent = @()
                }
                continue
            }

            # If we're in Implementation section, look for policy instruction headers
            if ($inImplementationSection) {
                if ($line -match '####\s+([A-Z0-9\.]+v\d+)\s+Instructions') {
                    if ($currentImplementationPolicy) {
                        $implementationInstructions[$currentImplementationPolicy] = ($currentImplementationContent -join "`n").Trim()
                        $currentImplementationContent = @()
                    }
                    $currentImplementationPolicy = $matches[1]
                    continue
                }
                if ($currentImplementationPolicy) {
                    $currentImplementationContent += $line
                }
            }
        }

        if ($inImplementationSection -and $currentImplementationPolicy) {
            $implementationInstructions[$currentImplementationPolicy] = ($currentImplementationContent -join "`n").Trim()
        }

        # Now parse the policies as before
        $inPoliciesSection = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ($line.Trim() -match '^### Policies') {
                $inPoliciesSection = $true
                continue
            }
            if ($inPoliciesSection -and ($line.Trim() -match '^## ' -or $line.Trim() -match '^# ')) {
                if ($currentPolicy) {
                    $currentPolicy += Parse-PolicyContent -Content ($currentContent -join "`n")
                    $policies += [PSCustomObject]$currentPolicy
                    $currentPolicy = $null
                    $currentContent = @()
                }
                $inPoliciesSection = $false
                $expectTitle = $false
                continue
            }
            if ($inPoliciesSection) {
                if ($line -match $policyHeaderPattern) {
                    if ($currentPolicy) {
                        $currentPolicy += Parse-PolicyContent -Content ($currentContent -join "`n")
                        $policies += [PSCustomObject]$currentPolicy
                        $currentContent = @()
                    }
                    $currentPolicy = @{
                        PolicyId = $matches[1]
                        Title    = ""
                        Implementation = ""
                    }
                    $expectTitle = $true
                    continue
                }
                elseif ($expectTitle -and $currentPolicy) {
                    if ($line.Trim() -ne "") {
                        $currentPolicy.Title = $line.Trim()
                        $expectTitle = $false
                    }
                }
                elseif ($currentPolicy) {
                    $currentContent += $line
                }
            }
        }

        if ($currentPolicy) {
            $currentPolicy += Parse-PolicyContent -Content ($currentContent -join "`n")
            $policies += [PSCustomObject]$currentPolicy
        }

        # Attach implementation instructions to the policies
        foreach ($policy in $policies) {
            if ($implementationInstructions.ContainsKey($policy.PolicyId)) {
                $policy.Implementation = $implementationInstructions[$policy.PolicyId]
            }
        }

        if ($policies.Count -gt 0) {
            $policiesByProduct[$product] = $policies
        }
    }

    return $policiesByProduct
}
# Usage examples:
<#
# Example usage of the Update-BaselineConfig function
$ScubaBaselines = Get-BaselinePolicies -GitHubDirectoryUrl "https://github.com/cisagov/ScubaGear/tree/main/PowerShell/ScubaGear/baselines"

# Example 1: Update from local baseline directory
Update-BaselineConfig -ConfigFilePath ".\ScubaConfig_en-US.json" -BaselineDirectory "C:\path\to\baselines"

# Example 2: Update from GitHub repository
Update-BaselineConfig -ConfigFilePath ".\ScubaConfig_en-US.json" -GitHubDirectoryUrl "https://github.com/cisagov/ScubaGear/tree/main/PowerShell/ScubaGear/baselines"

# Example 3: Use advanced version with intelligent mapping
Update-BaselineConfigAdvanced -ConfigFilePath ".\ScubaConfig_en-US.json" -GitHubDirectoryUrl "https://github.com/cisagov/ScubaGear/tree/main/PowerShell/ScubaGear/baselines" -UseIntelligentMapping

# Example 4: Use with custom exclusion rules
$customRules = @{
    "customType" = @("custom pattern 1", "custom pattern 2")
}
Update-BaselineConfigAdvanced -ConfigFilePath ".\ScubaConfig_en-US.json" -BaselineDirectory "C:\path\to\baselines" -UseIntelligentMapping -CustomExclusionRules $customRules
#>

