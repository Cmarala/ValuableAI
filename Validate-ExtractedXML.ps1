<#
.SYNOPSIS
    Validates SharePoint page extraction XML output
    
.DESCRIPTION
    This script validates the XML output from the SharePoint page extraction script
    against the defined schema and performs additional completeness checks.
    
.PARAMETER XmlFilePath
    Path to the extracted XML file to validate
    
.PARAMETER SchemaPath
    Path to the XSD schema file (optional, defaults to SharePointPageExtraction.xsd)
    
.PARAMETER DetailedReport
    Generate a detailed validation report
    
.EXAMPLE
    .\Validate-ExtractedXML.ps1 -XmlFilePath "PageData.xml"
    
.EXAMPLE
    .\Validate-ExtractedXML.ps1 -XmlFilePath "PageData.xml" -SchemaPath "SharePointPageExtraction.xsd" -DetailedReport
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$XmlFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$SchemaPath = "SharePointPageExtraction.xsd",
    
    [Parameter(Mandatory=$false)]
    [switch]$DetailedReport
)

function Write-ValidationLog {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        "Info" { "Cyan" }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-XmlWellFormed {
    param([string]$XmlPath)
    
    try {
        Write-ValidationLog "Checking XML well-formedness..." "Info"
        $xmlDoc = [xml](Get-Content $XmlPath -Raw)
        Write-ValidationLog "✓ XML is well-formed" "Success"
        return @{ IsValid = $true; Document = $xmlDoc; Error = $null }
    }
    catch {
        Write-ValidationLog "✗ XML is not well-formed: $($_.Exception.Message)" "Error"
        return @{ IsValid = $false; Document = $null; Error = $_.Exception.Message }
    }
}

function Test-XmlSchema {
    param(
        [xml]$XmlDocument,
        [string]$SchemaPath
    )
    
    if (-not (Test-Path $SchemaPath)) {
        Write-ValidationLog "Schema file not found: $SchemaPath" "Warning"
        return @{ IsValid = $false; Errors = @("Schema file not found") }
    }
    
    try {
        Write-ValidationLog "Validating against XML schema..." "Info"
        
        # Create schema set
        $schemaSet = New-Object System.Xml.Schema.XmlSchemaSet
        $schemaSet.Add("http://schemas.sharepoint.com/page-extraction/2016", $SchemaPath) | Out-Null
        
        # Add schema to document
        $XmlDocument.Schemas = $schemaSet
        
        # Collect validation errors
        $validationErrors = @()
        $validationHandler = {
            param($sender, $e)
            $validationErrors += $e.Message
        }
        
        # Validate
        $XmlDocument.Validate($validationHandler)
        
        if ($validationErrors.Count -eq 0) {
            Write-ValidationLog "✓ XML validates against schema" "Success"
            return @{ IsValid = $true; Errors = @() }
        }
        else {
            Write-ValidationLog "✗ XML schema validation failed with $($validationErrors.Count) errors" "Error"
            foreach ($error in $validationErrors) {
                Write-ValidationLog "  - $error" "Error"
            }
            return @{ IsValid = $false; Errors = $validationErrors }
        }
    }
    catch {
        Write-ValidationLog "✗ Schema validation error: $($_.Exception.Message)" "Error"
        return @{ IsValid = $false; Errors = @($_.Exception.Message) }
    }
}

function Test-ContentCompleteness {
    param([xml]$XmlDocument)
    
    Write-ValidationLog "Checking content completeness..." "Info"
    
    $issues = @()
    $warnings = @()
    
    try {
        $ns = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
        $ns.AddNamespace("sp", "http://schemas.sharepoint.com/page-extraction/2016")
        
        # Check required sections
        $requiredSections = @(
            "ExtractionInfo",
            "SiteInformation", 
            "PageMetadata",
            "PageContent",
            "WebPartZones",
            "AllWebParts",
            "PageStructure"
        )
        
        foreach ($section in $requiredSections) {
            $node = $XmlDocument.SelectSingleNode("//sp:$section", $ns)
            if (-not $node) {
                $issues += "Missing required section: $section"
            }
        }
        
        # Check page metadata completeness
        $pageMetadata = $XmlDocument.SelectSingleNode("//sp:PageMetadata", $ns)
        if ($pageMetadata) {
            $requiredMetadata = @("PageId", "PageName", "PageTitle", "PageType")
            foreach ($field in $requiredMetadata) {
                $fieldNode = $pageMetadata.SelectSingleNode("sp:$field", $ns)
                if (-not $fieldNode -or [string]::IsNullOrWhiteSpace($fieldNode.InnerText)) {
                    $issues += "Missing or empty required metadata field: $field"
                }
            }
        }
        
        # Check web parts
        $webPartsNode = $XmlDocument.SelectSingleNode("//sp:AllWebParts", $ns)
        if ($webPartsNode) {
            $webPartCount = [int]$webPartsNode.GetAttribute("count")
            $actualWebParts = $webPartsNode.SelectNodes("sp:WebPart", $ns)
            
            if ($webPartCount -ne $actualWebParts.Count) {
                $issues += "Web part count mismatch: declared $webPartCount, found $($actualWebParts.Count)"
            }
            
            # Check each web part for required fields
            foreach ($webPart in $actualWebParts) {
                $webPartId = $webPart.SelectSingleNode("sp:WebPartId", $ns)
                $webPartTitle = $webPart.SelectSingleNode("sp:Title", $ns)
                $webPartType = $webPart.SelectSingleNode("sp:TypeName", $ns)
                
                if (-not $webPartId -or [string]::IsNullOrWhiteSpace($webPartId.InnerText)) {
                    $issues += "Web part missing WebPartId"
                }
                if (-not $webPartTitle -or [string]::IsNullOrWhiteSpace($webPartTitle.InnerText)) {
                    $warnings += "Web part missing or empty Title"
                }
                if (-not $webPartType -or [string]::IsNullOrWhiteSpace($webPartType.InnerText)) {
                    $issues += "Web part missing TypeName"
                }
                
                # Check for content
                $content = $webPart.SelectSingleNode("sp:Content", $ns)
                if ($content) {
                    $hasContent = $false
                    foreach ($contentType in @("HtmlContent", "TextContent", "XmlContent", "RawContent")) {
                        $contentNode = $content.SelectSingleNode("sp:$contentType", $ns)
                        if ($contentNode -and -not [string]::IsNullOrWhiteSpace($contentNode.InnerText)) {
                            $hasContent = $true
                            break
                        }
                    }
                    if (-not $hasContent) {
                        $warnings += "Web part '$($webPartTitle.InnerText)' has no extracted content"
                    }
                }
            }
        }
        
        # Check zones
        $zonesNode = $XmlDocument.SelectSingleNode("//sp:WebPartZones", $ns)
        if ($zonesNode) {
            $zoneCount = [int]$zonesNode.GetAttribute("count")
            $actualZones = $zonesNode.SelectNodes("sp:Zone", $ns)
            
            if ($zoneCount -ne $actualZones.Count) {
                $issues += "Zone count mismatch: declared $zoneCount, found $($actualZones.Count)"
            }
        }
        
        # Check page structure
        $structureNode = $XmlDocument.SelectSingleNode("//sp:PageStructure", $ns)
        if ($structureNode) {
            $structureCount = [int]$structureNode.GetAttribute("count")
            $actualElements = $structureNode.SelectNodes("sp:StructureElement", $ns)
            
            if ($structureCount -ne $actualElements.Count) {
                $issues += "Page structure count mismatch: declared $structureCount, found $($actualElements.Count)"
            }
            
            # Check for proper ordering
            $orders = @()
            foreach ($element in $actualElements) {
                $orderNode = $element.SelectSingleNode("sp:Order", $ns)
                if ($orderNode) {
                    $orders += [int]$orderNode.InnerText
                }
            }
            
            $sortedOrders = $orders | Sort-Object
            if (($orders -join ",") -ne ($sortedOrders -join ",")) {
                $warnings += "Page structure elements are not in sequential order"
            }
        }
        
        # Check page content
        $pageContent = $XmlDocument.SelectSingleNode("//sp:PageContent", $ns)
        if ($pageContent) {
            $hasAnyContent = $false
            $contentTypes = @("AspxContent", "WikiContent", "PublishingContent", "RawPageContent")
            
            foreach ($contentType in $contentTypes) {
                $contentNode = $pageContent.SelectSingleNode("sp:$contentType", $ns)
                if ($contentNode -and -not [string]::IsNullOrWhiteSpace($contentNode.InnerText)) {
                    $hasAnyContent = $true
                    break
                }
            }
            
            if (-not $hasAnyContent) {
                $warnings += "No page content was extracted"
            }
        }
        
        # Report results
        if ($issues.Count -eq 0 -and $warnings.Count -eq 0) {
            Write-ValidationLog "✓ Content completeness check passed" "Success"
        }
        else {
            if ($issues.Count -gt 0) {
                Write-ValidationLog "✗ Content completeness issues found:" "Error"
                foreach ($issue in $issues) {
                    Write-ValidationLog "  - $issue" "Error"
                }
            }
            
            if ($warnings.Count -gt 0) {
                Write-ValidationLog "⚠ Content completeness warnings:" "Warning"
                foreach ($warning in $warnings) {
                    Write-ValidationLog "  - $warning" "Warning"
                }
            }
        }
        
        return @{
            HasIssues = ($issues.Count -gt 0)
            Issues = $issues
            Warnings = $warnings
        }
    }
    catch {
        Write-ValidationLog "Error during completeness check: $($_.Exception.Message)" "Error"
        return @{
            HasIssues = $true
            Issues = @("Completeness check failed: $($_.Exception.Message)")
            Warnings = @()
        }
    }
}

function Get-ExtractionSummary {
    param([xml]$XmlDocument)
    
    try {
        $ns = New-Object System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
        $ns.AddNamespace("sp", "http://schemas.sharepoint.com/page-extraction/2016")
        
        # Extraction info
        $extractionInfo = $XmlDocument.SelectSingleNode("//sp:ExtractionInfo", $ns)
        $scriptVersion = $extractionInfo.SelectSingleNode("sp:ScriptVersion", $ns)?.InnerText
        $extractionDate = $extractionInfo.SelectSingleNode("sp:Timestamp", $ns)?.InnerText
        
        # Page info
        $pageMetadata = $XmlDocument.SelectSingleNode("//sp:PageMetadata", $ns)
        $pageName = $pageMetadata.SelectSingleNode("sp:PageName", $ns)?.InnerText
        $pageTitle = $pageMetadata.SelectSingleNode("sp:PageTitle", $ns)?.InnerText
        $pageType = $pageMetadata.SelectSingleNode("sp:PageType", $ns)?.InnerText
        
        # Counts
        $webPartsNode = $XmlDocument.SelectSingleNode("//sp:AllWebParts", $ns)
        $webPartCount = if ($webPartsNode) { [int]$webPartsNode.GetAttribute("count") } else { 0 }
        
        $zonesNode = $XmlDocument.SelectSingleNode("//sp:WebPartZones", $ns)
        $zoneCount = if ($zonesNode) { [int]$zonesNode.GetAttribute("count") } else { 0 }
        
        $structureNode = $XmlDocument.SelectSingleNode("//sp:PageStructure", $ns)
        $structureCount = if ($structureNode) { [int]$structureNode.GetAttribute("count") } else { 0 }
        
        # Content analysis
        $pageContent = $XmlDocument.SelectSingleNode("//sp:PageContent", $ns)
        $hasWikiContent = $pageContent.SelectSingleNode("sp:WikiContent", $ns)?.InnerText.Length -gt 0
        $hasPublishingContent = $pageContent.SelectSingleNode("sp:PublishingContent", $ns)?.InnerText.Length -gt 0
        $hasRawContent = $pageContent.SelectSingleNode("sp:RawPageContent", $ns)?.InnerText.Length -gt 0
        
        return @{
            ScriptVersion = $scriptVersion
            ExtractionDate = $extractionDate
            PageName = $pageName
            PageTitle = $pageTitle
            PageType = $pageType
            WebPartCount = $webPartCount
            ZoneCount = $zoneCount
            StructureElementCount = $structureCount
            HasWikiContent = $hasWikiContent
            HasPublishingContent = $hasPublishingContent
            HasRawContent = $hasRawContent
        }
    }
    catch {
        Write-ValidationLog "Error generating summary: $($_.Exception.Message)" "Error"
        return @{}
    }
}

# Main validation execution
try {
    Write-ValidationLog "SharePoint Page Extraction XML Validator" "Info"
    Write-ValidationLog "=======================================" "Info"
    Write-ValidationLog "XML File: $XmlFilePath" "Info"
    Write-ValidationLog "Schema File: $SchemaPath" "Info"
    Write-ValidationLog "" "Info"
    
    # Check if files exist
    if (-not (Test-Path $XmlFilePath)) {
        Write-ValidationLog "XML file not found: $XmlFilePath" "Error"
        exit 1
    }
    
    # Test well-formedness
    $wellFormedResult = Test-XmlWellFormed -XmlPath $XmlFilePath
    if (-not $wellFormedResult.IsValid) {
        Write-ValidationLog "Validation failed - XML is not well-formed" "Error"
        exit 1
    }
    
    # Test schema validation
    $schemaResult = Test-XmlSchema -XmlDocument $wellFormedResult.Document -SchemaPath $SchemaPath
    
    # Test content completeness
    $completenessResult = Test-ContentCompleteness -XmlDocument $wellFormedResult.Document
    
    # Generate summary
    Write-ValidationLog "" "Info"
    Write-ValidationLog "EXTRACTION SUMMARY" "Info"
    Write-ValidationLog "==================" "Info"
    
    $summary = Get-ExtractionSummary -XmlDocument $wellFormedResult.Document
    
    if ($summary.Count -gt 0) {
        Write-ValidationLog "Script Version: $($summary.ScriptVersion)" "Info"
        Write-ValidationLog "Extraction Date: $($summary.ExtractionDate)" "Info"
        Write-ValidationLog "Page Name: $($summary.PageName)" "Info"
        Write-ValidationLog "Page Title: $($summary.PageTitle)" "Info"
        Write-ValidationLog "Page Type: $($summary.PageType)" "Info"
        Write-ValidationLog "Web Parts: $($summary.WebPartCount)" "Info"
        Write-ValidationLog "Zones: $($summary.ZoneCount)" "Info"
        Write-ValidationLog "Structure Elements: $($summary.StructureElementCount)" "Info"
        Write-ValidationLog "Has Wiki Content: $($summary.HasWikiContent)" "Info"
        Write-ValidationLog "Has Publishing Content: $($summary.HasPublishingContent)" "Info"
        Write-ValidationLog "Has Raw Content: $($summary.HasRawContent)" "Info"
    }
    
    # Final validation result
    Write-ValidationLog "" "Info"
    Write-ValidationLog "VALIDATION RESULTS" "Info"
    Write-ValidationLog "==================" "Info"
    
    $overallValid = $wellFormedResult.IsValid -and $schemaResult.IsValid -and (-not $completenessResult.HasIssues)
    
    if ($overallValid) {
        Write-ValidationLog "✓ VALIDATION PASSED - XML is valid and complete" "Success"
        if ($completenessResult.Warnings.Count -gt 0) {
            Write-ValidationLog "Note: $($completenessResult.Warnings.Count) warning(s) found (see above)" "Warning"
        }
        exit 0
    }
    else {
        Write-ValidationLog "✗ VALIDATION FAILED" "Error"
        
        if (-not $wellFormedResult.IsValid) {
            Write-ValidationLog "  - XML is not well-formed" "Error"
        }
        if (-not $schemaResult.IsValid) {
            Write-ValidationLog "  - XML does not validate against schema" "Error"
        }
        if ($completenessResult.HasIssues) {
            Write-ValidationLog "  - Content completeness issues found" "Error"
        }
        
        exit 1
    }
}
catch {
    Write-ValidationLog "Validation script error: $($_.Exception.Message)" "Error"
    exit 1
}