# .SYNOPSIS
# Export-PublicFolderStatistics.ps1
#    Generates a CSV file that contains the list of public folders and their individual sizes
#
# .DESCRIPTION
#
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
    # File to export to
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Full path of the output file to be generated. If only filename is specified, then the output file will be generated in the current directory.")]
    [ValidateNotNull()]
    [string] $ExportFile,
    
    # Server to connect to for generating statistics
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Public folder server to enumerate the folder hierarchy.")]
    [ValidateNotNull()]
    [string] $PublicFolderServer
    )

#load hashtable of localized string
Import-LocalizedData -BindingVariable PublicFolderStatistics_LocalizedStrings -FileName Export-PublicFolderStatistics.strings.psd1
    
################ START OF DEFAULTS ################

$WarningPreference = 'SilentlyContinue';
$script:Exchange14MajorVersion = 14;
$script:Exchange12MajorVersion = 8;

################ END OF DEFAULTS #################

# Function that determines if to skip the given folder
function IsSkippableFolder()
{
    param($publicFolder);
    
    $publicFolderIdentity = $publicFolder.Identity.ToString();

    for ($index = 0; $index -lt $script:SkippedSubtree.length; $index++)
    {
        if ($publicFolderIdentity.StartsWith($script:SkippedSubtree[$index]))
        {
            return $true;
        }
    }
    
    return $false;
}

# Function that gathers information about different public folders
function GetPublicFolderDatabases()
{
    $script:ServerInfo = Get-ExchangeServer -Identity:$PublicFolderServer;
    $script:PublicFolderDatabasesInOrg = @();
    if ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange14MajorVersion)
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase -IncludePreExchange2010);
    }
    elseif ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange12MajorVersion)
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase -IncludePreExchange2007);
    }
    else
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase);
    }
}

# Function that executes statistics cmdlet on different public folder databases
function GatherStatistics()
{   
    # Running Get-PublicFolderStatistics against each server identified via Get-PublicFolderDatabase cmdlet
    $databaseCount = $($script:PublicFolderDatabasesInOrg.Count);
    $index = 0;
    
    if ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange12MajorVersion)
    {
        $getPublicFolderStatistics = "@(Get-PublicFolderStatistics ";
    }
    else
    {
        $getPublicFolderStatistics = "@(Get-PublicFolderStatistics -ResultSize:Unlimited ";
    }

    While ($index -lt $databaseCount)
    {
        $serverName = $($script:PublicFolderDatabasesInOrg[$index]).Server.Name;
        $getPublicFolderStatisticsCommand = $getPublicFolderStatistics + "-Server $serverName)";
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.RetrievingStatistics -f $serverName);
        $publicFolderStatistics = Invoke-Expression $getPublicFolderStatisticsCommand;
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.RetrievingStatisticsComplete -f $serverName,$($publicFolderStatistics.Count));
        RemoveDuplicatesFromFolderStatistics $publicFolderStatistics;
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.UniqueFoldersFound -f $($script:FolderStatistics.Count));
        $index++;
    }
}

# Function that removed redundant entries from output of Get-PublicFolderStatistics
function RemoveDuplicatesFromFolderStatistics()
{
    param($publicFolders);
    
    $index = 0;
    While ($index -lt $publicFolders.Count)
    {
        $publicFolderEntryId = $($publicFolders[$index].EntryId);
        $folderSizeFromStats = $($publicFolders[$index].TotalItemSize.Value.ToBytes());
        $folderPath = $script:IdToNameMap[$publicFolderEntryId];
        $existingFolder = $script:FolderStatistics[$publicFolderEntryId];
        if (($existingFolder -eq $null) -or ($folderSizeFromStats -gt $existingFolder[0]))
        {
            $newFolder = @();
            $newFolder += $folderSizeFromStats;
            $newFolder += $folderPath;
            $script:FolderStatistics[$publicFolderEntryId] = $newFolder;
        }
       
        $index++;
    }    
}

# Function that creates folder objects in right way for exporting
function CreateFolderObjects()
{   
    $index = 1;
    foreach ($publicFolderEntryId in $script:FolderStatistics.Keys)
    {
        $existingFolder = $script:NonIpmSubtreeFolders[$publicFolderEntryId];
        $publicFolderIdentity = "";
        if ($existingFolder -ne $null)
        {
            $result = IsSkippableFolder($existingFolder);
            if (!$result)
            {
                $publicFolderIdentity = "\NON_IPM_SUBTREE\" + $script:FolderStatistics[$publicFolderEntryId][1];
                $folderSize = $script:FolderStatistics[$publicFolderEntryId][0];
            }
        }  
        else
        {
            $publicFolderIdentity = "\IPM_SUBTREE" + $script:FolderStatistics[$publicFolderEntryId][1];
            $folderSize = $script:FolderStatistics[$publicFolderEntryId][0];
        }  
        
        if ($publicFolderIdentity -ne "")
        {
            if(($index%10000) -eq 0)
            {
                Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessedFolders -f $index);
            }
            
            # Create a folder object to be exported to a CSV
            $newFolderObject = New-Object PSObject -Property @{FolderName = $publicFolderIdentity; FolderSize = $folderSize}
            $retValue = $script:ExportFolders.Add($newFolderObject);
            $index++;
        }
    }   
}

####################################################################################################
# Script starts here
####################################################################################################

# Array of folder objects for exporting
$script:ExportFolders = $null;

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolderStatistics)
$script:FolderStatistics = @{};

# Hash table that contains the folder list (NON_IPM_SUBTREE via Get-PublicFolder)
$script:NonIpmSubtreeFolders = @{};

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolder)
$script:IpmSubtreeFolders = @{};

# Hash table EntryId to Name to map FolderPath
$script:IdToNameMap = @{};

# Recurse through IPM_SUBTREE to get the folder path foreach Public Folder
# Remarks:
# This is done so we can overcome a limitation of Get-PublicFolderStatistics
# where it fails to display Unicode chars in the FolderPath value, but 
# Get-PublicFolder properly renders these characters (as MapiIdentity)
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtree;
$ipmSubtreeFolderList = Get-PublicFolder "\" -Server $PublicFolderServer -Recurse -ResultSize:Unlimited;
$ipmSubtreeFolderList | %{ $script:IdToNameMap.Add($_.EntryId, $_.Identity.ToString()) };
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtreeComplete -f $($ipmSubtreeFolderList.Count));

# Folders that are skipped while computing statistics
$script:SkippedSubtree = @("\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK", "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
                           "\NON_IPM_SUBTREE\schema-root", "\NON_IPM_SUBTREE\OWAScratchPad",
                           "\NON_IPM_SUBTREE\StoreEvents", "\NON_IPM_SUBTREE\Events Root");

Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtree;
$nonIpmSubtreeFolderList = Get-PublicFolder "\NON_IPM_SUBTREE" -Server $PublicFolderServer -Recurse -ResultSize:Unlimited;
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtreeComplete -f $($nonIpmSubtreeFolderList.Count));
foreach ($nonIpmSubtreeFolder in $nonIpmSubtreeFolderList)
{
    $script:NonIpmSubtreeFolders.Add($nonIpmSubtreeFolder.EntryId, $nonIpmSubtreeFolder); 
}

# Determining the public folder database deployment in the organization
GetPublicFolderDatabases;

# Gathering statistics from each server
GatherStatistics;

# Allocating space here
$script:ExportFolders = New-Object System.Collections.ArrayList -ArgumentList ($script:FolderStatistics.Count + 3);

# Creating folder objects for exporting to a CSV
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ExportStatistics -f $($script:FolderStatistics.Count));
CreateFolderObjects;

# Creating folder objects for all the skipped root folders
$newFolderObject = New-Object PSObject -Property @{FolderName = "\IPM_SUBTREE"; FolderSize = 0};
# Ignore the return value
$retValue = $script:ExportFolders.Add($newFolderObject);
$newFolderObject = New-Object PSObject -Property @{FolderName = "\NON_IPM_SUBTREE"; FolderSize = 0};
$retValue = $script:ExportFolders.Add($newFolderObject);
$newFolderObject = New-Object PSObject -Property @{FolderName = "\NON_IPM_SUBTREE\EFORMS REGISTRY"; FolderSize = 0};
$retValue = $script:ExportFolders.Add($newFolderObject);

# Export the folders to CSV file
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ExportToCSV;
$script:ExportFolders | Sort-Object -Property FolderName | Export-CSV -Path $ExportFile -Force -NoTypeInformation -Encoding "Unicode";

# SIG # Begin signature block
# MIIdwwYJKoZIhvcNAQcCoIIdtDCCHbACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4fC/MhAZju1h7x8R0LS3cba2
# fpigghhlMIIEwzCCA6ugAwIBAgITMwAAAMZ4gDYBdRppcgAAAAAAxjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODUz
# WhcNMTgwOTA3MTc1ODUzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkY1MjgtMzc3Ny04QTc2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEArQsjG6jKiCgU
# NuPDaF0GhCh1QYcSqJypNAJgoa1GtgoNrKXTDUZF6K+eHPNzXv9v/LaYLZX2GyOI
# 9lGz55tXVv1Ny6I1ueVhy2cUAhdE+IkVR6AtCo8Ar8uHwEpkyTi+4Ywr6sOGM7Yr
# wBqw+SeaBjBwON+8E8SAz0pgmHHj4cNvt5A6R+IQC6tyiFx+JEMO1qqnITSI2qx3
# kOXhD3yTF4YjjRnTx3HGpfawUCyfWsxasAHHlILEAfsVAmXsbr4XAC2HBZGKXo03
# jAmfvmbgbm3V4KBK296Unnp92RZmwAEqL08n+lrl+PEd6w4E9mtFHhR9wGSW29C5
# /0bOar9zHwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFNS/9jKwiDEP5hmU8T6/Mfpb
# Ag8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAJhbANzvo0iL5FA5Z5QkwG+PvkDfOaYsTYksqFk+MgpqzPxc
# FwSYME/S/wyihd4lwgQ6CPdO5AGz3m5DZU7gPS5FcCl10k9pTxZ4s857Pu8ZrE2x
# rnUyUiQFl5DYSNroRPuQYRZZXs2xK1WVn1JcwcAwJwfu1kwnebPD90o1DRlNozHF
# 3NMaIo0nCTRAN86eSByKdYpDndgpVLSoN2wUnsh4bLcZqod4ozdkvgGS7N1Af18R
# EFSUBVraf7MoSxKeNIKLLyhgNxDxZxrUgnPb3zL73zOj40A1Ibw3WzJob8vYK+gB
# YWORl4jm6vCwAq/591z834HDNH60Ud0bH+xS7PowggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhEwggP5
# oAMCAQICEzMAAACOh5GkVxpfyj4AAAAAAI4wDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNjExMTcyMjA5MjFaFw0xODAy
# MTcyMjA5MjFaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDQh9RCK36d2cZ61KLD4xWS
# 0lOdlRfJUjb6VL+rEK/pyefMJlPDwnO/bdYA5QDc6WpnNDD2Fhe0AaWVfIu5pCzm
# izt59iMMeY/zUt9AARzCxgOd61nPc+nYcTmb8M4lWS3SyVsK737WMg5ddBIE7J4E
# U6ZrAmf4TVmLd+ArIeDvwKRFEs8DewPGOcPUItxVXHdC/5yy5VVnaLotdmp/ZlNH
# 1UcKzDjejXuXGX2C0Cb4pY7lofBeZBDk+esnxvLgCNAN8mfA2PIv+4naFfmuDz4A
# lwfRCz5w1HercnhBmAe4F8yisV/svfNQZ6PXlPDSi1WPU6aVk+ayZs/JN2jkY8fP
# AgMBAAGjggGAMIIBfDAfBgNVHSUEGDAWBgorBgEEAYI3TAgBBggrBgEFBQcDAzAd
# BgNVHQ4EFgQUq8jW7bIV0qqO8cztbDj3RUrQirswUgYDVR0RBEswSaRHMEUxDTAL
# BgNVBAsTBE1PUFIxNDAyBgNVBAUTKzIzMDAxMitiMDUwYzZlNy03NjQxLTQ0MWYt
# YmM0YS00MzQ4MWU0MTVkMDgwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1
# ApUwVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jcmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEF
# BQcBAQRVMFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9w
# a2lvcHMvY2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNV
# HRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBEiQKsaVPzxLa71IxgU+fKbKhJ
# aWa+pZpBmTrYndJXAlFq+r+bltumJn0JVujc7SV1eqVHUqgeSxZT8+4PmsMElSnB
# goSkVjH8oIqRlbW/Ws6pAR9kRqHmyvHXdHu/kghRXnwzAl5RO5vl2C5fAkwJnBpD
# 2nHt5Nnnotp0LBet5Qy1GPVUCdS+HHPNIHuk+sjb2Ns6rvqQxaO9lWWuRi1XKVjW
# kvBs2mPxjzOifjh2Xt3zNe2smjtigdBOGXxIfLALjzjMLbzVOWWplcED4pLJuavS
# Vwqq3FILLlYno+KYl1eOvKlZbiSSjoLiCXOC2TWDzJ9/0QSOiLjimoNYsNSa5jH6
# lEeOfabiTnnz2NNqMxZQcPFCu5gJ6f/MlVVbCL+SUqgIxPHo8f9A1/maNp39upCF
# 0lU+UK1GH+8lDLieOkgEY+94mKJdAw0C2Nwgq+ZWtd7vFmbD11WCHk+CeMmeVBoQ
# YLcXq0ATka6wGcGaM53uMnLNZcxPRpgtD1FgHnz7/tvoB3kH96EzOP4JmtuPe7Y6
# vYWGuMy8fQEwt3sdqV0bvcxNF/duRzPVQN9qyi5RuLW5z8ME0zvl4+kQjOunut6k
# LjNqKS8USuoewSI4NQWF78IEAA1rwdiWFEgVr35SsLhgxFK1SoK3hSoASSomgyda
# Qd691WZJvAuceHAJvDCCB3owggVioAMCAQICCmEOkNIAAAAAAAMwDQYJKoZIhvcN
# AQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAw
# BgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEx
# MB4XDTExMDcwODIwNTkwOVoXDTI2MDcwODIxMDkwOVowfjELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUg
# U2lnbmluZyBQQ0EgMjAxMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AKvw+nIQHC6t2G6qghBNNLrytlghn0IbKmvpWlCquAY4GgRJun/DDB7dN2vGEtgL
# 8DjCmQawyDnVARQxQtOJDXlkh36UYCRsr55JnOloXtLfm1OyCizDr9mpK656Ca/X
# llnKYBoF6WZ26DJSJhIv56sIUM+zRLdd2MQuA3WraPPLbfM6XKEW9Ea64DhkrG5k
# NXimoGMPLdNAk/jj3gcN1Vx5pUkp5w2+oBN3vpQ97/vjK1oQH01WKKJ6cuASOrdJ
# Xtjt7UORg9l7snuGG9k+sYxd6IlPhBryoS9Z5JA7La4zWMW3Pv4y07MDPbGyr5I4
# ftKdgCz1TlaRITUlwzluZH9TupwPrRkjhMv0ugOGjfdf8NBSv4yUh7zAIXQlXxgo
# tswnKDglmDlKNs98sZKuHCOnqWbsYR9q4ShJnV+I4iVd0yFLPlLEtVc/JAPw0Xpb
# L9Uj43BdD1FGd7P4AOG8rAKCX9vAFbO9G9RVS+c5oQ/pI0m8GLhEfEXkwcNyeuBy
# 5yTfv0aZxe/CHFfbg43sTUkwp6uO3+xbn6/83bBm4sGXgXvt1u1L50kppxMopqd9
# Z4DmimJ4X7IvhNdXnFy/dygo8e1twyiPLI9AN0/B4YVEicQJTMXUpUMvdJX3bvh4
# IFgsE11glZo+TzOE2rCIF96eTvSWsLxGoGyY0uDWiIwLAgMBAAGjggHtMIIB6TAQ
# BgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUSG5k5VAF04KqFzc3IrVtqMp1ApUw
# GQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB
# /wQFMAMBAf8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0f
# BFMwUTBPoE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJv
# ZHVjdHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcB
# AQRSMFAwTgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kv
# Y2VydHMvTWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDCBnwYDVR0gBIGX
# MIGUMIGRBgkrBgEEAYI3LgMwgYMwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvZG9jcy9wcmltYXJ5Y3BzLmh0bTBABggrBgEFBQcC
# AjA0HjIgHQBMAGUAZwBhAGwAXwBwAG8AbABpAGMAeQBfAHMAdABhAHQAZQBtAGUA
# bgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAZ/KGpZjgVHkaLtPYdGcimwuWEeFj
# kplCln3SeQyQwWVfLiw++MNy0W2D/r4/6ArKO79HqaPzadtjvyI1pZddZYSQfYtG
# UFXYDJJ80hpLHPM8QotS0LD9a+M+By4pm+Y9G6XUtR13lDni6WTJRD14eiPzE32m
# kHSDjfTLJgJGKsKKELukqQUMm+1o+mgulaAqPyprWEljHwlpblqYluSD9MCP80Yr
# 3vw70L01724lruWvJ+3Q3fMOr5kol5hNDj0L8giJ1h/DMhji8MUtzluetEk5CsYK
# wsatruWy2dsViFFFWDgycScaf7H0J/jeLDogaZiyWYlobm+nt3TDQAUGpgEqKD6C
# PxNNZgvAs0314Y9/HG8VfUWnduVAKmWjw11SYobDHWM2l4bf2vP48hahmifhzaWX
# 0O5dY0HjWwechz4GdwbRBrF1HxS+YWG18NzGGwS+30HHDiju3mUv7Jf2oVyW2ADW
# oUa9WfOXpQlLSBCZgB/QACnFsZulP0V3HjXG0qKin3p6IvpIlR+r+0cjgPWe+L9r
# t0uX4ut1eBrs6jeZeRhL/9azI2h15q/6/IvrC4DqaTuv/DDtBEyO3991bWORPdGd
# Vk5Pv4BXIqF4ETIheu9BCrE/+6jMpF3BoYibV3FWTkhFwELJm3ZbCoBIa/15n8G9
# bW1qyVJzEw16UM0xggTIMIIExAIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5n
# IFBDQSAyMDExAhMzAAAAjoeRpFcaX8o+AAAAAACOMAkGBSsOAwIaBQCggdwwGQYJ
# KoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQB
# gjcCARUwIwYJKoZIhvcNAQkEMRYEFFv8SgfDrIjAIeWZwkorKCX4hc0jMHwGCisG
# AQQBgjcCAQwxbjBsoESAQgBFAHgAcABvAHIAdAAtAFAAdQBiAGwAaQBjAEYAbwBs
# AGQAZQByAFMAdABhAHQAaQBzAHQAaQBjAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAD1fh1U0
# ftNnLfpSIZcBTusLq1CxkAHskc12iFosspvJS8wD28A3O89bnO0ErqwgPUkjFR/G
# rEJHZ3YfkrmsbYJ+xrvx7hoJ4SvUIp+KgjwO8dG5xcJb6mMcAGPN3NlydIqrLCu0
# mFLOZYgayj1kOKnCHoNSd5oRPpynlCwdAjWe+FqGfXD22+8kAOiawoYmYtUhlYAM
# 6bqlOcdbWx0qn66K2BwBcz71qMQAl0Tpu7P/Rpg9PKpg3veBlRKkLXLSTZwPB26g
# UmMUItdgT/rCJejiCalEmyWyp12OkBFURde+YiJXFtMP+IGHAHrrASlqCHUS5fPt
# UtEgh1sk1B8HKhWhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAxniANgF1GmlyAAAAAADGMAkGBSsO
# AwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEP
# Fw0xNzAzMzExNDQ5MThaMCMGCSqGSIb3DQEJBDEWBBSAzu2PKsBvChnweEthnJmW
# J3cn7zANBgkqhkiG9w0BAQUFAASCAQBY1/IU7e85LAPIf/RlXxZEJSBdoMuUB6dm
# +3JajiHIryjWJYP9yhqxWaC+KSUPHAplfpIWNbzGPemt3FKD/uvNVRcBM/vXDAAB
# x+Aa6k1JK+j3bCt5Pqt2j+aGfHdRmy+6f1iEyO1dsvWMzkMEsXrD5OnMSvaIXQxZ
# XLeVYt8yV0d4R5t1thfb2klZEJZw87tI24ukuo7N0ZZ5che9UtEdvSPlk+tEKvy4
# XzCGp3eORiu24Sx3SfUHawJOsgpbIT/8nlEj23owUmeG9ed+DzZkLLm7V+ILqNyd
# hNIlYmTpJy2+Xmfmy4PkgQABOupQFKdbfmjZUZtXetfhjsoXsORI
# SIG # End signature block
