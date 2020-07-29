# .SYNOPSIS
# PublicFolderToMailboxMapGenerator.ps1
#    Generates a CSV file that contains the mapping of public folder branch to mailbox
#
# .DESCRIPTION
#
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
param(
    # Mailbox size 
    [Parameter(
	Mandatory=$true,
        HelpMessage = "Size (in Bytes) of any one of the Public folder mailboxes in destination. (E.g. For 1GB enter 1 followed by nine 0's)")]
    [long] $MailboxSize,

    # File to import from
    [Parameter(
        Mandatory=$true,
        HelpMessage = "This is the path to a CSV formatted file that contains the folder names and their sizes.")]
    [ValidateNotNull()]
    [string] $ImportFile,

    # File to export to
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Full path of the output file to be generated. If only filename is specified, then the output file will be generated in the current directory.")]
    [ValidateNotNull()]
    [string] $ExportFile
    )

# Folder Node's member indices
# This is an optimization since creating and storing objects as PSObject types
# is an expensive operation in powershell
# CLASSNAME_MEMBERNAME
$script:FOLDERNODE_PATH = 0;
$script:FOLDERNODE_MAILBOX = 1;
$script:FOLDERNODE_TOTALITEMSIZE = 2;
$script:FOLDERNODE_AGGREGATETOTALITEMSIZE = 3;
$script:FOLDERNODE_PARENT = 4;
$script:FOLDERNODE_CHILDREN = 5;
$script:MAILBOX_NAME = 0;
$script:MAILBOX_UNUSEDSIZE = 1;
$script:MAILBOX_ISINHERITED = 2;

$script:ROOT = @("`\", $null, 0, 0, $null, @{});

#load hashtable of localized string
Import-LocalizedData -BindingVariable MapGenerator_LocalizedStrings -FileName PublicFolderToMailboxMapGenerator.strings.psd1

# Function that constructs the entire tree based on the folderpath
# As and when it constructs it computes its aggregate folder size that included itself
function LoadFolderHierarchy() 
{
    foreach ($folder in $script:PublicFolders)
    {
        $folderSize = [long]$folder.FolderSize;
        if ($folderSize -gt $MailboxSize)
        {
            Write-Host "[$($(Get-Date).ToString())]" ($MapGenerator_LocalizedStrings.MammothFolder -f $folder, $folderSize, $MailboxSize);
            return $false;
        }

        # Start from root
        $parent = $script:ROOT;
        foreach ($familyMember in $folder.FolderName.Split('\', [System.StringSplitOptions]::RemoveEmptyEntries))
        {            
            # Try to locate the appropriate subfolder
            $child = $parent[$script:FOLDERNODE_CHILDREN].Item($familyMember);
            if ($child -eq $null)
            {
                # Create and add subfolder to parent's children
                $child = @($folder.FolderName, $null, $folderSize, $folderSize, $parent, @{});
                $parent[$script:FOLDERNODE_CHILDREN].Add($familyMember, $child);
            }

            # Add child's individual size to parent's aggregate size
            $parent[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] += $folderSize;
            $parent = $child;
        }
    }

    return $true;
}

# Function that assigns content mailboxes to public folders
# $node: Root node to be assigned to a mailbox
# $mailboxName: If not $null, we will attempt to accomodate folder in this mailbox
function AllocateMailbox()
{
    param ($node, $mailboxName)

    if ($mailboxName -ne $null)
    {
        # Since a mailbox was supplied by the caller, we should first attempt to use it
        if ($node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] -le $script:PublicFolderMailboxes[$mailboxName][$script:MAILBOX_UNUSEDSIZE])
        {
            # Node's contents (including branch) can be completely fit into specified mailbox
            # Assign the folder to mailbox and update mailbox's remaining size
            $node[$script:FOLDERNODE_MAILBOX] = $mailboxName;
            $script:PublicFolderMailboxes[$mailboxName][$script:MAILBOX_UNUSEDSIZE] -= $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE];
            if ($script:PublicFolderMailboxes[$mailboxName][$script:MAILBOX_ISINHERITED] -eq $false)
            {
                # This mailbox was not parent's content mailbox, but was created by a sibling
                $script:AssignedFolders += New-Object PSObject -Property @{FolderPath = $node[$script:FOLDERNODE_PATH]; TargetMailbox = $node[$script:FOLDERNODE_MAILBOX]};
            }

            return $mailboxName;
        }
    }

    $newMailboxName = "Mailbox" + ($script:NEXT_MAILBOX++);
    $script:PublicFolderMailboxes[$newMailboxName] = @($newMailboxName, $MailboxSize, $false);

    $node[$script:FOLDERNODE_MAILBOX] = $newMailboxName;
    $script:AssignedFolders += New-Object PSObject -Property @{FolderPath = $node[$script:FOLDERNODE_PATH]; TargetMailbox = $node[$script:FOLDERNODE_MAILBOX]};
    if ($node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE] -le $script:PublicFolderMailboxes[$newMailboxName][$script:MAILBOX_UNUSEDSIZE])
    {
        # Node's contents (including branch) can be completely fit into the newly created mailbox
        # Assign the folder to mailbox and update mailbox's remaining size
        $script:PublicFolderMailboxes[$newMailboxName][$script:MAILBOX_UNUSEDSIZE] -= $node[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE];
        return $newMailboxName;
    }
    else
    {
        # Since node's contents (including branch) could not be fitted into the newly created mailbox,
        # put it's individual contents into the mailbox
        $script:PublicFolderMailboxes[$newMailboxName][$script:MAILBOX_UNUSEDSIZE] -= $node[$script:FOLDERNODE_TOTALITEMSIZE];
    }

    $subFolders = @(@($node[$script:FOLDERNODE_CHILDREN].GetEnumerator()) | Sort @{Expression={$_.Value[$script:FOLDERNODE_AGGREGATETOTALITEMSIZE]}; Ascending=$true});
    $script:PublicFolderMailboxes[$newMailboxName][$script:MAILBOX_ISINHERITED] = $true;
    foreach ($subFolder in $subFolders)
    {
        $newMailboxName = AllocateMailbox $subFolder.Value $newMailboxName;
    }

    return $null;
}

# Function to check if further optimization can be done on the output generated
function TryAccomodateSubFoldersWithParent()
{
    $numAssignedFolders = $script:AssignedFolders.Count;
    for ($index = $numAssignedFolders - 1 ; $index -ge 0 ; $index--)
    {
        $assignedFolder = $script:AssignedFolders[$index];

        # Locate folder's parent
        for ($jindex = $index - 1 ; $jindex -ge 0 ; $jindex--)
        {
            if ($assignedFolder.FolderPath.StartsWith($script:AssignedFolders[$jindex].FolderPath))
            {
                # Found first ancestor
                $ancestor = $script:AssignedFolders[$jindex];
                $usedMailboxSize = $MailboxSize - $script:PublicFolderMailboxes[$assignedFolder.TargetMailbox][$script:MAILBOX_UNUSEDSIZE];
                if ($usedMailboxSize -le $script:PublicFolderMailboxes[$ancestor.TargetMailbox][$script:MAILBOX_UNUSEDSIZE])
                {
					# If the current mailbox can fit into its ancestor mailbox, add the former's contents to ancestor
					# and remove the mailbox assigned to it.Update the ancestor mailbox's size accordingly
                    $script:PublicFolderMailboxes[$assignedFolder.TargetMailbox][$script:MAILBOX_UNUSEDSIZE] = $MailboxSize;
                    $script:PublicFolderMailboxes[$ancestor.TargetMailbox][$script:MAILBOX_UNUSEDSIZE] -= $usedMailboxSize;
                    $assignedFolder.TargetMailbox = $null;
                }

                break;
            }
        }
    }
    
    if ($script:AssignedFolders.Count -gt 1)
    {
        $script:AssignedFolders = $script:AssignedFolders | where {$_.TargetMailbox -ne $null};
    }
}

# Parse the CSV file
Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.ProcessFolder;
$script:PublicFolders = Import-CSV $ImportFile;

# Check if there is atleast one public folder in existence
if (!$script:PublicFolders)
{
    Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.ProcessEmptyFile;
    return;
}

Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.LoadFolderHierarchy;
$loadHierarchy = LoadFolderHierarchy;
if ($loadHierarchy -ne $true)
{
    Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.CannotLoadFolders;
    return;
}

# Contains the list of instantiated public folder maiboxes
# Key: mailbox name, Value: unused mailbox size
$script:PublicFolderMailboxes = @{};
$script:AssignedFolders = @();
$script:NEXT_MAILBOX = 1;

Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.AllocateFolders;
$ignoreReturnValue = AllocateMailbox $script:ROOT $null;

Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.AccomodateFolders;
TryAccomodateSubFoldersWithParent;

Write-Host "[$($(Get-Date).ToString())]" $MapGenerator_LocalizedStrings.ExportFolderMap;
$script:NEXT_MAILBOX = 2;
$previous = $script:AssignedFolders[0];
$previousOriginalMailboxName = $script:AssignedFolders[0].TargetMailbox;
$numAssignedFolders = $script:AssignedFolders.Count;

# Prepare the folder object that is to be finally exported
# During the process, rename the mailbox assigned to it.  
# This is done to prevent any gap in generated mailbox name sequence at the end of the execution of TryAccomodateSubFoldersWithParent function
for ($index = 0 ; $index -lt $numAssignedFolders ; $index++)
{
    $current = $script:AssignedFolders[$index];
    $currentMailboxName = $current.TargetMailbox;
    if ($previousOriginalMailboxName -ne $currentMailboxName)
    {
        $current.TargetMailbox = "Mailbox" + ($script:NEXT_MAILBOX++);
    }
    else
    {
        $current.TargetMailbox = $previous.TargetMailbox;
    }

    $previous = $current;
    $previousOriginalMailboxName = $currentMailboxName;
}

# Export the folder mapping to CSV file
$script:AssignedFolders | Export-CSV -Path $ExportFile -Force -NoTypeInformation -Encoding "Unicode";

# SIG # Begin signature block
# MIIdzAYJKoZIhvcNAQcCoIIdvTCCHbkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0w+0nbMI2yy4o1qD8p9+TQu7
# pwygghhlMIIEwzCCA6ugAwIBAgITMwAAAMp9MhZ8fv0FAwAAAAAAyjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwOTA3MTc1ODU1
# WhcNMTgwOTA3MTc1ODU1WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAj3CeDl2ll7S4
# 96ityzOt4bkPI1FucwjpTvklJZLOYljFyIGs/LLi6HyH+Czg8Xd/oDQYFzmJTWac
# A0flGdvk8Yj5OLMEH4yPFFgQsZA5Wfnz/Cg5WYR2gmsFRUFELCyCbO58DvzOQQt1
# k/tsTJ5Ns5DfgCb5e31m95yiI44v23FVpKnTY9CUJbIr8j28O3biAhrvrVxI57GZ
# nzkUM8GPQ03o0NGCY1UEpe7UjY22XL2Uq816r0jnKtErcNqIgglXIurJF9QFJrvw
# uvMbRjeTBTCt5o12D4b7a7oFmQEDgg+koAY5TX+ZcLVksdgPNwbidprgEfPykXiG
# ATSQlFCEXwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGb30hxaE8ox6QInbJZnmt6n
# G7LKMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAGyg/1zQebvX564G4LsdYjFr9ptnqO4KaD0lnYBECEjMqdBM
# 4t+rNhN38qGgERoc+ns5QEGrrtcIW30dvMvtGaeQww5sFcAonUCOs3OHR05QII6R
# XYbxtAMyniTUPwacJiiCSeA06tLg1bebsrIY569mRQHSOgqzaO52EzJlOtdLrGDk
# Ot1/eu8E2zN9/xetZm16wLJVCJMb3MKosVFjFZ7OlClFTPk6rGyN9jfbKKDsDtNr
# jfAiZGVhxrEqMiYkj4S4OyvJ2uhw/ap7dbotTCfZu1yO57SU8rE06K6j8zWB5L9u
# DmtgcqXg3ckGvdmWVWBrcWgnmqNMYgX50XSzffQwggYHMIID76ADAgECAgphFmg0
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
# bW1qyVJzEw16UM0xggTRMIIEzQIBATCBlTB+MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5n
# IFBDQSAyMDExAhMzAAAAjoeRpFcaX8o+AAAAAACOMAkGBSsOAwIaBQCggeUwGQYJ
# KoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQB
# gjcCARUwIwYJKoZIhvcNAQkEMRYEFL2ZIcpj2id/Zmychk7lWo4odUJnMIGEBgor
# BgEEAYI3AgEMMXYwdKBMgEoAUAB1AGIAbABpAGMARgBvAGwAZABlAHIAVABvAE0A
# YQBpAGwAYgBvAHgATQBhAHAARwBlAG4AZQByAGEAdABvAHIALgBwAHMAMaEkgCJo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUA
# BIIBAHs0iUFvAUCHHTyx7mO2Ch74POLc4r+ahxqRvQG4M8iK3+4QXpZJphM4SbCb
# 1g3hlUwwYaRQoc/UcZiw78u4J4zxEG6W5tGL1TfCMVkN5EKBvuWZdG4xL9bD9SCE
# a9ENxr6mcPYTPbnseXGEKY3jWerzsNlml6MN6J57QltZz7Icu3f1VQwTmmPBuy0w
# LYUp+5iXUZtv7vfcqMuBh4lFCDI5/zfc5Ei68iJvuGWNCxU91QvI8s3sFyrC8c9A
# k5OJTqHoFXXfC5wH/3CBAyhrzJSkEvmlduvhosQJerzC0cj2ljrppa1u+hkcYza3
# uCN9rr/yS6cq4moMQHasTdeeEaShggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhEC
# AQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8G
# A1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAyn0yFnx+/QUDAAAA
# AADKMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0xNzAzMzExNDQ5MzZaMCMGCSqGSIb3DQEJBDEWBBRfG5VvxaWj
# lrOon4wdA9mEr+hcnTANBgkqhkiG9w0BAQUFAASCAQCNn5z9HJ8JPOB6YIK9lnDt
# DD6+V7cm5KX36nKTT/lUPas3HyFYhbV/LZXOM4dwtBhaUIgQQHyvCFy+sw+5nsd+
# MOzkj0nv1Fh07+T5lSDzYVYnOOLj2Y5Z2lo/oFtjxBt6ARyrPrED6dyfkWjqdC6u
# FfR/i6JqWqYZSE99+ZPZNgrboot+XZvnb7zTuw6asDte0woAWLp3C4O429qH31L+
# U9nVdSqTS2ViAzvEvMRpWGEET2tKmQGAhdD8gJh+AaSi6FS9z+ah2HJGAcQA/SNt
# OUAHpn6krnUrxREBKNwB0uZSXUJHazMHt6JcvLDXFS0048//DoSkZbcYYIX/jx4K
# SIG # End signature block
