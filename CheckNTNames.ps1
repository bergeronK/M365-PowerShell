# CheckNTNames.ps1
# PowerShell V2 script to check proposed pre-Windows 2000 names.
# To be used when the standard for sAMAccountNames is changed, and
# users are to be updated in bulk. This script ensures that the
# new values conform to the following:
# 1. The sAMAccountName values must be unique in the domain.
# 2. The value can be no more than 20 characters long.
# 3. The attributes on which the new value will be based must not be missing.
# 4. The following characters are not allowed: " [ ] : ; | = + * ? < > / \ ,
# 5. Values must not include a leading blank.
# Author: Richard L. Mueller
# Version 1.0 - February 20, 2015

Write-Host "Please Standby..."

# Load the PowerShell Active Directory Module.
Import-Module ActiveDirectory

# Flag to indicate whether objects will be updated.
# Change $Update to $True to have sAMAccountName values updated.
$Update = $False

# Setup the log file.
$LogFile = "CheckNTNames.log"
Add-Content -Path $LogFile -Value "------------------------------------------------" `
    -ErrorAction Stop
Add-Content -Path $LogFile -Value "CheckNTNames.ps1 Version 1.0 (February 20, 2015)"
Add-Content -Path $LogFile -Value $("Started: " + (Get-Date).ToString())
Add-Content -Path $LogFile -Value "Update flag: $Update"
Add-Content -Path $LogFile -Value "Log file: $LogFile"
Add-Content -Path $LogFile -Value "------------------------------------------------"

# Initialize counters.
$Changed = 0
$Missing = 0
$NotUnique = 0
$Script:Long = 0
$Script:Invalid = 0
$OK = 0
$Errors = 0
$Total = 0

# Flag to abort if there are too many errors attempting to update users.
$Abort = $False

# Function to remove invalid characters.
Function RemoveInvalid($NewNTName, $DN)
{
    [regex]$Reg = "(`"|:|;|=|<|>|/|,|\?|\[|\]|\||\+|\*|\\)"
    If ($Reg.Matches($NewNTName).Count -eq 0)
    {
        If ($NewNTName.Length -gt 0) {$NewNTName = $NewNTName.Trim()}
        Return $NewNTName
    }
    Else
    {
        Add-Content -Path $LogFile -Value $("## Invalid Characters Removed: "`
            + "Name $NewNTName")
        # Remove invalid characters and trim leading and trailing blanks.
        $NewNTName = $NewNTName.Replace("`"", "").Replace("[", "").Replace("]", "")
        $NewNTName = $NewNTName.Replace(":", "").Replace(";", "").Replace("|", "")
        $NewNTName = $NewNTName.Replace("=", "").Replace("+", "").Replace("*", "")
        $NewNTName = $NewNTName.Replace("?", "").Replace("<", "").Replace(">", "")
        $NewNTName = $NewNTName.Replace("\", "").Replace("/", "")
        If ($NewNTName.Length -gt 0) {$NewNTName = $NewNTName.Trim()}
        Add-Content -Path $LogFile -Value $("    Revised Name: $NewNTName")
        Add-Content -Path $LogFile -Value $("    DN: $DN")
        $Script:Invalid = $Script:Invalid + 1
        Return $NewNTName
    }
}

# Function to truncate names longer than 20 characters.
Function Truncate($NewNTName, $DN)
{
    If ($NewNTName.Length -le 20) {Return $NewNTName}
    Else
    {
        Add-Content -Path $LogFile -Value $("## Truncated: sAMAccountName " `
            + $NewNTName + " is too long")
        # Truncate to first 20 characters.
        $NewNTName = $NewNTName.SubString(0, 20).Trim()
        Add-Content -Path $LogFile -Value $("    Revised sAMAccountName: $NewNTName")
        Add-Content -Path $LogFile -Value $("    DN: $DN")
        $Script:Long = $Script:Long + 1
        Return $NewNTName
    }
}

# Function to determine new sAMAccountName value based on first and last names,
# in form "f.last" where "f" is the first initial of the first name and
# "last" is the last name.
Function Get-Name1($FirstName, $LastName, $DN)
{
    # Remove invalid characters from all values and trim leading and trailing spaces.
    $FirstName = RemoveInvalid $FirstName $DN
    $LastName = RemoveInvalid $LastName $DN
    # Check that neither parameter is missing.
    If (($FirstName) -and ($LastName))
    {
        # Determine proposed value of sAMAccountName.
        # First letter of first name, following by ".", followed by last name.
        # Make the value all lower case.
        Return $FirstName.ToLower() + "." `
            + $LastName.ToLower()
    }
    Else {Return "#=Problem"}
}

# Function to determine new sAMAccountName value based on first and last names and
# the middle initial, in form "fmlast", where "f" is the first initial of the first name,
# "m" is the middle initial, and "last" is the last name. If there is a first and last
# name, but not a middle initial, the sAMAccountName will be "flast".
Function Get-Name2($FirstName, $Middle, $LastName, $DN)
{
    # Remove invalid characters from all values and trim leading and trailing spaces.
    $FirstName = RemoveInvalid $FirstName $DN
    $LastName = RemoveInvalid $LastName $DN
    $Middle = RemoveInvalid $Middle $DN
    # Check that neither first nor last name are missing.
    # It is acceptable for the middle initial to be missing.
    If (($FirstName) -and ($LastName))
    {
        If ($Middle)
        {
            # Determine proposed value of sAMAccountName.
            # First letter of first name, following by the middle initial,
            # followed by last name.
            # Make the value all lower case.
            Return $FirstName.SubString(0, 1).ToLower() `
                + $Middle.SubString(0, 1).ToLower() + $LastName.ToLower()
        }
        Else
        {
            # No middle initial.
            # Make the value all lower case.
            Return $FirstName.SubString(0, 1).ToLower() + $LastName.ToLower()
        }
    }
    Else {Return "#=Problem"}
}

# Function to pad with blanks and right justify formatted integer values.
Function RightJustify($Value, $Size)
{
    # Format integer value for readability and pad with 10 blanks on the left.
    $Padded = "          $('{0:n0}' -f $Value)"
    # Right justify as much as needed to accomodate largest value.
    Return $Padded.SubString($Padded.Length - $Size, $Size)
}

# Function to add counter totals to the log file.
Function Add-Totals
{
    # Maximum integer length for right justifying the output.
    $TotalSize = $('{0:n0}' -f $Total).Length

    Add-Content -Path $LogFile -Value "------------------------------------------------"
    Add-Content -Path $LogFile -Value $("Finished: " + (Get-Date).ToString())
    Add-Content -Path $LogFile `
        -Value "New Name Truncated:                  $(RightJustify $Script:Long $TotalSize)"
    Add-Content -Path $LogFile `
        -Value "Invalid Characters Removed:          $(RightJustify $Script:Invalid $TotalSize)"
    Add-Content -Path $LogFile -Value "                                    ------------"
    If ($Update -eq $True)
    {
        Add-Content -Path $LogFile `
            -Value "Users Renamed:                       $(RightJustify $Changed $TotalSize)"
    }
    Else
    {
        Add-Content -Path $LogFile `
            -Value "Users Can be Renamed:                $(RightJustify $Changed $TotalSize)"
    }
    Add-Content -Path $LogFile `
        -Value "Users no Change Needed (Skipped):    $(RightJustify $OK $TotalSize)"
    Add-Content -Path $LogFile `
        -Value "Users with Missing Values (Skipped): $(RightJustify $Missing $TotalSize)"
    Add-Content -Path $LogFile `
        -Value "New Name not Unique (Skipped):       $(RightJustify $NotUnique $TotalSize)"
    If ($Update -eq $True)
    {
        Add-Content -Path $LogFile `
            -Value "Number of Errors Updating:           $(RightJustify $Errors $TotalSize)"
    }
    Add-Content -Path $LogFile -Value "                                    ------------"
    Add-Content -Path $LogFile `
        -Value "Total Number of Users Processed:     $('{0:n0}' -f $Total)"
}

# Hash table of sAMAccountNames to check for uniqueness.
$Names = @{}

# Array of objects to be checked for uniqueness a second time.
$Dups = @()

# Retrieve all objects in the domain with sAMAccountName values.
$Objects = Get-ADObject -LDAPFilter "(sAMAccountName=*)" -Properties sAMAccountName

# Populate the hash table. Key is sAMAccountName, value is distinguishedName.
ForEach ($Object In $Objects)
{
    $Names.Add($Object.sAMAccountName, $Object.distinguishedName)
}

# Retrieve all user objects in the domain that are to have new sAMAccountName values.
# If instead you only want to consider users in a specified organizational unit,
# add the -SearchBase parameter.
# For example: -SearchBase "ou=Sales,ou=West,dc=MyDomain,dc=com"
$Users = Get-ADUser -Filter * `
    -Properties GivenName, Initials, Surname, sAMAccountName, distinguishedName

# Check each user.
ForEach ($User In $Users)
{
    $Total = $Total + 1
    # Determine the new sAMAccountName, using Function Get-Name1. This function
    # should remove invalid characters and trim leading or trailing spaces from values.
    $NewName = Get-Name1 $User.GivenName $User.Surname $User.distinguishedName
    # If you use Get-Name2, use the following instead:
    # $NewName = Get-Name2 `
    #     $User.GivenName $User.Initials $User.Surname $User.distinguishedName
    # Check for problem (missing values).
    If ($NewName -ne "#=Problem")
    {
        # Truncate, if necessary, to 20 characters.
        $NewName = Truncate $NewName $User.distinguishedName
        # Check if user sAMAccountName already in correct format.
        $OldName = $User.sAMAccountName
        If ($OldName.ToLower() -ne $NewName.ToLower())
        {
            # Check that new name is unique in the domain.
            If ($Names.ContainsKey("$NewName") -eq $False)
            {
                # New NT name is OK.
                # Only update if the flag is set.
                If ($Update -eq $True)
                {
                    # Catch any possible errors, such as lacking permissions.
                    Try
                    {
                        Set-ADUser -Identity $User.distinguishedName `
                            -Replace @{sAMAccountName=$NewName.ToString()}
                        # Add the new name to the hash table.
                        $Names.Add($NewName, $User.distinguishedName)
                        # Remove the old name from the hash table.
                        $Names.Remove($OldName)
                        Add-Content -Path $LogFile `
                            -Value $("Renamed: From $OldName to $NewName")    
                        Add-Content -Path $LogFile -Value $("    DN: " `
                            + $User.distinguishedName)
                        $Changed = $Changed + 1
                    }
                    Catch
                    {
                        Add-Content -Path $LogFile -Value $("## Error: " `
                            + " failed to assign sAMAccountName $NewName")
                        Add-Content -Path $LogFile `
                            -Value $("    DN: " + $User.distinguishedName)
                        Add-Content -Path $LogFile -Value "    Error Message: $_"
                        $Errors = $Errors + 1
                        # Allow only 10 errors before aborting the script.
                        If ($Errors -gt 10)
                        {
                            $Abort = $True
                            Add-Content -Path $LogFile `
                                -Value "## More Than 10 Errors"
                            Add-Content -Path $LogFile `
                                -Value "    Script Aborted"
                            Write-Host "More Than 10 Errors Encountered!!"
                            Write-Host "Script Aborted."
                            # Break out of the ForEach loop.
                            Break
                        }
                    }
                }
                Else
                {
                    # User could be updated, if update flag set.
                    # Add the new name to the hash table.
                    $Names.Add($NewName, $User.distinguishedName)
                    Add-Content -Path $LogFile -Value $("Can be Renamed:" `
                        + " From $OldName to $NewName")
                    Add-Content -Path $LogFile -Value $("    DN: " + $User.distinguishedName)
                    $Changed = $Changed + 1
                } # End If $Update -eq $True.
            }
            Else
            {
                # New sAMAccountName not unique in the domain.
                Add-Content -Path $LogFile -Value $("## Not Unique: " `
                    + "sAMAccountName $NewName not unique in domain")
                Add-Content -Path $LogFile -Value $("    DN: " + $user.distinguishedName)
                Add-Content -Path $LogFile -Value $("    Conflicts with: " `
                    + $Names[$NewName])
                # The value will be checked again later.
                $Dups = $Dups + $User.distinguishedName
                $NotUnique = $NotUnique + 1
            } # End If $NewName unique.
        }
        Else
        {
            # sAMAccountName already in desired format.
            Add-Content -Path $LogFile -Value $("## OK: " `
                + "sAMAccountName $NewName already in correct format")
            Add-Content -Path $LogFile -Value $("    DN: " + $user.distinguishedName)
            $OK = $OK + 1
        } # End $NewName already correct.
    }
    Else
    {
        # Problem determining new sAMAccountName.
        Add-Content -Path $LogFile -Value $("## Missing Values: Attributes missing")
        Add-Content -Path $LogFile -Value $("    DN: " + $user.distinguishedName)
        $Missing = $Missing + 1
    } # End missing attributes.
} # End ForEach $User.

# Check any duplicates found to make sure they are still duplicates.
# This check only helps if $Update is set to $True, and we are not
# aborting due to too many errors.
If (($Dups.Count -gt 0) -and ($Update -eq $True) -and ($Abort -eq $False))
{
    Add-Content -Path $LogFile -Value "------------------------------------------------"
    Add-Content -Path $LogFile `
        -Value "Duplicate values for sAMAccountName will be checked again"
    Add-Content -Path $LogFile -Value "------------------------------------------------"

    # Update hash table of sAMAccountNames to check for uniqueness.
    $Names = @{}

    # Retrieve all objects in the domain with sAMAccountName values. This now
    # includes all new values, and does not include old values that were updated.
    $Objects = Get-ADObject -LDAPFilter "(sAMAccountName=*)" -Properties sAMAccountName

    #Populate the hash table. Key is sAMAccountName, value is distinguishedName.
    ForEach ($Object In $Objects)
    {
        $Names.Add($Object.sAMAccountName, $Object.distinguishedName)
    }
    # Consider each user previously found to be in conflict with another object.
    ForEach ($UserDN In $Dups)
    {
        # Retrieve the user attributes, as required by function Get-Name1 or Get-Name2.
        $User = Get-ADUser -Identity $UserDN `
            -Properties GivenName, Initials, Surname, sAMAccountName, distinguishedName
        # Determine the new sAMAccountName, using the function Get-Name1.
        $NewName = Get-Name1 $User.GivenName $User.Surname $User.distinguishedName
        # If you use Get-Name2, use the following instead:
        # $NewName = Get-Name2 $User.GivenName $User.Initials $User.Surname `
        #     $User.distinguishedName
        # Truncate, if necessary, to 20 characters.
        $NewName = Truncate $NewName $User.distinguishedName
        # In case of conflicts, consider variations of sAMAccountName values.
        If ($NewName.Length -le 19)
        {
            $NewName2 = "$NewName`2"
            $NewName3 = "$NewName`3"
            $NewName4 = "$NewName`4"
        }
        Else
        {
            # The length cannot exceed 20 characters.
            $NewName2 = $NewName.SubString(0, 19).Trim() + "2"
            $NewName3 = $NewName.SubString(0, 19).Trim() + "3"
            $NewName4 = $NewName.SubString(0, 19).Trim() + "4"
        }
        # Check that new name is unique in the domain.
        # Other checks have already been performed.
        # Check all variations of sAMAccountName.
        If ($Names.ContainsKey("$NewName") -eq $False) {$NTName = $NewName}
        Else
        {
            If ($Names.ContainsKey("$NewName2") -eq $False) {$NTName = $NewName2}
            Else
            {
                If ($Names.ContainsKey("$NewName3") -eq $False) {$NTName = $NewName3}
                Else
                {
                    If ($Names.ContainsKey("$NewName4") -eq $False) {$NTName = $NewName4}
                    Else {$NTName = "#=Conflict"}
                }
            }
        }
        If ($NTName -ne "#=Conflict")
        {
            # New NT name is OK.
            # The previous problem will either be resolved, or an error
            # will be raised, so reduce the $NotUnique count by 1.
            $NotUnique = $NotUnique - 1
            # Catch any possible errors, such as lacking permissions.
            Try
            {
                $OldName = $User.sAMAccountName
                Set-ADUser -Identity $User.distinguishedName `
                    -Replace @{sAMAccountName=$NTName.ToString()}
                # Add the new name to the hash table.
                $Names.Add($NTName, $User.distinguishedName)
                # Remove old name from the hash table.
                $Names.Remove($OldName)
                Add-Content -Path $LogFile `
                    -Value $("Renamed (Second Attempt): From $OldName to $NTName")    
                Add-Content -Path $LogFile -Value $("    DN: " `
                    + $User.distinguishedName)
                # The previous problem has been resolved.
                $Changed = $Changed + 1
            }
            Catch
            {
                Add-Content -Path $LogFile -Value $("## Error: " `
                    + " failed to assign sAMAccountName $NTName")
                Add-Content -Path $LogFile `
                    -Value $("    DN: " + $User.distinguishedName)
                Add-Content -Path $LogFile -Value "    Error Message: $_"
                $Errors = $Errors + 1
                # Allow only 10 errors before aborting the script.
                If ($Errors -gt 10)
                {
                    Add-Content -Path $LogFile `
                        -Value "## More Than 10 Errors"
                    Add-Content -Path $LogFile `
                        -Value "    Script Aborted"
                    Write-Host "More Than 10 Errors Encountered!!"
                    Write-Host "Script Aborted."
                    # Break out of the ForEach loop.
                    Break
                }
            }
        }
        Else
        {
            # New sAMAccountName, and variations, are not unique in the domain.
            Add-Content -Path $LogFile -Value $("## Not Unique (Second Attempt): " `
                + " sAMAccountName $NewName not unique in domain")
            Add-Content -Path $LogFile -Value $("    DN: " + $User.distinguishedName)
            Add-Content -Path $LogFile -Value "    Conflicts with users:"
            Add-Content -Path $LogFile -Value $("    $Names[$NewName]")
            Add-Content -Path $LogFile -Value $("    $Names[$NewName2]")
            Add-Content -Path $LogFile -Value $("    $Names[$NewName3]")
            Add-Content -Path $LogFile -Value $("    $Names[$NewName4]")
        } # End attempt to assign new name.
    } # End ForEach $UserDN to be reconsidered.
} # End duplicate sAMAccountNames to be reconsidered.

# Add totals to the log file.
Add-Totals

Write-Host "Done. See log file: $LogFile"
