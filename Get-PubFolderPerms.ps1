$pfs = Get-publicfolder -recurse
$perms = % ($_ in $pfs) {Get-PublicFolderclientpermission $_.publicfolder | select identity,user,@{name = "accessrights";expression={$_.accessrights}}}
$perms | Export-csv .\PFPerms.csv -NoTypeInformation