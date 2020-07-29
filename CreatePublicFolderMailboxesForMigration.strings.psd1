ConvertFrom-StringData @'
###PSLOC
InvalidCsvEmptyMapping = Invalid CSV: folder mapping is empty.
InvalidNumberOfConcurrentUsers = Invalid argument for parameter 'EstimatedNumberOfConcurrentUsers': The argument {0} is less than the minimum allowed range of 1. Supply an argument that is greater than or equal to 1 and then try again.
InvalidCsvDuplicateMapping = Invalid CSV: duplicate mapping found for folder {0}.
InvalidCsvMissingRootFolder = Invalid CSV: missing root folder mapping.
DeploymentNotLockedForMigration = Existing Public Folder deployment is not locked for migration. The script cannot continue unless all Public Folder mailboxes are deleted first. Please, make sure the existing mailboxes have no data before deleting them.
PrimaryMailboxNameNotMatching = The primary Public Folder mailbox name '{0}' on the input CSV does not match the name of the existing primary mailbox '{1}'. Either change the name of the primary mailbox on the input CSV or delete existing primary mailbox. Please, make sure the existing mailboxes have no data before deleting them.
PublicFolderMailboxesAlreadyExistTitle = Public Folder mailboxes already exist.
PublicFolderMailboxesAlreadyExistMessage = Would you like to reuse existing Public Folder mailboxes as migration targets?
CreatingMailboxesForMigrationActivity = Creating target mailboxes for Public Folder Migration
UpdatingMailboxesActivity = Updating existing Public Folder mailboxes
CreatingMailboxesToServeHierarchyActivity = Creating additional Public Folder mailboxes to serve hierarchy
CreatingMailboxesProgressStatus = Total created: {0:N0}/{1:N0}. Creating mailbox '{2}'...
UpdatingMailboxesProgressStatus = Total updated: {0:N0}/{1:N0}. Updating mailbox '{2}'...
ConfirmationYesOption = &Yes
ConfirmationNoOption = &No
ConfirmTooManyUsersTitle = Too many concurrent users.
ConfirmTooManyUsersMessage = The estimated number of users connecting simultaneously is over the {0:N0} limit supported. The number Public Folder mailboxes serving hierarchy will default to the maximum supported.  Do you really want to proceed?
ConfirmMailboxOperationsTitle = Public Folder mailbox updates.
ConfirmMailboxOperationsMessage = Creating {0:N0} Public Folder mailbox(es) and updating {1:N0}. Total mailboxes to serve hierarchy will be {2:N0}. Would you like to proceed?
SkipPrimaryCreation = Skipping the creation of Primary Public Folder mailbox as it already exists.
FinalSummary = Total mailboxes created: {0:N0}. Total mailboxes updated: {1:N0}. Total serving hierarchy: {2:N0}.
MailboxesCreatedSummary = Here is a list of Public Folder mailboxes created:\r\n
###PSLOC
'@