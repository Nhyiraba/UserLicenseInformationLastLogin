# UserLicenseInformationLastLogin

This project practically get user last Sign-in information for office 365 objects Enables you generate a csv file the last usage for the all the object in the anviroment

You can also check for specific type of objects,

DiscoveryMailbox, EquipmentMailbox, GroupMailbox, RoomMailbox, SchedulingMailbox, SharedMailbox, TeamMailbox and UserMailbox

## Report types for the parameter ReportTypeSelection 

AssignedLicenses
LastSignInDetail
AdminRoles
AllReportDetail (default)

# How to use

Run powershell

## Example 1 : Only one mailbox type
.\GetUserLicenseInformationLastLogin.ps1 -RecipientTypes UserMailbox

## Example 1 : Only one mailbox type
.\GetUserLicenseInformationLastLogin.ps1 -RecipientTypes UserMailbox, SharedMailbox

## Example 1 : Getting multiple mailbox types different report type
.\GetUserLicenseInformationLastLogin.ps1 -RecipientTypes RoomMailbox, UserMailbox,SharedMailbox -ReportTypeSelection LastSignInDetail

## Report Destination folder, 

### if the is not specified the current working directory where script is located will be used
.\GetUserLicenseInformationLastLogin.ps1 -RecipientTypes RoomMailbox, UserMailbox,SharedMailbox -ReportTypeSelection LastSignInDetail -ReportDestinationFolderPath "Enter Your Preferre Path"
