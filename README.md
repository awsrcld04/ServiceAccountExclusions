
# ServiceAccountExclusions

DESCRIPTION: 
- Generate list of service accounts to exclude from processing

> NOTES: "v1.0" was completed in 2011. ServiceAccountExclusions was written to work in on-premises Active Directory environments. The purpose of ServiceAccountExclusions was/is to find service accounts to exclude from the running of scripts/tasks. ServiceAccountExclusions was designed to work with other tools.

## Requirements:

Operating System Requirements:
- Windows Server 2003 or higher (32-bit)
- Windows Server 2008 or higher (32-bit)

Additional software requirements:
Microsoft .NET Framework v3.5

Active Directory requirements:
One of following domain functional levels
- Windows Server 2003 domain functional level
- Windows Server 2008 domain functional level

Additional requirements:
Domain administrative access is required to perform operations by ServiceAccountExclusions


## Operation and Configuration:

Command-line parameters:
- run (Required parameter)

Configuration file: configServiceAccountExclusions.txt
- Located in the same directory as ServiceAccountExclusions.exe

Configuration file parameters:

ExclusionGroupLocation: Specifies an OU location in Active Directory to create a group called ServiceAccountExclusions to hold the exclusions create by the tool

Output:
- Located in the Log directory inside the installation directory; log files are in tab-delimited format
- Path example: (InstallationDirectory)\Log\

Additional detail:
- All servers will be scanned. Servers are pinged before the scan attempt takes place to ensure the servers are online.
- The ServiceAccountExclusions group will be recreated every 14 days
