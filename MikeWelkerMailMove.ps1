#
# MoveItems.ps1
#
# By David Barrett, Microsoft Ltd. 2012. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

# Parameters

param (
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,
	
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Source folder (from which to move messages)")]
	[string]$SourceFolder,
	
	[Parameter(Position=2,Mandatory=$True,HelpMessage="Target folder (messages will be moved here)")]
	[string]$TargetFolder,
	
	[bool]$ProcessSubfolders = $false,
	[bool]$DeleteSourceFolder = $false,
	[string]$Username,
	[string]$Password,
	[string]$Domain,
	[bool]$Impersonate,
	[string]$EwsUrl,
	[bool]$IgnoreSSLCertificate,
	[string]$EWSManagedApiPath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	[bool]$LogVerbose = $false
);

# Define our functions

Function ShowParams()
{
	Write-Host "MoveItems -Mailbox <string>";
	Write-Host "		  -SourceFolder <string>";
	Write-Host "		  -TargetFolder <string>";
	Write-Host "				   [-DeleteSourceFolder <bool>]";
	Write-Host "                   [-Username <string> -Password <string> [-Domain <string>]]";
	Write-Host "                   [-Impersonate <bool>]";
	Write-Host "                   [-EwsUrl <string>]";
	Write-Host "                   [-IgnoreSSLCertificate <bool>]";
	Write-Host "                   [-EWSManagedApiPath <string>]";
	Write-Host "";
	Write-Host "Required:";
	Write-Host " -Mailbox : Mailbox SMTP email address";
	Write-Host " -SourceFolder : Full path to source folder in the mailbox (items are moved from this folder)";
	Write-Host " -TargetFolder : Full path to target folder in the mailbox (items are moved to this folder)";
	Write-Host "";
	Write-Host "Optional:";
	Write-Host " -ProcessSubfolders : If true, subfolders of the source folder will also be processed (default is false)";
	Write-Host " -DeleteSourceFolder : If true, the source folder will be deleted once items moved (so long as it is empty)";
	Write-Host " -Username : Username for the account being used to connect to EWS (if not specified, current user is assumed)";
	Write-Host " -Password : Password for the specified user (required if username specified)";
	Write-Host " -Domain : If specified, used for authentication (not required even if username specified)";
	Write-Host " -Impersonate : Set to $true to use impersonation.";
	Write-Host " -EwsUrl : Forces a particular EWS URl (otherwise autodiscover is used, which is recommended)";
	Write-Host " -IgnoreSSLCertificate : If $true, then any SSL errors will be ignored";
	Write-Host " -EWSManagedApiDLLFilePath : Full and path to the DLL for EWS Managed API (if not specified, default path for v1.2 is used)";
	Write-Host " -LogVerbose: Show detailed output";
	Write-Host "";
}

Function LoadEWSManagedAPI()
{
	# Find and load the managed API
	
	if ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		if ( { Test-Path $EWSManagedApiPath } )
		{
			Add-Type -Path $EWSManagedApiPath
			return $true
		}
		Write-Host ( [string]::Format("Managed API not found at specified location: {0}", $EWSManagedApiPath) ) -ForegroundColor Yellow
	}
	
	$a = Get-ChildItem -Recurse "C:\Program Files (x86)\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	if (!$a)
	{
		$a = Get-ChildItem -Recurse "C:\Program Files\Microsoft\Exchange\Web Services" -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and ( $_.Name -eq "Microsoft.Exchange.WebServices.dll" ) }
	}
	
	if ($a)	
	{
		# Load EWS Managed API
		Write-Host ([string]::Format("Using managed API {0} found at: {1}", $a.VersionInfo.FileVersion, $a.VersionInfo.FileName)) -ForegroundColor Gray
		Add-Type -Path $a.VersionInfo.FileName
		return $true
	}
	return $false
}

Function TrustAllCerts() {
    <#
    .SYNOPSIS
    Set certificate trust policy to trust self-signed certificates (for test servers).
    #>

    ## Code From http://poshcode.org/624
    ## Create a compilation environment
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

    $TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            { 
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@ 
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

Function MoveItems()
{
	# Process all the items in the given source folder, and move them to the target
	
	if ($args -eq $null)
	{
		throw "No folders specified for MoveItems";
	}
	$SourceFolderObject, $TargetFolderObject, $SourceFolderPath, $TargetFolderPath = $args[0];
	$SourceFolderPath = $SourceFolderPath + '\' + $SourceFolderObject.DisplayName;
	$TargetFolderPath = $TargetFolderPath + '\' + $TargetFolderObject.DisplayName;
	
	if ($SourceFolderObject.Id -eq $TargetFolderObject.Id)
	{
		Write-Host "Cannot copy from/to the same folder" -foregroundcolor Red;
		return;
	}
	
	Write-Host "Moving from", $SourceFolderPath "to", $TargetFolderPath -foregroundcolor White;
	
	# Set parameters - we will process in batches of 500 for the FindItems call
	$Offset=0;
	$PageSize=500;
	$MoreItems=$true;
	
	while ($MoreItems)
	{
		$View = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize, $Offset, [Microsoft.Exchange.Webservices.Data.OffsetBasePoint]::Beginning);
		$View.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly);
		$FindResults=$SourceFolderObject.FindItems($View);
		
		ForEach ($Item in $FindResults.Items)
		{
			if ($LogVerbose) { Write-Host "Processing", $Item.Id.UniqueId -foregroundcolor Gray; }
			try
			{
				$Item.Move($TargetFolderObject.Id) | out-null;
			}
			catch
			{
				Write-Host "Failed to move item", $Item.Id.UniqueId -foregroundcolor Red
			}
		}
		$MoreItems=$FindResults.MoreAvailable;
		$Offset+=$PageSize;
	}

	# Now process any subfolders
	if ($ProcessSubFolders)
	{
		if ($SourceFolderObject.ChildFolderCount -gt 0)
		{
			# Deal with any subfolders first
			$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000);
			$SourceFindFolderResults = $SourceFolderObject.FindFolders($FolderView);
			ForEach ($SourceSubFolderObject in $SourceFindFolderResults.Folders)
			{
				$Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SourceSubFolderObject.DisplayName);
				$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(2);
				$FindFolderResults = $TargetFolderObject.FindFolders($Filter, $FolderView);
				if ($FindFolderResults.TotalCount -eq 0)
				{
					$TargetSubFolderObject = New-Object Microsoft.Exchange.WebServices.Data.Folder($service);
					$TargetSubFolderObject.DisplayName = $SourceSubFolderObject.DisplayName;
					$TargetSubFolderObject.Save($TargetFolderObject.Id);
				}
				else
				{
					$TargetSubFolderObject = $FindFolderResults.Folders[0];
				}
				MoveItems($SourceSubFolderObject, $TargetSubFolderObject, $SourceFolderPath, $TargetFolderPath);
			}
		}
	}
}


Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath = $args[0];
	
	$Folder = $RootFolder;
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\';
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0);
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly;
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i]);
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View);
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null;
					Write-Host "Failed to find " $PathElements[$i];
					Write-Host "Requested folder path: " $FolderPath;
					break;
				}
				
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id)
			}
		}
	}
	
	$Folder;
}



# The following is the main script

 
# Check if we need to ignore any certificate errors
# This needs to be done *before* the managed API is loaded, otherwise it doesn't work consistently (i.e. usually doesn't!)
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -foregroundColor Yellow
    TrustAllCerts
}
 
# Load EWS Managed API
if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}
 
# Create Service Object.  We use Exchange 2010 schema
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)


# Set credentials if specified, or use logged on user.
 if ($Username -and $Password)
 {
     if ($Domain)
     {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain);
     } else {
         $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password);
     }
     
} else {
     $service.UseDefaultCredentials = $true;
 }
 

# Set EWS URL if specified, or use autodiscover if no URL specified.
if ($EwsUrl)
{
	$service.URL = New-Object Uri($EwsUrl);
}
else
{
	try
	{
		Write-Host "Performing autodiscover for $Mailbox";
		$service.AutodiscoverUrl($Mailbox, {$True});
	}
	catch
	{
		throw;
	}
}
 
# Set impersonation if specified
if ($Impersonate)
{
	$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox);
}

# Check we can bind to the source folder (if not, stop now)
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot);
$SourceFolderObject = GetFolder($MailboxRoot, $SourceFolder);

if ($SourceFolderObject)
{
	# We have the source folder, now check we can get the target folder
	if ($LogVerbose) { Write-Host "Source folder located: " $SourceFolderObject.DisplayName; }
	$TargetFolderObject = GetFolder($MailboxRoot, $TargetFolder);
	if ($TargetFolderObject)
	{
		# Found target folder, now initiate move
		if ($LogVerbose) { Write-Host "Target folder located: " $TargetFolderObject.DisplayName; }
		MoveItems($SourceFolderObject, $TargetFolderObject, '', '');
		
		# If delete parameter is set, check if the source folder is now empty (and if so, delete it)
		if ($DeleteSourceFolder)
		{
			$SourceFolderObject.Load();
			if (($SourceFolderObject.TotalCount -eq 0) -And ($SourceFolderObject.ChildFolderCount -eq 0))
			{
				# Folder is empty, so can be safely deleted
				try
				{
					$SourceFolderObject.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete);
					Write-Host $SourceFolder "successfully deleted" -foregroundcolor Green;
				}
				catch
				{
					Write-Host "Failed to delete " $SourceFolder -foregroundcolor Red;
				}
			}
			else
			{
				# Folder is not empty
				Write-Host $SourceFolder "could not be deleted as it is not empty." -foregroundcolor Red;
			}

		}
	}
}

