# This script is used in conjunction with Microsoft Intune. If you want to automate
# the deployment of a SharePoint Document Library via the OneDrive app, you can do so
# through Intune. However, Microsoft has removed the ability to get the site ID through
# the popup when adding the site to OneDrive. This PowerShell script is a replacement
# for that functionality.


# Question choices
  $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
  $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
  $cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel",""
  $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)


# Install MSOnline PowerShell Module if needed
  If (Get-Module -ListAvailable -Name "MSOnline") {
      Write-Host 'MSOnline is already installed'
  }Else {
      Install-Module MSOnline
  }

# Install SharePoint PowerShell Module if needed
  If (Get-Module -ListAvailable -Name "SharePointPnPPowerShellOnline") {
      Write-Host 'SharePointPnPPowerShellOnline is already installed'
  }Else {
      Install-Module SharePointPnPPowerShellOnline
  }


#Queries to find out if you are using MFA
  $title = "Multifactor Authentication"
  $message = "Are you connecting to a tenant that uses MFA?"
  $result = $host.ui.PromptForChoice($title, $message, $options, 1)

#Exits if cancel was chosen
  switch ($result) {
      0 {}
      1 {}
      2 {exit}
  }


# Connect to Microsoft Online
  Connect-msolservice

# Prompt user for variables
  $tenant = Read-Host -Prompt 'Enter your Tenant short name here' # xxxx.sharepoint.com
  $siteName = Read-Host -Prompt 'Input the portion of the SharePoint URL that comes after /sites/' #SharePoint site name
  $docLib = Read-Host -Prompt 'Enter the name of the document library (ex. `Documents`)' #SharePoint Document Library

#Connects to SharePoint using whichever method is required
  switch ($result) {
      0 {
          # Connect to SharePoint using MFA
          Connect-PnPOnline https://$tenant.sharepoint.com/sites/$siteName -SPOManagementShell
      }
      1 {
          #Connect to SharePoint without MFA
          Connect-PnPOnline https://$tenant.sharepoint.com/sites/$siteName
      }
  }

# Capture and convert the Tenant ID
  $tenantid = Get-MSOLCompanyInformation | select objectID
  $tenantid = $tenantid.objectID
  $tenantid = $tenantid -replace '-','%2D'

# Capture and convert the Site ID
  $siteid = Get-PnPSite -Includes Id | select id
  $siteid = $siteid.Id -replace '-','%2D'
  $siteid = '%7B' + $siteid + '%7D'

# Capture and convert the Web ID
  $webid = Get-pnpweb -Includes Id | select id
  $webid = $webid.Id -replace '-','%2D'
  $webid = '%7B' + $webid + '%7D'

# Capture and convert the List ID
  $listid = Get-PnPList $docLib -Includes Id | select id
  $listid = $listid.Id -replace '-','%2D'
  $listid = '%7B' + $pnplist + '%7D'
  $listid = $listid.toUpper()

# Builds the Full URL
  $libraryid = 'tenantId=' + $tenantid + '&siteId=' + $siteid + '&webId=' + $webid + '&listId=' + $listid + '&webUrl=https%3A%2F%2F' + $tenant + '%2Esharepoint%2Ecom%2Fsites%2F' + $siteName + '&version=1'

# Display the complete URL to Copy and Paste
  Write-Output 'Library ID: ' $libraryid
