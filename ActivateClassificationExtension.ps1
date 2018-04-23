$credentials = Get-Credential

# Import the list of sites where we want to apply 
$sitesToProcess = import-csv "sites.csv"

# details of custom action/SPFx extension
[guid]$spfxExtId = "4017f67b-80c7-4631-b0e5-57bd266bc5c1"
[string]$spfxExtName = "ClassificationExtensionApplicationCustomizer"
[string]$spfxExtTitle = "ClassificationExtensionApplicationCustomizer"
[string]$spfxExtGroup = "Ashbay"
[string]$spfxExtDescription = "Adds security classification to every site"
[string]$spfxExtLocation = "ClientSideExtension.ApplicationCustomizer"
[string]$spfxExtension_Properties = "{""ClassificationPropertyBag"":""sc_x005f_BusinessImpact"", ""DefaultClassification"":"""", ""DefaultHandlingUrl"":""/SitePages/Handling-instructions.aspx""}"

function Add-CustomActionForSPFxExt ([string]$url, $clientContext) {
    Write-Output "-- About to add custom action to: $url"

    # NOTE - using direct CSOM here (rather than Add-PnPCustomAction) for now, due to https://github.com/SharePoint/PnP-PowerShell/issues/1048
    $rootWeb = $clientContext.Web
    $clientContext.ExecuteQuery()
    $customActions = $rootWeb.UserCustomActions
    $clientContext.Load($customActions)
    $clientContext.ExecuteQuery()

    $custAction = $customActions.Add()
    $custAction.Name = $spfxExtName
    $custAction.Title = $spfxExtTitle
    $custAction.Description = $spfxExtDescription
    $custAction.Location = $spfxExtLocation
    $custAction.ClientSideComponentId = $spfxExtId
    $custAction.ClientSideComponentProperties = $spfxExtension_Properties
    $custAction.Update()
    $clientContext.ExecuteQuery()

    Write-Output "-- Successfully added extension" 	
	
    Write-Output "Processed: $url"
}
function Remove-CustomActionForSPFxExt ([string]$extensionName, [string]$url, $clientContext) {
    Write-Output "-- About to remove custom action with name '$($extensionName)' from: $url"

    $actionsToRemove = Get-PnPCustomAction -Web $clientContext.Web | Where-Object {$_.Location -eq $spfxExtLocation -and $_.Name -eq $extensionName }
    Write-Output "-- Found $($actionsToRemove.Count) extensions with name $extensionName on this web." 	
    foreach ($action in $actionsToRemove) {
        Remove-PnPCustomAction -Identity $action.Id
        Write-Output "-- Successfully removed extension $extensionName from web $url." 	
    }

    Write-Output "Processed: $url"
}

# -- end functions --

foreach ($site in $sitesToProcess) {
    $ctx = $null
    $url = $site.Url
    try {
        Connect-PnPOnline -Url $url -Credentials $credentials
        Write-Output ""
        Write-Output "Authenticated to: $url"
        $ctx = Get-PnPContext
    }
    catch {
        Write-Error "Failed to authenticate to $url"
        Write-Error $_.Exception
    }

	# Make sure have a context before continuing
    if ($ctx) {
		# Find out if the extension is already added
		$existingActions = Get-PnPCustomAction -Web $ctx.Web | Where-Object {$_.Location -eq $spfxExtLocation -and $_.Name -eq $spfxExtName }
		
		# Count how many existing extensions we found
		$count = $($existingActions.Count)
		
		# Don't re-install extension if it is already there
        if ($count -ge 1) {
			#This assumes that you don't want to duplicate extensions. If you do, feel free to change the logic below
            if ($count -eq 1) {
                Write-Output "Extension is already applied"
            }
            else {
                Write-Warning "Extension is duplicated!"
            }
        }
        else {
			# Add the extension
			Add-CustomActionForSPFxExt $url $ctx
			Write-Output "-- Successfully added extension $spfxExtName to web $url."
        }
		
        #Add-CustomActionForSPFxExt $url $ctx
        #Remove-CustomActionForSPFxExt $spfxExtName $site $ctx
        #Get-PnPCustomAction -Web $ctx.Web | Where-Object {$_.Location -eq "ClientSideExtension.ApplicationCustomizer" }
    }
}