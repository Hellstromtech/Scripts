#Parameters
$TenantAdminURL = "https://mullergasequipment-admin.sharepoint.com/"
#Site Collection feature "Open Documents in Client Applications by Default"
$FeatureId = "fa6a1bcc-fb4b-446b-8460-f4de5f7411d5"
 
#Connect to Admin Center
$Cred = Get-Credential
Connect-PnPOnline -Url $TenantAdminURL -Credentials $Cred
 
#Get All Site collections - Exclude: Seach Center, Mysite Host, App Catalog, Content Type Hub, eDiscovery and Bot Sites
$SitesCollections = Get-PnPTenantSite | Where -Property Template -NotIn ("SRCHCEN#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
 
#Loop through each site collection
ForEach($Site in $SitesCollections)
{
    #Connect to site collection
    Write-host -f Yellow "Trying to Activate the feature on site:"$Site.Url   
    $SiteConn = Connect-PnPOnline -Url $Site.Url -Credentials $Cred
 
    #Get the Feature
    $Feature = Get-PnPFeature -Scope Site -Identity $FeatureId -Connection $SiteConn
 
    #Check if feature is activated
    If($Feature.DefinitionId -eq $null)
    {
        #Enable site collection feature
        Enable-PnPFeature -Scope Site -Identity $FeatureId -Force -Connection $SiteConn
 
        Write-host -f Green "`tFeature Activated Successfully!"
    }
    Else
    {
        Write-host -f Cyan "`tFeature is already active!"
    }
    Disconnect-PnPOnline -Connection $SiteConn
}
