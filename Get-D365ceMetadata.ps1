#https://github.com/valerymoskalenko/D365CE-PowerShell-scripts

$tenantDomain = 'contoso.com' 
[uri]$url =  'https://contosocrm.api.crm.dynamics.com'
$ApplicationClientId = '699ee32d-a659-000-000-445a3dc7f0fb' 
$ApplicationClientSecretKey = 'rNfjoieiru984tyolijfeirjfowert85t4t5fvvjeriJbXE='; 

$DataEntity = 'account' #product'#'salesorderdetail' #'salesorder' #'account' #'contact'  #'customeraddress', 'businessunit' 'transactioncurrency' 'lead'

$ErrorActionPreference = "Stop" 
Write-Host "Authorization..." -ForegroundColor Yellow
Add-Type -AssemblyName System.Web
[string]$absoluteURL = $url.AbsoluteUri.Remove($url.AbsoluteUri.Length-1,1)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  

$Body = @{
    "client_id" = $ApplicationClientId
    "client_secret" = $ApplicationClientSecretKey
    "grant_type" = 'client_credentials'
    "scope" = "$absoluteURL/.default"
}


$login = $null
$login = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantDomain/oauth2/v2.0/token" -Body $Body -ContentType 'application/x-www-form-urlencoded' -Verbose
 
$Bearer = $null
[string]$Bearer = $login.access_token 
 
 
 
Write-Host "Getting data..." -ForegroundColor Yellow
 
$headers = @{
    "Accept" = "application/xml"
    "Accept-Charset" = "UTF-8"
    "Authorization" = "Bearer $Bearer"
    "Host" = "$($url.Host)"
}

Write-Host "Results..." -ForegroundColor Yellow
[System.UriBuilder] $urlBuilder = $url
$urlBuilder.Path = '/api/data/v9.0/$metadata'

[xml] $result=$null
[xml] $result = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose

$Entity = $result.Edmx.DataServices.Schema.EntityType | where {$_.Name -eq $DataEntity}

function getEdmDataType {    
    param ([parameter(Mandatory=$true)] [string]$dataType)

    switch($dataType)
    {
        Edm.Guid             {'guid'}
        Edm.String           {'str'}
        Edm.Int32            {'int'}
        Edm.Int64            {'int64'}
        Edm.Decimal          {'real'}
        Edm.Double           {'real'}
        Edm.Boolean          {'boolean'}
        Edm.Binary           {'binary'}
        Edm.Object           {'str'} 
        Edm.DateTimeOffset   {'utcDateTime'}
        Edm.DateTime         {'utcDateTime'}
        Edm.Date             {'Date'}

        default              {Write-Error "Unknown type at $dataType" }
    }
}


Write-Host "Generate a list of Attributes" -ForegroundColor Yellow
#Generate a list of Attributes
$attrList = @{}
[int] $i = 0;
foreach($el in $Entity.Property | Sort-Object -Property Name )
{
    if($el.Name.StartsWith('_') -and $el.Name.EndsWith('_value')) #checks for Lookup type
    {
        $elName = $el.Name.Substring(1, $el.Name.Length-7)  #Transform from '_owningbusinessunit_value' to 'owningbusinessunit'
    } else {
        $elName = $el.Name;
    }

    $processAttr = $true;

    #process Attribute
    if($processAttr)
    {

        $dataType = getEdmDataType($el.Type);

        $urlBuilder.path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')/Attributes(LogicalName='$elName')"
        $urlBuilder.Query = $null
        $AttrReq = $null
        $AttrReq = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose   
        $AttrType = $AttrReq.AttributeType

        [string]$OptionSet = $null
        [int]$StringMaxLengh = $null

        if($AttrType -in ('Picklist','State','Status' ))
        {  #https://vss2.crm.dynamics.com/api/data/v9.0/EntityDefinitions(LogicalName='account')/Attributes(LogicalName='paymenttermscode')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$expand=OptionSet,GlobalOptionSet
        
            if($AttrType -eq 'Status')
            {
                Write-Host ".. Expanding Status $elName" -ForegroundColor Yellow    
                $urlBuilder.Path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')/Attributes(LogicalName='$elName')/Microsoft.Dynamics.CRM.StatusAttributeMetadata"
            } 
            elseif($AttrType -eq 'State') 
            {
                Write-Host ".. Expanding State $elName" -ForegroundColor Yellow    
                $urlBuilder.Path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')/Attributes(LogicalName='$elName')/Microsoft.Dynamics.CRM.StateAttributeMetadata"
            } 
            else 
            {
                Write-Host ".. Expanding PickList $elName" -ForegroundColor Yellow    
                $urlBuilder.Path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')/Attributes(LogicalName='$elName')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
            }

            $urlBuilder.Query = '$expand=OptionSet,GlobalOptionSet'
            $AttrReqOptionSet = $null
            $AttrReqOptionSet = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose

            foreach($OptionSetValue in $AttrReqOptionSet.OptionSet.Options)
            {
                $OptionSet += ($OptionSetValue.Label.UserLocalizedLabel | where {$_.LanguageCode -eq 1033}).Label + ' = ' + $OptionSetValue.Value + '; '            
            }
            Write-Host ".... OptionSet values are $OptionSet" -ForegroundColor DarkGray

            #Add Formatted Value, like "_ownerid_value@OData.Community.Display.V1.FormattedValue":  "Jamie Reding (Sample Data)",
            $attrFormattedName = ($el.Name + '@OData.Community.Display.V1.FormattedValue');
            $attrElement = $null;
            $attrElement = @{
                Name              = 'Label for ' + $el.Name
                DataMemberAttr    = $attrFormattedName
                LogicAppAttr      = '_' + $el.Name + '_label'
                Type              = 'String'
                DataType          = 'str'
                DisplayName       = ('Formatted value for ' + ($AttrReq.DisplayName.UserLocalizedLabel | where {$_.LanguageCode -eq 1033}).Label)
                AttributeTypeAttr = 'String'
                IsPrimaryId       = ($AttrReq.IsPrimaryId)
                IsPrimaryName     = ($AttrReq.IsPrimaryName)
                IsValidForCreate  = ($AttrReq.IsValidForCreate)
                IsValidForRead    = ($AttrReq.IsValidForRead)
                IsValidForUpdate  = ($AttrReq.IsValidForUpdate)
                IsCustomAttribute = ($AttrReq.IsCustomAttribute)
                IsSearchable      = ($AttrReq.IsSearchable)
            }
            $attrList[$attrFormattedName] = $attrElement
        }  
        elseif($AttrType -eq 'String')
        {
            $StringMaxLengh = $AttrReq.MaxLength
            $dataType = $dataType + '[' + $StringMaxLengh + ']'
            $AttrReq.AttributeType = $AttrReq.AttributeType + '[' + $StringMaxLengh + ']'
        }
        elseif($AttrType -eq 'Memo')
        {
            $StringMaxLengh = $AttrReq.MaxLength
            $dataType = 'Notes' #AX Data type
        }

        elseif($AttrType -in ('Uniqueidentifier','Owner','Lookup'))
        {
            $dataType = 'guid'#AX Data type
        
            #Add Formatted Value, like "_ownerid_value@OData.Community.Display.V1.FormattedValue":  "Jamie Reding (Sample Data)",
            $attrFormattedName = ($el.Name + '@OData.Community.Display.V1.FormattedValue');
            $attrElement = $null;
            $attrElement = @{
                Name              = 'Types for ' + $el.Name
                DataMemberAttr    = $attrFormattedName
                LogicAppAttr      = $el.Name.Replace('_value','_type');
                Type              = 'String'
                DataType          = 'str'
                DisplayName       = ('Types for ' + ($AttrReq.DisplayName.UserLocalizedLabel | where {$_.LanguageCode -eq 1033}).Label)
                AttributeTypeAttr = 'String'
                IsPrimaryId       = ($AttrReq.IsPrimaryId)
                IsPrimaryName     = ($AttrReq.IsPrimaryName)
                IsValidForCreate  = ($AttrReq.IsValidForCreate)
                IsValidForRead    = ($AttrReq.IsValidForRead)
                IsValidForUpdate  = ($AttrReq.IsValidForUpdate)
                IsCustomAttribute = ($AttrReq.IsCustomAttribute)
                IsSearchable      = ($AttrReq.IsSearchable)
            }
            $attrList[$attrFormattedName] = $attrElement
        }
        elseif($AttrType -eq 'Money')
        {
            if($el.Name.EndsWith('_base')) #Base means local currency, i.e. in AX terms it's AmountMST
            {
                $dataType = 'AmountMST'#AX Data type
            } else {
                $dataType = 'AmountCur'#AX Data type
            }
        }


        ###GetSchemaName
        $urlBuilder.Path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')/Attributes(LogicalName='$elName')"
        $urlBuilder.Query = $null

        $SchemaName = $null
        $SchemaName = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose
        $EntityNameAttr = $null
        if($SchemaName.Targets)
        {
            ###GetEntityName
            $targetEntity = $SchemaName.Targets[0]
            $urlBuilder.Path = "/api/data/v9.0/EntityDefinitions(LogicalName='$targetEntity')"
            $EntityNameAttr = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose

            foreach($targetEntity in $SchemaName.Targets)
            {
                if($OptionSet)
                {
                    $OptionSet += '; ' 
                }
                $OptionSet += $targetEntity
            }
        }

        [string]$descriptionField = $null;
        $descriptionField = ($AttrReq.Description.UserLocalizedLabel | where {$_.LanguageCode -eq 1033}).Label;
        if ($descriptionField -ne '') { $descriptionField = $descriptionField.Replace(',',';') }

        $attrElement = $null;
        $attrElement = @{
            Name              = $el.Name
            DataMemberAttr    = $el.Name
            LogicAppAttr      = $el.Name
            Type              = $el.Type
            DataType          = $dataType  #getEdmDataType($el.Type)
            Description       = $descriptionField
            DisplayName       = ($AttrReq.DisplayName.UserLocalizedLabel | where {$_.LanguageCode -eq 1033}).Label
            AttributeTypeAttr = $AttrReq.AttributeType
            OptionSetValues   = $OptionSet
            StringMaxLengh    = $StringMaxLengh
            SchemaName        = if($SchemaName.IsCustomAttribute){$SchemaName.SchemaName} else {$SchemaName.LogicalName}
            EntityName        = if($EntityNameAttr.EntitySetName){$EntityNameAttr.EntitySetName} else {''}
            IsPrimaryId       = ($AttrReq.IsPrimaryId)
            IsPrimaryName     = ($AttrReq.IsPrimaryName)
            IsValidForCreate  = ($AttrReq.IsValidForCreate)
            IsValidForRead    = ($AttrReq.IsValidForRead)
            IsValidForUpdate  = ($AttrReq.IsValidForUpdate)
            IsCustomAttribute = ($AttrReq.IsCustomAttribute)
            IsSearchable      = ($AttrReq.IsSearchable)
        }
        $attrList[$el.Name] = $attrElement

    }    
    #Debug stop condition
    $i += 1;
    Write-Host ".. Next attribute is #$i" -ForegroundColor Yellow
    #if($i -ge 23) {break}
    #if($el.Name -eq '_parentid_value') {break};
}

#Get general information about DataEntity
$urlBuilder.path = "/api/data/v9.0/EntityDefinitions(LogicalName='$DataEntity')"
$urlBuilder.Query = $null
$EntityReq = $null
$EntityReq = Invoke-RestMethod -Method Get -Uri $urlBuilder.Uri.AbsoluteUri -Headers $headers -ContentType 'application/xml; charset=utf-8' -Verbose   
$EntityCSName = $EntityReq.CollectionSchemaName

#Generate variables
[String]$className = 'CieIS'+$EntityCSName +'DataContract'
[String[]]$vararr = @( $EntityCSName );


Write-Host "Generate CSV with Metadata description" -ForegroundColor Yellow

$vararr += ('Logic App Attr,Display Name,Attribute Type,OptionSet values, ' `
        + 'Attribute,Name,Description,Type, Data Type, String max length, ' `
        + 'IsPrimaryId,IsPrimaryName,IsValidForCreate,IsValidForRead,IsValidForUpdate,IsCustomAttribute,IsSearchable');
foreach($key in $attrList.Keys | Sort-Object )
{
    $element = $null;
    $element = $attrList[$key]
    #Write-Host "Name=" $element.Name  "Type=" $element.Type  -ForegroundColor Yellow
    #$dataType = $element.DataType

    $vararr += ($element.LogicAppAttr +',"'+ $element.DisplayName  +'","' + $element.AttributeTypeAttr +'","' + $element.OptionSetValues +'","'  `
        + $element.DataMemberAttr +'","' + $element.Name +'","'+$element.Description +'","' `
        + $element.Type +'","' + $element.DataType +'",'+ $element.StringMaxLengh +',' `
        + $element.IsPrimaryId +',"'+ $element.IsPrimaryName +'","'+ $element.IsValidForCreate +'","'+ $element.IsValidForRead +'","'+ $element.IsValidForUpdate +'","'+ $element.IsCustomAttribute +'","'+ $element.IsSearchable +'"' );

}

$vararr | Out-File -FilePath c:\temp\a.txt -Encoding utf8 -Force
notepad c:\temp\a.txt

