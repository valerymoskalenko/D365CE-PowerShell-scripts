# D365CE-PowerShell-scripts
PowerShell scripts for Dynamics 365 CE (CRM)

## Get-D365ceMetadata.ps1
PowerShell script for Dynamics 365 CRM
Generates CSV file with field list for selected DataEntity with description and metadata for each field. 
Please find how to connect to CRM instance by the Azure Application here https://vmoskalenkoblog.wordpress.com/2018/06/25/reading-odata-from-dynamics-365/

Please find CSV example below (I privde below just few rows from the output)

|LogicAppAttr|Attribute|Name|DisplayName|Description|Type|DataType| AttributeTypeAttr|String max length| Options values|IsPrimaryId|IsPrimaryName|IsValidForCreate|IsValidForRead|IsValidForUpdate|IsCustomAttribute|IsSearchable|
|--|--|--|--|--|--|--|--|--|--|--|--|--|--|--|--|--|
|_accountid_value|_accountid_value|_accountid_value|Account|Unique identifier of the account with which the lead is associated.|Edm.Guid|guid|Lookup|0|account|False|False|False|True|False|False|False|
|_accountid_type|_accountid_value@OData.Community.Display.V1.FormattedValue|Types for _accountid_value|Types for Account||String|str|String|||False|False|False|True|False|False|False|
|_campaignid_value|_campaignid_value|_campaignid_value|Source Campaign|Choose the campaign that the lead was generated from to track the effectiveness of marketing campaigns and identify  communications received by the lead.|Edm.Guid|guid|Lookup|0|campaign|False|False|True|True|True|False|False|
|_campaignid_type|_campaignid_value@OData.Community.Display.V1.FormattedValue|Types for _campaignid_value|Types for Source Campaign||String|str|String|||False|False|True|True|True|False|False|
|_contactid_value|_contactid_value|_contactid_value|Contact|Unique identifier of the contact with which the lead is associated.|Edm.Guid|guid|Lookup|0|contact|False|False|False|True|False|False|False|
|_contactid_type|_contactid_value@OData.Community.Display.V1.FormattedValue|Types for _contactid_value|Types for Contact||String|str|String|||False|False|False|True|False|False|False|
|_createdby_value|_createdby_value|_createdby_value|Created By|Shows who created the record.|Edm.Guid|guid|Lookup|0|systemuser|False|False|False|True|False|False|False|
|_createdby_type|_createdby_value@OData.Community.Display.V1.FormattedValue|Types for _createdby_value|Types for Created By||String|str|String|||False|False|False|True|False|False|False|
|address1_addressid|address1_addressid|address1_addressid|Address 1: ID|Unique identifier for address 1.|Edm.Guid|guid|Uniqueidentifier|0||True|False|True|True|True|False|False|
|address1_addressid|address1_addressid@OData.Community.Display.V1.FormattedValue|Types for address1_addressid|Types for Address 1: ID||String|str|String|||True|False|True|True|True|False|False|
|address1_addresstypecode|address1_addresstypecode|address1_addresstypecode|Address 1: Address Type|Select the primary address type.|Edm.Int32|int|Picklist|0|Default Value = 1; |False|False|True|True|True|False|False|
|_address1_addresstypecode_label|address1_addresstypecode@OData.Community.Display.V1.FormattedValue|Label for address1_addresstypecode|Formatted value for Address 1: Address Type||String|str|String|||False|False|True|True|True|False|False|
|address1_city|address1_city|address1_city|City|Type the city for the primary address.|Edm.String|str|String|80||False|False|True|True|True|False|False|
|address1_composite|address1_composite|address1_composite|Address 1|Shows the complete primary address.|Edm.String|Notes|Memo|1000||False|False|False|True|False|False|False|
|address1_country|address1_country|address1_country|Country/Region|Type the country or region for the primary address.|Edm.String|str|String|80||False|False|True|True|True|False|False|
|address1_county|address1_county|address1_county|Address 1: County|Type the county for the primary address.|Edm.String|str|String|50||False|False|True|True|True|False|False|
|companyname|companyname|companyname|Company Name|Type the name of the company associated with the lead. This becomes the account name when the lead is qualified and converted to a customer account.|Edm.String|str|String|100||False|False|True|True|True|False|True|
|confirminterest|confirminterest|confirminterest|Confirm Interest|Select whether the lead confirmed interest in your offerings. This helps in determining the lead quality.|Edm.Boolean|boolean|Boolean|0||False|False|True|True|True|False|False|
|createdon|createdon|createdon|Created On|Date and time when the record was created.|Edm.DateTimeOffset|utcDateTime|DateTime|0||False|False|False|True|False|False|False|
|decisionmaker|decisionmaker|decisionmaker|Decision Maker?|Select whether your notes include information about who makes the purchase decisions at the lead's company.|Edm.Boolean|boolean|Boolean|0||False|False|True|True|True|False|False|
|description|description|description|Description|Type additional information to describe the lead, such as an excerpt from the company's website.|Edm.String|Notes|Memo|2000||False|False|True|True|True|False|False|
|donotbulkemail|donotbulkemail|donotbulkemail|Do not allow Bulk Emails|Select whether the lead accepts bulk email sent through marketing campaigns or quick campaigns. If Do Not Allow is selected, the lead can be added to marketing lists, but will be excluded from the email.|Edm.Boolean|boolean|Boolean|0||False|False|True|True|True|False|False|
|donotemail|donotemail|donotemail|Do not allow Emails|Select whether the lead allows direct email sent from Microsoft Dynamics 365.|Edm.Boolean|boolean|Boolean|0||False|False|True|True|True|False|False|
|leadid|leadid|leadid|Lead|Unique identifier of the lead.|Edm.Guid|guid|Uniqueidentifier|0||True|False|True|True|False|False|False|
|leadid|leadid@OData.Community.Display.V1.FormattedValue|Types for leadid|Types for Lead||String|str|String|||True|False|True|True|False|False|False|
|leadqualitycode|leadqualitycode|leadqualitycode|Rating|Select a rating value to indicate the lead's potential to become a customer.|Edm.Int32|int|Picklist|0|Hot = 1; Warm = 2; Cold = 3; |False|False|True|True|True|False|False|
|_leadqualitycode_label|leadqualitycode@OData.Community.Display.V1.FormattedValue|Label for leadqualitycode|Formatted value for Rating||String|str|String|||False|False|True|True|True|False|False|
|leadsourcecode|leadsourcecode|leadsourcecode|Lead Source|Select the primary marketing source that prompted the lead to contact you.|Edm.Int32|int|Picklist|0|Advertisement = 1; Employee Referral = 2; External Referral = 3; Partner = 4; Public Relations = 5; Seminar = 6; Trade Show = 7; Web = 8; Word of Mouth = 9; Other = 10; |False|False|True|True|True|False|False|
|_leadsourcecode_label|leadsourcecode@OData.Community.Display.V1.FormattedValue|Label for leadsourcecode|Formatted value for Lead Source||String|str|String|||False|False|True|True|True|False|False|
