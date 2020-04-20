# redfish-spec-generator
parsing BMC server and compare to redfish spec with odata.id then try to modify spec
# how to use it ?
1. prepare your own spec and add a row of 
|------------------------------------------------------|
|Type URI | /redfish/v1/AccountService/Accounts/benchin|
|------------------------------------------------------|
in every response property table for parser to check properties exists or not

2. check your spec whether there API support list in your spec if so its format should be match to like below:

|Resource         | Resource URI                 | Redfish Schema                         |
|-----------------------------------------------------------------------------------------|
|TelemetryService | /redfish/v1/TelemetryService |TelemetryService.v1_1_1.TelemetryService|


As you complished it, parser will check if every resource URI match to its schema on the spec

3. launch redfish parser it will ask you enter 
    a. your server ip
    b. spec file path+file name with absolute path
    c. path of your redfishSupportAPI.txt which include all of the URI you want to check list on that txt file
    
4. After redfish parser finish its task, it will generate 5 documents in the same foler of parser program
    including 
    a. missingURI.docx => URI exists in server but not in redfish spec
    b. modifiedSpec.docx => add missing property onto the spec and save it for another file
    c. generateSpec.docx => using catched URI and properties write into the document
    d. missing_redfish_schema.txt => redfish schema exist in server but not in redfish spec
    e. all_redfish_property.txt => write all properties parser catched into document


