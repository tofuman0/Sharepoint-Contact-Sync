# Sharepoint Contact Sync

Utility made to synchronise contacts from either Active Directory or CSV file. You will need to create a certificate to use for 365 authentication (see: https://learn.microsoft.com/en-us/azure/active-directory/develop/howto-create-self-signed-certificate).

A configuration file named "contactsync.json" will be created in the folder that the application is run from if it doesn't exist already. Ensure you edit this configuration file with the appropriate settings before starting the application.

Run with -cert cert_location.pfx cert_password to set the certificate in the configuration. The password is obfuscated so that it isn't easily known by just looking at the configuration file. However the method used to obfuscate isn't complex and it easily reversed so precautions must be taken to ensure the configuration file isn't accessible by unauthorised users.

Run with -getfieldvalues to obtain the field names. As the names in 365 aren't the names that are used to store the data. 365 shows the description not the actual name. Fields with spaces will often have `_x0020_` (hex value for space in unicode) in place of a spaces.

Running the application without any arguments will set the application to try to synchronise the data.

***NOTE:*** *To use a CSV file it is assumed that it has the following columns "Full Name" (separated with a space), "Job Title", "Phone No.", "Mobile No.", "Extension", "Division" and that they are separated with commas*

```
Usage: Sharepoint Contact Sync.exe [-getfieldvalues] [-cert cert_path cert_password]

Options:

    -getfieldvalues     Returns with field values in sharepoint list.
    -cert               Sets the certificate to use and its password in configuration file.
	-defaultconfig      Replaces/Creates with default configuration file.

Note: Populate the SharepointFieldNames values in the configuration file from the values retrieved from -getfieldvalues.
```

## Configuration

**SiteURL:** Tenant site eg: https://tenant.sharepoint.com/Sites/Site

**ListName:** Name of list in sharepoint

**ApplicationID:** Application ID from application that has been created in 365 Azure portal

**CertificatePath:** Path of certifcate

**CertificatePassword:** Obfuscated certificate password. Use -cert switch to set.

**TenantURL:** Tenant URL eg: tenant.onmicrosoft.com

**DomainController:** Fully qualified domain controller eg: dc.domain.local

**LDAPPath:** LDAP path for users eg: CN=Users,DC=Domain,DC=local

**CSVPath:** CSV path

**DataFetchType:** Type of fetch type. Options: csv, ldap

**ExecuteLimit:** Execution limit of queries. Queries will be split up by this limit. Attempting to query too many results will fail

**RequestTimeout:** Limit to wait for query to execute in seconds

**ClearListFirst:** Whether or not to clear existing results in list when updating lists. Options: true, false

**SharepointFieldNames:** Grab property names by using the -getfieldvalues
