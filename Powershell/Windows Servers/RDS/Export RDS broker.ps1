Import-Module ExportImportRdsDeployment


#files end up in c:\ temp unless otherwise specified
Export-RDCollectionsFromConnectionBroker -ConnectionBroker localhost -Verbose