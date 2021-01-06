# Python 3.6

# This script is in development please use against sample data first.
# Issue to be mindful of: https://support.esri.com/en/bugs/nimbus/QlVHLTAwMDExMDU4MA== the bug was closed in 2.5, but I was able to repro
# in Pro 2.6.2, and 2.7 Alpha.

# Script could require Standard or Advanced license!

import arcgis
from arcgis.gis import GIS
import arcpy
import shutil
import zipfile
import os
import time

# Define variables
portalURL = r'' # Portal URL 
username = '' # Portal Username
password = '' # Portal Password
itemUrl = ''  # REST URL for Hosted Feature Service
layers = '0,1'  # Layer indices from REST endpoint
outPath = r'C:\temp'  # Download folder directory
sde_conn = r""  # Folder location for the SDE file
webGIS = "enterprise" # Use "online" for ArcGIS Online or "enterprise" for ArcGIS Enterprise.

# Connect to GIS
print("Connecting to GIS...")
gis = GIS(portalURL, username, password, verify_cert=False)

# Connect to the Feature Service and define the Feature Layer Collection.
print("Connecting to Feature Service and checking sync capabilities...")
survey_flc = arcgis.features.FeatureLayerCollection(itemUrl, gis)
# Confirm Sync is enabled.
sync = survey_flc.properties.syncEnabled

if not sync:
    print("Sync is disabled, temporarily enabling it...")
    disable_sync = 1
    updateSync = {"syncEnabled": True, "syncCapabilities": {
        "supportsAsync": True,
        "supportsRegisteringExistingData": True,
        "supportsSyncDirectionControl": True,
        "supportsPerLayerSync": True,
        "supportsPerReplicaSync": True,
        "supportsSyncModelNone": True,
        "supportsRollbackOnFailure": True,
        "supportsAttachmentsSyncDirection": True,
        "supportsBiDirectionalSyncForServer": True
    }}
    survey_flc.manager.update_definition(updateSync)
    survey_flc = arcgis.features.FeatureLayerCollection(itemUrl, gis)

# Define a default extent when creating the replica:
print("Defining default extent for replica...")
extents = survey_flc.properties['fullExtent']
extents_str = ",".join(format(x, "10.3f") for x in [extents['xmin'], extents['ymin'], extents['xmax'], extents['ymax']])

# Define a default geometry filter when creating the replica:
print("Defining default geometry filter for replica...")
geom_filter = {'geometryType': 'esriGeometryEnvelope'}
geom_filter.update({'geometry': extents_str})

# Create the replica: 
# Doc Link: https://developers.arcgis.com/python/guide/checking-out-data-from-feature-layers-using-replicas/
print("Creating replica..")
# Test leaving the spatial reference blank and see what happens
replica = survey_flc.replicas.create('syncSurvey', layers, '', geom_filter, '102100', 'esriTransportTypeUrl',
                                     True, False, False, 'bidirectional',
                                     'none', 'filegdb', None, False, outPath, None,
                                     'server', None)

print("Extracting replica...")
zfile = zipfile.ZipFile(replica)
gdb_name = zfile.namelist()[0]
zfile.extractall(outPath)

if "disable_sync" in globals():
    print("Disabling sync...")
    # if disable_sync == 1:
    updateSync = {"syncEnabled": False}
    survey_flc.manager.update_definition(updateSync)
# ArcGIS Datastore creates a GDB with a length of 45 characters, whereas ArcGIS Online creates a GDB with a length of 36 characters.
if webGIS == "online":
    surveyGDB = outPath + "\\" + gdb_name[:36]
elif webGIS == "enterprise":
    surveyGDB = outPath + "\\" + gdb_name[:45]
else:
    print("Error: Please enter WebGIS Infrastructure")

# Start working with the downloaded data:
workspace = arcpy.env.workspace = surveyGDB
arcpy.env.overwriteOutput = True
arcpy.env.maintainAttachments = True
arcpy.env.preserveGlobalIds = True
# List the Feature Classes & Tables
# Tables also returns Attachment tables
print("Identifying File Geodatabase content...")
featureClasses = arcpy.ListFeatureClasses()
tables = arcpy.ListTables()
fc_name = featureClasses[0]

arcpy.env.workspace = sde_conn
# List the feature classes in the EGDB
GDB_Features = arcpy.ListFeatureClasses()

# Match the feature class name from the FGDB to a Feature Class in the EGDB
print("Identifying Feature Classes in the Enterprise Geodatabase...")
matching = [s for s in GDB_Features if featureClasses[0] in s]
#string_matching = ' '.join([str(elem) for elem in matching])

if len(matching) > 0:
    print("Data is already present in Enterprise Geodatabase, updating data...")
    arcpy.env.workspace = surveyGDB
    print("Filtering Feature Class(s) for new records...")
    arcpy.management.JoinField(surveyGDB + "\\" + fc_name, "globalid", sde_conn + "\\" + fc_name, "globalid",
                               "LAST_SYNC")
    selection = arcpy.management.SelectLayerByAttribute(surveyGDB + "\\" + fc_name, "NEW_SELECTION",
                                                        "LAST_SYNC IS NOT NULL", None)
    arcpy.management.DeleteFeatures(selection)
    print("Copying updated Feature Class(s) and table(s) to Enterprise Geodatabase...")
    temp_FC = fc_name + "_temp"
    try:
        copied_features = arcpy.management.Copy(surveyGDB + "\\" + fc_name, sde_conn + "\\" + temp_FC)
    except:
        print("Initial copy failed attempting to copy data over using XML workspace document")
        for features in featureClasses:
            arcpy.Rename_management(features, features + '_temp')
        for table in tables:
            arcpy.Rename_management(table, table + '_temp')
        tempData = outPath + "\\" + "TempData.xml"
        arcpy.management.ExportXMLWorkspaceDocument(surveyGDB, tempData, "DATA", "BINARY", "METADATA")
        print("XML Workspace document created, attempting to import...")
        arcpy.management.ImportXMLWorkspaceDocument(sde_conn, tempData, "DATA")
        print("Data imported successfully continuing with append workflow...")
    arcpy.env.workspace = sde_conn
    print("Appending Feature Class(s)...")
    for features in featureClasses:
        interest = arcpy.ListFeatureClasses(wild_card='*' + features + '_temp')
        FCOI = interest[0]
        target = arcpy.ListFeatureClasses(wild_card='*' + features)
        Target_FC = target[0]
        arcpy.management.AddIndex(sde_conn + "\\" + FCOI, "globalid", "GlobalId", "UNIQUE", "NON_ASCENDING")
        arcpy.DisableEditorTracking_management(Target_FC)
        arcpy.management.Append(FCOI, Target_FC, "TEST", '', '', '')
        arcpy.Delete_management(FCOI)
        if webGIS == "online":
            arcpy.management.EnableEditorTracking(Target_FC, "Creator", "CreationDate", "Editor",
                                                  "EditDate", "NO_ADD_FIELDS", "UTC")
        elif webGIS == "enterprise":
            arcpy.management.EnableEditorTracking(Target_FC, "created_user", "created_date",
                                                  "last_edited_user", "last_edited_date",
                                                  "NO_ADD_FIELDS", "UTC")
    print("Deleting intermediate Feature Classes...")
    FGDB_Tables = tables

    for i in range(len(FGDB_Tables) - 1, -1, -1):
        if FGDB_Tables[i].endswith("__ATTACH"):
            del (FGDB_Tables[i])
    arcpy.env.workspace = sde_conn
    for tables in FGDB_Tables:
        TOI = arcpy.ListTables(wild_card='*' + tables + '_temp')
        Target_TOI = arcpy.ListTables(wild_card='*' + tables)
        print(Target_TOI)
        if TOI:
            print("Appending table(s)")
            for input in TOI:
                for target in Target_TOI:

                    arcpy.DisableEditorTracking_management(target)
                    arcpy.management.Append(sde_conn + "\\" + input, sde_conn + "\\" + target, "TEST", '', '', '')
                    print("Deleting temp tables")
                    arcpy.Delete_management(input)
            for Table_editors in Target_TOI:
                if webGIS == "online":
                    arcpy.management.EnableEditorTracking(Table_editors, "Creator", "CreationDate", "Editor",
                                                          "EditDate",
                                                          "NO_ADD_FIELDS", "UTC")
                elif webGIS == "enterprise":
                    arcpy.management.EnableEditorTracking(Table_editors, "created_user", "created_date",
                                                          "last_edited_user", "last_edited_date", "NO_ADD_FIELDS",
                                                          "UTC")
            print("Deleting temp tables")

    print("Adding sync timestamp...")
    arcpy.management.CalculateField(sde_conn + "\\" + fc_name, "LAST_SYNC", "datetime.datetime.now()", "PYTHON3", '',
                                    "TEXT")

else:
    print("Data is not present in the Enterprise Geodatabase copying from File Geodatabase...")
    try:
        arcpy.management.Copy(surveyGDB + "\\" + fc_name, sde_conn + "\\" + fc_name)
    except:
        print("Initial copy failed attempting to copy data over using XML workspace document")
        backup = outPath + "\\" + "FGDBBackup.xml"
        arcpy.management.ExportXMLWorkspaceDocument(surveyGDB, backup, "DATA", "BINARY", "METADATA")
        print("XML workspace document created, starting the import process...")
        arcpy.management.ImportXMLWorkspaceDocument(sde_conn, backup, "DATA")
        print("Data imported successfully...")
    arcpy.env.workspace = sde_conn
    arcpy.management.AddIndex(sde_conn + "\\" + fc_name, "globalid", "GlobalId", "UNIQUE", "NON_ASCENDING")
    print("Adding last sync field and time...")
    arcpy.management.AddField(sde_conn + "\\" + fc_name, "LAST_SYNC", "DATE", None, None, None, '', "NULLABLE", "NON_REQUIRED", '')
    arcpy.management.CalculateField(sde_conn + "\\" + fc_name, "LAST_SYNC", "datetime.datetime.now()", "PYTHON3", '', "TEXT")

#print("Cleaning up downloaded File Geodatabase...")
#arcpy.Compact_management(surveyGDB)
#os.remove(replica)
#os.remove(surveyGDB)
#shutil.rmtree(replica)
#shutil.rmtree(surveyGDB)
print("Finished syncing survey data with Enterprise Geodatabase.")

# Error messages encountered:

# arcgisscripting.ExecuteError: ERROR 160371: The current version does not support editing (base, consistent, or closed)
# when performing final calculate field for initial import. Data is not versioned, ran tool manually in Pro after the error
# and it worked as expected.
