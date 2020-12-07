import os
import arcpy
import requests
import arcgis
from arcgis import GIS

gis = GIS(username="Username", password="Password")

sr = gis.content.search("AttachmentTesting", "Feature Layer")

surveyFL = sr[0]
FL = surveyFL.layers[0]

directory = r'C:\temp\'
URL = r"https://services.arcgis.com/.../FeatureServer/0"
token = 'Variable not needed just hardcoding for testing'

def update_attachment(url, token, oid, attachment, attachID, keyword):
    att_url = '{}/{}/updateAttachment'.format(url, oid)
    start, extension = arcpy.os.path.splitext(attachment)

    jpg_list = ['.jpg', '.jpeg']
    png_list = ['.png']
    if extension in jpg_list:
        files = {'attachment': (os.path.basename(attachment), open(attachment, 'rb'), 'image/jpeg')}
    elif extension in png_list:
        files = {'attachment': (os.path.basename(attachment), open(attachment, 'rb'), 'image/png')}
    else:
        files = {'attachment': (os.path.basename(attachment), open(attachment, 'rb'), 'application/octect-stream')}

    print(keyword)

    params = {'token': token,'f': 'json', 'attachmentId': attachID, 'keywords': keyword}
    r = requests.post(att_url, params, files=files)
    return r

# List folders in the directory, assumes directory contains folders with attachments associated with each OID, and each folder name is the OID 
# those attachmetns were downloaded from.

fldr = os.listdir(directory)
for object in fldr:
    attach = directory + '\\' + object
    OID = object
    attachment = os.listdir(attach)

    attachid = FL.attachments.get_list(oid=OID)
    for atch in attachment:
        for p_id in attachid:
            print(p_id['id'])
            ment = attach + '\\' + atch
            print(ment)
            #update_attachment(URL, token, OID, ment, p_id, "PythonTestScript")
