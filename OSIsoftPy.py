//  Copyright 2016  OSIsoft, LLC
// 
//  Licensed under the Apache License, Version 2.0 (the "License");
//  you may not use this file except in compliance with the License.
//  You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
//  Unless required by applicable law or agreed to in writing, software
//  distributed under the License is distributed on an "AS IS" BASIS,
//  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//  See the License for the specific language governing permissions and
//  limitations under the License.

import sys
import clr

sys.path.append(r'C:\Program Files\PIPC\AF\PublicAssemblies\4.0')  
clr.AddReference('OSIsoft.AFSDK')

from OSIsoft.AF import *  
from OSIsoft.AF.PI import *
from OSIsoft.AF.Search import * 
from OSIsoft.AF.Asset import *  
from OSIsoft.AF.Data import *  
from OSIsoft.AF.Time import *  
from OSIsoft.AF.UnitsOfMeasure import *

class OSIsoftPy(object):

    ## CONNECT TO PI SERVER
    def connect_to_Server(serverName):  
        piServers = PIServers()  
        global piServer  
        piServer = piServers[serverName]                                                    #Write PI Server Name
        piServer.Connect(False)                                                             #Connect to PI Server
        print ('Connected to server: ' + serverName)
        
    ## CONNECT TO AF SERVER AND PRINT ATTRIBUTE VALUE
    def connect_to_AF(AFserverName, Database, Tech, Plant, Unit, Attribute): 
        afServers = PISystems()  
        afServer = afServers[AFserverName]                                                  #Write AF Server Name
        afServer.Connect()                                                                  #Connect to AF Server
        DB = afServer.Databases.get_Item(Database)                                          #Define architecture
        element = DB.Elements.get_Item(Tech).Elements.get_Item(Plant).Elements.get_Item(Plant + " " + Unit)
        attribute = element.Attributes.get_Item(Attribute)
        attval = attribute.GetValue()
        print ('Element Name: {0}'.format(element.Name))                                    #Print Attributr Value
        print ('Attribute Name: {0} \nValue: {1} \nUOM: {2}'.format(attribute.Name, attval.Value, attribute.DefaultUOM))

    ## WRITE TAG VALUE IN PI TAG
    def write_tag(tagname, value, datetime):  
        writept = PIPoint.FindPIPoint(piServer, tagname)                                    #Select PI Server and Tag name
        val = AFValue(value, AFTime(datetime))                                              #Select Value and Timestamp
        writept.UpdateValue(val, AFUpdateOption.Replace, AFBufferOption.BufferIfPossible)   #Write value
        print ('Tag "' + tagname + '" updated.')                                            #Print Tag Name updated

    ## GET SNAPSHOT TAG VALUE
    def get_tag_snapshot(tagname):  
        tag = PIPoint.FindPIPoint(piServer, tagname)  
        lastData = tag.Snapshot()                                                           #Get Snapshot
        print ('Last Value in PI Tag ' + tagname + ' = ' + str(lastData))                   #Print Tag Value
        return lastData.Value, lastData.Timestamp

    ## GET SAMPLED VALUES
    def sampled_values(tagname, initdate, enddate, span):
        tag = PIPoint.FindPIPoint(piServer, tagname)                                        #Select PI Server and Tag name
        timerange = AFTimeRange(initdate, enddate)                                          #Select Time Range (Osisoft PI format)
        sampled = tag.InterpolatedValues(timerange, AFTimeSpan.Parse(span), '', False)      #Get Sampled Values (IMPORTANT: Define Span)
        print('\nShowing sampled values in PI Tag {0}'.format(tagname))                     #Print Sampled Values
        for event in sampled:  
            print('{0} value: {1}'.format(event.Timestamp.LocalTime, event.Value)) 

    ## GET RECORDED VALUES
    def recorded_values(tagname, initdate, enddate):
        tag = PIPoint.FindPIPoint(piServer, tagname)                                        #Select PI Server and Tag name
        timerange = AFTimeRange(initdate, enddate)                                          #Select Time Range (Osisoft PI format)
        recorded = tag.RecordedValues(timerange, AFBoundaryType.Inside, "", False)          #Get Recorded Values in Time Range
        print('\nShowing recorded values in PI Tag {0}'.format(tagname))                    #Print Recorded Values
        for event in recorded:  
            print('{0} value: {1}'.format(event.Timestamp.LocalTime, event.Value)) 
            
    ## FIND TAGS
    def find_tags(mask):
        points = PIPoint.FindPIPoints(piServer, mask, None, None)                           #Select PI Server and Mask
        points = list(points)
        return [print(i.get_Name()) for i in points]                                        #Print coincidences

    ## DELETE TAG VALUES IN PI TAG
    def delete_values(tagname, initdate, enddate):
        deleteval = PIPoint.FindPIPoint(piServer, tagname)                                  #Select PI Server and Tag Name
        timerange = AFTimeRange(initdate, enddate)                                          #Select Time Range (Osisoft PI format)
        recorded = deleteval.RecordedValues(timerange, AFBoundaryType.Inside, "", False)    #Get Recorded Values in Time Range
        deleteval.UpdateValues(recorded, AFUpdateOption.Remove)                             #Delete Recorded Values in Time Range
        print ('\nTag Values selected of PI Tag "' + tagname + '" have been deleted.')      #Print Tag Name updated
       
    ## UPDATE AF ATTRIBUTES
    def update_AF_attribute(AFserverName, Database, Elem, Tech, Plant, Unit, Attribute1, Value1, Attribute2, Value2):
        afServers = PISystems()
        afServer = afServers[AFserverName]
        DB = afServer.Databases.get_Item(Database)
        element = DB.Elements.get_Item(Elem).Elements.get_Item(Tech).Elements.get_Item(Plant).Elements.get_Item(Unit)
        attribute = element.Attributes.get_Item(Attribute1)
        attribute.SetValue(AFValue(Value1))
        attribute = element.Attributes.get_Item(Attribute2)
        attribute.SetValue(AFValue(Value2))
