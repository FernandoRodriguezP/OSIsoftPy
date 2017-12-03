#region Copyright
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
#endregion

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

    ##FUNCION PARA REALIZAR LA CONEXION CON UN SERVIDOR PI
    def connect_to_Server(serverName):  
        piServers = PIServers()  
        global piServer  
        piServer = piServers[serverName]  
        piServer.Connect(False)
        print ('Conectado al servidor: ' + serverName)

    ##FUNCION PARA ESCRIBIR UN VALOR EN UN TAG
    def write_tag(tagname, value, datetime):  
        writept = PIPoint.FindPIPoint(piServer, tagname)  #Seleccionas el servidor y el tag deseado
        val = AFValue(value, AFTime(datetime)) #Seleccionas el valor y el timestamp
        writept.UpdateValue(val, AFUpdateOption.Replace, AFBufferOption.BufferIfPossible)   #Escribes el valor
        print ('Tag "' + tagname + '" actualizado.')

    ##FUNCION PARA IMPRIMIR EL ULTIMO VALOR ALMACENADO DE UN TAG
    def get_tag_snapshot(tagname):  
        tag = PIPoint.FindPIPoint(piServer, tagname)  
        lastData = tag.Snapshot()
        print ('Ultimo valor almacenado en el PI Tag ' + tagname + ' = ' + str(lastData))
        return lastData.Value, lastData.Timestamp

    ##FUNCION PARA MOSTRAR LOS VALORES GRABADOS EN UN TAG DENTRO DE UN RANGO TEMPORAL (¡CON INTERVALO!)
    def sampled_values(tagname, initdate, enddate, span):
        tag = PIPoint.FindPIPoint(piServer, tagname)   
        timerange = AFTimeRange(initdate, enddate)
        sampled = tag.InterpolatedValues(timerange, AFTimeSpan.Parse(span), '', False)  
        print('\nMostrando valores almacenados en el PI Tag {0}'.format(tagname))  
        for event in sampled:  
            print('{0} value: {1}'.format(event.Timestamp.LocalTime, event.Value)) 

    ##FUNCION PARA BUSCAR TAGS
    def find_tags(mask):
        points = PIPoint.FindPIPoints(piServer, mask, None, None)
        points = list(points) #casting it to a Python list
        return [print(i.get_Name()) for i in points]

    ##FUNCION PARA MOSTRAR TODOS LOS VALORES GRABADOS EN UN TAG DENTRO DE UN RANGO TEMPORAL
    def recorded_values(tagname, initdate, enddate):
        tag = PIPoint.FindPIPoint(piServer, tagname)   
        timerange = AFTimeRange(initdate, enddate)  
        recorded = tag.RecordedValues(timerange, AFBoundaryType.Inside, "", False)  
        print('\nMostrando valores almacenados en el PI Tag {0}'.format(tagname))  
        for event in recorded:  
            print('{0} value: {1}'.format(event.Timestamp.LocalTime, event.Value)) 

    ##FUNCION PARA BORRAR VALORES EN UN TAG
    def delete_values(tagname, initdate, enddate):
        deleteval = PIPoint.FindPIPoint(piServer, tagname)  #Seleccionas el servidor y el tag deseado
        timerange = AFTimeRange(initdate, enddate)
        recorded = deleteval.RecordedValues(timerange, AFBoundaryType.Inside, "", False) 
        deleteval.UpdateValues(recorded, AFUpdateOption.Remove)
        print ('\nValor seleccionado del PI Tag "' + tagname + '" eliminado.')

##FUNCION PARA ACCEDER A UNA BASE DE DATOS AF
    def connect_to_AF(AFserverName, Database, Tech, Plant, Unit, Attribute): 
        afServers = PISystems()  
        afServer = afServers[AFserverName]
        afServer.Connect()
        DB = afServer.Databases.get_Item(Database)
        element = DB.Elements.get_Item(Tech).Elements.get_Item(Plant).Elements.get_Item(Plant + " " + Unit)
        #element = DB.Elements.get_Item(Tech).Elements.get_Item(Plant)
        attribute = element.Attributes.get_Item(Attribute)
        attval = attribute.GetValue()
          
        print ('Element Name: {0}'.format(element.Name))  
        print ('Attribute Name: {0} \nValue: {1} \nUOM: {2}'.format(attribute.Name, attval.Value, attribute.DefaultUOM))
