# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import math
from System.Windows import Window, Thickness, FontWeights, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Controls import Label, StackPanel, Grid, ColumnDefinition, RowDefinition, ScrollViewer, ScrollBarVisibility
import data
import acessories
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

FEAT_TO_METER_MUTIPLY = 0.304785126

PIPE_CATEGORY = [
    DB.BuiltInCategory.OST_PipeCurves,
]

ASSESSORIES_CATEGORY = [
    DB.BuiltInCategory.OST_PipeAccessory, 
    DB.BuiltInCategory.OST_PipeConnections,
    DB.BuiltInCategory.OST_PipeFitting, 
]

LENGHTNAME = 'lenght'
VOLNAME = 'vol'


PVC_DENSITY = 1400  # kg/m³
PVC_DENSITY_Map = {
    "UnMEP_PVC Branco - Série Normal": 1320,
}

EnergiaPerMeter = 1.2
CO2PerKg = 2.5
WatherPerKg = .57

finalDataDecode = {
    LENGHTNAME: "Comprimento de tubulação",
    VOLNAME: "Volume de tubulação",
}
FINAL_DATA_MASS_KEY= "Massa de tubulação"

DECODE_FINALDATA_UNIT = {
    LENGHTNAME: 'M',
    VOLNAME: 'M³',
    FINAL_DATA_MASS_KEY: 'Kg/M³'
}

def _get_elements_of_category(getCategories):
    results = None
    for category in getCategories:
        searchResult = DB.FilteredElementCollector(revit.doc)\
              .OfCategory(category)\
              .WhereElementIsNotElementType()\
              .ToElements()
        if (results is not None):
             results.AddRange(searchResult)
        else:  
            results = searchResult
    return results

def coletar_tubos():
    tubos = _get_elements_of_category(PIPE_CATEGORY)
    data = {}
    # totalVol = 0
    # totalLenght = 0

    for tubo in tubos:
        try:
            lenghtParam = tubo.LookupParameter("Comprimento")
            outer_diamParam = tubo.LookupParameter('Diâmetro externo')
            inner_diamParam = tubo.LookupParameter('Diâmetro interno')
            tubSegTypeParam = tubo.LookupParameter('Segmento de tubulação')

            if (lenghtParam and outer_diamParam and inner_diamParam and tubSegTypeParam):
                instanciaTubo = revit.doc.GetElement(tubo.GetTypeId())
                descTypeParam = instanciaTubo.LookupParameter('Descrição')
                if (descTypeParam):
                    name = descTypeParam.AsValueString()
                else:
                    name = tubSegTypeParam.AsValueString()
                lenght = lenghtParam.AsDouble() * FEAT_TO_METER_MUTIPLY
                outer_diam = outer_diamParam.AsDouble() * FEAT_TO_METER_MUTIPLY
                inner_diam = inner_diamParam.AsDouble() * FEAT_TO_METER_MUTIPLY
                # dataKey = '{} {}'.format(name, outer_diamParam.AsValueString())
                dataKey = '{}'.format(name)
                if ( dataKey not in data):
                    data[dataKey] = {
                        LENGHTNAME: 0,
                        VOLNAME: 0,
                    }
                data[dataKey][LENGHTNAME] += lenght
                tuboVol = math.pi * (outer_diam*outer_diam - inner_diam*inner_diam) * lenght / 4
                data[dataKey][VOLNAME] += tuboVol
        except Exception as e:
            pass

    return data
    # return [totalVol, totalLenght]

def coletar_assesorios():
    tubos = _get_elements_of_category(ASSESSORIES_CATEGORY)
    data = {}
    # totalVol = 0
    # totalLenght = 0
    for tubo in tubos:
        try:
            descriptionParam = tubo.LookupParameter("UnMEP: Descrição do Material")

            if (descriptionParam):
                key = descriptionParam.AsValueString()
                if(key not in data):
                    data[key] = 0
                data[key] += 1
            else:
                instanciaTubo = revit.doc.GetElement(tubo.GetTypeId())
                descTypeParam = instanciaTubo.LookupParameter("UnMEP: Descrição do Material")
                if (descTypeParam):
                    key = descTypeParam.AsValueString()
                    if(key not in data):
                        data[key] = 0
                    data[key] += 1
                # else:
                #     print(descriptionParam, descTypeParam)
        except Exception as e:
            # print('error', e)
            pass

    return data
    # return [totalVol, totalLenght]

class DataGridWindow(Window):
    colsName = ["A1-A3", "A4", "A5", "C1", "C2", "C3", "C4", "D"]
    def __init__(self, data_dict):
        super(DataGridWindow, self).__init__()
        self.Title = "Resumo dos Tubos"
        self.MaxWidth = 1280
        self.MaxHeight = 720

        # Create grid
        grid = Grid()
        grid.Margin = Thickness(18)

        colDef = ColumnDefinition()
        colDef.SetValue(ColumnDefinition.WidthProperty, GridLength(8.0, GridUnitType.Star))
        # colDef.width = GridLength(8, GridUnitType.Star)
        grid.ColumnDefinitions.Add(colDef)


        colDef = ColumnDefinition()
        # colDef.width = GridLength(1, GridUnitType.Star)
        colDef.SetValue(ColumnDefinition.WidthProperty, GridLength(2.0, GridUnitType.Star))
        grid.ColumnDefinitions.Add(colDef)

        for _ in range(len(DataGridWindow.colsName)):
            colDef = ColumnDefinition()
            # colDef.width = GridLength(1, GridUnitType.Star)
            colDef.SetValue(ColumnDefinition.WidthProperty, GridLength(1.0, GridUnitType.Star))
            grid.ColumnDefinitions.Add(colDef)
        
        current_row = 0

        grid.RowDefinitions.Add(RowDefinition())
        grid.Children.Add(self._make_label("Metrica", current_row, 0, bold=True))
        grid.Children.Add(self._make_label("Unidade", current_row, 1, bold=True))
        cur_col = 2
        for fase in DataGridWindow.colsName:
            grid.Children.Add(self._make_label(fase, current_row, cur_col, bold=True))
            cur_col += 1
        current_row+=1

        for (key, data_val) in data_dict.items():
            grid.RowDefinitions.Add(RowDefinition())
            grid.Children.Add(self._make_label("{}".format(key), current_row, 0, bold=True))
            grid.Children.Add(self._make_label("{}".format(data_val['unit']), current_row, 1, bold=True))
            totals = data_val['total']
            cur_col = 2
            for fase in DataGridWindow.colsName:
                total = 0
                if (fase in totals):
                    total = totals[fase]
                total_str = "{:.2f}".format(total) if type(total) == float else "{}".format(total)
                grid.Children.Add(self._make_label(total_str, current_row, cur_col))
                cur_col += 1
            current_row+=1

            for (subKey, susbVal) in data_val['data'].items():
                grid.RowDefinitions.Add(RowDefinition())
                grid.Children.Add(self._make_label("\t{}:".format(subKey), current_row, 0))
                
                grid.Children.Add(self._make_label("{}".format(data_val['unit']), current_row, 1, bold=True))
                cur_col = 2
                
                for fase in DataGridWindow.colsName:
                    val = 0
                    if (fase in susbVal):
                        val = susbVal[fase]
                    val_str = "{:.2f}".format(val) if type(val) == float else "{}".format(val)
                    grid.Children.Add(self._make_label(val_str, current_row, cur_col))
                    cur_col += 1
                current_row+=1

        scroll_viewer = ScrollViewer()
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.Content = grid

        self.Content = scroll_viewer

    def _make_label(self, text, row, col, bold=False):
        label = Label()
        label.Content = text
        label.Margin = Thickness(5)
        label.HorizontalAlignment = HorizontalAlignment.Left
        label.VerticalAlignment = VerticalAlignment.Center
        if bold:
            label.FontWeight = FontWeights.Bold
        Grid.SetRow(label, row)
        Grid.SetColumn(label, col)
        return label

def _help_set_data_on_dict(dict, resultKey, val_to_add, baseUnit):
    dict['total'] += val_to_add
    dict["Unit"] = baseUnit # :(
    if resultKey not in dict['data']:
        dict['data'][resultKey] = 0
    dict['data'][resultKey] += val_to_add

resultados = coletar_tubos()

assesorios = coletar_assesorios()

def map_pipes_to_accessories():
    pipes = _get_elements_of_category(PIPE_CATEGORY)

    result = {}

    for pipe in pipes:
        try:
            instanciaTubo = revit.doc.GetElement(pipe.GetTypeId())
            desc_param = instanciaTubo.LookupParameter('Descrição')
            if not desc_param:
                continue
            pipe_desc = desc_param.AsString()

            # Initialize list for this pipe type
            if pipe_desc not in result:
                result[pipe_desc] = {}

            pipe_connector_set = pipe.ConnectorManager.Connectors
            for conn in pipe_connector_set:
                for ref in conn.AllRefs:
                    elem = ref.Owner
                    acc_desc = elem.LookupParameter("UnMEP: Descrição do Material")
                    if acc_desc:
                        val = acc_desc.AsValueString()
                        if (val not in result[pipe_desc]):
                            result[pipe_desc][val] = 0
                        result[pipe_desc][val] += 1
                        
                    # else:
                    #     print(elem.Category, elem.Category.Id.IntegerValue)
        except Exception as e:
            # print("Erro:", e)
            pass

    return result

pipe_acessories_map = map_pipes_to_accessories()

# import os
# filename = 'assesorios.txt'
# # print(assesorios)
# try:
#     desktop = os.path.join(os.path.expanduser("~"), "Desktop")
#     filepath = os.path.join(desktop, filename)
#     with open(filepath, "w") as f:
#         f.write("Relatório de Acessórios\n")
#         f.write("========================\n\n")
#         for key, value in assesorios.items():
#             f.write("{}: {}\n".format(key, value))
#     os.startfile(filepath) 
# except Exception as e:
#     pass

metricsData = {
    "Comprimento de tubulação": {
        "data": {},
        "total": 0,
    },
    "Volume de tubulação": {
        "data": {},
        "total": 0,
    },
    "Massa de tubulação": {
        "data": {},
        "total":0
    },
    "Acessorios": {
        'total': 0,
        "Unit" : 'Unidade',
        "data" : {}
    }
}


for resultKey, resultVal in resultados.items():
    # TotalMass = resultados.volTotal * FEAT_TO_METER_MUTIPLY * PVC_DENSITY
    TotalMass = resultVal[VOLNAME] * PVC_DENSITY

    for (key, val) in finalDataDecode.items():
        toAddVal = resultVal[key]
        _help_set_data_on_dict(metricsData[val], resultKey, toAddVal, DECODE_FINALDATA_UNIT[key])
    
    _help_set_data_on_dict(metricsData[FINAL_DATA_MASS_KEY], resultKey, TotalMass, DECODE_FINALDATA_UNIT[FINAL_DATA_MASS_KEY])

    # window = DataGridWindow({
    #     "Comprimento total": resultVal["totalLenght"],
    #     "Volume total": resultVal["totalVol"],
    #     "Peso total": resultVal,
    # })
acessoriosDict = metricsData['Acessorios']
for key, value in assesorios.items():
    acessoriosDict['total'] += value
    acessoriosDict['data'][key] = value


finalData = {}

for peformaceName, peformaceData in data.dataPerKg.items():
    finalData[peformaceName] = {
        'unit': peformaceData['unit'],
        "total": {},
        "data": {}
    }
    curData = finalData[peformaceName]['data']
    for pipeData in metricsData[FINAL_DATA_MASS_KEY]['data']:
        if (pipeData not in peformaceData['data']):
            print('{} not in database'.format(pipeData))
            continue
        if pipeData not in finalData:
            curData[pipeData] = {}
        pipePeformaceData = peformaceData['data'][pipeData]
        for estage in pipePeformaceData:
            if estage not in curData[pipeData]:
                curData[pipeData][estage] = 0
            curData[pipeData][estage] += metricsData[FINAL_DATA_MASS_KEY]['data'][pipeData] * pipePeformaceData[estage]

        if (pipeData in pipe_acessories_map):
            curData['{} - Acessorios'.format(pipeData)] = {}
            curAccessoryData = curData['{} - Acessorios'.format(pipeData)]
            for accessory in pipe_acessories_map[pipeData]:
                if (accessory in acessories.dataAcessories):
                    accessoryMass = acessories.dataAcessories[accessory] * pipe_acessories_map[pipeData][accessory]
                    for estage in pipePeformaceData:
                        if estage not in curAccessoryData:
                            curAccessoryData[estage] = 0
                        curAccessoryData[estage] += accessoryMass * pipePeformaceData[estage]

    peformaceKeySet = set().union(*[finalData[peformaceName]['data'][a].keys() for a in finalData[peformaceName]['data']])

    finalData[peformaceName]['total'] = {key: 0 for key in peformaceKeySet}
    curData = finalData[peformaceName]['total']
    for perfData in finalData[peformaceName]['data']:
        for key in peformaceKeySet:
            curData[key] += finalData[peformaceName]['data'][perfData][key]

window = DataGridWindow(finalData)
window.ShowDialog()
# else:
#     forms.alert("Nenhum elemento PVC com o parâmetro encontrado.", title="Resultado da Busca", warn_icon=True)
