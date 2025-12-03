# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms, script
import math
from System.Windows import Window, Thickness, FontWeights, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Controls import Label, Grid, ColumnDefinition, RowDefinition, ScrollViewer, ScrollBarVisibility, StackPanel
from System.Windows.Media import RotateTransform, Brushes
import data
import acessories
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
import csv
import os

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


# PVC_DENSITY = 1400  # kg/m³
PVC_DENSITY_Map = {
    "Tubo de PVC Esgoto Série Normal": 1400,
    'Tubo de PVC Marrom Soldável':1300,
}

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
    colsName_tooltip = ["Produção", "Transporte", "Instalação", "Demolição", "Transporte", "Processamento \nde resíduos", "Destinação final", "Estágio \nde Recuperação"]
    def __init__(self, data_dict, filepath, hasCreated):
        super(DataGridWindow, self).__init__()
        self.Title = "Resumo dos Tubos"
        self.MaxWidth = 1280
        self.MaxHeight = 720

        scroll_viewer = ScrollViewer()
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.Content = scroll_viewer
        # scroll_viewer.Content = grid
        
        stack = StackPanel()
        stack.Margin = Thickness(10)
        stack.VerticalAlignment = VerticalAlignment.Top
        scroll_viewer.Content = stack

        grid = self._make_baseTable(data_dict)
        stack.Children.Add(grid)
        
        label = Label()
        label.Margin = Thickness(5, 20, 0, 0)
        label.HorizontalAlignment = HorizontalAlignment.Left
        label.VerticalAlignment = VerticalAlignment.Center
        stack.Children.Add(label)
        if hasCreated:
            label.Content = "Um arquivos foi criado com esses dados em {}".format(filepath)
            label.Foreground = Brushes.Green
        else:
            label.Content = "Não foi possivel criar um arquivo com esses dados"
            label.Foreground = Brushes.Red


        grid = self._build_system_boundaries_table()
        stack.Children.Add(grid)

    
    def _make_baseTable(self, data_dict):
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
        cur_col = 2
        for n, fase in enumerate(DataGridWindow.colsName):
            label = self.create_vertical_header(self.colsName_tooltip[n])
            Grid.SetRow(label, current_row)
            Grid.SetColumn(label, cur_col)
            grid.Children.Add(label)
            cur_col+=1
        current_row+=1

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
        
        grid.RowDefinitions.Add(RowDefinition())
        return grid

    def create_vertical_header(self, text):
        tb = Label()
        tb.Width = 120
        tb.Height = 100
        tb.Content = text
        tb.LayoutTransform = RotateTransform(270)
        tb.Margin = Thickness(5)
        tb.HorizontalAlignment = HorizontalAlignment.Stretch
        tb.VerticalAlignment = VerticalAlignment.Center
        return tb

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

    def _build_system_boundaries_table(self):
        grid = Grid()

        # Define colunas
        for col in range(3):  # Estágio, Módulo, Declarado
            grid.ColumnDefinitions.Add(ColumnDefinition())

        # Dados da tabela
        rows = [
            ("Estágio de Produção", "A1 - Fornecimento de matéria-prima"),
            ("", "A2 - Transporte"),
            ("", "A3 - Fabricação"),
            ("Estágio de Construção", "A4 - Transporte"),
            ("", "A5 - Instalação"),
            ("Estágio de Uso", "B1 - Uso"),
            ("", "B2 - Manutenção"),
            ("", "B3 - Reparo"),
            ("", "B4 - Substituição"),
            ("", "B5 - Reforma"),
            ("", "B6 - Consumo de energia operacional"),
            ("", "B7 - Consumo de água operacional"),
            ("Estágio de Fim de Vida", "C1 - Desconstrução/Demolição"),
            ("", "C2 - Transporte"),
            ("", "C3 - Processamento de resíduos"),
            ("", "C4 - Destinação final"),
            ("Estágio de Recuperação", "D - Benefícios e cargas além dos limites do sistema"),
        ]

        # Adiciona linhas
        for i in range(len(rows)):
            grid.RowDefinitions.Add(RowDefinition())

        # Preenche a grade
        for i, (stage, module) in enumerate(rows):
            # Estágio
            lbl_stage = Label()
            lbl_stage.Content = stage
            lbl_stage.FontWeight = FontWeights.Bold if stage else FontWeights.Normal
            lbl_stage.Margin = Thickness(5)
            lbl_stage.VerticalAlignment = VerticalAlignment.Center
            grid.Children.Add(lbl_stage)
            Grid.SetRow(lbl_stage, i)
            Grid.SetColumn(lbl_stage, 0)

            # Módulo
            lbl_module = Label()
            lbl_module.Content = module
            lbl_module.Margin = Thickness(5)
            grid.Children.Add(lbl_module)
            Grid.SetRow(lbl_module, i)
            Grid.SetColumn(lbl_module, 1)

        return grid

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
    TotalMass = resultVal[VOLNAME] * PVC_DENSITY_Map[resultKey]

    for (key, val) in finalDataDecode.items():
        toAddVal = resultVal[key]
        _help_set_data_on_dict(metricsData[val], resultKey, toAddVal, DECODE_FINALDATA_UNIT[key])
    
    _help_set_data_on_dict(metricsData[FINAL_DATA_MASS_KEY], resultKey, TotalMass, DECODE_FINALDATA_UNIT[FINAL_DATA_MASS_KEY])

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


def write_csv(filename, data_dict):
    with open(filename, mode="w") as f:
        writer = csv.writer(f)

        header = ["Metrica", "Unidade"] + DataGridWindow.colsName
        writer.writerow(header)

        for key, data_val in data_dict.items():
            totals = data_val["total"]
            row = [key, data_val["unit"]]
            for fase in DataGridWindow.colsName:
                val = totals.get(fase, 0)
                row.append("{:.2f}".format(val) if isinstance(val, float) else str(val))
            writer.writerow(row)

            for subKey, subVal in data_val["data"].items():
                row = ["{}:".format(subKey), data_val["unit"]]
                for fase in DataGridWindow.colsName:
                    val = subVal.get(fase, 0)
                    row.append("{:.2f}".format(val) if isinstance(val, float) else str(val))
                writer.writerow(row)
hasCreate = False
filepath = ''
try:
    filename = 'Data.csv'
    
    desktop = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    filepath = os.path.join(desktop, filename)
    write_csv(filepath, finalData)
    hasCreate = True
    import subprocess

    subprocess.run(['start', '', filepath], shell=True)

except Exception:
    pass

window = DataGridWindow(finalData, filepath, hasCreate)
window.ShowDialog()
