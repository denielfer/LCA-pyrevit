# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms
import math

def buscar_elementos_pvc(parametro_alvo):
    colecao = DB.FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
    elementos_pvc = []

    for elemento in colecao:
        try:
            # Verifica se o elemento tem material PVC
            materiais = elemento.GetMaterialIds(False)
            for mat_id in materiais:
                material = revit.doc.GetElement(mat_id)
                if "PVC" in material.Name.upper():
                    # Verifica se o parâmetro existe
                    param = elemento.LookupParameter(parametro_alvo)
                    if param and param.HasValue:
                        elementos_pvc.append((elemento.Id, param.AsString()))
        except Exception:
            continue

    return elementos_pvc

def getNome(eid):
    if eid and eid.IntegerValue != -1:
        elemento = revit.doc.GetElement(eid)
        if elemento:
            return elemento.ToString()
        return "({}, id: {})".format(elemento.__str__(), eid)
        # return elemento.ToString()

    return "(id: {})".format(eid)


def _get_tubos_and_connections():
    getCategories = [
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeAccessory, 
        DB.BuiltInCategory.OST_PipeConnections,
        DB.BuiltInCategory.OST_PipeFitting, 
    ]
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
    tubos = _get_tubos_and_connections()
    totalVol = 0
    totalLenght = 0

    selectedTubos = []

    for tubo in tubos[:10]:
        # parametros = tubo.Parameters
        # lista = []
        print("__________")
        print("__________")
        print(dir(tubo))
        paramLookupList = ["Diâmetro", "Width", "Length", "Comprimento", "Titulo da página",
                           "Segmento de tubulação", "Área"]
        for paramLook in paramLookupList:
            param = tubo.LookupParameter(paramLook)
            if param:
                print('{}: '.format(paramLook), param.AsValueString())
            else:
                print('{}: '.format(paramLook), param)
        lenght = tubo.LookupParameter("Comprimento")
        outer_diam = tubo.LookupParameter('Diâmetro externo')
        inner_diam = tubo.LookupParameter('Diâmetro interno')
        print("lenght: ", lenght, " / out din:", outer_diam, " / inner din: ", inner_diam)
        # print("lenght: ", dir(lenght), " / out din:", dir(outer_diam), " / inner din: ", dir(inner_diam))

        if (lenght and outer_diam and inner_diam):
            lenght = lenght.AsDouble()
            outer_diam = outer_diam.AsDouble()
            inner_diam = inner_diam.AsDouble()
            totalLenght += lenght
            tuboVol = math.pi * (outer_diam*outer_diam - inner_diam*inner_diam) * lenght / 4
            totalVol += tuboVol
            print("lenght: ", lenght)
        print('PipeType: ', dir(tubo.PipeType))
        print("__________")
        print("__________")


        # for param in parametros: 
        #     eid = param.AsElementId()
        #     if (eid < 0):
        #         continue
        #     nome =  str(eid)# getNome(param.AsElementId())
        #     tipo = param.StorageType
        #     valor = None

        #     if param.HasValue:
        #         if tipo == DB.StorageType.String:
        #             valor = param.AsString()
        #         elif tipo == DB.StorageType.Double:
        #             valor = param.AsDouble()
        #         elif tipo == DB.StorageType.Integer:
        #             valor = param.AsInteger()
        #         elif tipo == DB.StorageType.ElementId:
        #             valor = getNome(eid)

        #     if (nome is not None and tipo is not None ):
        #         lista.append("{} -> {}: {}".format(param, nome, valor))


        # selectedTubos.append((tubo.Id, ';;'.join(lista)))

    return [totalVol, totalLenght]

# Nome do parâmetro que você quer buscar
# parametro_desejado = "Length"
resultados = coletar_tubos()


# Convert feet to meters
FEAT_TO_METER_MUTIPLY = 0.3048
PVC_DENSITY = 1400  # kg/m³

EnergiaPerMeter = 1.2
CO2PerKg = 2.5
WatherPerKg = .57


print()
  
if resultados:
    # TotalMass = resultados.volTotal * FEAT_TO_METER_MUTIPLY * PVC_DENSITY
    TotalMass = resultados[0] * FEAT_TO_METER_MUTIPLY * PVC_DENSITY
    print(resultados, TotalMass)
    # print(dir(resultados), TotalMass)
else:
    forms.alert("Nenhum elemento PVC com o parâmetro encontrado.", title="Resultado da Busca", warn_icon=True)
