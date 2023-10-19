# Prácticas de SAP 2000. 
# Análisis de Estructuras. Grado en Ingeniería Civil. 
# Autor: Alejandro Martínez Castro (amcastro@ugr.es).
# Departamento de Mecánica de Estructuras e Ingeniería Hidráulica. Universidad de Granada. 

#Practica 2. 
#Uso de la API de SAP 2000. Actualizado a 19/10/2023 para SAP2000 23, con paquete comtypes.
# Adaptación a la nueva sintaxis de la API de SAP 2000 para creación de frames,
# etc ('err' se coloca al final ahora)

import os
import sys
import comtypes.client

# Parámetros
l_pilar_1 = 10  #Longitud del pilar_1
l_pilar_2 = 2   #Longitud de pilar 2
l_viga = 10     #Longitud de viga
elast = 2.1E8   #Módulo de elasticidad en kN / m^2

ang = 30        #Ángulo del apoyo no concordante

canto_pilar_1 = 50E-3
ancho_pilar_1 = 50E-3
espesor_alas_pilar_1 = 4E-3
espesor_almas_pilar_1 =4E-3

canto_pilar_2 = 50E-3
ancho_pilar_2 = 50E-3
espesor_alas_pilar_2 = 3E-3
espesor_almas_pilar_2 =3E-3

canto_viga = 100E-3
ancho_viga = 50E-3
espesor_alas_viga = 5E-3
espesor_almas_viga =5E-3

q_viga = 10 #Carga uniforme en kN/m en la viga

#PASO 1: 
#Crear el objeto SAP2000
#========================
#    La variable "ObjetoSap" podemos definirla con otro nombre si deseamos.
#    Es un objeto de tipo SAP2000, creado mediante comtypes, que es un Módulo de python.
ObjetoSap = comtypes.client.CreateObject('SAP2000v1.Helper')
ObjetoSap = ObjetoSap.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

ObjetoSap = ObjetoSap.CreateObjectProgID("CSI.SAP2000.API.SapObject")

#PASO 2:
#Iniciar la aplicación Sap2000
#=============================
#    Este comando es el responsable de que "súbtitamente" al ejecutar este script se arranque el programa SAP2000
ObjetoSap.ApplicationStart()

#PASO 3:
#crear el objeto "ModeloSap"
#===========================
#    En este paso se va a crear un objeto, denominado (arbitrariamente) "ModeloSap". 
#    Los objetos son variables extendidas, y forman la base de la programación orientada a objetos.
#    
#    "SapModel" es un miembro fijo y programado en la API de SAP. Para invocarlo, se añade un punto al nombre que
#    hemos definido para el objeto. En este caso, "ObjetoSap". 
#    Nótese que si en vez de ObjetoSap lo hubiésemos denominado "ObjetoSAP2000", entonces hubiésemos escrito
#    ModeloSap = ObjetoSAP2000.SapModel

#    Note la diferencia entre los nombres arbitrarios con los que podemos definir los objetos,
#    y los miembros dentro de una clase, que son nombres fijos, programados dentro de la API de SAP2000. 
ModeloSap = ObjetoSap.SapModel

#PASO 4:
#Inicializar un nuevo modelo
#============================
#    Es equivalente a hacer clik en "Nuevo Modelo" en el entorno de ventanas de SAP2000
ModeloSap.InitializeNewModel()

#PASO 5:
#Crear un nuevo modelo en blanco
#===============================
#    La variable "err" puede cambiar de nombre. Contiene un valor 0 si no hay error, y 1 si se produce un error. 
err = ModeloSap.File.NewBlank()

#PASO 6:
#Cambiar el sistema de unidades
#==============================
#    El número 6 es el correspondiente al sistema kN_m_C
#    Otras opciones son: (consultar la ayuda del comando SetPresentUnits)
#         lb_in_F = 1
#         lb_in_F = 1
#         lb_ft_F = 2
#         kip_in_F = 3
#         kip_ft_F = 4
#         kN_mm_C = 5
#         kN_m_C = 6
#         kgf_mm_C = 7
#         kgf_m_C = 8
#         N_mm_C = 9
#         N_m_C = 10
#         Ton_mm_C = 11
#         Ton_m_C = 12
#         kN_cm_C = 13
#         kgf_cm_C = 14
#         N_cm_C = 15
#         Ton_cm_C = 16

kN_m_C = 6 
err = ModeloSap.SetPresentUnits(kN_m_C)

#PASO 7:
#Definir las propiedades del material
#====================================
#    Primero se asigna un códgo numérico al material. En este caso, el 2. También una etiqueta, en este caso 'ACERO'
MATERIAL_ACERO = 2
err = ModeloSap.PropMaterial.SetMaterial('ACERO', MATERIAL_ACERO)

#    En segundo lugar, se asigna el comportamiento, mediante PropMaterial.SetMPIsotropic
#    Nótese que el parámetro "elast" se definió al principio como variable de Python.
err = ModeloSap.PropMaterial.SetMPIsotropic('ACERO', elast, 0.3, 0.0000055)

#PASO 8:
#Definir las propiedades de la secciones tubulares rectangulares
#===============================================================
#   La sintaxis viene definida por SetTube. Consultar en la ayuda de la API.
#   SetTube('etiqueta_seccion','etiqueta_material',canto,ancho,espesor_alas,espesor_almas)
err = ModeloSap.PropFrame.SetTube('pilar_1', 'ACERO', canto_pilar_1, ancho_pilar_1, espesor_alas_pilar_1, espesor_almas_pilar_1)
err = ModeloSap.PropFrame.SetTube('pilar_2', 'ACERO', canto_pilar_2, ancho_pilar_2, espesor_alas_pilar_2, espesor_almas_pilar_2)
err = ModeloSap.PropFrame.SetTube('viga', 'ACERO', canto_viga, ancho_viga, espesor_alas_viga, espesor_almas_viga)

#PASO 9:
#Añadir objetos de tipo frame mediante coordenadas
#=================================================
#    Asignación de nombres a los Frames: se dejan libres para que SAP asigne correlativamente 1,2,3...
FrameName1 = ' '
FrameName2 = ' '
FrameName3 = ' '

#    Coordenada Z del punto extremo de pilar_2: en función de los parámetros l_pilar_1 y l_pilar_2.
l3 = l_pilar_1 - l_pilar_2

#    Generación de los elementos Frame mediante coordenadas. 
[FrameName1, err] = ModeloSap.FrameObj.AddByCoord(0, 0, 0, 0, 0, l_pilar_1, FrameName1, 'pilar_1', '1', 'Global')
[FrameName2, err] = ModeloSap.FrameObj.AddByCoord(0, 0, l_pilar_1, l_viga, 0, l_pilar_1, FrameName2, 'viga', '2', 'Global')
[FrameName3, err] = ModeloSap.FrameObj.AddByCoord(l_viga, 0, l_pilar_1, l_viga, 0, l3, FrameName3, 'pilar_2', '3', 'Global')

#    Refrescar vista, actualizar Zoom (En este momento se ve el pórtico en pantalla SAP)
err = ModeloSap.View.RefreshView(0, False)

#PASO 10:
#Definición de un sistema de coordenadas nodal, girado alfa grados, para el extremo de pilar_2
#=============================================================================================
#    Primero, se selecciona el punto extremo de pilar_2
PointName1 = ' '
PointName2 = ' '
[PointName1, PointName2, err] = ModeloSap.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
#    Después, se gira el sistema nodal 30 grados en sentido contrario a las agujas del reloj.
#    SetLocalAxes('Etiqueta_del_punto',giro_x,giro_y,giro_z,codigo=0 para objetos)
alfa = -ang;
err = ModeloSap.PointObj.SetLocalAxes(PointName2,0,alfa,0,0)

#PASO 11:
#Definición de las condiciones de apoyo.
#======================================
PointName1 = ' '
PointName2 = ' '
Restraint = [True, True, True, True, True, True]
[PointName1, PointName2,err] = ModeloSap.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
err = ModeloSap.PointObj.SetRestraint(PointName1, Restraint)

Restraint = [False, False, True, False, False, False]
[PointName1, PointName2,err] = ModeloSap.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
err = ModeloSap.PointObj.SetRestraint(PointName2, Restraint)


#PASO 12:
#Definición del Load Pattern denominado 'CARGA'
#==============================================
#     El tipo de carga del Pattern es un parámetro de la función Add para LoadPatterns. 
#        LTYPE_DEAD = 1
#        LTYPE_SUPERDEAD = 2
#        LTYPE_LIVE = 3
#        LTYPE_REDUCELIVE = 4
#        LTYPE_QUAKE = 5
#        LTYPE_WIND= 6
#        LTYPE_SNOW = 7
#        LTYPE_OTHER = 8
#        LTYPE_MOVE = 9
#        LTYPE_OTHER = 8
#        etc (ver más opciones en Add para Load Patterns)
LTYPE = 3
err = ModeloSap.LoadPatterns.Add('CARGA', LTYPE , 0, True)

#PASO 13:
#Asignar la carga distribuida q_viga en 'CARGA' en dirección de la gravedad
#==========================================================================
#    La variable 'FrameName2' es la viga
err = ModeloSap.FrameObj.SetLoadDistributed(FrameName2, 'CARGA', 1, 10, 0, 1, q_viga, q_viga)

#PASO 14:
#Guardar el modelo en un fichero
#================================
#    Se pueden introducir otros path con operaciones del módulo de Python 'os' (Operative System)
APIPath = 'C:\API'
if not os.path.exists(APIPath):
 try:
     os.makedirs(APIPath)
 except OSError:
     pass
err = ModeloSap.File.Save(APIPath + os.sep + 'Practica2_API_Python.sdb')

#refresh view, update (initialize) zoom
err = ModeloSap.View.RefreshView(0, False)

#PASO 15:
#Seleccionar grados de libertad activos para Pórtico Plano
#==========================================================
DOF = [True,False,True,False,True,False] #Corresponde a UX,UY,UZ,RX,RY,RZ; 
err = ModeloSap.Analyze.SetActiveDOF(DOF)

#PASO 16:
#Seleccionar los casos de carga que se van a analizar.
#=====================================================
#    En este caso sólo se va a analizar el caso 'CARGA'
err = ModeloSap.Analyze.SetRunCaseFlag('DEAD',False)
err = ModeloSap.Analyze.SetRunCaseFlag('MODAL',False)

#PASO 17:
#Correr el modelo
#=================
err = ModeloSap.Analyze.RunAnalysis()



#PASO 18:
#Obtener resultados
#====================
[PointName1, PointName2, err] = ModeloSap.FrameObj.GetPoints(FrameName3, PointName1, PointName2)

NumberResults = 0
Obj = []
Elm = []
ACase = []
StepType = []
StepNum = []
U1 = []
U2 = []
U3 = []
R1 = []
R2 = []
R3 = []
ObjectElm = 0;
err = ModeloSap.Results.Setup.DeselectAllCasesAndCombosForOutput()
err = ModeloSap.Results.Setup.SetCaseSelectedForOutput('CARGA')

[NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, err] = ModeloSap.Results.JointDispl(PointName2, ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)

print ('El desplazamiento U1 en el plano inclinado ', ang , 'grados, es: ', U1, ' metros')

#close Sap2000
err = ObjetoSap.ApplicationExit(False)
ModeloSap = 0;
ObjetoSap = 0;
