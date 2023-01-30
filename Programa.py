#Importamos las librerias necesarias.
import pandas as pd
import os
import time
import warnings
import multiprocessing
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)

current_directory = os.getcwd()
print(current_directory)
#os.chdir('C:/Users/ep_jbarrientost/OneDrive - Colbun S.A/Escritorio/programa version con van anual/')
# Creamos las funciones asociadas.

# Función para pasar una fecha a string.
def marca(encabezados,fecha):
    marcador = 0
    for i in range(len(encabezados)):
        if encabezados[i] == fecha:
            marcador = i
            break
    return marcador

# Función para pasar una fecha a string.
def fecha_to_string(inicio,mes_inicio):

    if 0<mes_inicio <10:
        fecha = str(inicio) + '-0' + str(mes_inicio)
    elif mes_inicio < 0:
        print('No existe mes con número negativo')
        quit()
    else:
        fecha = str(inicio) + '-' + str(mes_inicio)
    return str(fecha)

#Funcion para importar datos.
def importar_datos(planilla):
    hojas = pd.ExcelFile(planilla)
    hojas = hojas.sheet_names
    datos = pd.read_excel(planilla, sheet_name = hojas[0])
    
    return datos

 # Funcion para extraer tabla con información importante del proyecto
def info(datos,columna):
    # Datos importantes
    tabla_info = datos.iloc[0:10,columna]
    proyecto = tabla_info.name
    FP = datos.iloc[3,columna]
    subestacion = datos.iloc[0,columna].strip('S/E').strip()
    potencia_parque = datos.iloc[2,columna]
    capex = (datos.iloc[4,columna]*potencia_parque)/1000 # en MMUSD
    gx_anual = potencia_parque* FP *8.76 
    deg_anual = datos.iloc[8,columna]
    deg_mensual = datos.iloc[9,columna]
    p_suficiencia = datos.iloc[10,columna]
    opex_fijo =  datos.iloc[5,columna]
    terrenos_fijo = datos.iloc[6,columna]
        # Fechas
    horizonte = datos.iloc[7,columna]
    inicio = datos.iloc[1,columna].year
    mes_inicio = datos.iloc[1,columna].month
    fin = inicio + horizonte +1
    mes_fin = mes_inicio
    año_proyeccion = datos.iloc[11,columna].year
    mes_proyeccion = datos.iloc[11,columna].month
    año_van = datos.iloc[12,columna].year
    mes_van = datos.iloc[12,columna].month

    return tabla_info,FP,subestacion,potencia_parque,capex,gx_anual,deg_anual,deg_mensual,p_suficiencia,opex_fijo,terrenos_fijo,inicio,mes_inicio,fin,mes_fin,año_proyeccion,mes_proyeccion,horizonte,proyecto,año_van,mes_van

# Función que crea una tabla con la producción de energía en [GWh] y otra con la producción de energía durante el día, madrugada y noche.
def tablas_energia(datos, FP,columna):
    # Tabla de producción de energía en [GWh]
    meses =['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    horas = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23]

        # Variable intermedia
    valores = datos.iloc[16:304,columna].to_list()

    for i in range(len(valores)):
        valores[i] = round((valores[i]* (FP/0.3))/1000,2)

    pde = pd.DataFrame(columns=meses, index=horas)

    for i in range(len(pde.columns)):
        pde[pde.columns[i]] = valores[i*24:(i+1)*24]


    # Tabla con la producción de madrugada, dia y noche, y el total
    produccion_durante_dia_por_meses = pd.DataFrame(columns=meses, index=['Madrugada [GWh]','Día [GWh]','Noche [GWh]','Total [GWh]'])
    for i in range(0,12): 
        produccion_durante_dia_por_meses.iloc[0,i] = pde.iloc[0:8,i].sum()
        produccion_durante_dia_por_meses.iloc[1,i] = pde.iloc[8:18,i].sum()
        produccion_durante_dia_por_meses.iloc[2,i] = pde.iloc[18:24,i].sum()
        produccion_durante_dia_por_meses.iloc[3,i] = produccion_durante_dia_por_meses.iloc[0:3,i].sum()

    produccion_durante_dia_por_meses

    return pde, produccion_durante_dia_por_meses

# Funcion para convertir tabla a forma procentual la produccion de energia
def tabla_energia_porcentual(pde, produccion_durante_dia_por_meses):
    meses =['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    horas = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23]
    produccion_durante_dia_por_meses_en_porcentaje = pd.DataFrame(columns=meses, index=['Madrugada [%]','Día[%]','Noche[%]'])
    for i in range(0,12):
        produccion_durante_dia_por_meses_en_porcentaje.iloc[0,i] = pde.iloc[0:8,i].sum()/produccion_durante_dia_por_meses.iloc[3,i]
        produccion_durante_dia_por_meses_en_porcentaje.iloc[1,i] = pde.iloc[8:18,i].sum()/produccion_durante_dia_por_meses.iloc[3,i]
        produccion_durante_dia_por_meses_en_porcentaje.iloc[2,i] = pde.iloc[18:24,i].sum()/produccion_durante_dia_por_meses.iloc[3,i]
    return produccion_durante_dia_por_meses_en_porcentaje

# Funcion que arma la tabla de generacion que separa en año y por mes la prodruccion por seccion del día.
def generacion_seccion_dia(produccion_durante_dia_por_meses_en_porcentaje, produccion_durante_dia_por_meses):
    meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
    generacion_seccion_dia = pd.DataFrame(columns = meses, index = ['Gx Madrugada','Gx Día','Gx Noche']) 

    for i in range(0,12):
       generacion_seccion_dia.iloc[0,i] = (produccion_durante_dia_por_meses_en_porcentaje.iloc[0,i]*produccion_durante_dia_por_meses.iloc[3,i])/1000
       generacion_seccion_dia.iloc[1,i] = (produccion_durante_dia_por_meses_en_porcentaje.iloc[1,i]*produccion_durante_dia_por_meses.iloc[3,i])/1000
       generacion_seccion_dia.iloc[2,i] = (produccion_durante_dia_por_meses_en_porcentaje.iloc[2,i]*produccion_durante_dia_por_meses.iloc[3,i])/1000

    return generacion_seccion_dia

# Funcion para obtener tabla con valores de CPI
def cpi(planilla, año_proyeccion, mes_proyeccion,horizonte):
    hojas = pd.ExcelFile(planilla)
    hojas = hojas.sheet_names
    datos = pd.read_excel(planilla, sheet_name = hojas[0],header = 11, index_col = 0 )
    datos = datos.iloc[0:,:12]
    for q in range(len(datos.index)):
        if datos.index.values[q] == año_proyeccion:
            inicio = q
            break
    datos = datos.iloc[inicio:,:]
    anual = []
    promedio_año = []
    cantidad = []

    for w in range(len(datos.index.values)):
        if datos.index.values[w] == datos.index.values[0]:
            for q in range(mes_proyeccion,12):
                    cantidad.append(datos.loc[datos.index.values[0]][q]) 
            if len(cantidad)<12:
                año = []
                año_partida = []
                for q in range(mes_proyeccion,12):
                        año.append(datos.loc[datos.index.values[w]][q])
                for x in range(mes_proyeccion-1,12):
                        año_partida.append(datos.loc[datos.index.values[w]][x])

                len_año_inicial = len(año)
                promedio_año.append(datos.loc[datos.index.values[w]].mean())
            # No puede haber division por 0.
                if len_año_inicial != 0:
                    anual.append((año_partida[-1]/año_partida[0])-1)
                    datos.loc[datos.index.values[w]] = ((1+anual[w])**(1/len_año_inicial))-1 
                else:
                    anual.append((año_partida[-1]/año_partida[0])-1)
                    datos.loc[datos.index.values[w]] = ((1+anual[w]))-1

            else:
                promedio_año.append(datos.loc[datos.index.values[w]].mean())
                anual.append((datos.loc[datos.index.values[w]][-1]/datos.loc[datos.index.values[w]][0])-1)
                datos.loc[datos.index.values[w]] = ((1+anual[w])**(1/12))-1
        else:
            año = []
            anual.append((datos.loc[datos.index.values[w]].mean()/promedio_año[w-1])-1)
            promedio_año.append(datos.loc[datos.index.values[w]].mean())
            datos.loc[datos.index.values[w]] = ((1+anual[w])**(1/12))-1
    primer_año = datos.index.values[0]
    for i in range(horizonte*2):
        anual.append(anual[-1])
        datos.loc[datos.index.values[-1]+1] = ((1+anual[-1])**(1/12))-1

    for i in range(1,16):
        anual.insert(0,0)
        datos.loc[primer_año-(i)] = 0

    datos.sort_index(inplace = True)

    for i in range(len(datos.index)):
        for q in range(len(datos.columns)):
            if datos.index.values[i] == año_proyeccion:
                if q < mes_proyeccion:
                    datos.iloc[i,q] = 0
        months = ['Janaury','February','March','April','May','June','July','August','September','October','November','December','Anual']
        datos_iniciales_cpi = pd.DataFrame(columns= months)
        datos_iniciales_cpi['Janaury'] = datos.iloc[:,0]
        datos_iniciales_cpi['February'] = datos.iloc[:,1]
        datos_iniciales_cpi['March'] = datos.iloc[:,2]
        datos_iniciales_cpi['April'] = datos.iloc[:,3]
        datos_iniciales_cpi['May'] = datos.iloc[:,4]
        datos_iniciales_cpi['June'] = datos.iloc[:,5]
        datos_iniciales_cpi['July'] = datos.iloc[:,6]
        datos_iniciales_cpi['August'] = datos.iloc[:,7]
        datos_iniciales_cpi['September'] = datos.iloc[:,8]
        datos_iniciales_cpi['October'] = datos.iloc[:,9]
        datos_iniciales_cpi['November'] = datos.iloc[:,10]
        datos_iniciales_cpi['December'] = datos.iloc[:,11]
        datos_iniciales_cpi['Anual'] = anual

    return datos_iniciales_cpi

# Función para obtener las barras que están presentes en el proyecto.
def barras(datos,columna,subestacion):
    barras = []
    barras.append(subestacion)
    vector_temporal = datos.fillna(0).iloc[:,columna]

    for i in range(13,16):
        
        if vector_temporal[i] != 0:
            barras.append(datos.iloc[i,columna])
    
    del vector_temporal
    for i in range(len(barras)):
        barras[i] = barras[i].strip('S/E').strip().strip('.csv')
    return barras

# Función para obtener los nombres y rutas de las barras disponibles en la base de datos.
def db_barras():
    barras_db = os.listdir('cmg')
    barras_db= [i for i in barras_db if i.endswith('.csv') and i.startswith('CMg')]

    folder_path = 'cmg'
    rutas = []
#obtener rutas de cada archivo de la caprta 
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv') and filename.startswith('CMg'):
             rutas.append(os.path.join(folder_path, filename))
    return barras_db,rutas

# Funcion para buscar la ruta de la barra en la base de datos.
def buscador_ruta(barra,barras_db,rutas):
    for i in range(len(barras_db)):
        if barra in barras_db[i]:
        #print('esta en el indice {}, la ruta es {}'.format(i,rutas[i]))
            barra = rutas[i]
            break
    return barra

# Funcion para obtener los datos marginales con hidrologias desde la BBDD.
def barra_cmg_hidro(planilla):

    barra = pd.read_csv(planilla, sep=',', encoding='latin-1')
    columnas = barra.columns.to_list()
    barra = barra.groupby([columnas[0],columnas[8],columnas[6],columnas[2]])[columnas[-1]].mean().to_frame()
    barra =barra.pivot_table(index=[columnas[8],columnas[6]],columns=[columnas[0],columnas[2]],values=columnas[-1])
    barra.reset_index(inplace=True)

    encabezados = barra.columns.to_list()[2:]

    for i in range(len(encabezados)):
        encabezados[i] = fecha_to_string(encabezados[i][0],encabezados[i][1])
    encabezados.insert(0,'Momento')
    encabezados.insert(1,'Hidro')
    barra.columns = encabezados

    return barra

# Funcion para obtener la tabla con la hidrologia correpsondiente.
def barra_cmg(barra_cmg_hidro,hidrologia,horizonte,fecha_inicio,inicio):

    barra_cmg  = pd.DataFrame(columns = barra_cmg_hidro.columns[2:], index = ['Madrugada','Día','Noche'])

    madrugada = barra_cmg_hidro.query('Momento == "Madrugada"').iloc[hidrologia].to_list()[2:]
    dia = barra_cmg_hidro.query('Momento == "Dia"').iloc[hidrologia].to_list()[2:]
    noche = barra_cmg_hidro.query('Momento == "Noche"').iloc[hidrologia].to_list()[2:]
    
    barra_cmg.iloc[0] = madrugada
    barra_cmg.iloc[1] = dia
    barra_cmg.iloc[2] = noche

    madrugada = madrugada[-1]
    dia = dia[-1]
    noche = noche[-1]

    año = int(barra_cmg_hidro.columns[-1][:4])
    mes = int(barra_cmg_hidro.columns[-1][-2:])
    

    
    # Aqui se ajustan los años. ojo con la tabla de CMg hidrologias, si tiene mas años puede generar problemas, solo está condicionada cunado los años son menores al horizonte.
    dif = horizonte*12 - len(barra_cmg.columns[barra_cmg.columns.tolist().index(fecha_inicio):])
    if dif > 0:
        for i in range(dif+1):
            barra_cmg[fecha_to_string(año,mes)] = [madrugada,dia,noche]
            mes += 1
            if mes > 12:
                mes = 1
                año += 1
        
    año = int(barra_cmg_hidro.columns[2][:4])
    mes = int(barra_cmg_hidro.columns[2][-2:])

    # Agrega meses anteriores
    if abs(año - inicio) < 5:
        for i in range((5-abs(año-inicio))*12):
            if mes == 1:
                año -= 1
                mes = 12
            else:
                mes -= 1

            barra_cmg.insert(0,fecha_to_string(año,mes),[0,0,0], allow_duplicates = False) 
    
            
    del madrugada, dia, noche , año , mes

    return barra_cmg

# Función que genera la tabla de generacion por seccion de dia.
def gen(generacion_seccion_dia,fecha,años_mes_strings,potencia_parque,deg_anual):
    gen = pd.DataFrame(columns = años_mes_strings, index = ['Potencia [MW]','Madrugada [TWh]','Día [TWh]','Noche [TWh]','Total Gx por dia [TWh] '])
    marcador = marca(años_mes_strings,fecha)

# Tiene distintos valores dependiendo del inicio del proyecto, por ejemplo, los 4 meses antes, 
# produce en forma parcial, hasta que llega a el mes de inicio, donde produce a su maxima capacidad. 
# Luego de un año, empieza a producir con una degradacion.
# También, antes de los cuatro meses de que empiece, los valores son 0.

    for i in range(len(gen.columns)):
        if i < marcador:
            if marcador-3>i>= marcador-4:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]*0.2
                        gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]*0.2
                        gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]*0.2
            elif marcador-2>i>= marcador-3:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]*0.4
                        gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]*0.4
                        gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]*0.4
            elif  marcador-1>i>= marcador-2:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]*0.6
                        gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]*0.6
                        gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]*0.6
            elif marcador>i>= marcador-1:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]*0.8
                        gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]*0.8
                        gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]*0.8
            elif i< marcador-4:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = 0
                        gen.iloc[2,i] = 0
                        gen.iloc[3,i] = 0
            else:
                for x in range(12):
                    if i%int(gen.columns[i][-2:]) == x:
                        gen.iloc[0,i] = 0
                        gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]
                        gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]
                        gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]
        elif i>marcador+12:
            for x in range(12):
                if i%int(gen.columns[i][-2:]) == x:
                    gen.iloc[0,i] = potencia_parque
                    gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]*(1-deg_anual)
                    gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]*(1-deg_anual)
                    gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]*(1-deg_anual)
        else:
            for x in range(12):
                if i%int(gen.columns[i][-2:]) == x:
                    gen.iloc[0,i] = potencia_parque
                    gen.iloc[1,i] = generacion_seccion_dia.iloc[0,x]
                    gen.iloc[2,i] = generacion_seccion_dia.iloc[1,x]
                    gen.iloc[3,i] = generacion_seccion_dia.iloc[2,x]

        gen.iloc[4,i] = gen.iloc[1,i] + gen.iloc[2,i] + gen.iloc[3,i]

    return gen

# Función para generar la tabla de generación promedio anual.
def generacion_promedio_anual(gen):
    gen_promedio_anual  = pd.DataFrame(columns =[i for i in range(int(gen.columns[0][0:4]),int(gen.columns[-1][0:4])+1)], index = ['Potencia promedio [MW]','Madrugada promedio [TWh]','Día promedio [TWh]','Noche promedio [TWh]','Total Gx por dia promedio [TWh] '])

    lista_anios_presentes = [i for i in range(int(gen.columns[0][0:4]) ,int(gen.columns[-1][0:4])+1)]

    for i in range(len(lista_anios_presentes)):
        lista = []
        for x in range(len(gen.columns)):
            if int(gen.columns[x][0:4]) == lista_anios_presentes[i]:
                lista.append(gen.columns[x])


        valores = gen[lista].sum(axis=1)
        gen_promedio_anual[lista_anios_presentes[i]][0] = valores[0]/12
        gen_promedio_anual[lista_anios_presentes[i]][1] = valores[1]
        gen_promedio_anual[lista_anios_presentes[i]][2] = valores[2]
        gen_promedio_anual[lista_anios_presentes[i]][3] = valores[3]
        gen_promedio_anual[lista_anios_presentes[i]][4] = valores[4]
    return gen_promedio_anual  

# Funcion que crea la tabla de PPA mes-año
def PPA_mes_año(gx_anual,fecha,años_mes_strings):
    PPA  =  pd.DataFrame(columns= años_mes_strings,index = ['PPA Madrugada','PPA Dia','PPA Noche','PPA Total'])
    marcador = marca(años_mes_strings,fecha)
    dias_meses = [31,28,31,30,31,30,31,31,30,31,30,31]
    for i in range(len(PPA.columns)):
        if i < marcador:
                for x in range(12):
                    if i%int(PPA.columns[i][-2:]) == x:
                        PPA.iloc[0,i] = 0
                        PPA.iloc[1,i] = 0
                        PPA.iloc[2,i] = 0

        else:
            for x in range(12):
                if i%int(PPA.columns[i][-2:]) == x:
                    PPA.iloc[0,i] = gx_anual*(dias_meses[x]/365)*(8/24)*(1/1000)
                    PPA.iloc[1,i] = gx_anual*(dias_meses[x]/365)*(10/24)*(1/1000)
                    PPA.iloc[2,i] = gx_anual*(dias_meses[x]/365)*(6/24)*(1/1000)
                
        PPA.iloc[3,i] = PPA.iloc[0,i] + PPA.iloc[1,i] + PPA.iloc[2,i]
        
    return PPA

# Funcion que genera la tabla de PPA anual.
def PPA_anual(PPA):
    PPA_anual  = pd.DataFrame(columns =[i for i in range(int(PPA.columns[0][0:4]),int(PPA.columns[-1][0:4])+1)], index = ['PPA Madrugada','PPA Dia','PPA Noche','PPA Total'])
    lista_anios_presentes = [i for i in range(int(PPA.columns[0][0:4]),int(PPA.columns[-1][0:4])+1)]

    for i in range(len(lista_anios_presentes)):
        lista = []
        for x in range(len(PPA.columns)):
            if int(PPA.columns[x][0:4]) == lista_anios_presentes[i]:
                lista.append(PPA.columns[x])
  
    
        valores = PPA[lista].sum(axis=1)
        PPA_anual[lista_anios_presentes[i]][0] = valores[0]
        PPA_anual[lista_anios_presentes[i]][1] = valores[1]
        PPA_anual[lista_anios_presentes[i]][2] = valores[2]
        PPA_anual[lista_anios_presentes[i]][3] = valores[3]
        
    return PPA_anual,lista_anios_presentes

# Funcion que genera el vector de precios PPA por año.
def precio_PPA(datos_iniciales_cpi,ppa,año_proyeccion,inicio,horizonte,mes_inicio):

    # Para este caso, toma la tabla datos_iniciales cpi, y realiza una comparacion. 
    # Si el año esta presente en la tabla de datos_inciiales cpi, 
    # Agrega a precio ppa el valor corrspondiente.
    # Si el año no esta presente en la tabla CPI y es menor al menor año de la tabla cpi, rellena con 0.
    # Si el año no está presente en la tabla CPI y es mayor al mayor año de la tabla cpi, 
    # rellena con el ultimo valor del ultimo año de la tabla cpi hasta rellenar los años faltantes.

    porcentaje_cpi = pd.DataFrame(columns = datos_iniciales_cpi.index ,index = ['%'])
      
    for i in range(len(porcentaje_cpi.columns)):
        if porcentaje_cpi.columns[i] < año_proyeccion:
            porcentaje_cpi[porcentaje_cpi.columns[i]][0] = 0
        elif porcentaje_cpi.columns[i] >= año_proyeccion:
            porcentaje_cpi[porcentaje_cpi.columns[i]][0] = datos_iniciales_cpi['Anual'][porcentaje_cpi.columns[i]]
        else:
            print('Revisar funcion precio PPA en primeros condicionales')


    #Calculo de vector de CPI anual |||| revisar por si acaso
    cpi_anual = pd.DataFrame(columns = datos_iniciales_cpi.index ,index = ['%'])
    for i in range(len(cpi_anual.columns)):
        if cpi_anual.columns[i] <= año_proyeccion:
            cpi_anual[cpi_anual.columns[i]][0] = 1
        elif cpi_anual.columns[i] > año_proyeccion:
            cpi_anual[cpi_anual.columns[i]][0] = cpi_anual[cpi_anual.columns[i-1]][0]*(1+porcentaje_cpi[cpi_anual.columns[i-1]][0])


    if mes_inicio == 1:
        columnas_ppa = datos_iniciales_cpi.index[ :datos_iniciales_cpi.index.to_list().index(inicio+horizonte)]
    else:
        columnas_ppa = datos_iniciales_cpi.index[ :datos_iniciales_cpi.index.to_list().index(inicio+horizonte)+1]
    

    
    precio_PPA = pd.DataFrame(columns = columnas_ppa,index = ['Precio PPA [MMUSD]'])
    for i in precio_PPA.columns:
        if i < inicio:
            precio_PPA[i][0] = 0
        elif i >= inicio:
            precio_PPA[i][0] = ppa*cpi_anual[i][0]
        else:
            print('Revisar funcion precio PPA en segundo condicional')

   
    marcador_precio_ppa = marca(datos_iniciales_cpi.index.to_list(),inicio-5)

    # Cortamos los datasets a los años que necesitamos hacer la proyeccion
    precio_PPA = precio_PPA.loc[:,precio_PPA.columns.to_list()[marcador_precio_ppa:]]
    cpi_anual = cpi_anual.loc[:,cpi_anual.columns.to_list()[marcador_precio_ppa:]]


    
    
    return precio_PPA,cpi_anual

# Funcion para obtener los vectores de CPI y de factor CPI.
def factorCPI(datos_iniciales_cpi):

    columnas_cpi = []
    for i in range(len(datos_iniciales_cpi.index)):
        # se puede arreglar mas
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'01')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'02')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'03')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'04')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'05')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'06')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'07')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'08')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'09')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'10')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'11')
        columnas_cpi.append(str(datos_iniciales_cpi.index[i])+'-'+'12')



    cpi_horizontal = pd.DataFrame(columns = columnas_cpi, index = ['CPI'])

    n_mes ={'Janaury':'01','February':'02','March':'03','April':'04','May':'05','June':'06','July':'07','August':'08','September':'09','October':'10','November':'11','December':'12'}

    datos_cpi_mod = datos_iniciales_cpi.copy()
    datos_cpi_mod.drop('Anual',axis=1,inplace=True)
    

    # Pasamos los datos a un formato horizontal
    # ojo con este, puede resolver mi problema
    for i in range(len(datos_cpi_mod.index)):
        for x in range(len(datos_cpi_mod.columns)):
            cpi_horizontal[str(datos_cpi_mod.index[i])+'-'+n_mes[datos_cpi_mod.columns[x]]] = datos_cpi_mod.iloc[i,x]


    cpi_final = cpi_horizontal.copy()
    del cpi_horizontal

    
    factor_cpi = pd.DataFrame(columns=cpi_final.columns,index = ['Factor CPI'])

    for i in range(len(cpi_final.columns)):
        if i == 0:
            factor_cpi[factor_cpi.columns[i]] = 1
        else:
            factor_cpi[factor_cpi.columns[i]] = factor_cpi[cpi_final.columns[i-1]]*(1+cpi_final[cpi_final.columns[i]].to_list()[0])
        
    return cpi_final,factor_cpi

# Función para modificar los costos marginales de cada barra
def barra_cmg_mod(barra_cmg,fecha,tasa_nominal,factor_cpi):
    
    factor_random = 0.11346861973958
    tasa_2 = 0.01
    #Tasa 1 es un valor definido
    tasa_1 = 0.0825
    tasa_desconocida = tasa_nominal - tasa_1
    valor_desconocido = 1 + (tasa_desconocida/tasa_2)*factor_random
    marcador = marca(barra_cmg.columns.to_list(),fecha)
    
    # No creamos nueva variables para barra_cmg, si no que volvemosa utilizar la misma.
    for i in range(len(barra_cmg.columns)):
        if i == 0:
            if (int(barra_cmg.columns[i][0:4]) >= int(barra_cmg.columns[marcador][0:4]) ) and (int(barra_cmg.columns[i][5:7]) >= int(barra_cmg.columns[marcador][5:7])):
                #barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]*0.858164225
                barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]*valor_desconocido
                #barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]
            else:
                barra_cmg[barra_cmg.columns[i]] = 0
        else:
            if int(barra_cmg.columns[i][0:4]) >= int(barra_cmg.columns[marcador][0:4]):
                #barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]*0.858164225
                barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]*valor_desconocido
                #barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i]]*factor_cpi[barra_cmg.columns[i]][0]
            else:
                barra_cmg[barra_cmg.columns[i]] = barra_cmg[barra_cmg.columns[i-1]]

    return barra_cmg

# Funcion para realizar el calculo de EBITDA, el cual arroja un dataframe con el valor ebitda y otros valores de interés
def EBITDA(factor_cpi,gen,p_suficiencia,fecha,fin,mes_fin, PPA,PPA_anual,precio_PPA,inicio,horizonte,columna,opex_fijo,terrenos_fijo,potencia_parque,años_mes_strings,inyeccion_cmg,retiro_cmg,lista_anios_presentes,cpi_anual,datos):
 
    
    ebitda_mes = pd.DataFrame(columns = años_mes_strings,index=['Ventas Potencia [MMUSD]','Inyección [MMUSD]','Retiro [MMUSD]','Balance I/R [MMUSD]'])
    
    marcador_ebitda_cpi = marca(factor_cpi.columns.to_list(),ebitda_mes.columns.to_list()[0])
    marcador_ebitda_cpi_fin = marca(factor_cpi.columns.to_list(),ebitda_mes.columns.to_list()[-1])

    marcador_gen = marca(factor_cpi.columns[marcador_ebitda_cpi:marcador_ebitda_cpi_fin+1],gen.columns[0])
    
    # Rellenamos vector Gen[potencia] si es necesario

    gen_vector  = gen.loc['Potencia [MW]'].to_list()
    for i in range(marcador_gen):
        gen_vector.insert(0,0)
    
    ventas_potencia = factor_cpi.values[0][marcador_ebitda_cpi:marcador_ebitda_cpi_fin+1]*gen_vector*p_suficiencia*0.008

    marcador = marca(ebitda_mes.columns.to_list(),fecha)

    for i in range(len(ventas_potencia)):   
        if i < marcador:
            ventas_potencia[i] = 0

    ebitda_mes.loc['Ventas Potencia [MMUSD]'] = ventas_potencia
    for i in range(len(ebitda_mes.columns)):
        ebitda_mes.iloc[1,i] = gen.iloc[1,i]*inyeccion_cmg.iloc[0,i] + gen.iloc[2,i]*inyeccion_cmg.iloc[1,i] + gen.iloc[3,i]*inyeccion_cmg.iloc[2,i] 
        ebitda_mes.iloc[2,i] = -(PPA.iloc[0,i]*retiro_cmg.iloc[0,i] + PPA.iloc[1,i]*retiro_cmg.iloc[1,i] + PPA.iloc[2,i]*retiro_cmg.iloc[2,i])
        ebitda_mes.iloc[3,i] = ebitda_mes.iloc[1,i] + ebitda_mes.iloc[2,i] + ebitda_mes.iloc[0,i]
    
    # EBITDA por año
    ebitda_año = pd.DataFrame(columns = lista_anios_presentes,index=['Ventas Potencia [MMUSD]','Ventas PPA[MMUSD]','Inyección [MMUSD]','Retiro [MMUSD]','Balance I/R [MMUSD]'])
    
    for i in range(len(lista_anios_presentes)):
        lista = []
        for x in range(len(ebitda_mes.columns)):
            if int(ebitda_mes.columns[x][0:4]) == lista_anios_presentes[i]:
                lista.append(ebitda_mes.columns[x])
    
        valores = ebitda_mes[lista].sum(axis=1)

        ebitda_año[lista_anios_presentes[i]][0] = valores[0]
        ebitda_año[lista_anios_presentes[i]][1] = PPA_anual[PPA_anual.columns[i]][3]*precio_PPA[precio_PPA.columns[i]][0]
        ebitda_año[lista_anios_presentes[i]][2] = valores[1]
        ebitda_año[lista_anios_presentes[i]][3] = valores[2]
        ebitda_año[lista_anios_presentes[i]][4] = valores[1] + valores[2]


    if  len(datos.iloc[:,columna]) > 304:

        datos.iloc[:,columna].fillna('a',inplace=True)
        
        if datos.iloc[304,columna] != 'a':

            for i in range(len(lista_anios_presentes)):
                if inicio == lista_anios_presentes[i]:
                    new_columns = (lista_anios_presentes[i:])
                    break
                opex =  datos.iloc[304:333,columna].to_list()[0:-1]
                terrenos = datos.iloc[334:363,columna].to_list()[0:-1]
                

        else:

            opex = [-(opex_fijo*potencia_parque)/1000 for i in range(inicio,inicio+horizonte-1)]
            terrenos = [-(terrenos_fijo*potencia_parque)/1000000 for i in range(inicio,inicio+horizonte-1)]

    
    elif len(datos.iloc[:,columna]) < 304:
        opex = [-(opex_fijo*potencia_parque)/1000 for i in range(inicio,inicio+horizonte-1)]
        terrenos = [-(terrenos_fijo*potencia_parque)/1000000 for i in range(inicio,inicio+horizonte-1)]

        
    for i in range(5):
            opex.insert(0,0)
            terrenos.insert(0,0)
    
    if len(opex) < len(ebitda_año.columns):
        dif = abs(len(opex) - len(ebitda_año.columns))
        for i in range(dif):
            opex.append(opex[-1])
            terrenos.append(terrenos[-1])


    for i in range(len(opex)):
        opex[i] = opex[i]*cpi_anual.iloc[0,i]
        terrenos[i] = terrenos[i]*cpi_anual.iloc[0,i]

#Agrego los valores de opex y terrenos a la tabla de ebitda año , ay que no hay problemas.
    ebitda_año.loc['OPEX [MMUSD]'] = opex
    ebitda_año.loc['Terrenos [MUSD]'] = terrenos


    total_costos = []
    for i in range(len(ebitda_año.columns)):
        total_costos.append( ebitda_año[ebitda_año.columns[i]][5]+ ebitda_año[ebitda_año.columns[i]][6])

    ebitda_año.loc['Total Costos [MMUSD]'] = total_costos

# Se agrega el valor de EBITDA
    valor_ebitda = []
    for i in range(len(ebitda_año.columns)):
        valor_ebitda.append( ebitda_año[ebitda_año.columns[i]][0]+ ebitda_año[ebitda_año.columns[i]][1]+ ebitda_año[ebitda_año.columns[i]][4]+ ebitda_año[ebitda_año.columns[i]][7])


    ebitda_año.loc['EBITDA [MMUSD]'] = valor_ebitda

    del valor_ebitda

    return ebitda_año,ebitda_mes

# Funcion que crea el vector capex
def capex_vector(ebitda_año,cpi_anual,capex,inicio):
    vector_capex = [i for i in range(ebitda_año.columns[0],ebitda_año.columns[-1]+1)]

    for i in range(len(vector_capex)):
        if vector_capex[i] == inicio-1:
            vector_capex[i] = -(0.7*capex)*cpi_anual.iloc[0,i]
        elif vector_capex[i] == inicio:
            vector_capex[i] = -(0.3*capex)*cpi_anual.iloc[0,i]
        else:
            vector_capex[i] = 0
    return vector_capex

# Funcion que crea tabla impuesto_fcf y vector_capex 
def impuesto(ebitda_año,cpi_anual,capex,inicio):
    impuesto_fcf = pd.DataFrame(columns= ebitda_año.columns,index =['Utilidad Tributaria [MMUSD]','Depreciación [MMUSD]','Utilidad antes de impuestos [MMUSD]','Pérdidas Acumuladas [MMUSD]','Impuesto [MMUSD]'])
    impuesto_fcf.loc['Utilidad Tributaria [MMUSD]'] = ebitda_año.loc['EBITDA [MMUSD]']

# Creamos vector capex 
    vector_capex = capex_vector(ebitda_año,cpi_anual,capex,inicio)

# Se definen estos porcentajes del desglose del capex.
    porcentaje_OOCC = 0.085897571
    porcentaje_equipos = 0.866463351
    porcentaje_terrenos = 0.047639079

    depreciacion = pd.DataFrame(columns= ebitda_año.columns,index = ebitda_año.columns)

    for i in range(len(depreciacion.columns)):
        for j in range(len(depreciacion.index)):
            if vector_capex[j] == 0 :
                depreciacion.iloc[j,i] = 0
            elif vector_capex[j] != 0:
                for i in range(16):
                    if i <3:
                        depreciacion.iloc[j,i+j+1] = (vector_capex[j]/3)*porcentaje_equipos + (vector_capex[j]/16)*porcentaje_OOCC + (vector_capex[j]/6)*porcentaje_terrenos
                    elif i >=3 and i <6:
                        depreciacion.iloc[j,i+j+1] = (vector_capex[j]/16)*porcentaje_OOCC + (vector_capex[j]/6)*porcentaje_terrenos
                    elif i >=6 and i <=16:
                        depreciacion.iloc[j,i+j+1] = (vector_capex[j]/16)*porcentaje_OOCC

    depreciacion.fillna(0,inplace=True)       

    depreciacion.loc['Total Depreciacion año [MMUSD]'] = depreciacion.sum(axis=0)

    impuesto_fcf.loc['Depreciación [MMUSD]'] = depreciacion.loc['Total Depreciacion año [MMUSD]']
    impuesto_fcf.loc['Utilidad antes de impuestos [MMUSD]'] = impuesto_fcf.sum(axis=0)

    for q in range(len(impuesto_fcf.columns)):
        if q == 0:
            if  impuesto_fcf.iloc[3,q] < 0:
                impuesto_fcf.iloc[4,q] = impuesto_fcf.iloc[3,q]
            else:
                impuesto_fcf.iloc[3,q] = 0
        else:
            if impuesto_fcf.iloc[2,q] + impuesto_fcf.iloc[3,q-1] < 0:
                impuesto_fcf.iloc[3,q] = impuesto_fcf.iloc[2,q] + impuesto_fcf.iloc[3,q-1]
            else:
                impuesto_fcf.iloc[3,q] = 0

# Calculo de impuesto
    impuesto_fcf.loc['Impuesto [MMUSD]'] = 0
    for q in range(len(impuesto_fcf.columns)):
        if q == 0:
            if -(impuesto_fcf.iloc[2,q])*0.27 < 0:
                impuesto_fcf.iloc[4,q] = -(impuesto_fcf.iloc[2,q])*0.27
            elif -(impuesto_fcf.iloc[2,q])*0.27 > 0 :
                impuesto_fcf.iloc[4,q] = 0
        else:
            if -(impuesto_fcf.iloc[2,q]+impuesto_fcf.iloc[3,q-1])*0.27 < 0:
                impuesto_fcf.iloc[4,q] = -(impuesto_fcf.iloc[2,q] + impuesto_fcf.iloc[3,q-1])*0.27
            elif -(impuesto_fcf.iloc[2,q]+impuesto_fcf.iloc[3,q-1])*0.27 > 0:
                impuesto_fcf.iloc[4,q] = 0
    
    return impuesto_fcf, vector_capex

# Funcion que crea tabla de flujo de caja.
def flujo_caja(ebitda_año,impuesto_fcf,vector_capex):
    # Creamos tabla de flujo de caja.
    flujo = pd.DataFrame(columns= ebitda_año.columns,index = ['EBITDA [MMUSD]','CAPEX [MMUSD]']) 
# Ingresamos los valores a la tabla.
    flujo.loc['EBITDA [MMUSD]'] = ebitda_año.loc['EBITDA [MMUSD]']
    flujo.loc['Impuesto a la Renta [MMUSD]'] = impuesto_fcf.loc['Impuesto [MMUSD]']
    flujo.loc['CAPEX [MMUSD]'] = vector_capex

    flujo.loc['Flujo de caja [MMUSD]'] = flujo.sum(axis=0)
    return flujo

# Funcion para calcular el VAN.
def VAN(flujo,tasa_nominal,datos_iniciales_cpi,año_van,mes_van,año_proyeccion,mes_proyeccion,tasa_mensual):
    van_vector = []

    if año_van == año_proyeccion and mes_van == 1 and mes_proyeccion == 1: 
        marcador_van = marca(flujo.columns.to_list(),año_van)
        for i in range(len(flujo.columns.to_list()[marcador_van:])):
            van_vector.append(1/(1+tasa_nominal)**(i) * flujo.loc[flujo.index[-1]].to_list()[marcador_van:][i])
        vans = sum(van_vector)

    elif año_van > año_proyeccion:

        dif = (año_van - año_proyeccion)*12 + mes_van - mes_proyeccion
    
        factor = (1/(1+tasa_mensual)**(dif-1))
        marcador_van = marca(flujo.columns.to_list(),año_van)
        
        for i in range(len(flujo.columns.to_list()[marcador_van:])):
            van_vector.append(1/(1+tasa_nominal)**(i) * flujo.loc[flujo.index[-1]].to_list()[marcador_van:][i])
        vans = sum(van_vector)*factor
        
    return vans

# Función que busca la tasa mensual del cpi, que se utilizará en el caso en que la fecha de proyección este antes o despues de la fecha del van.
def buscar_tasa(datos_iniciales_cpi,año_proyeccion,mes_proyeccion):
 #Podemos ver tambien si es mayor al año o menor, pero eso lo haria mas adelante.
   tasa_men = datos_iniciales_cpi.query('Year=={}'.format(año_proyeccion)).iloc[0].to_list()[mes_proyeccion-1]

   return tasa_men


def CalculoPPACorte(columna,barrita):

    datos = importar_datos('Entrada.xlsx')

    tabla_info,FP,subestacion,potencia_parque,capex,gx_anual,deg_anual,deg_mensual,p_suficiencia,opex_fijo,terrenos_fijo,inicio,mes_inicio,fin,mes_fin,año_proyeccion,mes_proyeccion,horizonte,proyecto,año_van,mes_van= info(datos,columna)
    fecha_inicio = fecha_to_string(inicio,mes_inicio)
    fecha_fin = fecha_to_string(fin,mes_fin)
    archivo = pd.ExcelWriter('Resultados/Resultados {}.xlsx'.format(proyecto), engine='xlsxwriter')

    # Valores con lo que se deben calcular cada flujo
    hidrologias = [i for i in range(0,13)]
    tasas = [0.07]#,0.0825,0.09]


    # Tablas de energía
    produccion_de_energia,produccion_durante_dia_por_meses = tablas_energia(datos, FP, columna)
    # Tabla de energía en porcentaje por bloques y meses.
    produccion_durante_dia_por_meses_en_porcentaje = tabla_energia_porcentual(produccion_de_energia,produccion_durante_dia_por_meses)
    # Tabla de generación por sección y día.
    generacion_seccion_dias  = generacion_seccion_dia(produccion_durante_dia_por_meses_en_porcentaje,produccion_durante_dia_por_meses)
    # tabla CPI
    datos_iniciales_cpi = cpi('CPI.xlsx',año_proyeccion,mes_proyeccion,horizonte)
    # Tasa Mensual
    tasa_mensual = buscar_tasa(datos_iniciales_cpi,año_proyeccion,mes_proyeccion)
    # Barras presentes en el archivo.
    barras_proyecto = barras(datos,columna,subestacion)
    #barras_proyecto = [barras_proyecto[0]]
    barras_db, rutas = db_barras()
   
    barra_local = buscador_ruta(barras_proyecto[0],barras_db,rutas)

    # Tabla de costos marginales de la subestacion donde se conecta la planta.
    #inyeccion_cmg_hidro, años_mes_strings = local_bar_cmg(hidro,subestacion)  # Es independiente 
    inyeccion_cmg_hidro = barra_cmg_hidro(barra_local)  # Es independiente

    
    
    # Tabla de costos marginales de la barra de retiro para todas las hidrologias
    retiro_cmg_hidro = barra_cmg_hidro(buscador_ruta(barras_proyecto[barrita],barras_db,rutas))  # Es independiente

        
    for t in range(len(tasas)):
            lcoe_a = []
            tasa_nominal = tasas[t]
            PPA_LCOE = pd.DataFrame(columns = ['Hidrología 1968','Hidrología 1998','Hidrología 2011','Hidrología 2012','Hidrología 2013','Hidrología 2014','Hidrología 2015','Hidrología 2016','Hidrología 2017','Hidrología 2018','Hidrología 2019','Hidrología 2020','Hidrología 2021','Promedio'],index = ['PPA/LCOE [USD/MWh]'])    

            for h in range(len(hidrologias)):
                hidrologia = hidrologias[h]
                val = []
                ppa = 200
                p = 199
                i = 0
                q = 0
                z = 0
                print(' Proyecto: {}, Barra: {}, Tasa: {}, Hidrologia: {}'.format(proyecto,barras_proyecto[barrita],tasa_nominal,hidrologia+1))
                ti = time.time()
                while True:
                    # Tabla de costos marginales de la hidrologia para barra de inyeccion.
                    inyeccion_cmg = barra_cmg(inyeccion_cmg_hidro,hidrologia,horizonte,fecha_inicio,inicio)
                    años_mes_strings = inyeccion_cmg.columns.to_list()
                    # Tabla de costos marginales de la hidrologia para barra de retiro.
                    retiro_cmg = barra_cmg(retiro_cmg_hidro,hidrologia,horizonte,fecha_inicio,inicio)
                    # Creación de tabla de generacion de cada mes de cada año.
                    gens = gen(generacion_seccion_dias,fecha_inicio,años_mes_strings,potencia_parque,deg_anual)
                    # Creación de tabla de generacion de promedio anual, la cual se calcula con la tabla anterior.
                    gen_promedio_anual = generacion_promedio_anual(gens)
                    # Creación de tabla PPA Mes-Año, que a su vez esta divididoo en seccion del día.
                    PPA = PPA_mes_año(gx_anual,fecha_inicio,años_mes_strings)
                    # Creación de tabla de PPA Anual
                    PPA_anuals, lista_anios_presentes = PPA_anual(PPA)
                    # Creación de vector Precio PPA
                    precio_PPAs, cpi_anual = precio_PPA(datos_iniciales_cpi,ppa,año_proyeccion,inicio,horizonte,mes_inicio)
                    # Creación de vector factor CPI.
                    cpi_final,factor_cpi = factorCPI(datos_iniciales_cpi)

                    # Modificación de tabla de valores marginales de barra de inyección.
                    inyeccion_cmg = barra_cmg_mod(inyeccion_cmg,fecha_inicio,tasa_nominal,factor_cpi)      
                    # Modificación de tabla de valores marginales de barra de retiro.
                    retiro_cmg = barra_cmg_mod(retiro_cmg,fecha_inicio,tasa_nominal,factor_cpi)
                    # Creación de tabla de EBITDA por año.
                    ebitda_año,ebitda_mes  =  EBITDA(factor_cpi,gens,p_suficiencia,fecha_inicio,fin,mes_fin, PPA,PPA_anuals,precio_PPAs,inicio,horizonte,columna,opex_fijo,terrenos_fijo,potencia_parque,años_mes_strings,inyeccion_cmg,retiro_cmg,lista_anios_presentes,cpi_anual,datos)
                     # Creación tabla impuesto fcf y vector capex
                    impuesto_fcf, vector_capex = impuesto(ebitda_año,cpi_anual,capex,inicio)
                    # Creación tabla de flujo de caja.
                    flujo  = flujo_caja(ebitda_año,impuesto_fcf,vector_capex)
                    # Calculo de VAN 
                    van = VAN(flujo,tasa_nominal,datos_iniciales_cpi,año_van,mes_van,año_proyeccion,mes_proyeccion,tasa_mensual)
                    
                    val.append(van)

                    

                    if round(van,1) == 0 and hidrologia == 0 :
                        gsalida, gen_anual_salida, PPA_salida, flujo_salida  = gens, gen_promedio_anual, precio_PPAs,flujo
                        gsalida ,gen_anual_salida, PPA_salida, flujo_salida = gsalida.reset_index(),gen_anual_salida.reset_index(), PPA_salida.reset_index(), flujo_salida.reset_index()

                    if round(van,1) == 0 and i != 0: 
                        print('El VAN es 0')
                        print('Proyecto: {} Barra: {}, Tasa: {}, Hidrologia: {}'.format(proyecto,barras_proyecto[barrita],tasa_nominal,hidrologia+1))
                        print('EL LCOE es {} '.format(ppa))
                        lcoe_a.append(ppa)
                        break
                    
                    elif van > 0:
                        ppa = ppa - p

                    elif van < 0:
                        ppa = ppa + p

                    if abs(van) < 100:
                        q=+1
                        if q > 5:
                            p = p*2
                            q = 0

                        else:
                            p = p/2
        
                    else:
                        p = p/3
                    
                        

                    if i > 4:
                        if int(val[0]) == int(val[1]) == int(val[2]) and (100> abs(van) > 1):

                            if van >0:
                                p = p -2
                            else:
                                p = p +2

                        elif int(val[0]) == int(val[1]) == int(val[2]) and (200 >= abs(van) > 100):

                            if van > 0:
                                p = p - 5
                            else:
                                p = p + 5

                        elif int(val[0]) == int(val[1]) == int(val[2]) and (300 >= abs(van) > 200):

                            if van > 0:
                                p = p - 5
                            else:
                                p = p + 5
                            

                        elif int(val[0]) == int(val[1]) == int(val[2]) and (500 >= abs(van) > 300):

                            if van > 0:
                                p = p - 5
                            else:
                                p = p + 5

                        elif int(val[0]) == int(val[1]) == int(val[2]) and abs(van) > 500:

                            if van > 0:
                                p = p - 5
                            else:
                                p = p + 5

                    if i > 3:
                        val.pop(0)

                    i += 1  
                    
                tf = time.time()
                print('Tiempo de ejecución: {} segundos'.format(tf-ti))
            
            lcoe_a.append(sum(lcoe_a)/len(lcoe_a))
            PPA_LCOE.loc['PPA/LCOE [USD/MWh]'] = lcoe_a
            PPA_LCOE = PPA_LCOE.reset_index()
            
            resultado = pd.concat([ gen_anual_salida, PPA_salida,flujo_salida,PPA_LCOE],axis = 0, ignore_index = True)
            resultado.to_excel(archivo, sheet_name = 'LCOE {}% PPA {}'.format(round(tasa_nominal*100,1),barras_proyecto[barrita][0:5]))

    archivo.save()
    
    
def cuenta_barra(datos,columna):
    barras = [1]
    vector_temporal = datos.fillna(0).iloc[:,columna]
    for i in range(13,16):
        if vector_temporal[i] != 0:
            barras.append(datos.iloc[i,columna])
    return len(barras)


if __name__ == '__main__':

    datos = importar_datos('Entrada.xlsx')


    with multiprocessing.Pool() as pool:   
         results = [pool.starmap(CalculoPPACorte, [(columna, barrita) for columna in range(len(datos.columns)) for barrita in range(cuenta_barra(datos,columna))])]





