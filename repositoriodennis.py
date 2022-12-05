import pandas as pd

archivo=r'C:\Users\denni\Downloads\BASE DE DATOS CAMPAÑA PROXIMOS A VENCER DICIEMBRE 2022.xlsx'
df = pd.read_excel(archivo, sheet_name='PROXIMOS A VENCER DICIEMBRE', engine='openpyxl')

filtered_df = df[df['RESPONSABLE']=='CAMILO MARTINEZ']

list_name = list(filtered_df['NOMBRE'])

list_soat = list(filtered_df['FECHA VENCIMIENTO DEL SOAT '])

fecha=[]
for i in list_soat:
    list_fecha = i.strftime('%d / %m / %Y')
    fecha.append(list_fecha)

list_placa = list(filtered_df['CEDULA'])

list_correo= list(filtered_df['CORREO'])

list_comp = list(filtered_df['ASEGURADORA'])


for i in range (0,len(list_name)):
        corre=list_correo[i]
        name=list_name[i]
        date=fecha[i]
        placa=list_placa[i]
        firm=list_comp[i]
        print("Correo:",corre,"\n" " Buen día Sr", name, "\n" "Seguros Bolivar y Davivienda pensando en su protección quiere brindarle una alternativa de póliza para su vehículo de placa", placa, "que se encuentra actualmente con la compañía", firm ,"con fecha de vigencia" , date,"Al tener el crédito de vehículo le estamos dando el beneficio para que su póliza sea descontada mensualmente sin ninguna tasa de interés de financiación junto con la cuota del crédito del banco Davivienda. La primera cuota se estaría efectuando en 30 días de acuerdo a su fecha de facturación. Si desea conocer sobre las diferentes alternativas de póliza y sus servicios de asistencia puede contestar este correo y con gusto me comunicaré con usted." "\n" "Quedo pendiente de sus comentarios." "\n")