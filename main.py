# from selenium import webdriver
# from selenium.webdriver.common.by import By
# import pandas as pd
# from datetime import date
# import time
# import openpyxl
# import xlsxwriter
#
# driver = webdriver.Chrome()
#
# # Lista de URLs a extraer
# # url_list = [ "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=32025219",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0016422042",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=29000420",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=29000422",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0014101624",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0014719218",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0029320333",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=29000702",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=29002120",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0019650753",
# #     "http://www.unicajainmuebles.com/listadoCoincidencias.do?referencia=0141XX"
# #              ]
#
# # Lee el archivo Excel y obtiene los URLs de la columna "Referencia"
# df = pd.read_excel('Inmueblesdisponibles20-03-2023.xlsx', sheet_name='InmueblesVentaConApis', usecols=['Enlace Ficha Comercial Portal Inmobiliario'])
#
# # Convierte los URLs en una lista
# url_list = df['Enlace Ficha Comercial Portal Inmobiliario'].tolist()
#
#
#
#
# # Lista para almacenar los datos extraídos de todas las páginas
# data = []
#
# # Recorrer cada URL de la lista
# for url in url_list:
#     # Navegar a la URL
#     driver.get(url)
#     time.sleep(15)
#
#     try:
#         # Obtener los datos de la página
#         title = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[1]")
#         provincia = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[1]/label[2]")
#         municipio = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[2]/label[2]")
#         codigoPostal = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[3]/label[2]")
#         tipo = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[4]/label[2]")
#         superficie = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[5]/label[2]")
#         price = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[2]/div/div[1]/div/div")
#         descripcion = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[4]/div")
#         main_photo = driver.find_element(By.XPATH, "//*[@id='listado']/div/div[1]/div[2]/img")
#         image_source = main_photo.get_attribute("src")
#
#         elements = title + provincia + municipio + codigoPostal + tipo + superficie + descripcion + price + [image_source]
#
#         # Crear un diccionario de datos para esta página
#         page_data = {}
#
#         # Verificar si se ha extraído información de cada elemento y agregarla al diccionario de datos
#         page_data["Title"] = title[0].text if title else ""
#         page_data["Provincia"] = provincia[0].text if provincia else ""
#         page_data["Municipio"] = municipio[0].text if municipio else ""
#         page_data["CodigoPostal"] = codigoPostal[0].text if codigoPostal else ""
#         page_data["Tipo"] = tipo[0].text if tipo else ""
#         page_data["Superficie"] = superficie[0].text if superficie else ""
#         page_data["Price"] = price[0].text if price else ""
#         page_data["Main Photo"] = image_source
#
#         # Agregar el diccionario de datos a la lista de datos
#         data.append(page_data)
#
#         # Imprimir los datos en la consola
#         for e in elements:
#             if isinstance(e, webdriver.remote.webelement.WebElement):
#                 print(e.text)
#             else:
#                 print(e)
#
#     except:
#         print("No se pudo extraer información de la URL:", url)
#         data.append({})  # Agregar un diccionario vacío a la lista de datos para esta URL
#         continue
#
# # Cerrar el navegador
# driver.quit()
#
# # Crear un DataFrame de pandas a partir de la lista de datos
# df = pd.DataFrame(data)
#
# # Escribir el DataFrame en un archivo de Excel
# file_name = "real_estate_data.xlsx"
# writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
# df.to_excel(writer, index=False)
#
# # Obtener el objeto workbook y worksheet
# workbook = writer.book
# worksheet = writer.sheets['Sheet1']
#
# # Obtener el número de filas y columnas
# num_rows, num_cols = df.shape
#
# # Iterar sobre las celdas del DataFrame
# for row in range(num_rows):
#     for col in range(num_cols):
#         cell_value = df.iloc[row, col]
#         if pd.isna(cell_value):
#             # Si el valor de la celda es NaN, dejarla en blanco en el archivo de Excel
#             worksheet.write_blank(row + 1, col, None)
#         else:
#             worksheet.write(row + 1, col, cell_value)
#
# # Guardar el archivo de Excel
# writer._save()




from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

# Configurar el navegador
driver = webdriver.Chrome()


# Lee el archivo Excel y obtiene los URLs de la columna "Referencia"
df = pd.read_excel('Inmueblesdisponibles20-03-2023.xlsx', sheet_name='InmueblesVentaConApis', usecols=['Enlace Ficha Comercial Portal Inmobiliario'])

# Convierte los URLs en una lista
url_list = df['Enlace Ficha Comercial Portal Inmobiliario'].tolist()

# Lista para almacenar los datos extraídos de todas las páginas
data = []

# Recorrer cada URL de la lista
for url in url_list:
    # Navegar a la URL
    driver.get(url)
    time.sleep(5)

    try:
        # Obtener los datos de la página
        title = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[1]")
        provincia = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[1]/label[2]")
        municipio = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[2]/label[2]")
        codigoPostal = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[3]/label[2]")
        tipo = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[4]/label[2]")
        superficie = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[3]/div[5]/label[2]")
        price = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[2]/div/div[1]/div/div")
        descripcion = driver.find_elements(By.XPATH, "//*[@id='listado']/div/div[1]/div[4]/div")
        main_photo = driver.find_element(By.XPATH, "//*[@id='listado']/div/div[1]/div[2]/img")
        image_source = main_photo.get_attribute("src")

        elements = title + provincia + municipio + codigoPostal + tipo + superficie + descripcion + price + [image_source]

        # Crear un diccionario de datos para esta página
        page_data = {}

        # Agregar los datos al diccionario de datos
        page_data["Title"] = title[0].text if title else ""
        page_data["Provincia"] = provincia[0].text if provincia else ""
        page_data["Municipio"] = municipio[0].text if municipio else ""
        page_data["CodigoPostal"] = codigoPostal[0].text if codigoPostal else ""
        page_data["Tipo"] = tipo[0].text if tipo else ""
        page_data["Superficie"] = superficie[0].text if superficie else ""
        page_data["Price"] = price[0].text if price else ""
        page_data["Main Photo"] = image_source

        # Agregar el diccionario de datos a la lista de datos
        data.append(page_data)

        for e in elements:
                if isinstance(e, webdriver.remote.webelement.WebElement):
                    print(e.text)
                else:
                    print(e)

    except:
        print("No se pudo extraer información de la URL:", url)
        data.append({})  # Agregar un diccionario vacío a la lista de datos para esta URL
        continue

    # Guardar los datos en un archivo de Excel cada 100 filas
    if len(data) % 100 == 0:
        # Crear un DataFrame de pandas a partir de la lista de datos
        df = pd.DataFrame(data)

        # Escribir el DataFrame en un archivo de Excel
        file_name = "data_" + str(len(data) // 100) + ".xlsx"
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        df.to_excel(writer, index=False)

        # Obtener el objeto workbook y worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']


        # Iterar sobre las columnas y ajustar el ancho
        for i, col_name in enumerate(df.columns):
            col_width = max(len(str(col_name)), df[col_name].astype(str).map(len).max())
            worksheet.set_column(i, i, col_width)

        # Guardar el archivo de Excel
        writer._save()

# Cerrar el navegador
driver.quit()


