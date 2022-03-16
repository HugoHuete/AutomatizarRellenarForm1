# Creado por Hugo Huete Arroliga
# Fecha: 5/3/2022

from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd


def main():
    datos = pd.read_excel("Registros.xlsx", "Ingresos")

    global driver
    driver = webdriver.Chrome()
    sleep(2)
    driver.get("https://commcarehq.org")
    sleep(4)
    login("alguien@correo.com", "contraseña")

    id_inicial = 30
    print("\n")
    for num_fila in range(len(datos)):
        seleccionar_usuario(datos["Codigo"][num_fila])
        id_inicial = rellenar(datos, num_fila, id_inicial + 120)


def login(username, password):
    # Ingresar usuario
    usuario_login = driver.find_element(By.ID, "id_auth-username")
    usuario_login.clear()
    usuario_login.send_keys(username)

    # Ingresar contraseña
    contraseña = driver.find_element(By.ID, "id_auth-password")
    contraseña.clear()
    contraseña.send_keys(password)

    # Click en boton
    boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
    boton_login.click()
    sleep(2.5)

    # Seleccionar boton de COMMCARE
    commcare = driver.find_element(
        By.XPATH, "//div[@aria-label='COMMCARE MEAL PROGRESA CARIBE']")
    commcare.click()
    sleep(0.8)

    # Seleccionar sistema de recolección
    sistema_de_recoleccion = driver.find_element(
        By.XPATH, "//div[@class='col-xs-6 col-sm-4 col-md-3']")
    sistema_de_recoleccion.click()
    sleep(2)


def seleccionar_usuario(usuario):
    """
    Busca en la lista de usuarios el usuario a ingresar y abrir formulario
    """
    
    # Registro de bienes y servicios
    print("Ingresando al registro de bienes y servicios...", end = "")
    intentos = 1
    while True:
        try:
            registro_bienes_y_servicios = driver.find_element(
            By.XPATH, "//*[contains(@style,'module0_form2_en.png')]")
            registro_bienes_y_servicios.click()
            print("Exitoso.")
            sleep(1)
            break
        except:
            print(f"\nIntento #{intentos} de entrar al registro de bienes y servicios resultó fallido. Reintentando...")
            intentos += 1
            sleep(1)
            continue

    # Buscar usuario en el searchtext
    print(f"Buscando al usuario {usuario}... ")
    intentos = 1
    while True:
        try:
            buscar_usuario = driver.find_element(By.ID, "searchText")
            buscar_usuario.clear()
            buscar_usuario.send_keys(f'"{usuario}"')

            boton_buscar = driver.find_element(By.ID, "case-list-search-button")
            boton_buscar.click()
            sleep(2)
            break
        except:
            print(f"Intento #{intentos} de ingresar usuario en el searchtext resultó fallido. Reintentando...")
            sleep(1)
            intentos += 1
            continue


    # Buscar usuario en los resultados
    intentos = 1
    while True:
        try:
            usuario_correcto = driver.find_element(
            By.XPATH, f"//td[contains(text(),{usuario})]")
            usuario_correcto.click()
            print(f"Busqueda del usuario {usuario} terminada. Ahora intentando ingresar al formulario.")
            sleep(1)
            break
        except:
            print(f"Intento #{intentos} de encontrar al usuario {usuario} resultó fallido. Reintentando...")
            intentos += 1
            sleep(1)
            continue

    # Probar ingresar al formulario
    intentos = 1
    while True:
        try:
            seleccionar = driver.find_element(By.ID, "select-case")
            seleccionar.click()
            print("Ingreso al formulario terminado.")
            sleep(2)
            break
        except:
            print(f"Intento #{intentos} de ingresar al formulario resultó fallido. Reintentando...")
            intentos += 1
            sleep(1)
            continue




def rellenar(datos, num_fila, id_inicial):
    """
    Busca cada uno de los elementos del formulario e ingresa el valor que corresponde
    """

    usuario = datos['Codigo'][num_fila]
    # Fecha capacitacion

    print(f"Rellenando formulario del usuario {usuario}... Buscando el id inicial en el html...")

    fecha = datos["Fecha de entrega"][num_fila]
    text_fecha = f"{fecha.month}/{fecha.day}/{fecha.year}"
    for i in range(1000):
        try:
            id_inicial += 1
            campo_fecha = driver.find_element(By.ID, f"date{id_inicial}")
            campo_fecha.clear()
            campo_fecha.send_keys(text_fecha)
            sleep(0.5)
            campo_fecha.send_keys(Keys.ENTER)
            sleep(1.5)
            break
        except:
            continue

    print(f"id_inicial encontrado con el número: {id_inicial}")

    # Seleccionar tipo participante
    # tp_part_opc = {"Técnico": 0, "Promotor": 1, "Productor": 2, "Prestador de servicios": 3,
    #                "Administracióndecooperativa": 4, "Centro de acopio cooperativa": 5, "Otro": 6}

    # while True:
    #     try:
    #         campo_tp_part = driver.find_element(
    #             By.ID, f"group-select{id_inicial-1}-choice-{tp_part_opc[datos['Tipo de participante'][num_fila]]}")
    #         campo_tp_part.click()
    #         break
    #     except:
    #         campo_fecha = driver.find_element(By.ID, f"date{id_inicial}")
    #         campo_fecha.send_keys(Keys.ENTER)
    #         sleep(1)
    #         continue

    # Tipo de bien o servicio
    tp_bien_opc = { "Material genético": 0, 
                    "Kits de insumos para la implementación de los planes de fertilización y fitosanitarios": 1,
                    "Kits de herramientas para la implementación de mejores prácticas para el manejo de las plantaciones de cacao": 2,
                    "Materiales didácticos y capacitación": 3,
                    "Injertación":5,
                    "Parcelas demostrativas":7,
                    "Planes de finca":8,
                    "Otro": 9}

    campo_tp_capt = driver.find_element(
        By.ID, f"group-select{id_inicial+1}-choice-{tp_bien_opc[datos['TIPO DE BIEN O SERVICIO'][num_fila]]}")
    campo_tp_capt.click()

    if datos["TIPO DE BIEN O SERVICIO"][num_fila] == "Otro":
        while True:
            sleep(1)
            try:
                campo_rol_facil_text = driver.find_element(
                By.ID, f"str{id_inicial+28}")
                campo_rol_facil_text.send_keys(
                datos["Otro tipo de bien"][num_fila])
                break
            except:
                continue

    # Descripción del bien recibido
    campo_descrip_bien = driver.find_element(By.ID, f"str{id_inicial+3}")
    campo_descrip_bien.send_keys(datos["Descripcion"][num_fila])

    # Unidad medida
    campo_und_medida = driver.find_element(
        By.ID, f"group-select{id_inicial+4}-choice-0")
    campo_und_medida.click()

    # Cantidad
    campo_cant = driver.find_element(By.ID, f"float{id_inicial+5}")
    campo_cant.send_keys(str(datos["Cantidad"][num_fila]))

    # Costo por und en córdobas
    campo_costo_cord = driver.find_element(By.ID, f"float{id_inicial+6}")
    campo_costo_cord.send_keys(
        str(datos["Costo / Unidad en cordobas C$"][num_fila]))
    campo_costo_cord.send_keys(Keys.ENTER)

    # Tipo de cambio
    campo_tipo_camb = driver.find_element(By.ID, f"float{id_inicial+8}")
    campo_tipo_camb.send_keys(str(datos["Tipo de Cambio"][num_fila]))
    campo_tipo_camb.send_keys(Keys.ENTER)


    # Aporte del productor
    campo_aport_prod = driver.find_element(By.ID, f"float{id_inicial+10}")
    campo_aport_prod.send_keys(str(datos["Aporte del productor US$"][num_fila]))

    # Aporte de otros
    campo_aport_otros = driver.find_element(By.ID, f"float{id_inicial+11}")
    campo_aport_otros.send_keys(str(datos["Aporte de otros"][num_fila]))
    campo_aport_otros.send_keys(Keys.ENTER)
    sleep(1)


    # Submit
    boton_submit = driver.find_element(
        By.XPATH, "//button[@class='submit btn btn-primary']")
    boton_submit.click()
    print(f"Ingreso de datos del {usuario} fue finalizado exitosamente..!!!!\n\n")
    sleep(1)
    return id_inicial


if __name__ == "__main__":
    main()
