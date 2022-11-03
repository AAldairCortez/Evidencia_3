import random
import openpyxl
import datetime
import sys
import sqlite3
from sqlite3 import Error

turno_dict = {1:"Matutino", 2:"Vespertino", 3:"Nocturno"}
id_val_cliente = []
id_val_sala = []
fecha_evento = []
encontradas = []
disponibles = []

libro = openpyxl.Workbook()
hoja = libro["Sheet"]
hoja.title = "PRIMERA"

try:
    with sqlite3.connect("Evidencia_3.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
        mi_cursor = conn.cursor()
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS cliente (id_cliente INTEGER PRIMARY KEY, nombre_cliente TEXT NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS sala (id_sala INTEGER PRIMARY KEY, nombre_sala TEXT NOT NULL, cupo TEXT NOT NULL);")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS evento (folio INTEGER PRIMARY KEY, nombre_evento TEXT NOT NULL, turno INTEGER NOT NULL, fecha_evento timestamp, id_cliente INTEGER NOT NULL, id_sala INTEGER NOT NULL, FOREIGN KEY(id_cliente) REFERENCES cliente(id_cliente), FOREIGN KEY(id_sala) REFERENCES sala(id_sala));")
########################************INSERTS EXTRA ************########################
        mi_cursor.execute("INSERT INTO cliente VALUES (1, 'Jesus Almada');") 
        mi_cursor.execute("INSERT INTO cliente VALUES (2, 'Armando Cosas');") 
        mi_cursor.execute("INSERT INTO cliente VALUES (3, 'Oscar Corleone');") 
        mi_cursor.execute("INSERT INTO sala VALUES (4, 'Sala 1', 345);") 
        mi_cursor.execute("INSERT INTO sala VALUES (5, 'Sala 2', 543);") 
        mi_cursor.execute("INSERT INTO sala VALUES (6, 'Sala 3', 987);") 
        print("Tablas creadas exitosamente")

except Error as e:
    print (e)
except:
    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
finally:
    conn.close()

def agregar_evento():
    global evento
    print("\nRegistro de un evento")
    print("*" *36) 
    print("Revisaremos que seas un usuario registrado")

    try:
        with sqlite3.connect("Evidencia_3.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("SELECT * FROM cliente")
            busqueda_cliente = mi_cursor.fetchall()
            mi_cursor.execute("SELECT * FROM sala")
            busqueda_sala = mi_cursor.fetchall()                       
        for filtro_idcliente, filtro_nomcliente in busqueda_cliente:
            id_val_cliente.append(filtro_idcliente)
            continue
        print(busqueda_cliente)
        r_Cliente = int(input("Ingresa tu clave: "))
        if r_Cliente in id_val_cliente:  
            for filtro_idsala, filtro_nomsala, filtro_cupo_sala in busqueda_sala:
                id_val_sala.append(filtro_idsala)
                continue
            print(busqueda_sala)
            r_Sala = int(input("Ingresa el ID de la sala que quieres usar: "))
            if r_Sala in id_val_sala:
                nombreEvento=input("Ingresa el nombre del evento: ")
                turno=int(input("Ingresa un turno (1:Matutino, 2:Vespertino, 3:Nocturno): "))
                if turno in turno_dict.keys():
                            fechaEvento=input("Ingresa la fecha del evento en formato dd/mm/aaaa: ")
                            fechaEvento = datetime.datetime.strptime(fechaEvento,"%d/%m/%Y").date()
                            fecha_actual =datetime.date.today()
                            diasAntes = fechaEvento.day - fecha_actual.day
                            mi_cursor.execute("SELECT * FROM evento")
                            busqueda_evento = mi_cursor.fetchall()
                            if busqueda_evento:
                                print(busqueda_evento)
                                for folio_consulta, nevento_consulta, turno_consulta, fecha_consulta, r_Cliente_consulta, r_Sala_consulta in busqueda_evento:
                                    fecha_evento.append(turno_consulta, fecha_consulta, r_Sala_consulta)
                                    continue
                                print(fecha_evento)
                                if fechaEvento in fecha_evento:
                                    print("\n**La fecha y turno no estan disponibles para ese dia, por favor selecciona otra**\n")
                                else:
                                        if diasAntes < 2:
                                            print("\n**Para reservar una fecha debe hacerlo con al menos 2 dias de anticipación\n**")     
                                        elif diasAntes >= 2:
                                            folio = (random.randint(1000,9999))
                                            valores = (folio, nombreEvento, turno, fechaEvento, r_Cliente, r_Sala)
                                            mi_cursor.execute("INSERT INTO evento VALUES(?,?,?,?,?,?)", valores)
                                            print("\n**Su reservación ha sido éxitosa**")
                            else:
                                folio = (random.randint(1000,9999))
                                valores_evento = (folio, nombreEvento, turno, fechaEvento, r_Cliente, r_Sala)
                                print(valores_evento)
                                mi_cursor.execute("INSERT INTO evento VALUES(?,?,?,?,?,?)", valores_evento)
                                print("\n**Su reservación ha sido éxitosa**")
                else:
                    print("\n*Turno fuera de los disponibles, por favor ingrese un turno valido*\n") 
            else:
                print("\n**No existe esa sala**\n")
        else:
            print("\n**No estas aun registrado como cliente**\n")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()  

def editarReservacion():
    global evento
    print("\nEdita el nombre de un evento")
    print("*" *36)

    try:
        with sqlite3.connect("Evidencia_3.db") as conn:
            mi_cursor = conn.cursor()
            editar_evento=int(input("Ingrese el folio de su evento: "))
            criterios = {"folio":editar_evento}
            mi_cursor.execute("SELECT * FROM evento WHERE folio = :folio", criterios)
            folio_editar = mi_cursor.fetchall()
            print(f"Reserva a cambiar: {folio_editar}")
            nuevo_nevento=input("Nuevo nombre del evento reservado : ")
            nuevo_criterio = {"nombre_evento":nuevo_nevento} 
            mi_cursor.execute("UPDATE evento SET nombre_evento = :nombre_evento WHERE folio = :folio", nuevo_criterio, criterios)
            print("**Cambio realizado** ")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
        
def consultar():
    print("\nConsulta de reservaciones")
    print("*" *54)

    fecha_cons = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
    fecha_cons = datetime.datetime.strptime(fecha_cons,"%d/%m/%Y").date()

    try:
        with sqlite3.connect("Evidencia_3.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            criterio = {"fecha":fecha_cons}
            mi_cursor.execute("SELECT * FROM evento WHERE DATE(fecha_evento) = :fecha;", criterio)
            bus_fecha = mi_cursor.fetchall()

            print("\n")
            print("**"*34)
            print("**" + " "*8 + f" REPORTE DE RESERVACIONES PARA EL DÍA {fecha_cons}" + " " *8 + "**")
            print("**"*34)
            print("{:<15} {:<15} {:<15} {:<15}".format('SALA','NOMBRE','EVENTO', 'TURNO' ))
            print("**"*34)
            for folio, nombre_evento, turno, fecha_evento, id_cliente, id_sala in bus_fecha():
                    if fecha_cons == fecha_evento:
                        print("{:<15} {:<15} {:<15} {:<15}".format (id_sala, id_cliente , nombre_evento, turno ))
            print("*"*25 + " FIN DEL REPORTE  " + "*"*25)
                    
    except ValueError:
        print(f"\n**El valor proporcionado no es compatible con la operación solicitada**\n")
        
def agregar_cliente():
    global cliente
    print("\nRegistro de un cliente")
    print("*" *36)
    while True:
        nombreCliente=input("Introduce el nombre: ").title()
        if nombreCliente.strip() == "":
            print("*El nombre no puede quedar vacio, por favor proporcione uno*")
            continue
        else:
            try:
                with sqlite3.connect("Evidencia_3.db") as conn:
                    mi_cursor = conn.cursor()
                    id_cliente = (random.randint(1000,9999))
                    valores = (id_cliente, nombreCliente)
                    mi_cursor.execute("INSERT INTO cliente VALUES (?,?)", valores) 
                    print("**Registro echo**")
            except Error as e:
                print (e)
            except:
                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
            finally:
                conn.close()
            break

def registroSala():
    global sala
    print("\nRegistro de una sala")
    print("*" *36)
    reg_sala = True
    while reg_sala: 
        nombreSala=input("Introduce el nombre de la sala: ").title()
        if nombreSala.strip() == "":
            print("\n*El nombre no puede quedar vacio, por favor proporcione uno*")
            continue
        while reg_sala:   
            try: 
                cupoSala=int(input("Introduce el cupo de la sala: "))
                if cupoSala == 0:
                    print("\n**El cupo de la sala no puede ser 0**")
                    continue
                else:
                    try:
                        with sqlite3.connect("Evidencia_3.db") as conn:
                            mi_cursor = conn.cursor()
                            id_sala = (random.randint(1000,9999))
                            sala_valores = (id_sala, nombreSala, cupoSala)
                            print(sala_valores)
                            mi_cursor.execute("INSERT INTO sala VALUES (?,?,?)", sala_valores) 
                            print("**Registro echo**")
                    except Error as e:
                        print (e)
                    except:
                        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                    finally:
                        conn.close()
                        reg_sala = False
                    break
            except ValueError:
                print("**La respuesta no es valida**")

def rep_fechas():
    print("\nReporte de reservaciones")
    print("*" *36)
    fecha_consulta = input("Ingresa la fecha del evento en formato dd/mm/aaaa: ")
    fecha_consulta = datetime.datetime.strptime(fecha_consulta,"%d/%m/%Y").date()
    
def exp_reporte():
    print("\nReporte de reservaciones")
    print("*" *36)
    fecha_solicitada = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
    fecha_solicitada = datetime.datetime.strptime(fecha_solicitada,"%d/%m/%Y").date()

    try:
        with sqlite3.connect("Evidencia_3.db", detect_types = sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            criterio = {"fecha":fecha_solicitada}
            mi_cursor.execute("SELECT * FROM evento WHERE DATE(fecha_evento) = :fecha;", criterio)
            busqueda_fecha = mi_cursor.fetchall()

            if busqueda_fecha:
                evento_parte=[(id_sala, id_cliente, nombre_evento, turno)]
                hoja["B1"].value = f"REPORTE DE RESERVACIONES PARA EL DÍA {fecha_solicitada}"
                hoja["A2"].value = "SALA"
                hoja["B2"].value = "CLIENTE"
                hoja["C2"].value = "EVENTO"
                hoja["D2"].value = "TURNO"
                for folio, nombre_evento, turno, fecha_evento, id_cliente, id_sala in busqueda_fecha():
                    evento_parte=[(id_sala, id_cliente, nombre_evento, turno)]
                    for valor in evento_parte:
                        hoja.append(valor)
                    libro.save("ExcelEvidencia3.xlsx")
                    print("**Libro creado**")
            else:
                print("**No existe el evento**")

    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
    
def eli_reserva():
    print("\nReporte de reservaciones")
    print("*" *36)
    try:
        with sqlite3.connect("Evidencia_3.db") as conn:
                mi_cursor = conn.cursor()
                evento_eliminar=int(input("Ingrese el folio del evento a eliminar: "))
                criterios = {"folio":evento_eliminar}
                mi_cursor.execute("SELECT * FROM evento WHERE folio = :folio", criterios)
                folio_eliminar = mi_cursor.fetchall()
                print(f"Reserva a eliminar: {folio_eliminar}")
                mi_cursor.execute("DELETE FROM evento WHERE folio = :folio", criterios)
                print("**Cambio realizado** ")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
        
def sub_menu_reserva():
    while True:
        print("\n**MENU RESERVACION DE UN EVENTO**")
        print("*" *36 )
        print("1 - Registrar nueva reservacion.")
        print("2 - Modificar descripcion de una reservacion.")
        print("3 - Consultar disponibilidad de salas para una fecha.")
        print("4 - Eliminar una reservacion.")
        print("5 - Salir")
        respuesta_reserva = input("\nIndique la opcion deseada: ")
        try:
            respuesta_int2 = int(respuesta_reserva)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int2 == 1:
            agregar_evento()

        elif respuesta_int2 == 2:
            editarReservacion()

        elif respuesta_int2 == 3:
            rep_fechas()

        elif respuesta_int2 == 4:
            eli_reserva() 

        elif respuesta_int2 == 5:
            break

        else: 
            print("\n*Su respuesta no corresponde con ninguna de las opciones*.")

def reportes():
    while True:
        print("\n**MENU REPORTES**")
        print("*" *36)
        print("1 - Reporte en pantalla de reservaciones para una fecha.")
        print("2 - Exportar reporte tabular en Excel.")
        print("3 - Salir.")
        respuesta_reportes = input("\n Indique la opcion deseada: ")
        
        try:
            respuesta_int3 = int(respuesta_reportes)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int3 == 1:
            consultar()

        elif respuesta_int3 == 2:
            exp_reporte()

        elif respuesta_int3 == 3:
            break

def menu():
    while True:
        print("\n**MENU DE OPERACIONES**")
        print("*" *36 )
        print("1 - Reservaciones")
        print("2 - Reportes.")
        print("3 - Registrar un cliente")
        print("4 - Registrar una sala ")
        print("5 - Salir")
        respuesta = input("\nIndique la opcion deseada: ")

        try:
            respuesta_int = int(respuesta)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int == 1:
            sub_menu_reserva()

        elif respuesta_int == 2:
            reportes()

        elif respuesta_int == 3:
            agregar_cliente()

        elif respuesta_int == 4: 
            registroSala()

        elif respuesta_int == 5:
            print("\n**TERMINO EL MENU DE OPERACIONES**")
            print("*" *36)
            break
        else: 
            print("\n*Su respuesta no corresponde con ninguna de las opciones*.")

menu()

