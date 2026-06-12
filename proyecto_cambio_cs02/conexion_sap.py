CONEXION_SAP = input("ingrese la ruta de conexion")

DESCONEXION_SAP = input("INGRESE RUTA DE DESCONEXION")

def conectar_sap():
    print("conectado a sap por medio del log")
    def desconectar_sap():
        print("desconectado de sap por el log")
        return DESCONEXION_SAP
    return CONEXION_SAP, desconectar_sap()
