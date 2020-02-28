import pyodbc
import shutil

direccion_servidor = '10.120.25.80'
nombre_bd = 'AnovoASR'
nombre_usuario = 'ENVIRONMENT_PRD'
password = '@env-PRD-2015$#'

try:
    conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                              direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
    print("Conectado")

    with conexion.cursor() as cursor:
        cursor.execute("EXEC usp_OrdenesClaimsInvalidacionGarantia")
        pendientes = cursor.fetchall()
        
        for gspn,ost,ano,glosa in pendientes:
            print(str(ost) + ' ' + str(ano))

            sqlfile = ("exec usp_claims_obtener_rutapdfdoa ?,?")
            valfile = (ost,ano)
            cursor.execute(sqlfile,valfile)
            ejecutar=cursor.fetchone()

            rutapdf = ''.join(str(e) for e in ejecutar)
            shutil.copy(rutapdf, 'C:\\evidenciasPDF')
            print("PDF Obtenido")

        else:
            print("No hay datos")

except Exception as e:
    print("Ocurrió un error al conectar a SQL Server: ", e)

finally:
    conexion.close()
    print("Conexion Cerrada")