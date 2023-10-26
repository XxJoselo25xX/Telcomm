import mysql.connector

class Registro_datos:
    def __init__(self):
        self.conexion = mysql.connector.connect(
            host='localhost',
            database='base_datos',
            user='root',
            password='admin'
        )
        self.cursor = self.conexion.cursor()  # Inicializa el cursor aquí

    def inserta_producto(self, numero, nombre, unidad, puesto):
        sql = '''INSERT INTO productos (NUMERO, NOMBRE, UNIDAD, PUESTO) 
                 VALUES(%s, %s, %s, %s)'''
        values = (numero, nombre, unidad, puesto)
        self.cursor.execute(sql, values)
        self.conexion.commit()

    def buscar_productos(self):
        sql = "SELECT * FROM productos"
        self.cursor.execute(sql)
        registro = self.cursor.fetchall()
        return registro

    def busca_producto(self, numero_producto):
        sql = "SELECT * FROM productos WHERE NUMERO = {}".format(numero_producto)
        self.cursor.execute(sql)
        numerox = self.cursor.fetchall()
        return numerox

    def elimina_productos(self, numero):
        sql = '''DELETE FROM productos WHERE NUMERO = {}'''.format(numero)
        self.cursor.execute(sql)
        nom = self.cursor.rowcount
        self.conexion.commit()
        return nom

    def actualiza_productos(self, numero, nombre, unidad, puesto):
        sql = '''UPDATE productos SET  NOMBRE ='{}' , UNIDAD = '{}', PUESTO = '{}'
                 WHERE NUMERO = '{}' '''.format(nombre, unidad, puesto, numero)
        self.cursor.execute(sql)
        act = self.cursor.rowcount
        self.conexion.commit()
        return act
