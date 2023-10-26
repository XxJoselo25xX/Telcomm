from conexionBD import *
from PyQt5.QtWidgets import QMainWindow, QApplication, QTableWidgetItem, QFileDialog
from PyQt5.QtGui import QPixmap
from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt, QByteArray, QBuffer, QIODevice, QDateTime, QTimer, QTime
import mysql.connector
import array
import openpyxl
import os
from bg_rc import *
class MiApp(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi('GUI.ui', self) 
          # Diccionario para realizar un seguimiento de los registros recientes
        self.registros_recientes = {}


                 # Crea un temporizador para actualizar la hora cada segundo
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.actualizar_hora)
        self.timer.start(10000)  # El temporizador se activará cada 1000 milisegundos (1 segundo)

        # Llama a la función para mostrar la hora actual
        self.actualizar_hora()

        self.datosTotal = Registro_datos()
        self.bt_refrescar.clicked.connect(self.m_productos)
        self.bt_agregar.clicked.connect(self.insert_productos)
        self.bt_buscar.clicked.connect(self.buscar_producto)
        self.bt_borrar.clicked.connect(self.eliminar_producto)
        self.bt_actualizar.clicked.connect(self.modificar_productos)
        self.ln_buscar_img.returnPressed.connect(self.search_data)
        self.pushButton.clicked.connect(self.load_image)
        self.btn_clear.clicked.connect(self.clear_data)
        self.btn_save.clicked.connect(self.save_data)

        self.tabla_productos.setColumnWidth(0, 98)
        self.tabla_productos.setColumnWidth(1, 250)
        self.tabla_productos.setColumnWidth(2, 300)
        self.tabla_productos.setColumnWidth(3, 300)

        self.tabla_borrar.setColumnWidth(0, 98)
        self.tabla_borrar.setColumnWidth(1, 250)
        self.tabla_borrar.setColumnWidth(2, 300)
        self.tabla_borrar.setColumnWidth(3, 300)
        self.tabla_borrar.setColumnWidth(4, 300)

        self.tabla_buscar.setColumnWidth(0, 98)
        self.tabla_buscar.setColumnWidth(1, 250)
        self.tabla_buscar.setColumnWidth(2, 300)
        self.tabla_buscar.setColumnWidth(3, 300)
        self.tabla_buscar.setColumnWidth(4, 300)

    def m_productos(self):
        datos = self.datosTotal.buscar_productos()
        i = len(datos)
        self.tabla_productos.setRowCount(i)
        for tablerow, row in enumerate(datos):
            self.tabla_productos.setItem(tablerow, 0, QTableWidgetItem(str(row[1])))
            self.tabla_productos.setItem(tablerow, 1, QTableWidgetItem(row[2]))
            self.tabla_productos.setItem(tablerow, 2, QTableWidgetItem(row[3]))
            self.tabla_productos.setItem(tablerow, 3, QTableWidgetItem(row[4]))

    def insert_productos(self):
        numero = self.numeroA.text()
        nombre = self.nombreA.text()
        unidad = self.unidadA.text()
        puesto = self.puestoA.text()

        self.datosTotal.inserta_producto(numero, nombre, unidad, puesto)
        self.numeroA.clear()
        self.nombreA.clear()
        self.unidadA.clear()
        self.puestoA.clear()

    def modificar_productos(self):
        id_producto = self.id_producto.text()
        id_producto = str("'" + id_producto + "'")
        nombreXX = self.datosTotal.busca_producto(id_producto)

        if nombreXX is not None:
            self.id_buscar.setText("ACTUALIZAR")
            numeroM = self.numero_actualizar.text()
            nombreM = self.nombre_actualizar.text()
            unidadM = self.unidad_actualizar.text()
            puestoM = self.puesto_actualizar.text()

            act = self.datosTotal.actualiza_productos(numeroM, nombreM, unidadM, puestoM)
            if act == 1:
                self.id_buscar.setText("ACTUALIZADO")
                self.numero_actualizar.clear()
                self.nombre_actualizar.clear()
                self.unidad_actualizar.clear()
                self.puesto_actualizar.clear()
                self.id_producto.clear()
            elif act == 0:
                self.id_buscar.setText("ERROR")
            else:
                self.id_buscar.setText("INCORRECTO")
        else:
            self.id_buscar.setText("NO EXISTE")

    def buscar_producto(self):
        numero_producto = self.numeroB.text()
        numero_producto = str("'" + numero_producto + "'")

        datosB = self.datosTotal.busca_producto(numero_producto)
        i = len(datosB)

        self.tabla_buscar.setRowCount(i)
        for tablerow, row in enumerate(datosB):
            self.tabla_buscar.setItem(tablerow, 0, QTableWidgetItem(str(row[1])))
            self.tabla_buscar.setItem(tablerow, 1, QTableWidgetItem(row[2]))
            self.tabla_buscar.setItem(tablerow, 2, QTableWidgetItem(row[3]))
            self.tabla_buscar.setItem(tablerow, 3, QTableWidgetItem(row[4]))

    def eliminar_producto(self):
        eliminar = self.numero_borrar.text()
        eliminar = str("'" + eliminar + "'")
        resp = self.datosTotal.elimina_productos(eliminar)
        datos = self.datosTotal.buscar_productos()
        i = len(datos)
        self.tabla_borrar.setRowCount(i)
        for tablerow, row in enumerate(datos):
            self.tabla_borrar.setItem(tablerow, 0, QTableWidgetItem(str(row[1])))
            self.tabla_borrar.setItem(tablerow, 1, QTableWidgetItem(row[2]))
            self.tabla_borrar.setItem(tablerow, 2, QTableWidgetItem(row[3]))
            self.tabla_borrar.setItem(tablerow, 3, QTableWidgetItem(row[4]))

        if resp is None:
            self.borrar_ok.setText("NO EXISTE")
        elif resp == 0:
            self.borrar_ok.setText("NO EXISTE")
        else:
            self.borrar_ok.setText("SE ELIMINO")

    def search_data(self):
        numero_buscar = self.ln_buscar_img.text()

        # Establece la conexión a la base de datos MySQL
        conexion = mysql.connector.connect(
            host='localhost',
            user='root',
            password='admin',
            database='base_datos')

        # Crea un cursor para ejecutar la consulta SQL
        cursor = conexion.cursor()

        # Ejecuta la consulta SQL
        consulta = "SELECT * FROM productos WHERE NUMERO = %s"
        cursor.execute(consulta, (numero_buscar,))

        # Obtiene el resultado de la consulta
        producto = cursor.fetchone()  # Lee la primera fila de resultados

        # Guarda el registro en Excel si se encuentra un producto
        if producto:
            self.guardar_registro_en_excel(producto[1], producto[2], producto[3])  # Pasa el valor de UNIDAD como tercer argumento

            # Actualiza los widgets en tu interfaz gráfica con los resultados obtenidos de la base de datos
            self.numero.setText(f' {producto[1]}')
            self.nombre.setText(f' {producto[2]}')
            self.unidad.setText(f' {producto[3]}')
            self.puesto.setText(f' {producto[4]}')

            # 'foto' sería el campo de la base de datos donde se almacenan los datos de la imagen
            foto = QPixmap()
            foto.loadFromData(producto[5])  # Asegúrate de que [5] sea el índice correcto para el campo de imagen
            self.imagen.setPixmap(foto)
            self.ln_buscar_img.clear()

        else:
            self.nombre.setText(' NONE')
            self.unidad.setText(' NONE')
            self.puesto.setText(' NONE')
            self.numero.setText(' NONE')
            self.imagen.clear()
            self.ln_buscar_img.clear()

   

    def guardar_registro_en_excel(self, numero_empleado, nombre, unidad):
        current_time = QDateTime.currentDateTime()

        # Verifica si el número de empleado ya está en el diccionario y si el tiempo ha pasado
        if numero_empleado in self.registros_recientes:
            last_time, _ = self.registros_recientes[numero_empleado]
            time_difference = last_time.secsTo(current_time)
            if abs(time_difference) < 10:
                print("Registro duplicado, no se guardó en el archivo Excel.")
                return
        # Actualiza el tiempo para el número de empleado
        self.registros_recientes[numero_empleado] = (current_time, unidad)

        # Obtiene el año, mes, día y hora en formato de cadena
        year = current_time.toString("yyyy")
        month = current_time.toString("MM")
        day = current_time.toString("dd")
        hour = current_time.toString("hh:mm:ss")

        # Obtiene el mes y año actual para el archivo Excel
        month_year = current_time.toString("MM-yyyy")

        # Define el nombre del archivo Excel
        excel_filename = f"registros_{month_year}.xlsx"

        # Crea o abre el archivo Excel
        if not os.path.exists(excel_filename):
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.append(["Año", "Mes", "Día", "Hora", "Número", "Nombre", "Unidad"])
        else:
            workbook = openpyxl.load_workbook(excel_filename)
            worksheet = workbook.active

        # Obtiene la hora de búsqueda en formato de cadena
        search_time = current_time.toString(Qt.ISODate)
        # Agrega los datos al archivo si no están duplicados en el último intervalo de tiempo
        worksheet.append([year, month, day, hour, numero_empleado, nombre, unidad])
        workbook.save(excel_filename)
        workbook.close()
        print("Registro guardado en el archivo Excel.")


    def load_image(self):
        filename, _ = QFileDialog.getOpenFileName(filter="Image Files (*.jpg *.png);;All Files (*)")
        if filename:
            pixmapImage = QPixmap(filename).scaled(300, 300, Qt.AspectRatioMode.KeepAspectRatio, Qt.SmoothTransformation)
            self.img_prevew.setPixmap(pixmapImage)

    def clear_data(self):
        self.in_numero.clear()
        self.in_nombre.clear()
        self.in_unidad.clear()
        self.in_puesto.clear()
        self.img_prevew.clear()

    def save_data(self):
        numero = self.in_numero.text()
        nombre = self.in_nombre.text()
        unidad = self.in_unidad.text()
        puesto = self.in_puesto.text()
        foto = self.img_prevew.pixmap()
        if foto:
            bArray = QByteArray()
            buff = QBuffer(bArray)
            buff.setData(bArray)  # Asignar los datos a tu búfer
            buff.open(QIODevice.WriteOnly)
            foto.save(buff, "PNG")

            # Convertir QByteArray a bytes
            byte_array = bytes(bArray)

            if self.datosTotal.busca_producto(numero):
                self.img_prevew.setText('El producto ya existe')
            elif len(numero) <= 0:
                self.img_prevew.setText('Número inválido')
            elif len(nombre) <= 0:
                self.img_prevew.setText('Nombre inválido')
            elif len(unidad) <= 0:
                self.img_prevew.setText('Unidad inválida')
            elif len(puesto) <= 0:
                self.img_prevew.setText('Puesto inválido')
            elif foto:
                # Guardar en la base de datos
                sql = "INSERT INTO productos (NUMERO, NOMBRE, UNIDAD, PUESTO, FOTO) VALUES (%s, %s, %s, %s, %s)"
                values = (numero, nombre, unidad, puesto, byte_array)
                self.datosTotal.cursor.execute(sql, values)
                self.datosTotal.conexion.commit()
                self.img_prevew.setText('Producto guardado correctamente')
                self.clear_data()
            else:
                self.img_prevew.setText('No hay foto')




        

    def actualizar_hora(self):
        # Obtiene la hora actual en formato de texto (solo horas y minutos)
        hora_actual = QTime.currentTime().toString('hh:mm')
        # Actualiza el texto del QLabel con la hora actual
        self.label_30.setText(hora_actual)


if __name__ == "__main__":
    app = QApplication([])
    mi_app = MiApp()
    mi_app.show()
    app.exec_()