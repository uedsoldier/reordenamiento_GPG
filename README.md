Manual de instalación de dependencias
1.	Instalación de Python en su versión más reciente. En este momento está disponible la versión 3.9 en el enlace https://www.python.org/downloads/
2.	Una vez instalado correctamente python, se realiza la instalación de los paquetes openpyxl y xlrd. Basta con introducir en la línea de comandos de windows lo siguiente:

pip install openpyxl xlrd

Presionar Enter y esperar a la instalación de ambos paquetes

3.	Abrir directorio en el que está alojado el programa mediante línea de comandos de Windows (cmd.exe). Por ejemplo:
C:/Users/PC/Escritorio/

4. Una vez que el directorio se encuentre activo en la línea de comandos, y asegurándonos de que el archivo de origen con el formato correcto se encuentre presente, se introduce el siguiente comando:
python GPG_reordenamiento.py GPG_origen.xlsx

En este caso, GPG_origen.xlsx es el nombre del archivo de origen, el cual puede tener cualquier nombre
Cabe señalar que el archivo de origen deberá cumplir con el formato indicado y esto es responsabilidad del personal de GPG que entregue dicho archivo. El no cumplir con esto, puede dar lugar a malfuncionamiento 
de la herramienta de adaptación.

Como salida del programa, se apreciarán mensajes como los siguientes a manera de ejemplo:

Directorio de archivos de carga masiva creado
Archivo origen: GPG_origen.xlsx
Archivo destino .xlsx: carga masiva/carga_masiva_20201207_144846.xlsx
Archivo destino .csv: carga masiva/carga_masiva_20201207_144846.csv
Abriendo hojas de cálculo
Hojas de cálculo abiertas
Cantidad de registros: 46
Obteniendo productos
Cantidad total de productos: 46
Cantidad familias SKU: 5

5. Una vez finalizada la ejecución del programa se creará (si no existe) una carpeta llamada carga masiva, en la cual se generará un archivo de carga masiva de datos compatible con la plataforma woocommerce
Cada vez que se ejecute el programa, se generará un archivo nuevo de carga masiva, diferenciado por un timestamp de fecha y hora en la cual se generó, por ejemplo:

carga_masiva_20201207_144846.csv --> Archivo de carga masiva generado el 7 de diciembre de 2020 a las 14:48:46

Este archivo, en caso de que el de origen tenga el formato correcto, ya podrá ser cargado directamente en la plataforma de comercio electrónico para gregar nuevos productos o actualizar los existentes.
