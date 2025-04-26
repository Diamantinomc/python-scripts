# Scripts de Python
## 1. Consumos
### Función del script
Extrae un archivo Excel desde el sistema Dynamics, utilizado en la empresa, filtra los datos, los oprdena y entrega un nuevo archivo Excel.

### Problemática
Cada vez que alguien retira artículos de la bodega se debe realizar el descuento en el stock disponible. Para ello el solicitante debe llenar una planilla con los datos del artículo; 
código, nombre, cantidad, firma del solicitante, del receptor y de quien entrega.
Con este documento se realiza el descuento por sistema.
Por requerimientos de jefatura estos documentos deben registrarse en una planilla en Excel. En temnparada alta los documentos son muchos y registrarlo es tedioso y quita mucho tiempo.
Mi trabajo no es en TI, por lo que realizar alguna mejora a nivel de sistemas escapa de mis labores y el computador no tenía los implementos necesarios para desarrollar.

### Solución
Presenté la problemática a un integrante de Soporte que me permitió instalar Python y un VsCode para poder hacer pruebas.
Al no tener acceso a la base de datos de la empresa, la solución a la que se llegó fue a manipular el cursor del mouse para poder llegar a la página de donde poder descargar el archivo.
Esta slución no es la más eficiente, pero en un principio soluciona la problemática.

## 2. Consumos_epp
### Función del script
Crea un dataset a partir de un archivo Excel y entrega un reporte de los artículos más utilizados

### Problemática
La compra de elementos de protección personal es mucho y al comprarse un elemento por tallas, por ejemplo, esta se compra por intuición, sin considerar realmente que talla se entregó más y a qué área.

### Solución
El modelo nos debería mostrar qué producto se entrag más, qué talla, qué configuración, a qué área, etc. Todo esto dentro de un rango de fecha.
