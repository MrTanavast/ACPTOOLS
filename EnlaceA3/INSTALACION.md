# Enlace contable a3: instalación y uso

Esta herramienta genera el fichero **SUENLACE.DAT** que lee a3ECO / a3CON (*Utilidades > Importar/Exportar > Enlace contable*) a partir de dos tipos de origen:

- **Diarios**: asientos sin IVA. Usa la misma plantilla de siempre, con las columnas A a H.
- **Libros de facturas emitidas**:
  - el CSV de Contapluz de la empresa **01692**, que se trata exactamente igual que el .DAT que ya se importó bien en a3;
  - cualquier otra empresa, con la **plantilla GENERAL**.

Todo va en **un único libro de Excel con macros (.xlsm)**, la "herramienta". **Tus diarios y libros de facturas siguen en sus propios archivos** y no necesitan macros. La herramienta los lee sin modificarlos.

---

## 1. Qué hay en la carpeta `dist/`

| Fichero | Qué es | Qué se hace con él |
|---|---|---|
| `modA3Nucleo.bas` | Formatos de A3, textos, importes, fechas, lectura de CSV y escritura del .DAT | Importar |
| `modA3Config.bas` | Hojas EMPRESAS, IVA y PERFILES | Importar |
| `modA3Diario.bas` | Motor de diarios | Importar |
| `modA3Facturas.bas` | Motor de facturas emitidas | Importar |
| `modA3Control.bas` | Excel de control | Importar |
| `modA3Inicio.bas` | Macros de arranque, plantillas y orquestación | Importar |
| `clsA3Ventana.cls` | La ventana: construye los controles y hace el trabajo | Importar |
| `clsA3Evento.cls` | Clase auxiliar de la ventana (recoge los clics) | Importar |
| `logo_acp.png` | Logo de ACP (opcional) | Dejarlo junto al .xlsm |

## 2. Montaje (una sola vez, unos 5 minutos)

1. Abre un **libro nuevo en blanco** en Excel y guárdalo como **Libro de Excel habilitado para macros (.xlsm)**, por ejemplo `EnlaceA3_ACP.xlsm`. Copia `logo_acp.png` en la misma carpeta.
2. Pulsa **Alt+F11** para abrir el editor de VBA.
3. **Crea primero el formulario**, antes de importar nada:
   - *Insertar > UserForm*.
   - En la ventana de Propiedades (F4), cambia **(Name)** de `UserForm1` a **`frmEnlaceA3`**.
   - **Déjalo vacío**: ni dibujes controles ni pegues código. Los botones y las casillas se crean solos al abrirse; todo el código va en `clsA3Ventana.cls`.
   - Se crea primero para que Excel active la referencia a los formularios, que necesitan las dos clases.
4. **Importa los 8 ficheros**: *Archivo > Importar archivo…*, uno a uno:
   - los 6 `.bas`;
   - `clsA3Ventana.cls` y `clsA3Evento.cls`.
5. *Depuración > Compilar VBAProject*. No debe salir ningún mensaje.
6. Cierra el editor. En Excel, pulsa **Alt+F8**, elige **PrepararLibro** y pulsa Ejecutar. Se crean estas hojas:
   - **INICIO**, con los botones de la herramienta;
   - **EMPRESAS**, que ya incluye la 01692 Patio Posadero, con sus cuentas;
   - **IVA**, con los tipos 4 %, 10 % y 21 % de la 01692;
   - **PERFILES**, con GENERAL y CONTAPLUZ;
   - **PLANTILLA_DIARIO** y **PLANTILLA_FACTURAS**.
7. **Guarda**. Este es el archivo "maestro": haz una copia para cada compañera.

> Si Excel avisa de que las macros están deshabilitadas: *Archivo > Opciones > Centro de confianza > Configuración > Configuración de macros*, o marca la carpeta como ubicación de confianza.

> **Para actualizar a una versión nueva:** en el editor de VBA, quita los módulos `modA3…` y las clases `clsA3…` (clic derecho > *Quitar…* > *No* exportar) e importa los de la carpeta `dist/` nueva. El formulario y las hojas (EMPRESAS, IVA, PERFILES) no se tocan: tu configuración se conserva.

> **Nunca copies el código desde el navegador o un PDF**: parte las líneas largas y mete pies de página, y VBA da "Se esperaba fin de la instrucción". Descarga los ficheros e impórtalos.

## 3. Uso diario

1. Abre la herramienta (`EnlaceA3_ACP.xlsm`). Si los datos están en Excel, **abre también ese libro**, o ábrelo desde la ventana con el botón *Abrir…*.
2. Pulsa **ABRIR ENLACE A3** en la hoja INICIO, o Alt+F8 > **EnlaceA3**.
3. En la ventana:
   - **Código de empresa** de A3. Si la empresa está en EMPRESAS, se rellenan solos el nombre, los dígitos, el perfil y la última carpeta.
   - **Tipo de enlace**: *Diario* o *Facturas emitidas*.
   - **Dígitos del plan** de cuentas. Por ejemplo, 570000000 son 9 dígitos.
   - **Origen**: una hoja de cualquier libro abierto, o un archivo CSV. Por defecto se propone la hoja activa del último libro de datos que abriste, nunca las plantillas de la herramienta.
   - **Filtros**:
     - periodo rápido (mes, trimestre o año) o fechas libres;
     - rango de documentos (INV1510 a INV1704) o de asientos;
     - series que se incluyen o se excluyen (`F;R` o `INV;REC`);
     - tipos de documento (facturas, tickets, abonos);
     - en diarios, excluir los asientos descuadrados.
   - **Salida**: la carpeta donde se guarda `SUENLACE.DAT`. Se recuerda por empresa. Puedes usar tu truco de `V:\A3DIARIOS\01692`.
4. **Analizar**: muestra el resultado y las incidencias **sin escribir nada**.
5. **GENERAR SUENLACE.DAT**: escribe el fichero y el **Excel de control** (`CONTROL_SUENLACE_<empresa>_<fecha>.xlsx`), que tiene estas hojas:
   - resumen y cuadre con el origen;
   - detalle de lo que va al fichero;
   - incidencias;
   - filas de origen con su estado;
   - el propio fichero, línea a línea.
6. En A3: **copia de seguridad** de la empresa, *Enlace contable*, carpeta y `SUENLACE.DAT`, **CH (chequear)** y después enlazar.

Si ya había un `SUENLACE.DAT` en la carpeta, se renombra a `SUENLACE_anterior_<fecha_hora>.DAT`. Nunca se pierde.

## 4. Reglas que aplica

**Diarios**

- Las líneas se agrupan por **nº de asiento + fecha**, aunque no estén seguidas.
- Primera línea **I**, intermedias **M** y última **U**. Se calcula después de quitar las filas omitidas.
- Una fila con Debe y Haber a 0 se omite, con aviso.
- Una fila con Debe **y** Haber a la vez genera dos apuntes, con aviso.
- Un asiento con cualquier error se **excluye entero** y queda listado para corregirlo:
  - fecha imposible;
  - importe no numérico;
  - cuenta con un número de dígitos distinto al del plan;
  - descuadre (esta comprobación se puede desactivar).
- Las cuentas admiten la notación de punto de A3: `572.1` pasa a `572000001`.

**Facturas emitidas** (el mismo procedimiento que usaba tu compañera, ampliado)

- Las líneas se agrupan por **nº de factura + fecha**. Así, una numeración que se reinicia (REC2) no se mezcla. Una factura anulada y otra emitida con el mismo número el mismo día tampoco se mezclan: se exporta la buena, con aviso.
- Se excluyen, y quedan listados:
  - los documentos **anulados**;
  - los de **importe 0**;
  - los que tienen importes no numéricos o celdas con error (#N/D);
  - los de **tipos de IVA sin cuenta** configurada (por ejemplo, 100 %);
  - las cuentas con un número de dígitos incorrecto;
  - los **descuadres** entre el total y base + cuotas, porque A3 los rechazaría.
- Las líneas con base y cuota a 0 se quitan de la factura.
- **Abonos** (Abono / Credit Note): registro tipo 2 con importes en positivo. Siempre restan ventas. Si un abono llega en positivo en el origen, se avisa.
- **Rectificativas** (solo en el perfil GENERAL): registro tipo 2 con el signo del origen. Si en el origen son negativas, disminuyen la factura original; si son positivas, la aumentan (A3 las recibe en negativo) y se avisa.
- Los tickets se tratan como facturas.
- El detalle lleva una línea por **cuenta de ventas + tipo de IVA**.
- Impreso: 01 (347). Las entregas intracomunitarias (subtipos 03 y 04) llevan 02 (349). También se puede indicar en una columna "Impreso".
- **Nº de factura**: A3 guarda 10 caracteres. Si el número es más largo, se quitan los separadores y se conserva el final (`F2026/000123` pasa a `2026000123`). Si dos facturas quedaran con el mismo número, se excluye la segunda.
- **NIF**:
  - se normaliza (sin guiones ni espacios, sin el prefijo ES);
  - si no es un NIF español válido, se avisa;
  - el nombre y el CP solo se informan si hay NIF. Si el CP no es de 5 cifras, se deja en blanco y se avisa.
- **Retención**: si viene con el signo cambiado, se corrige y se avisa. Si viene sin porcentaje, se calcula (1, 2, 7, 15, 19 o 24 %).
- **Porcentajes**: se aceptan celdas con formato % (21 % o 0,21).
- Todos los textos van en ASCII: sin tildes ni eñes.

## 5. Otras empresas (perfil GENERAL)

1. Con la empresa escrita y *Facturas emitidas* marcado, pulsa **Configurar empresa…**. Se da de alta con cuentas por defecto (430…, 700…, 477… por tipo), en naranja para revisarlas. Revisa **EMPRESAS** e **IVA**.
2. Adapta el libro de facturas del cliente a la **plantilla de facturas**, una fila por factura y tipo de IVA:
   - Botón *Nueva plantilla de FACTURAS* en INICIO.
   - Obligatorias: Fecha, Nº Factura, Base imponible, % IVA y Cuota IVA.
   - Opcionales: tipo (Factura/Ticket/Abono/Rectificativa), cliente, NIF, CP, total, cuenta de cliente o de ventas por línea, recargo, retención, subtipo, impreso y anulada.
3. Si otro cliente envía siempre el mismo formato de exportación, no hace falta adaptarlo a mano. Añade una **columna nueva en PERFILES**:
   - escribe, para cada campo, el nombre de la columna de su fichero;
   - puedes poner alternativas separadas con `|`;
   - pon `?` delante si la columna es opcional.

   Después, elige ese perfil en la ventana.

## 6. Pendiente o preparado

- **Facturas recibidas**: la estructura está preparada (tipo de factura 2 en la cabecera, subtipos de compras), pero la opción aparece desactivada hasta que haya libros de compras que importar.
- **Formato**: se genera el formato 4 de 254 posiciones, el que ya os ha funcionado. El PDF 9.80 describe el formato 5 (512 posiciones, datos SII). Se puede añadir si alguna empresa lo necesita.
