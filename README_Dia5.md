# Día 5. Macros y VBA
Dra. Elda C. Morales.

## Archivos y uso
1. Abre dia5.html en un navegador con JavaScript o visita el tutorial publicado.
2. Descarga Excel_Dia_5_Macros_Practicas.xlsx. Contiene datos y soluciones por fórmula, sin VBA incrustado ni botones activos.
3. En Excel de escritorio guarda una copia como Libro de Excel habilitado para macros (*.xlsm). No basta con cambiar la extensión.
4. En Programador > Visual Basic selecciona el proyecto de esa copia. Archivo > Importar archivo > Modulo_Dia5.bas.
5. Alternativa: Insertar > Módulo y pegar Codigo_Dia5_Para_Pegar.txt. Usa solo una vía. El TXT no incluye Attribute VB_Name.
6. Revisa el código, guarda y usa Depuración > Compilar VBAProject. Las políticas institucionales pueden limitar la ejecución de VBA.
7. Ejecuta PrepararTrabajo y escribe 0 para el caso base o de 1 a 5 para un reto. La macro conserva la fuente y crea la hoja Trabajo correspondiente.
8. Asigna botones de formulario a PrepararTrabajo y ResumenPorRegion en una zona libre de Guia. ActiveX no es compatible con Mac.
9. Compara los resultados con las hojas Resuelto y Resumen. Registra pruebas reales en Plan_Pruebas.
10. En UDF_Pruebas escribe =MargenSeguro(A6;B6) en D6 y copia hasta D13 después de importar el módulo en XLSM.

## Macros disponibles
- PrepararTrabajo: valida la fuente, copia datos y escribe fórmulas y formato.
- FormatoGrabadoMejorado: referencia explícita a Base_Trabajo y última fila calculada.
- ResumenPorRegion: pide caso y región y genera una consulta.
- EvaluarUmbral: pide un mínimo numérico y cuenta pedidos con Sales mayor o igual.
- DemostrarMsgBox: muestra respuestas Sí, No y Cancelar sin modificar datos.
- MargenSeguro: función personalizada sin efectos sobre celdas ni diálogos.

El módulo completo incluye funciones auxiliares. Importa o pega todo el contenido una sola vez. VBA usa palabras clave en inglés y comillas dobles duplicadas dentro de cadenas. Las fórmulas visibles del tutorial usan español y punto y coma, que debe adaptarse a la configuración de Excel.

## Alcance y comprobación
Los pedidos son sintéticos, enero-septiembre de 2026, con importes MXN. Cada fila representa un pedido. Los resueltos son una copia del caso con fórmulas equivalentes. Las hojas Resumen están vinculadas a las fuentes Datos. CSV UTF-8 con coma: Pedido como Texto y Fecha como fecha ISO.

Se revisaron estáticamente el código y los archivos y se probaron cálculos equivalentes y simuladores. El módulo no se compiló ni ejecutó en Excel durante su preparación. La compilación, los botones y la ejecución se verifican en la instalación de escritorio del estudiante. Un fallo durante la escritura puede dejar un resultado parcial; la macro no es transaccional ni deshace automáticamente los cambios.

El código no contiene eventos de apertura, acceso a red ni archivos externos. Sus procedimientos se invocan manualmente y trabajan dentro de ThisWorkbook. Antes de reemplazar un resultado existente pide confirmación, con No como predeterminado. Usa una copia y conserva la fuente.

El HTML es una simulación JavaScript: no ejecuta VBA, no controla Excel ni comprueba automáticamente el archivo del estudiante. El avance se guarda en el navegador y puede descargarse como JSON.

## Publicación
https://eldamorales.github.io/excel-decisiones/dia5.html
https://github.com/eldamorales/excel-decisiones
