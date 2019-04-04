# MatrizID
## Matriz de probabilidad - impacto ambiental 

En este proyecto se ha intentado abordar el análisis de impactos medioambientales a través de macros en Excel que generan una matriz de Leopold.
Se que suena horrible pero es una versión de prueba.

El ciclo de trabajo con este programa se basa en:
1.  Identificar los factores a intervenir (24 por ahora), después seleccionar las acciones que se llevarán a cabo (26 por ahora).*
1.  Identificar las acciones que se tomarán a cabo para cada uno de los factores seleccionados.*
1.  Medir de forma cualitativa las magnitudes _signo, intensidad, extensión, momento, persistencia, reversibilidad, sinergia, acumulación, efecto, preriodicidad y recuperabilidad_** los valores asociados a los diferentes niveles se encuentra en la pestaña Atributos y se pueden modificar(_ver apartado precacaución_).
1.  Se puede rellenar la pestaña Importancia una vez completados los pasos anteriores para calcular los sumatorios totales y el número y tipo de los impactos a mano. _lo implementaré si se vuelve solicitado_.
1.  Medir de forma cuantitativa el UIP y sus valores absoluto y relativo, así como el total en cada una de las fases y medios. 

***Los pasos 1 y 2 se realizan únicamente seleccionando botones**

****El paso 3 se realiza rellenando los valores con un desplegable**

Estoy desarrollando una aplicación robusta que lleve a cabo el análisis y cálculo de los factores ambientales que intervienen en la evaluación y estudio del terreno y los riesgos asociados.

##  MODO DE EMPLEO
1.  Descargar el fichero xlsm y ejecutar con microsoft excel.
1.  Habilitar las macros para excel.
1.  CTRL+R devuelve al estado inicial los botones, la pestaña MatrizIdentificacion, AccionesxFactores e Importancia.
1.  No es necesario escribir ningún dato, todo se rellena automáticamente.
1.  La pestaña importancia no hace las sumas absoluta y relativa aún, el valor ha de ser calculado a mano o creando la fórmula que sume el intervalo correspondiente.

## PRECAUCIÓN
* **Al cambiar los valores de los Atributos se recomienda no eliminar ni insertar filas ya que puede provocar un mal funcionamiento.**
* **Modificar el excel puede comprometer su integridad.**
* **Todas las pestañas salvo la de importancia, si la acción es a mano, han de modificarse entendiendo parte del código primero.**
* **Este excel bien usado puede ser una herramienta muy útil pero con un mal uso no servirá para nada. Como consejo, al descargar el fichero no lo abras y crea una copia del excel para cada proyecto que vayas a llevar a cabo**
