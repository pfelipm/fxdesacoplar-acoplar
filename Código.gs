/**
 * Esta función desacopla las filas de un intervalo de datos que contiene valores múltiples, delimitados por la 
 * secuencia de caracteres indicada, en algunas de sus columnas. Se ha diseñado principalmente para facilitar el
 * tratamiento estadístico de las respuestas a un formulario cuando alguna de sus preguntas admite múltiples opciones 
 * (casillas de verificación), que en ese caso están separadas por la secuencia delimitadora ", " (coma espacio).
 * Tras ser desacopladas, las respuestas (filas) con opciones múltiples se repiten en el intervalo resultante para cada
 * combinación posible de los valores múltiples únicos de las columnas especificadas.
 * @param {A1:D10} intervalo Intervalo de datos.
 * @param {VERDADERO} encabezado Indica si el rango tiene una fila de encabezado con etiquetas para cada columna ([VERDADERO] | FALSO).
 * @param {", "} separador Secuencia de caracteres que separa los valores múltiples. Opcional, si se omite se utiliza ", " (coma espacio).
 * @param {2} columna Número de orden, desde la izquierda, de la columna dentro del intervalo que contiene valores múltiples a descoplar.
 * @param {4} [más_columnas] Columnas adicionales, opcionales, que contienen valores múltiples a desacoplar, separadas por ";".
 *
 * @return Intervalo de datos desacoplados
 *
 * @customfunction
 *
 * Artículo:    https://pablofelip.online/desacoplar-acoplar/
 * Repositorio: https://github.com/pfelipm/fxdesacoplar-acoplar
 * 
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer (@pfelipm)
 */ 
function DESACOPLAR(intervalo, encabezado, separador, columna, ...masColumnas) {

  // Control de parámetros inicial
  
  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo.';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas.';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto.';
  let columnas = typeof columna != 'undefined' ? [columna, ...masColumnas].sort() : [...masColumnas].sort();
  if (columnas.length == 0) throw 'No se han indicado columnas a descoplar.';
  if (columnas.some(col => typeof col != 'number' || col < 1)) throw 'Las columnas deben indicarse mediante números enteros';
  if (Math.max(...columnas) > intervalo[0].length) throw 'Al menos una columna está fuera del intervalo.';

  // Se construye un conjunto (set) para evitar automáticamente duplicados en columnas con valores múltiples
    
  const colSet = new Set();
  columnas.forEach(col => colSet.add(col - 1));
    
  // Listos para comenzar
  
  if (encabezado) encabezado = intervalo.shift();
  
  const intervaloDesacoplado = [];
  
  // Recorramos el intervalo fila a fila
  
  intervalo.forEach(fila => {
                    
    // Enumerar los valores únicos en cada columna que se ha indicado contiene datos múltiples
         
    const opciones = []; 
    for (const col of colSet) {
    
       // Eliminar opciones duplicadas, si las hay, en cada columna gracias al uso de un nuevo conjunto
    
       const opcionesSet = new Set();
       String(fila[col]).split(separador).forEach(opcion => opcionesSet.add(opcion)); // split solo funciona con string, convertimos números
       opciones.push([...opcionesSet]); // también opciones.push(Array.from(opcionesSet))
    
    }
      
    // Ahora desacoplamos la respuesta (fila) mediante una IIFE recursiva 🔄
    // que genera un vector para todas las posibles combinaciones (vectores) de respuestas
    // de las columnas con valores múltiples
    // Ej:
    //     ENTRADA: vector = [ [a, b], [1, 2] ]
    //     SALIDA:  combinaciones = [ [a, 1], [a, 2], [b, 1], [b, 2] ]
  
    const combinaciones = (function combinar(vector) {
    
      if (vector.length == 1) {
        
        // Fin del proceso recursivo
        
        const resultado = [];
        vector[0].forEach(opcion => resultado.push([opcion]));
        return resultado;
      }
      
      else {
        
        // El resultado se calcula recursivamente
        
        const resultado = [];
        const subvector = vector.splice(0, 1)[0];
        const subresultado = combinar(vector);
      
        // Composición de resultados en la secuencia recursiva >> generación de vector de combinaciones
        
        subvector.forEach(e1 => subresultado.forEach(e2 => resultado.push([e1, ...e2])));
      
        return resultado;   
        
      }
    })(opciones);

    // Ahora hay que generar las filas repetidas para cada combinación de datos múltiples
    // Ej:
    //     ENTRADA: combinaciones = [ [a, 1], [a, 2], [b, 1], [b, 2] ]
    //     SALIDA:  respuestaDesacoplada = [ [Pablo, a, 1, Tarde], [Pablo, a, 2, Tarde], [Pablo, b, 1, Tarde], [Pablo, b, 2, Tarde] ]
  
    const respuestaDesacoplada = combinaciones.map(combinacion => {
                        
      let colOpciones = 0;
      const filaDesacoplada = [];
      fila.forEach((valor, columna) => {
      
        // Tomar columna de la fila original o combinación de datos generada anteriormente
        // correspondiente a cada una de las columnas con valores múltiples
      
        if (!colSet.has(columna)) filaDesacoplada.push(valor);
        else filaDesacoplada.push(combinacion[colOpciones++]);
      
      });
      return filaDesacoplada;
    });
  
    // Se desestructura (...) respuestaDesacoplada dado que combinaciones.map es [[]]

    intervaloDesacoplado.push(...respuestaDesacoplada);
  
  });
  
  // Si hay fila de encabezados, colocar en 1ª posición en la matriz de resultados

  return encabezado.map ? [encabezado, ...intervaloDesacoplado] : intervaloDesacoplado;
  
}

/**
 * Esta función acopla (combina) las filas de un intervalo de datos que corresponden a una misma entidad. Para ello, 
 * se debe indicar la columna (o columnas) *clave* que identifican los datos de cada entidad única. Los valores registrados
 * en el resto de columnas se agruparán, para cada una de ellas, utilizando como delimitador la secuencia de caracteres 
 * indicada. Se trata de una función que realiza una operación complementaria a DESACOPLAR(), aunque no perfectamente simétrica.
 * @param {A1:D10} intervalo Intervalo de datos.
 * @param {VERDADERO} encabezado Indica si el rango tiene una fila de encabezado con etiquetas para cada columna ([VERDADERO] | FALSO).
 * @param {", "} separador Secuencia de caracteres a emplear como separador de los valores múltiples. Opcional, si se omite se utiliza ", " (coma espacio).
 * @param {1} columna Número de orden, desde la izquierda, de la columna clave que identifica los datos de la fila como únicos.
 * @param {2} [más_columnas] Columnas clave adicionales, opcionales, que actúan como identificadores únicos, separadas por ";".
 *
 * @return Intervalo de datos desacoplados
 *
 * @customfunction
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer (@pfelipm)
 */ 
function ACOPLAR(intervalo, encabezado, separador, columna, ...masColumnas) {

  // Control de parámetros inicial
  
  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo.';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas.';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto.';
  let columnas = typeof columna != 'undefined' ? [columna, ...masColumnas].sort() : [...masColumnas].sort();
  if (columnas.length == 0) throw 'No se han indicado columnas clave.';
  if (columnas.some(col => typeof col != 'number' || col < 1)) throw 'Las columnas clave deben indicarse mediante números enteros';
  if (Math.max(...columnas) > intervalo[0].length) throw 'Al menos una columna clave está fuera del intervalo.';

  // Se construye un conjunto (set) para evitar automáticamente duplicados en columnas CLAVE
    
  const colSet = new Set();
  columnas.forEach(col => colSet.add(col - 1));
  
  // ...y en este conjunto se identifican las columnas susceptibles de contener valores que deben concatenarse
  
  const colNoClaveSet = new Set();
  for (let col = 0; col < intervalo[0].length; col++) {
  
    if (!colSet.has(col)) colNoClaveSet.add(col);
  
  }
  
  // Listos para comenzar
  
  if (encabezado) encabezado = intervalo.shift();
  
  const intervaloAcoplado = [];

  // 1ª pasada: recorremos el intervalo fila a fila para identificar entidades (concatenación de columnas clave) únicas
  
  const entidadesClave = new Set();
  intervalo.forEach(fila => {
    
    const clave = [];
    // ⚠️ A la hora de diferenciar dos entidades únicas (filas) usando una serie de columnas clave:
    //    a) No basta con concatenar los valores de las columnas clave como cadenas y simplemente compararlas. Ejemplo:
    //       clave fila 1 → col1 = 'pablo' col2 = 'felip'     >> Clave compuesta: 'pablofelip'
    //       clave fila 2 → col1 = 'pa'    col2 = 'blofelip'  >> Clave compuesta: 'pablofelip'
    //       ✖️ Misma clave compuesta, pero entidades diferentes
    //    b) No basta con con unir los valores de las columnas clave como cadenas utilizando un carácter delimitador. Ejemplo ('/'):
    //       clave fila 1 → col1 = 'pablo/' col2 = 'felip'    >> Clave compuesta: 'pablo//felip' 
    //       clave fila 2 → col1 = 'pablo'  col2 = '/felip'   >> Clave compuesta: 'pablo//felip'
    //       ✖️ Misma clave compuesta, pero entidades diferentes
    //    c) No es totalmente apropiado eliminar espacios antes y después de valores clave y unirlos usando un espacio delimitador (' '):
    //       clave fila 1 → col1 = ' pablo' col2 = 'felip'    >> Clave compuesta: 'pablo felip'
    //       clave fila 2 → col1 = 'pablo'  col2 = 'felip'    >> Clave compuesta: 'pablo felip'
    //       ✖️ Misma clave compuesta, pero entidades estrictamente diferentes (a menos que espacios anteriores y posteriores no importen)
    // 💡 En su lugar, se generan vectores con valores de columnas clave y se comparan sus versiones transformadas en cadenas JSON.
    for (const col of colSet) clave.push(String(fila[col])) 
    entidadesClave.add(JSON.stringify(clave));
                    
  });

  // 2ª pasada: obtener filas para cada clave única, combinar columnas no-clave y generar filas resultado

  for (const clave of entidadesClave) {

    const filasEntidad = intervalo.filter(fila => {
    
      const claveActual = [];
      for (const col of colSet) claveActual.push(String(fila[col]));
      return clave == JSON.stringify(claveActual);
     
    });

    // Acoplar todas las filas de cada entidad, concatenando valores en columnas no-clave con separador indicado

    const filaAcoplada = filasEntidad[0];  // Se toma la 1ª fila del grupo como base
    const noClaveSets = [];
    for (let col = 0; col < colNoClaveSet.size; col++) {noClaveSets.push(new Set())}; // Vector de sets para recoger valores múltiples   
    filasEntidad.forEach(fila => {
      
      let conjunto = 0;   
      for (const col of colNoClaveSet) {noClaveSets[conjunto++].add(String(fila[col]));}
                            
    });

    // Set >> Vector >> Cadena única con separador

    let conjunto = 0;
    for (const col of colNoClaveSet) {filaAcoplada[col] = [...noClaveSets[conjunto++]].join(separador);}

    intervaloAcoplado.push(filaAcoplada);
  
  }

  return encabezado.map ? [encabezado, ...intervaloAcoplado] : intervaloAcoplado;
  
}