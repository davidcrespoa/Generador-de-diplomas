function generarConstancias() {
  // 1. Inicialización
  const hoja = abrirHojaPorID(ID_HOJA);
  const datos = leerDatos(hoja);
  
  Logger.log("Iniciando proceso para " + datos.length + " registros.");

  // 2. Bucle
  for (const filaActual of datos) {
    
    // Obtenemos los datos de la fila
    const nombreEmpleado = obtenerNombre(filaActual);
    const cursoEmpleado = obtenerCurso(filaActual);
    const fechaCurso = obtenerFecha(filaActual);

    // Acciones con el documento
    const copiaDocumento = crearCopiaPlantilla(ID_PLANTILLA);
    const documento = abrirDocumento(copiaDocumento);

    // Reemplazar los marcadores
    reemplazarTexto(documento, "{{nombre}}", nombreEmpleado);
    reemplazarTexto(documento, "{{curso}}", cursoEmpleado);
    reemplazarTexto(documento, "{{fecha}}", fechaCurso);

    // Finalizar y renombrar
    guardarDocumento(documento);
    cambiarNombreDocumento(copiaDocumento, nombreEmpleado);

    // Organizar en carpeta
    moverArchivoACarpeta(copiaDocumento, ID_CARPETA_DESTINO);

    Logger.log("Certificado finalizado para: " + nombreEmpleado);
  };
  
  Logger.log("Proceso completado exitosamente.");
}

