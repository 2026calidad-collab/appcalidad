// =====================================================================
//  CONFIGURACIÓN DE LA APP DE SANEAMIENTO — Packing del Carmen SAC
//  Normalmente solo hay que pegar APPS_SCRIPT_URL.
//  Las LISTAS de abajo son los valores iniciales: después se administran desde la app
//  (Inicio → Administrar listas) y quedan guardadas en la hoja "Listas" del Google Sheets.
// =====================================================================
window.CONFIG = {
  // Pega aquí la URL de Google Apps Script (termina en /exec). Vacía = MODO PRUEBA.
  APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbyMnsye4Ss5AY5fn60yiteg8GCXj4cnnyblXvZnOW0rwfbHCUDLald1yWeflI9GIEcYpQ/exec",

  empresa: "Packing del Carmen SAC",

  // Usuarios de ejemplo SOLO para el modo prueba. Los usuarios reales se crean en la app:
  // ingresa con el usuario "admin" y la contraseña maestra (ADMIN_PIN del Apps Script) → ⚙ Administrar → Usuarios y permisos.
  usuariosPrueba: [
    { id: "U1", usuario: "indira", nombre: "Indira Cruz", cargo: "Asistente de Calidad", permisos: ["ADMIN", "VOBO", "EDITAR", "FT-07", "FT-14", "FT-17", "FT-18", "FT-19", "FT-20"], activo: true },
    { id: "U2", usuario: "operario", nombre: "Operario Saneamiento", cargo: "Saneamiento", permisos: ["FT-07", "FT-14", "FT-17", "FT-18", "FT-19", "FT-20"], activo: true }
  ],

  // Datos fijos de cada formato (código, fecha y versión del documento controlado)
  formatos: {
    "FT-17": { codigo: "PDC-SA-FT-17", fecha: "30.10.26", version: "6",  titulo: "Control de cebaderos" },
    "FT-20": { codigo: "PDC-SA-FT-20", fecha: "30.10.26", version: "10", titulo: "Control de visitas" },
    "FT-14": { codigo: "PDC-SA-FT-14", fecha: "30.10.26", version: "11", titulo: "Control de lentes" },
    "FT-18": { codigo: "PDC-SA-FT-18", fecha: "30.10.26", version: "04", titulo: "Trampas mecánicas" },
    "FT-19": { codigo: "PDC-SA-FT-19", fecha: "30.10.26", version: "11", titulo: "Insectocutores" },
    "FT-07": { codigo: "PDC-SA-FT-07", fecha: "30.10.26", version: "08", titulo: "Hermeticidad, infraestructura y aves" }
  },

  // ------------------------------------------------------------------
  //  CATÁLOGO DE LISTAS
  //  tipo: "lista"  → solo texto            (ej. motivos de visita)
  //        "codigo" → código + descripción  (ej. S/C = Sin consumo); puede marcar "alerta"
  //        "rango"  → rango de números + nombre (ej. 1-30 = Perímetro externo)
  //        "texto"  → un solo texto         (ej. pregunta de salud)
  //        "numero" → un solo número        (ej. total de cebaderos)
  //        "par"    → código + nombre, el código puede repetirse (ej. PR = Rosa Huamán)
  //        "area"   → localización + área + aspectos + frecuencia + monitoreo de aves (FT-07)
  //  alerta: true → ese valor cuenta como hallazgo / no conformidad y dispara correo
  // ------------------------------------------------------------------
  catalogo: [
    { id: "ft17.total", formato: "FT-17", nombre: "Cantidad de cebaderos", tipo: "numero", valor: 186 },
    { id: "ft17.ingrediente", formato: "FT-17", nombre: "Ingrediente activo del cebo", tipo: "texto", valor: "Brodifacoum" },
    { id: "ft17.zonas", formato: "FT-17", nombre: "Zonas de cebaderos (rango de números)", tipo: "rango",
      items: [["1-30", "Perímetro externo"], ["31-60", "Almacén general"], ["61-90", "Recepción"], ["91-120", "Mercado local"], ["121-150", "Comedor y SS.HH."], ["151-186", "Zona de motores y talleres"]] },
    { id: "ft17.consumo", formato: "FT-17", nombre: "Opciones de consumo", tipo: "codigo",
      items: [["S/C", "Sin consumo", false], ["P", "Parcial", true], ["T", "Total", true]] },
    { id: "ft17.incidencia", formato: "FT-17", nombre: "Opciones de incidencia", tipo: "codigo",
      items: [["Ninguna", "Ninguna", false], ["PR", "Presencia de roedor", true], ["HR", "Huella de roedor", true], ["E", "Excreta", true]] },
    { id: "ft17.recebado", formato: "FT-17", nombre: "Opciones de recebado", tipo: "codigo",
      items: [["No", "No", false], ["P", "Pellet", false], ["B", "Block", false], ["PA", "Pasta", false]] },

    { id: "ft20.motivos", formato: "FT-20", nombre: "Motivos de visita", tipo: "lista",
      items: ["Auditoría de cliente", "Auditoría de certificación", "Inspección SENASA", "Proveedor / servicio", "Mantenimiento externo", "Reunión / visita comercial", "Otro"] },
    { id: "ft20.normasBPM", formato: "FT-20", nombre: "Normas de Calidad – BPM (se muestran a la visita)", tipo: "lista",
      items: [
        "Uso completo del uniforme (mandil, toca).",
        "Lavado de manos.",
        "Evitar sonarse la nariz, toser o estornudar sobre o cerca del alimento.",
        "No ingresar alimentos o bebidas a sala de proceso.",
        "No usar joyería (aretes, collares, pulseras, relojes, anillos con piedra u otro objeto personal), excepto el aro matrimonial.",
        "No comer, fumar, masticar chicle, escupir, no cogerse el cabello, ni colocarse los dedos en las orejas, ojos, nariz, boca o cualquier otra parte del cuerpo."] },
    { id: "ft20.normasSeguridad", formato: "FT-20", nombre: "Normas de Seguridad (se muestran a la visita)", tipo: "lista",
      items: [
        "Transitar por las zonas autorizadas.",
        "Transitar con el personal que recepciona la visita.",
        "Respetar las señales de orden y seguridad.",
        "No utilizar instrumentos electrónicos, excepto si es autorizado por Gerencia de planta.",
        "No correr en áreas de proceso.",
        "En caso se le entreguen EPP, usarlos de manera adecuada.",
        "En todo momento debe permanecer con el carnet de identificación que le proporcione la empresa."] },
    { id: "ft20.salud", formato: "FT-20", nombre: "Pregunta de declaración de salud", tipo: "texto",
      valor: "Durante las últimas 72 horas, ¿ha tenido algún malestar como fiebre, diarrea, hepatitis, bronquitis, tos, heridas, vómitos o dolor estomacal?" },
    { id: "ft20.lentesCriterio", formato: "FT-20", nombre: "Criterio de lentes conformes", tipo: "texto",
      valor: "Marco, lunas y tornillos completos, sin rajaduras." },

    // ---- FT-14 Control de lentes ----
    { id: "ft14.areas", formato: "FT-14", nombre: "Áreas del personal", tipo: "codigo",
      items: [["AL", "Almacén", false], ["FR", "Frío", false], ["RE", "Recepción", false], ["LO", "Logística", false], ["AB", "Abastecimiento", false], ["MA", "Mantenimiento", false], ["PR", "Proceso", false], ["CA", "Calidad", false]] },
    { id: "ft14.personal", formato: "FT-14", nombre: "Personal que usa lentes (área + nombre)", tipo: "par", items: [] },

    // ---- FT-18 Trampas mecánicas ----
    { id: "ft18.areas", formato: "FT-18", nombre: "Áreas de trampas mecánicas", tipo: "codigo",
      items: [["AG", "Almacén General", false], ["AE", "Almacén Espárrago", false], ["ML", "Mercado Local", false], ["RC", "Recepción Cítrico", false], ["RML", "Rampa de Mercado Local", false], ["RAG", "Rampa de Almacén General", false], ["RP", "Recepción Palta", false], ["EP", "Entrada Principal", false], ["CO", "Comedor", false]] },
    { id: "ft18.trampas", formato: "FT-18", nombre: "Trampas mecánicas (N° + código de área)", tipo: "par",
      items: Array.from({ length: 23 }, (_, i) => [String(i + 1), "Por asignar"]) },

    // ---- FT-19 Insectocutores ----
    { id: "ft19.sectores", formato: "FT-19", nombre: "Insectocutores (N° y sector)", tipo: "codigo",
      items: [["1", "Ingreso a planta", false], ["3", "Oficina clientes", false], ["5", "Almacén", false], ["7", "Entrada a línea de palta nueva", false], ["8", "Volcado palta", false], ["9", "Cámara", false], ["10", "Comedor", false], ["12", "Ingreso de personal", false], ["13", "Recepción", false], ["14", "Armado de cajas", false], ["15", "Congelado", false]] },
    { id: "ft19.estado", formato: "FT-19", nombre: "Estado del insectocutor", tipo: "codigo",
      items: [["OK", "Bueno", false], ["R", "Roto", true], ["Q", "Quemado", true], ["M", "Malogrado", true]] },
    { id: "ft19.especies", formato: "FT-19", nombre: "Especies del conteo (especie + orden)", tipo: "codigo",
      items: [["MOSCA", "Díptero", false], ["MOSQUITO", "Díptero", false], ["ZANCUDO", "Díptero", false], ["POLILLA", "Lepidóptero", false], ["MARIPOSA", "Lepidóptero", false], ["CAPACHO", "Coleóptero", false], ["MARIQUITA", "Coleóptero", false], ["AVISPA", "Himenóptero", false], ["ABEJA", "Himenóptero", false], ["LIBÉLULA", "Odonata", false], ["OTROS", "Otros", false]] },
    { id: "ft19.limTrampa", formato: "FT-19", nombre: "Límite de insectos por trampa (No conforme si llega)", tipo: "numero", valor: 100 },
    { id: "ft19.limEspecie", formato: "FT-19", nombre: "Límite por especie (suma de todas las trampas)", tipo: "numero", valor: 50 },
    { id: "ft19.limTotal", formato: "FT-19", nombre: "Límite total por inspección", tipo: "numero", valor: 200 },

    // ---- FT-07 Hermeticidad, infraestructura y monitoreo de aves ----
    { id: "ft07.localizacion", formato: "FT-07", nombre: "Tipo de localización", tipo: "codigo",
      items: [["PR", "Proceso", false], ["C", "Cítrico", false], ["P", "Palta", false], ["G", "Granada", false], ["E", "Espárrago", false], ["EC", "Espárrago congelado", false], ["F", "Frío", false], ["A", "Almacén", false], ["M", "Mantenimiento", false], ["EX", "Exteriores", false]] },
    // área: [localización, área, aspecto a considerar, frecuencia, incluye monitoreo de aves]
    { id: "ft07.areas", formato: "FT-07", nombre: "Áreas a verificar", tipo: "area",
      items: [
        ["PR", "Oficina Administrativa", "Estado del techo, paredes, puertas, conexión eléctrica, pisos, ventilación, ventanas.", "Diario", false],
        ["EX", "SS.HH. y vestuarios de damas", "Estado del techo, paredes, puerta, piso, grifo, lavamanos, duchas, sanitarios, espejo, casilleros, divisiones de baños.", "Diario", false],
        ["EX", "SS.HH. y vestuarios de varones", "Estado del techo, paredes, puerta, piso, grifo, lavamanos, duchas, sanitarios, urinarios, espejo, casilleros, divisiones de baños.", "Diario", false],
        ["C", "Recepción", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico.", "Diario", false],
        ["G", "Recepción", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico.", "Diario", false],
        ["P", "Recepción", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico.", "Diario", false],
        ["E", "Recepción", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico, verificación de incidencias de aves.", "Diario", true],
        ["G", "Sopleteo", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico.", "Diario", false],
        ["E", "Escaldado", "Estado del techo, paredes, pisos, ventiladores, tablero eléctrico.", "Diario", false],
        ["PR", "Mercado Local", "Estado del techo, paredes, pisos, tablero eléctrico.", "Diario", false],
        ["C", "Volcado", "Estado del techo, paredes, pisos, tablero eléctrico.", "Diario", false],
        ["P", "Volcado", "Estado del techo, paredes, pisos, tablero eléctrico.", "Diario", false],
        ["F", "Cámaras de producto terminado", "Estado del techo, paredes, pisos, ventiladores, puertas.", "Diario", false],
        ["C", "Cámaras de producto terminado", "Estado del techo, paredes, pisos, ventiladores, puertas.", "Diario", false],
        ["F", "Túnel", "Estado del techo, paredes, pisos, ventiladores, puertas.", "Diario", false],
        ["C", "Túnel", "Estado del techo, paredes, pisos, ventiladores, puertas.", "Diario", false],
        ["C", "Zona de empaque", "Estado del techo, pisos, puertas.", "Diario", false],
        ["G", "Zona de empaque", "Estado del techo, pisos, puertas.", "Diario", false],
        ["P", "Zona de empaque", "Estado del techo, pisos, puertas.", "Diario", false],
        ["E", "Zona de empaque", "Estado del techo, pisos, puertas.", "Diario", false],
        ["EC", "Zona de empaque", "Estado del techo, pisos, puertas.", "Diario", false],
        ["C", "Caseta de mando", "Estado del techo, paredes, pisos, maquinarias, ventiladores, puertas, electricidad, tableros.", "Diario", false],
        ["P", "Caseta de mando", "Estado del techo, paredes, pisos, maquinarias, ventiladores, puertas, electricidad, tableros.", "Diario", false],
        ["A", "Armado de cajas", "Estado del techo, paredes, pisos, puertas, orden, maquinarias.", "Diario", false],
        ["A", "Almacén", "Estado del techo, paredes, pisos, ventiladores, puerta, orden, electricidad.", "Diario", false],
        ["M", "Taller de mantenimiento", "Estado del piso, orden, techo.", "Diario", false],
        ["M", "Zona de motores", "Estado del piso, limpieza (verificación de incidencia de aves).", "Diario", true],
        ["EX", "Lavandería", "Estado del piso, techo, paredes (verificación de incidencia de aves).", "Diario", true],
        ["EX", "Comedor", "Estado de pisos, paredes, techo, orden.", "Diario", false]] }
  ]
};
