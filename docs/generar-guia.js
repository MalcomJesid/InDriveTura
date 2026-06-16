const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ─── Colores ───────────────────────────────────────────────────────────────
const AZUL       = "1E3A5F";
const AZUL_CLARO = "2E75B6";
const VERDE      = "1E7E34";
const NARANJA    = "E67E22";
const GRIS_CLARO = "F5F7FA";
const GRIS_BORDE = "CCCCCC";
const AMARILLO   = "FFF9C4";
const VERDE_CLARO= "E8F5E9";
const ROJO_CLARO = "FDEDEC";
const AZUL_FONDO = "EBF3FB";
const BLANCO     = "FFFFFF";

// ─── Helpers ───────────────────────────────────────────────────────────────
const border = (color = GRIS_BORDE) => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (color = GRIS_BORDE) => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });
const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };

function titulo(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, bold: true, color: BLANCO, size: 36, font: "Arial" })]
  });
}

function subtitulo(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: AZUL_CLARO } },
    children: [new TextRun({ text, bold: true, color: AZUL, size: 28, font: "Arial" })]
  });
}

function subtitulo3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, color: AZUL_CLARO, size: 24, font: "Arial" })]
  });
}

function parrafo(text, options = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...options })]
  });
}

function espacio() {
  return new Paragraph({ children: [new TextRun({ text: "", size: 22 })] });
}

function saltoPlantilla() {
  return new Paragraph({ children: [new PageBreak()] });
}

function cajaInfo(emoji, titulo_texto, contenido_parrafos, fondo = AZUL_FONDO, borde_color = AZUL_CLARO) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: borders(borde_color),
          shading: { fill: fondo, type: ShadingType.CLEAR },
          margins: { top: 150, bottom: 150, left: 200, right: 200 },
          width: { size: 9360, type: WidthType.DXA },
          children: [
            new Paragraph({
              spacing: { before: 60, after: 60 },
              children: [new TextRun({ text: `${emoji}  ${titulo_texto}`, bold: true, size: 24, font: "Arial", color: AZUL })]
            }),
            ...contenido_parrafos
          ]
        })
      ]
    })
    ]
  });
}

function codigoBloque(lineas) {
  const parrafos = lineas.map(linea =>
    new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: linea, font: "Courier New", size: 18, color: "D4D4D4" })]
    })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: borders("333333"),
        shading: { fill: "1E1E1E", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: 9360, type: WidthType.DXA },
        children: parrafos
      })
    ]
  })
  ]
  });
}

function tablaSimple(encabezados, filas, anchos) {
  const total = anchos.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    children: encabezados.map((h, i) =>
      new TableCell({
        borders: borders(AZUL_CLARO),
        shading: { fill: AZUL, type: ShadingType.CLEAR },
        margins: cellMargins,
        width: { size: anchos[i], type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text: h, bold: true, color: BLANCO, size: 20, font: "Arial" })]
        })]
      })
    )
  });
  const dataRows = filas.map((fila, ri) =>
    new TableRow({
      children: fila.map((celda, ci) =>
        new TableCell({
          borders: borders(GRIS_BORDE),
          shading: { fill: ri % 2 === 0 ? BLANCO : GRIS_CLARO, type: ShadingType.CLEAR },
          margins: cellMargins,
          width: { size: anchos[ci], type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: celda, size: 20, font: "Arial" })]
          })]
        })
      )
    })
  );
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: anchos,
    rows: [headerRow, ...dataRows]
  });
}

function itemLista(texto, nivel = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level: nivel },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: texto, size: 22, font: "Arial" })]
  });
}

function itemCheck(texto, hecho = true) {
  const emoji = hecho ? "✅" : "⏳";
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: `${emoji}  ${texto}`, size: 22, font: "Arial" })]
  });
}

// ─── PORTADA ───────────────────────────────────────────────────────────────
function portada() {
  return [
    espacio(), espacio(), espacio(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 100 },
      children: [new TextRun({ text: "🚛", size: 96 })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 60 },
      children: [new TextRun({ text: "InDriveTura", bold: true, size: 72, color: AZUL, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: "Guia de Aprendizaje - Diario del Proyecto", size: 32, color: AZUL_CLARO, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 200 },
      children: [new TextRun({ text: "Aplicacion movil para trazabilidad de vehiculos recolectores de basura", size: 24, color: "666666", font: "Arial", italics: true })]
    }),
    espacio(),
    tablaSimple(
      ["Campo", "Informacion"],
      [
        ["Estudiante", "MalcomJesid"],
        ["Materia", "Seminario II"],
        ["Repositorio", "github.com/MalcomJesid/InDriveTura"],
        ["Tecnologia principal", "Ionic Angular + Capacitor"],
        ["Backend API", "apirecoleccion.gonzaloandreslucio.com"],
        ["Inicio del proyecto", "Mayo 2026"],
      ],
      [3500, 5860]
    ),
    espacio(), espacio(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Este documento se actualiza cada vez que aprendemos algo nuevo.", size: 20, color: "888888", font: "Arial", italics: true })]
    }),
    saltoPlantilla()
  ];
}

// ─── SECCION 1 — QUE ES ESTE PROYECTO ────────────────────────────────────
function seccion1() {
  return [
    subtitulo("SECCION 1 — ?Que es este proyecto?"),
    espacio(),
    cajaInfo("🎯", "La idea en palabras simples", [
      parrafo("Imagina que eres supervisor de los camiones de basura de tu ciudad. Cada manana salen 20 camiones a recoger la basura. Pero... ?como sabes si realmente pasaron por todas las calles? ?Y si un conductor dijo que fue a una calle pero no fue?"),
      espacio(),
      parrafo("Este proyecto resuelve ese problema. Creamos una aplicacion movil que:"),
      itemLista("El conductor lleva su celular con la app abierta."),
      itemLista("La app guarda automaticamente donde esta el camion cada 30 segundos."),
      itemLista("El conductor puede tomar fotos de evidencia (basureros llenos, problemas)."),
      itemLista("Todo se guarda en un servidor. El supervisor puede ver el recorrido completo."),
      itemLista("Si no hay internet, la app guarda todo y lo sube cuando vuelva la senal."),
    ], VERDE_CLARO, VERDE),
    espacio(),
    subtitulo3("?Que tecnologias usamos?"),
    tablaSimple(
      ["Tecnologia", "Que es", "Para que la usamos"],
      [
        ["Ionic Angular", "Framework para apps moviles", "Crear las pantallas de la app"],
        ["Capacitor", "Puente nativo", "Acceder al GPS y la camara del celular"],
        ["Supabase", "Backend en la nube", "Registro y login de conductores"],
        ["SQLite", "Base de datos local", "Guardar datos cuando no hay internet"],
        ["API del profesor", "Servidor central (Laravel)", "Registrar todos los recorridos y posiciones"],
        ["Git + GitHub", "Control de versiones", "Guardar el historial del codigo"],
      ],
      [2200, 2800, 4360]
    ),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 2 — LA API ────────────────────────────────────────────────────
function seccion2() {
  return [
    subtitulo("SECCION 2 — La API del profesor"),
    espacio(),
    cajaInfo("💡", "?Que es una API?", [
      parrafo("Piensa en una API como el menu de un restaurante. Tu no entras a la cocina a preparar tu comida. Le dices al mesero (API) lo que quieres, el va a la cocina (servidor), y te trae lo que pediste."),
      espacio(),
      parrafo("En nuestro caso: nuestra app le pide datos a la API del profesor, y la API le responde con informacion real de la base de datos."),
    ], AMARILLO, NARANJA),
    espacio(),
    subtitulo3("Todos los endpoints disponibles"),
    parrafo("Un endpoint es una direccion especifica de la API. Cada uno hace una cosa diferente."),
    espacio(),
    tablaSimple(
      ["Metodo", "Endpoint", "Para que sirve"],
      [
        ["GET",  "/api/perfiles/todas",                        "Obtener lista de conductores"],
        ["GET",  "/api/vehiculos",                             "Obtener lista de vehiculos"],
        ["GET",  "/api/rutas/todas",                           "Obtener todas las rutas"],
        ["POST", "/api/recorridos/iniciar",                    "Iniciar un recorrido"],
        ["POST", "/api/recorridos/{id}/finalizar",             "Finalizar un recorrido"],
        ["GET",  "/api/misrecorridos",                         "Ver mis recorridos"],
        ["POST", "/api/recorridos/{id}/posiciones",            "Registrar una coordenada GPS"],
        ["GET",  "/api/recorridos/{id}/posiciones",            "Ver posiciones de un recorrido"],
        ["POST", "/api/recorridos/posiciones/{id}/imagen",     "Subir foto de evidencia en Base64"],
        ["GET",  "/api/recorridos/posiciones/{id}/imagen",     "Ver foto de una posicion"],
      ],
      [1200, 3900, 4260]
    ),
    espacio(),
    subtitulo3("Estructura de respuesta de la API"),
    parrafo("La API SIEMPRE envuelve los datos en un objeto con la clave 'data'. Esto es importante para no cometer errores al programar."),
    espacio(),
    codigoBloque([
      "// Lo que devuelve la API (NO es un arreglo directo):",
      "{",
      '  "data": [',
      '    { "id": "abc-123", "nombre_perfil": "Grupo - 1" },',
      '    { "id": "def-456", "nombre_perfil": "Grupo - 2" }',
      "  ]",
      "}",
      "",
      "// Por eso en el codigo usamos: respuesta.data",
      "// NO: respuesta directamente",
    ]),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 3 — AMBIENTE ─────────────────────────────────────────────────
function seccion3() {
  return [
    subtitulo("SECCION 3 — Ambiente de desarrollo"),
    espacio(),
    parrafo("Antes de escribir una sola linea de codigo, necesitas tener instaladas las herramientas correctas. Es como un cocinero que necesita su cocina equipada antes de cocinar."),
    espacio(),
    subtitulo3("Herramientas instaladas en tu PC"),
    tablaSimple(
      ["Herramienta", "Version", "Para que sirve"],
      [
        ["Node.js",      "v22.14.0",   "Motor que ejecuta JavaScript fuera del navegador"],
        ["npm",          "v11.12.0",   "Instalador de librerias y herramientas"],
        ["Git",          "v2.49.0",    "Control de versiones del codigo"],
        ["Java JDK",     "OpenJDK 21", "Necesario para compilar para Android"],
        ["Android SDK",  "v35.0.2",    "Herramientas de Android"],
        ["Android Studio","Instalado", "Emulador y compilacion Android"],
        ["VS Code",      "v1.114.0",   "Editor de codigo"],
        ["Ionic CLI",    "v7.2.1",     "Crear y manejar proyectos Ionic"],
        ["Angular CLI",  "v21.2.13",   "Generar componentes y servicios"],
      ],
      [2500, 2000, 4860]
    ),
    espacio(),
    subtitulo3("Variables de entorno configuradas"),
    cajaInfo("⚙️", "JAVA_HOME configurado", [
      parrafo("Ruta: C:\\Program Files\\Eclipse Adoptium\\jdk-21.0.10.7-hotspot"),
      parrafo("Esto le dice a Android Studio y a las herramientas donde esta instalado Java en tu PC."),
    ], GRIS_CLARO, GRIS_BORDE),
    espacio(),
    subtitulo3("Comandos importantes de la terminal"),
    codigoBloque([
      "cd carpeta          # entrar a una carpeta",
      "cd ..               # subir un nivel",
      "pwd                 # ver en que carpeta estas",
      "ls                  # listar archivos",
      "ionic serve         # correr la app en el navegador",
      "ionic start nombre  # crear proyecto nuevo",
      "ng generate service ruta/nombre  # crear un servicio",
    ]),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 4 — GIT Y GITHUB ─────────────────────────────────────────────
function seccion4() {
  return [
    subtitulo("SECCION 4 — Git y GitHub"),
    espacio(),
    cajaInfo("💡", "Git vs GitHub — ?cual es la diferencia?", [
      parrafo("GIT es una herramienta que vive en TU computadora. Guarda el historial de cambios localmente."),
      parrafo("GITHUB es un sitio web en internet donde subes una copia de tu proyecto. Es el respaldo en la nube."),
      espacio(),
      parrafo("Analogia: Git es el fotografo que toma fotos de tu proyecto. GitHub es el album en la nube donde guardas esas fotos."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    subtitulo3("Los comandos que ya usaste"),
    codigoBloque([
      "git init                           # convertir carpeta en repositorio",
      "git remote add origin URL          # conectar con GitHub",
      "git status                         # ver que archivos cambiaron",
      "git add .                          # preparar TODOS los cambios",
      "git add archivo.ts                 # preparar UN archivo especifico",
      'git commit -m "mensaje"            # guardar los cambios con descripcion',
      "git push                           # subir commits a GitHub",
      "git pull                           # bajar cambios de GitHub",
      "git branch                         # ver las ramas existentes",
      "git checkout -b nombre             # crear rama nueva y cambiar a ella",
      "git checkout nombre                # cambiar a una rama existente",
      "git merge nombre                   # fusionar una rama con la actual",
      "git log                            # ver historial de commits",
    ]),
    espacio(),
    subtitulo3("Como escribir commits profesionales"),
    parrafo("Un buen commit dice QUE hiciste y POR QUE. Nunca escribas 'cambios' o 'arreglos'."),
    espacio(),
    tablaSimple(
      ["Prefijo", "Cuando usarlo", "Ejemplo"],
      [
        ["feat:",     "Nueva funcionalidad",         "feat: crear pantalla de seleccion de perfil"],
        ["fix:",      "Correccion de error",          "fix: corregir GPS cuando no hay permisos"],
        ["docs:",     "Documentacion",                "docs: actualizar README con instalacion"],
        ["style:",    "Formato, sin cambio de logica","style: ordenar imports en app.module.ts"],
        ["refactor:", "Reorganizacion de codigo",     "refactor: mover logica GPS a servicio"],
      ],
      [1500, 2800, 5060]
    ),
    espacio(),
    subtitulo3("Estructura de ramas del proyecto"),
    codigoBloque([
      "main             <- solo codigo listo para presentar. NUNCA trabajas aqui directo.",
      "develop          <- integracion del trabajo diario",
      "  |",
      "  +-- feature/auth        <- seleccion de perfil y login",
      "  +-- feature/vehicles    <- seleccion de vehiculo",
      "  +-- feature/routes      <- seleccion de ruta",
      "  +-- feature/tracking    <- GPS y recorrido activo",
      "  +-- feature/photos      <- camara y evidencias en Base64",
      "  +-- feature/offline     <- SQLite y sincronizacion",
      "  +-- feature/docs        <- documentacion",
    ]),
    espacio(),
    subtitulo3("Flujo de trabajo diario"),
    codigoBloque([
      "1. git checkout feature/nombre-funcionalidad",
      "2. ... escribes codigo ...",
      "3. git add .",
      '4. git commit -m "feat: descripcion de lo que hiciste"',
      "5. git push",
    ]),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 5 — IONIC Y ANGULAR ──────────────────────────────────────────
function seccion5() {
  return [
    subtitulo("SECCION 5 — Ionic y Angular"),
    espacio(),
    subtitulo3("La estructura del proyecto"),
    codigoBloque([
      "InDriveTura/",
      "  src/",
      "    app/",
      "      home/",
      "        home.page.ts       <- logica de la pantalla",
      "        home.page.html     <- vista (lo que se ve)",
      "        home.page.scss     <- estilos (como se ve)",
      "        home.module.ts     <- modulo de la pantalla",
      "      services/",
      "        api.ts             <- servicio que habla con la API",
      "      app.module.ts        <- registro central de Angular",
      "      app-routing.module.ts<- navegacion entre pantallas",
      "    index.html             <- HTML principal (solo uno)",
      "  package.json            <- lista de dependencias",
      "  capacitor.config.ts     <- configuracion de Capacitor",
    ]),
    espacio(),
    subtitulo3("Los tres conceptos fundamentales de Angular"),
    tablaSimple(
      ["Concepto", "Analogia", "Funcion real"],
      [
        ["Componente (Page)", "Una pantalla de la app",      "Tiene .ts (logica) + .html (vista) + .scss (estilos)"],
        ["Servicio (Service)","El asistente detras",          "Habla con la API, maneja datos, se comparte entre pantallas"],
        ["Modulo (Module)",   "El registro central",          "Le dice a Angular que componentes y herramientas existen"],
      ],
      [2500, 2500, 4360]
    ),
    espacio(),
    subtitulo3("Componentes de Ionic que ya usamos"),
    codigoBloque([
      "<ion-header>       <- barra superior de la pantalla",
      "<ion-toolbar>      <- contenedor dentro del header",
      "<ion-title>        <- titulo visible en la barra",
      "<ion-content>      <- area de contenido scrolleable",
      "<ion-list>         <- lista de elementos",
      "<ion-item>         <- un elemento de la lista",
      "<ion-label>        <- texto dentro de un item",
      "<ion-icon>         <- icono",
      "<ion-spinner>      <- animacion de carga",
    ]),
    espacio(),
    subtitulo3("Directivas de Angular en el HTML"),
    codigoBloque([
      "<!-- Mostrar u ocultar segun condicion -->",
      '<div *ngIf="cargando">Cargando...</div>',
      '<div *ngIf="!cargando">Contenido listo</div>',
      "",
      "<!-- Repetir por cada elemento de una lista -->",
      '<ion-item *ngFor="let perfil of perfiles">',
      "  {{ perfil.nombre_perfil }}",
      "</ion-item>",
      "",
      "<!-- Mostrar valor de una variable -->",
      "{{ perfil.nombre_perfil }}",
    ]),
    espacio(),
    subtitulo3("TypeScript basico"),
    codigoBloque([
      "// Variables con tipo declarado",
      "let nombre: string = 'Malcom'",
      "let edad: number = 25",
      "let activo: boolean = true",
      "let lista: string[] = []",
      "",
      "// Interface: define la estructura exacta de un objeto",
      "interface Perfil {",
      "  id: string",
      "  nombre_perfil: string",
      "}",
      "",
      "// Arreglo de objetos tipados",
      "let perfiles: Perfil[] = []",
    ]),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 6 — PRIMER CODIGO REAL ──────────────────────────────────────
function seccion6() {
  return [
    subtitulo("SECCION 6 — Tu primer codigo real"),
    espacio(),
    parrafo("En esta seccion construiste la primera funcionalidad real: consumir la API del profesor y mostrar la lista de conductores en pantalla."),
    espacio(),
    subtitulo3("Paso 1 — Registrar HttpClientModule en app.module.ts"),
    parrafo("Sin esto, Angular no puede hacer llamadas HTTP. Es como enchufar el cable de internet antes de navegar."),
    espacio(),
    codigoBloque([
      "// src/app/app.module.ts",
      "import { NgModule } from '@angular/core';",
      "import { BrowserModule } from '@angular/platform-browser';",
      "import { RouteReuseStrategy } from '@angular/router';",
      "import { HttpClientModule } from '@angular/common/http'; // <- AGREGADO",
      "",
      "import { IonicModule, IonicRouteStrategy } from '@ionic/angular';",
      "import { AppComponent } from './app.component';",
      "import { AppRoutingModule } from './app-routing.module';",
      "",
      "@NgModule({",
      "  declarations: [AppComponent],",
      "  imports: [",
      "    BrowserModule,",
      "    IonicModule.forRoot(),",
      "    AppRoutingModule,",
      "    HttpClientModule, // <- AGREGADO",
      "  ],",
      "  providers: [{ provide: RouteReuseStrategy, useClass: IonicRouteStrategy }],",
      "  bootstrap: [AppComponent],",
      "})",
      "export class AppModule {}",
    ]),
    espacio(),
    subtitulo3("Paso 2 — Crear el servicio api.ts"),
    parrafo("El servicio es el intermediario entre la pantalla y la API. La pantalla nunca habla directo con la API."),
    espacio(),
    codigoBloque([
      "// src/app/services/api.ts",
      "import { Injectable } from '@angular/core';",
      "import { HttpClient } from '@angular/common/http';",
      "import { Observable } from 'rxjs';",
      "",
      "// Define la estructura exacta de un perfil",
      "export interface ApiResponse {",
      "  data: Perfil[];",
      "}",
      "",
      "export interface Perfil {",
      "  id: string;",
      "  nombre_perfil: string;",
      "}",
      "",
      "@Injectable({ providedIn: 'root' })",
      "export class Api {",
      "",
      "  private baseUrl = 'https://apirecoleccion.gonzaloandreslucio.com/api';",
      "",
      "  // Angular inyecta HttpClient automaticamente",
      "  constructor(private http: HttpClient) {}",
      "",
      "  // Llama a GET /api/perfiles/todas",
      "  getPerfiles(): Observable<ApiResponse> {",
      "    return this.http.get<ApiResponse>(`${this.baseUrl}/perfiles/todas`);",
      "  }",
      "",
      "}",
    ]),
    espacio(),
    subtitulo3("Paso 3 — Modificar home.page.ts"),
    parrafo("La pantalla usa el servicio para pedir los datos y los guarda en una variable local."),
    espacio(),
    codigoBloque([
      "// src/app/home/home.page.ts",
      "import { Component, OnInit } from '@angular/core';",
      "import { Api, Perfil } from '../services/api';",
      "",
      "@Component({",
      "  selector: 'app-home',",
      "  templateUrl: 'home.page.html',",
      "  styleUrls: ['home.page.scss'],",
      "  standalone: false,",
      "})",
      "export class HomePage implements OnInit {",
      "",
      "  perfiles: Perfil[] = [];    // lista que llena la API",
      "  cargando: boolean = true;   // muestra spinner mientras espera",
      "  error: string = '';         // mensaje si algo falla",
      "",
      "  constructor(private apiService: Api) {}",
      "",
      "  // Se ejecuta automaticamente cuando la pantalla carga",
      "  ngOnInit() {",
      "    this.cargarPerfiles();",
      "  }",
      "",
      "  cargarPerfiles() {",
      "    this.apiService.getPerfiles().subscribe({",
      "      next: (respuesta) => {",
      "        this.perfiles = respuesta.data; // extrae el arreglo del objeto",
      "        this.cargando = false;",
      "      },",
      "      error: (err) => {",
      "        this.error = 'No se pudo conectar con la API';",
      "        this.cargando = false;",
      "        console.error(err);",
      "      }",
      "    });",
      "  }",
      "",
      "}",
    ]),
    espacio(),
    subtitulo3("Paso 4 — Modificar home.page.html"),
    codigoBloque([
      "<!-- src/app/home/home.page.html -->",
      '<ion-header [translucent]="true">',
      '  <ion-toolbar color="primary">',
      "    <ion-title>InDriveTura</ion-title>",
      "  </ion-toolbar>",
      "</ion-header>",
      "",
      '<ion-content [fullscreen]="true">',
      "",
      '  <!-- Mientras carga -->',
      '  <div *ngIf="cargando" style="text-align: center; margin-top: 40px;">',
      '    <ion-spinner name="crescent"></ion-spinner>',
      "    <p>Cargando conductores...</p>",
      "  </div>",
      "",
      "  <!-- Si hay error -->",
      '  <div *ngIf="error">',
      "    <p>{{ error }}</p>",
      "  </div>",
      "",
      "  <!-- Lista de perfiles -->",
      '  <ion-list *ngIf="!cargando && !error">',
      "    <ion-list-header>",
      "      <ion-label>Selecciona tu perfil</ion-label>",
      "    </ion-list-header>",
      '    <ion-item *ngFor="let perfil of perfiles" button>',
      '      <ion-icon name="person-circle-outline" slot="start"></ion-icon>',
      "      <ion-label>{{ perfil.nombre_perfil }}</ion-label>",
      "    </ion-item>",
      "  </ion-list>",
      "",
      "</ion-content>",
    ]),
    espacio(),
    cajaInfo("🏆", "Resultado logrado", [
      parrafo("La app muestra 20 conductores reales (Grupo-1 a Grupo-20) obtenidos directamente de la API del profesor. Arquitectura cliente-servidor funcionando en la vida real."),
    ], VERDE_CLARO, VERDE),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 7 — PROGRESO ─────────────────────────────────────────────────
function seccion7() {
  return [
    subtitulo("SECCION 7 — Progreso del proyecto"),
    espacio(),
    subtitulo3("Semana 1 — Completada"),
    itemCheck("Instalar y configurar el ambiente de desarrollo", true),
    itemCheck("Instalar Ionic CLI y Angular CLI", true),
    itemCheck("Configurar JAVA_HOME", true),
    itemCheck("Crear proyecto Ionic con NgModules", true),
    itemCheck("Ver app corriendo en ionic serve", true),
    itemCheck("Crear repositorio en GitHub", true),
    itemCheck("Hacer primer commit profesional", true),
    itemCheck("Crear estructura de 9 ramas", true),
    itemCheck("Crear servicio ApiService", true),
    itemCheck("Registrar HttpClientModule", true),
    itemCheck("Consumir GET /api/perfiles/todas", true),
    itemCheck("Mostrar lista de conductores en pantalla", true),
    espacio(),
    subtitulo3("Proximas semanas"),
    itemCheck("Semana 2 — Seleccion de perfil, vehiculo y ruta", false),
    itemCheck("Semana 3 — Iniciar y finalizar recorrido + GPS", false),
    itemCheck("Semana 4 — Camara y evidencias en Base64", false),
    itemCheck("Semana 5 — Modo offline con SQLite", false),
    itemCheck("Semana 6 — Sincronizacion, pruebas y documentacion", false),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 8 — PREGUNTAS DEL PROFESOR ──────────────────────────────────
function seccion8() {
  return [
    subtitulo("SECCION 8 — Como defender el proyecto"),
    espacio(),
    parrafo("Estas son las preguntas mas probables en la sustentacion. Estudia las respuestas hasta que puedas decirlas con tus propias palabras."),
    espacio(),
    subtitulo3("Sobre las tecnologias"),
    cajaInfo("❓", "?Que es Ionic y por que lo elegiste?", [
      parrafo("Ionic es un framework para crear aplicaciones moviles usando tecnologias web (HTML, CSS, JavaScript). Lo elegimos porque con un solo codigo fuente la app funciona en Android e iOS. Ademas usa Angular, que es un estandar en la industria."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    cajaInfo("❓", "?Cual es la diferencia entre Ionic, Angular y Capacitor?", [
      parrafo("Angular es el motor: maneja la logica, la navegacion y los datos. Ionic es la carroceria: provee los componentes visuales adaptados para movil (botones, listas, tarjetas). Capacitor son las ruedas: conecta el codigo web con las funciones nativas del celular como el GPS y la camara."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    cajaInfo("❓", "?Que es una API REST?", [
      parrafo("Es un conjunto de URLs (endpoints) que permiten comunicarse con un servidor. Cada URL hace una cosa especifica: obtener datos, crear datos, actualizar o eliminar. Se comunican usando el protocolo HTTP y los datos viajan en formato JSON."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    cajaInfo("❓", "?Que es JSON?", [
      parrafo("JSON (JavaScript Object Notation) es un formato de texto para intercambiar datos entre sistemas. Es como un formulario en texto plano que tanto el servidor como la app pueden leer y escribir facilmente. Ejemplo: {\"nombre\": \"Juan\", \"edad\": 30}"),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    subtitulo3("Sobre el codigo"),
    cajaInfo("❓", "?Que es un Observable y por que lo usas?", [
      parrafo("Un Observable es un flujo de datos asincrono. Cuando la app le pide datos a la API, no se queda congelada esperando. El Observable avisa cuando llegan los datos. Me suscribo con .subscribe() y cuando la API responde, Angular actualiza la pantalla automaticamente."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    cajaInfo("❓", "?Por que separas la logica en un servicio?", [
      parrafo("Porque cada parte del codigo debe tener una sola responsabilidad. La pantalla (componente) se encarga de mostrar informacion. El servicio se encarga de obtener informacion. Si manana cambia la URL de la API, solo cambio el servicio. Las pantallas no se tocan."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    cajaInfo("❓", "?Que es la inyeccion de dependencias?", [
      parrafo("Es cuando Angular crea y entrega automaticamente los objetos que necesito. En lugar de crear el servicio con 'new Api()', lo declaro en el constructor y Angular lo inyecta. Esto hace el codigo mas facil de probar y mantener."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    subtitulo3("Sobre Git"),
    cajaInfo("❓", "?Para que sirven las ramas en Git?", [
      parrafo("Las ramas me permiten trabajar en una funcionalidad nueva sin afectar el codigo que ya funciona. Cada funcionalidad tiene su propia rama. Si algo sale mal, elimino la rama. Cuando termino y funciona bien, fusiono con develop. main siempre tiene codigo estable."),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── SECCION 9 — GLOSARIO ─────────────────────────────────────────────────
function seccion9() {
  return [
    subtitulo("SECCION 9 — Glosario de terminos"),
    espacio(),
    parrafo("Si el profesor menciona alguno de estos terminos y no lo recuerdas, aqui estan las definiciones simples."),
    espacio(),
    tablaSimple(
      ["Termino", "Definicion simple"],
      [
        ["API",              "Direcciones de un servidor para pedir o enviar datos"],
        ["Endpoint",        "Una URL especifica de la API que hace una cosa concreta"],
        ["HTTP",            "Protocolo de comunicacion entre apps y servidores en internet"],
        ["GET",             "Metodo HTTP para PEDIR datos (sin enviar cuerpo)"],
        ["POST",            "Metodo HTTP para ENVIAR datos nuevos al servidor"],
        ["JSON",            "Formato de texto para intercambiar datos entre sistemas"],
        ["Observable",      "Flujo de datos asincrono. Los datos llegan cuando estan listos."],
        ["Componente",      "Una pantalla de Angular. Tiene logica (.ts) y vista (.html)"],
        ["Servicio",        "Clase Angular que maneja datos y llama a la API"],
        ["Modulo",          "Registro central que le dice a Angular que existe en la app"],
        ["TypeScript",      "JavaScript con tipos declarados. Detecta errores antes de correr"],
        ["Interface",       "Define la estructura exacta de un objeto en TypeScript"],
        ["Inyeccion",       "Angular entrega automaticamente los objetos que necesitas"],
        ["ngOnInit",        "Metodo que Angular ejecuta automaticamente al cargar la pantalla"],
        ["*ngIf",           "Directiva para mostrar/ocultar elementos segun una condicion"],
        ["*ngFor",          "Directiva para repetir un elemento por cada item de una lista"],
        ["Commit",          "Fotografia del estado del proyecto en un momento especifico"],
        ["Rama (Branch)",   "Linea de desarrollo independiente en Git"],
        ["Push",            "Subir commits locales a GitHub"],
        ["Pull",            "Bajar cambios de GitHub a tu computadora"],
        ["UUID",            "Identificador unico universal (ej: 18851282-1a08-42b7-9384-xxx)"],
        ["Base64",          "Forma de convertir una imagen en texto para enviarla por la API"],
        ["SQLite",          "Base de datos que vive en el celular. Funciona sin internet."],
        ["Capacitor",       "Puente entre el codigo web y las funciones nativas del celular"],
        ["GPS",             "Sistema de posicionamiento global. Coordenadas de ubicacion."],
        ["Offline First",   "Arquitectura donde la app funciona sin internet y sincroniza despues"],
      ],
      [2800, 6560]
    ),
    espacio(),
  ];
}

// ─── SECCION 10 — CONFIGURACIONES Y COMANDOS ─────────────────────────────
function seccion10() {
  return [
    subtitulo("SECCION 10 — Configuraciones y Comandos"),
    espacio(),
    parrafo("Esta seccion es tu hoja de referencia rapida. Todo lo que necesitas configurar o ejecutar esta aqui, ordenado por tema."),
    espacio(),

    subtitulo3("CONFIGURACION DEL AMBIENTE"),
    espacio(),
    cajaInfo("⚙️", "Variables de entorno — Windows", [
      parrafo("Las variables de entorno le dicen al sistema operativo donde estan las herramientas instaladas."),
      espacio(),
      parrafo("JAVA_HOME — apunta a tu instalacion de Java:"),
    ], GRIS_CLARO, GRIS_BORDE),
    espacio(),
    codigoBloque([
      "# Ver si JAVA_HOME esta configurado",
      "$env:JAVA_HOME",
      "",
      "# Configurar JAVA_HOME permanentemente (PowerShell como administrador)",
      '[System.Environment]::SetEnvironmentVariable("JAVA_HOME", "C:\\Program Files\\Eclipse Adoptium\\jdk-21.0.10.7-hotspot", "User")',
      "",
      "# Valor actual en tu PC:",
      "# C:\\Program Files\\Eclipse Adoptium\\jdk-21.0.10.7-hotspot",
    ]),
    espacio(),

    cajaInfo("🔒", "Problema de SSL en redes universitarias", [
      parrafo("En redes con proxy o antivirus, npm no puede verificar los certificados de seguridad. Solucion:"),
    ], ROJO_CLARO, "E74C3C"),
    espacio(),
    codigoBloque([
      "# Desactivar verificacion SSL en npm (permanente)",
      "npm config set strict-ssl false",
      "",
      "# Desactivar para una sola sesion de PowerShell",
      '$env:NODE_TLS_REJECT_UNAUTHORIZED = "0"',
      "",
      "# Verificar configuracion actual de npm",
      "npm config list",
    ]),
    espacio(),

    subtitulo3("COMANDOS NPM"),
    espacio(),
    tablaSimple(
      ["Comando", "Para que sirve"],
      [
        ["npm install",                    "Instalar todas las dependencias del package.json"],
        ["npm install -g @ionic/cli",      "Instalar Ionic CLI de forma global"],
        ["npm install -g @angular/cli",    "Instalar Angular CLI de forma global"],
        ["npm install nombre-paquete",     "Instalar una libreria en el proyecto actual"],
        ["npm list -g",                    "Ver todas las herramientas instaladas globalmente"],
        ["npm --version",                  "Ver version de npm instalada"],
        ["npm fund",                       "Ver paquetes que solicitan financiamiento (solo informativo)"],
      ],
      [4200, 5160]
    ),
    espacio(),

    subtitulo3("COMANDOS IONIC CLI"),
    espacio(),
    tablaSimple(
      ["Comando", "Para que sirve"],
      [
        ["ionic start Nombre blank --type=angular --capacitor", "Crear proyecto nuevo con plantilla vacia"],
        ["ionic serve",                                          "Correr app en navegador (localhost:8100)"],
        ["ionic build",                                          "Compilar app para produccion"],
        ["ionic cap add android",                                "Agregar plataforma Android al proyecto"],
        ["ionic cap sync",                                       "Sincronizar cambios web con plataformas nativas"],
        ["ionic cap run android",                                "Correr en emulador o dispositivo Android"],
        ["ionic cap open android",                               "Abrir proyecto en Android Studio"],
        ["ionic generate page nombre",                           "Crear una nueva pagina/pantalla"],
        ["ionic --version",                                      "Ver version de Ionic CLI instalada"],
      ],
      [4800, 4560]
    ),
    espacio(),

    subtitulo3("COMANDOS ANGULAR CLI"),
    espacio(),
    tablaSimple(
      ["Comando", "Para que sirve"],
      [
        ["ng generate page nombre",           "Crear nueva pantalla (page)"],
        ["ng generate service ruta/nombre",   "Crear nuevo servicio"],
        ["ng generate component nombre",      "Crear nuevo componente"],
        ["ng build",                           "Compilar el proyecto Angular"],
        ["ng version",                         "Ver version de Angular CLI y dependencias"],
        ["ng lint",                            "Verificar errores de estilo en el codigo"],
      ],
      [4200, 5160]
    ),
    espacio(),

    subtitulo3("COMANDOS GIT — REFERENCIA COMPLETA"),
    espacio(),
    cajaInfo("📋", "Flujo diario de trabajo", [
      parrafo("Estos son los comandos que ejecutas cada vez que trabajas en el proyecto:"),
    ], AZUL_FONDO, AZUL_CLARO),
    espacio(),
    codigoBloque([
      "# 1. Verificar en que rama estas",
      "git branch",
      "",
      "# 2. Cambiar a la rama de la funcionalidad",
      "git checkout feature/auth",
      "",
      "# 3. Ver que archivos cambiaste",
      "git status",
      "",
      "# 4. Ver exactamente que cambio dentro de los archivos",
      "git diff",
      "",
      "# 5. Preparar todos los cambios",
      "git add .",
      "",
      "# 6. Guardar con mensaje descriptivo",
      'git commit -m "feat: descripcion de lo que hiciste"',
      "",
      "# 7. Subir a GitHub",
      "git push",
    ]),
    espacio(),
    cajaInfo("🌿", "Comandos de ramas", [
      parrafo("Para crear, cambiar y fusionar ramas:"),
    ], VERDE_CLARO, VERDE),
    espacio(),
    codigoBloque([
      "# Crear rama nueva y cambiar a ella",
      "git checkout -b feature/nueva-funcionalidad",
      "",
      "# Cambiar a una rama existente",
      "git checkout develop",
      "",
      "# Ver todas las ramas (locales y remotas)",
      "git branch -a",
      "",
      "# Fusionar una rama con la rama actual",
      "git merge feature/auth",
      "",
      "# Subir rama nueva a GitHub por primera vez",
      "git push -u origin feature/nueva-funcionalidad",
      "",
      "# Eliminar rama local (cuando ya no la necesitas)",
      "git branch -d feature/terminada",
    ]),
    espacio(),
    cajaInfo("🔍", "Comandos de inspeccion", [
      parrafo("Para revisar el estado e historia del proyecto:"),
    ], AMARILLO, NARANJA),
    espacio(),
    codigoBloque([
      "# Ver historial de commits (resumido)",
      "git log --oneline",
      "",
      "# Ver historial con grafico de ramas",
      "git log --oneline --graph --all",
      "",
      "# Ver quien cambio cada linea de un archivo",
      "git blame archivo.ts",
      "",
      "# Descargar cambios sin aplicarlos",
      "git fetch",
      "",
      "# Descargar y aplicar cambios del servidor",
      "git pull",
      "",
      "# Ver diferencia con la version en GitHub",
      "git diff origin/develop",
    ]),
    espacio(),

    subtitulo3("CONFIGURACION DE GIT — PRIMERA VEZ"),
    espacio(),
    codigoBloque([
      "# Configurar nombre y correo (se muestra en cada commit)",
      'git config --global user.name "MalcomJesid"',
      'git config --global user.email "malconyfigue@gmail.com"',
      "",
      "# Verificar la configuracion",
      "git config --global user.name",
      "git config --global user.email",
      "",
      "# Ver toda la configuracion global",
      "git config --global --list",
    ]),
    espacio(),

    subtitulo3("ESTRUCTURA DE RAMAS DEL PROYECTO"),
    espacio(),
    tablaSimple(
      ["Rama", "Proposito", "Cuando se usa"],
      [
        ["main",              "Codigo estable final",              "Solo al entregar el proyecto"],
        ["develop",           "Integracion del trabajo diario",    "Recibe merges de las feature"],
        ["feature/auth",      "Login y seleccion de perfil",       "Semana 2"],
        ["feature/vehicles",  "Seleccion de vehiculo",             "Semana 2"],
        ["feature/routes",    "Seleccion de ruta",                 "Semana 2"],
        ["feature/tracking",  "GPS y recorrido activo",            "Semana 3"],
        ["feature/photos",    "Camara y evidencias Base64",        "Semana 4"],
        ["feature/offline",   "SQLite y sincronizacion",           "Semana 5"],
        ["feature/docs",      "Documentacion del proyecto",        "Semana 6"],
      ],
      [2200, 3500, 3660]
    ),
    espacio(),

    subtitulo3("ARCHIVOS DE CONFIGURACION DEL PROYECTO"),
    espacio(),
    cajaInfo("📄", "capacitor.config.ts — Configuracion de Capacitor", [
      parrafo("Este archivo le dice a Capacitor el nombre de la app, el ID del paquete y la carpeta de build."),
    ], GRIS_CLARO, GRIS_BORDE),
    espacio(),
    codigoBloque([
      "// capacitor.config.ts",
      "import { CapacitorConfig } from '@capacitor/cli';",
      "",
      "const config: CapacitorConfig = {",
      "  appId: 'com.indrivetura.app',      // ID unico de la app en Play Store",
      "  appName: 'InDriveTura',             // Nombre visible en el celular",
      "  webDir: 'www',                      // Carpeta de archivos compilados",
      "  server: {",
      "    androidScheme: 'https'            // Protocolo para Android",
      "  }",
      "};",
      "",
      "export default config;",
    ]),
    espacio(),
    cajaInfo("📄", "ionic.config.json — Configuracion de Ionic", [
      parrafo("Identifica el proyecto ante el Ionic CLI y define el tipo de app."),
    ], GRIS_CLARO, GRIS_BORDE),
    espacio(),
    codigoBloque([
      "// ionic.config.json",
      "{",
      '  "name": "InDriveTura",',
      '  "integrations": {',
      '    "capacitor": {}',
      "  },",
      '  "type": "angular"',
      "}",
    ]),
    espacio(),
    cajaInfo("📄", "package.json — Dependencias del proyecto", [
      parrafo("Lista todas las librerias que usa el proyecto. Al ejecutar npm install, descarga todo lo que esta aqui."),
    ], GRIS_CLARO, GRIS_BORDE),
    espacio(),
    codigoBloque([
      "// Estructura basica de package.json",
      "{",
      '  "name": "in-drive-tura",',
      '  "version": "0.0.1",',
      '  "scripts": {',
      '    "ng": "ng",',
      '    "start": "ng serve",',
      '    "build": "ng build",',
      '    "test": "ng test"',
      "  },",
      '  "dependencies": {',
      '    "@angular/core": "^19.0.0",',
      '    "@ionic/angular": "^8.0.0",',
      '    "@capacitor/core": "^7.0.0"',
      "    // ... mas dependencias",
      "  }",
      "}",
    ]),
    espacio(),

    subtitulo3("SOLUCIONES A PROBLEMAS COMUNES"),
    espacio(),
    tablaSimple(
      ["Problema", "Causa", "Solucion"],
      [
        ["npm error UNABLE_TO_VERIFY",  "Red con proxy/antivirus",      "npm config set strict-ssl false"],
        ["ionic: command not found",    "CLI no instalado",              "npm install -g @ionic/cli"],
        ["ng: command not found",       "Angular CLI no instalado",      "npm install -g @angular/cli"],
        ["Port 8100 already in use",    "Ya hay una instancia corriendo","Cerrar la terminal anterior o usar ionic serve --port=8101"],
        ["Cannot find module",          "Dependencias no instaladas",    "npm install dentro de la carpeta del proyecto"],
        ["JAVA_HOME not set",           "Variable no configurada",       "Seguir los pasos de configuracion de JAVA_HOME"],
      ],
      [2600, 2500, 4260]
    ),
    espacio(),
    saltoPlantilla()
  ];
}

// ─── DOCUMENTO COMPLETO ───────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: "2C2C2C" } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: BLANCO },
        paragraph: {
          spacing: { before: 360, after: 200 },
          outlineLevel: 0,
          shading: { fill: AZUL, type: ShadingType.CLEAR },
          indent: { left: 200, right: 200 }
        }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: AZUL },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: AZUL_CLARO },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: AZUL_CLARO } },
          children: [
            new TextRun({ text: "InDriveTura — Guia de Aprendizaje", bold: true, color: AZUL, font: "Arial", size: 18 }),
            new TextRun({ text: "   |   MalcomJesid   |   Seminario II", color: "888888", font: "Arial", size: 18 }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: AZUL_CLARO } },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Pagina ", color: "888888", font: "Arial", size: 18 }),
            new TextRun({ children: [PageNumber.CURRENT], color: "888888", font: "Arial", size: 18 }),
          ]
        })]
      })
    },
    children: [
      ...portada(),
      ...seccion1(),
      ...seccion2(),
      ...seccion3(),
      ...seccion4(),
      ...seccion5(),
      ...seccion6(),
      ...seccion7(),
      ...seccion8(),
      ...seccion9(),
      ...seccion10(),
    ]
  }]
});

const outputPath = "C:/Users/malco/Desktop/Proyectos programación/InDriveTura/docs/InDriveTura-Guia-Aprendizaje.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Documento generado: " + outputPath);
}).catch(err => {
  console.error("Error:", err.message);
});
