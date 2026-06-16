// =============================================
//  AUDITORIA SHAREPOINT - PULPAFRUIT SAS
//  app.js
//  Conecta a Microsoft Graph API y renderiza
//  el dashboard de auditoria por area
// =============================================

const CONFIG = {
  clientId:        '5c24b02d-ce93-4865-9b3d-409786cb175b',
  tenantId:        'f16b5c8e-ad45-4fc4-a55b-0af26c456817',
  scopes:          ['Sites.Read.All', 'User.Read'],
  diasInactividad: 90,
  patronFormato:   /^[A-Z]{2,4}-[A-Z]{2,4}-\d{3,4}$/
};

// Colores por area — paleta verde PULPAFRUIT + colores complementarios
const AREA_COLORS = [
  { bg: '#E8F5E7', color: '#2D7A28', dot: '#3DAA35' },
  { bg: '#E6F1FB', color: '#185FA5', dot: '#378ADD' },
  { bg: '#EEEDFE', color: '#534AB7', dot: '#7F77DD' },
  { bg: '#FAEEDA', color: '#854F0B', dot: '#BA7517' },
  { bg: '#E1F5EE', color: '#0F6E56', dot: '#1D9E75' },
  { bg: '#FAECE7', color: '#993C1D', dot: '#D85A30' },
  { bg: '#FCEBEB', color: '#A32D2D', dot: '#E24B4A' },
  { bg: '#F1EFE8', color: '#5F5E5A', dot: '#888780' },
  { bg: '#FBEAF0', color: '#72243E', dot: '#D4537E' },
  { bg: '#EAF3DE', color: '#3B6D11', dot: '#639922' },
  { bg: '#E6F7FA', color: '#0E6E7A', dot: '#1DA3B5' },
  { bg: '#FDF4E7', color: '#7A4F0B', dot: '#C47B1A' },
];

const AVATAR_COLORS = [
  { bg: '#E8F5E7', color: '#2D7A28' },
  { bg: '#E6F1FB', color: '#185FA5' },
  { bg: '#EEEDFE', color: '#534AB7' },
  { bg: '#FAEEDA', color: '#854F0B' },
  { bg: '#FAECE7', color: '#993C1D' },
  { bg: '#E1F5EE', color: '#0F6E56' },
  { bg: '#F1EFE8', color: '#5F5E5A' },
];

const AREA_META = {
  'comercio exterior':     { icon: '🌎', nombre: 'Comercio Exterior' },
  'laboratorio':           { icon: '🔬', nombre: 'Laboratorio' },
  'materias primas':       { icon: '🌿', nombre: 'Materias Primas' },
  'i&d':                   { icon: '💡', nombre: 'I&D' },
  'innovacion y desarrollo': { icon: '💡', nombre: 'I&D' },
  'proyectos':             { icon: '📋', nombre: 'Proyectos' },
  'ambiental':             { icon: '♻️', nombre: 'Ambiental' },
  'metrologia':            { icon: '📏', nombre: 'Metrologia' },
  'metrología':            { icon: '📏', nombre: 'Metrologia' },
  'area sst':              { icon: '⛑️', nombre: 'SST' },
  'sst':                   { icon: '⛑️', nombre: 'SST' },
  'area contable':         { icon: '💰', nombre: 'Contable' },
  'contable':              { icon: '💰', nombre: 'Contable' },
  'recursos humanos':      { icon: '👥', nombre: 'Recursos Humanos' },
  'rrhh':                  { icon: '👥', nombre: 'Recursos Humanos' },
  'form logistica':        { icon: '🚚', nombre: 'Logistica' },
  'logistica':             { icon: '🚚', nombre: 'Logistica' },
  'logística':             { icon: '🚚', nombre: 'Logistica' },
  'limpieza y desinfeccion': { icon: '🧹', nombre: 'Limpieza y Desinfeccion' },
  'limpieza y desinfección': { icon: '🧹', nombre: 'Limpieza y Desinfeccion' },
  'dev-formatos':          { icon: '🧹', nombre: 'Limpieza y Desinfeccion' },
};

let DATOS = { areas: [], usuarios: [], token: null };

// -----------------------------------------------
// GRAPH API
// -----------------------------------------------
async function graphGet(url, token) {
  const headers = { Authorization: `Bearer ${token}` };
  let results = [], nextUrl = url;
  while (nextUrl) {
    const resp = await fetch(nextUrl, { headers });
    const data = await resp.json();
    results = results.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }
  return results;
}

function esPrincipal(nombre) {
  return CONFIG.patronFormato.test(nombre.trim());
}

function calcularFrecuencia(items) {
  if (items.length < 2) return items.length === 1 ? 'Un registro' : 'Sin datos';
  const fechas = items.map(i => new Date(i.createdDateTime)).sort((a,b) => a-b);
  let totalH = 0;
  for (let i = 1; i < fechas.length; i++) totalH += (fechas[i] - fechas[i-1]) / 3600000;
  const prom = totalH / (fechas.length - 1);
  if (prom <= 24)  return 'Diaria';
  if (prom <= 168) return 'Semanal';
  if (prom <= 720) return 'Mensual';
  return 'Esporadica';
}

function getNombreArea(displayName) {
  const key = displayName.toLowerCase().trim();
  return AREA_META[key]?.nombre || displayName;
}

function getIconoArea(displayName) {
  const key = displayName.toLowerCase().trim();
  return AREA_META[key]?.icon || '📁';
}

// -----------------------------------------------
// CARGA DE DATOS REALES
// -----------------------------------------------
async function cargarDatos() {
  mostrarLoading(true);
  try {
    const token = await obtenerToken();
    if (!token) { cargarDatosEjemplo(); return; }
    DATOS.token = token;

    const sitios = await graphGet(
      'https://graph.microsoft.com/v1.0/sites?search=*&$select=id,displayName,webUrl,createdDateTime',
      token
    );

    const listasPorSitio = {};
    const itemsPorLista  = {};

    for (const sitio of sitios) {
      const listas = await graphGet(
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists?$select=id,displayName,list,lastModifiedDateTime,createdDateTime`,
        token
      );
      listasPorSitio[sitio.id] = listas;
      for (const lista of listas) {
        if (esPrincipal(lista.displayName.trim())) {
          try {
            itemsPorLista[`${sitio.id}__${lista.id}`] = await graphGet(
              `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items?$select=id,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy&$top=999`,
              token
            );
          } catch(e) { itemsPorLista[`${sitio.id}__${lista.id}`] = []; }
        }
      }
    }

    const resultado = procesarDatos(sitios, listasPorSitio, itemsPorLista);
    DATOS.areas    = resultado.areas;
    DATOS.usuarios = resultado.usuarios;
    renderizarTodo();
  } catch(e) {
    console.error('Error:', e);
    cargarDatosEjemplo();
  } finally {
    mostrarLoading(false);
  }
}

async function obtenerToken() {
  return null;
}

function procesarDatos(sitios, listasPorSitio, itemsPorLista) {
  const areas = [];
  const usuariosMap = {};

  sitios.forEach((sitio, idx) => {
    const listas  = listasPorSitio[sitio.id] || [];
    const color   = AREA_COLORS[idx % AREA_COLORS.length];
    const formatos = { activos: [], inactivos: [], nunca: [] };

    listas.forEach(lista => {
      const nombre = lista.displayName.trim();
      if (!esPrincipal(nombre)) return;

      const items      = itemsPorLista[`${sitio.id}__${lista.id}`] || [];
      const ultimaMod  = new Date(lista.lastModifiedDateTime);
      const diasSinUso = Math.floor((Date.now() - ultimaMod) / 86400000);
      const estado     = diasSinUso <= CONFIG.diasInactividad ? 'ACTIVA' : 'INACTIVA';
      const frecuencia = calcularFrecuencia(items);
      const sorted     = [...items].sort((a,b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      const ultimoItem = sorted[0];

      const formato = {
        codigo: nombre, estado, totalItems: items.length, frecuencia, diasSinUso,
        ultimaMod:     ultimaMod.toLocaleDateString('es-CO'),
        ultimoUsuario: ultimoItem?.lastModifiedBy?.user?.displayName || '—',
        emailUltimo:   ultimoItem?.lastModifiedBy?.user?.email || '—',
        fechaCreacion: new Date(lista.createdDateTime).toLocaleDateString('es-CO'),
      };

      if (items.length === 0)       formatos.nunca.push(formato);
      else if (estado === 'ACTIVA') formatos.activos.push(formato);
      else                          formatos.inactivos.push(formato);

      items.forEach(item => {
        const u = item.createdBy?.user;
        if (!u?.displayName) return;
        if (!usuariosMap[u.displayName]) {
          usuariosMap[u.displayName] = {
            nombre: u.displayName, email: u.email || '—',
            registros: 0, areas: new Set(), formatos: new Set(),
            ultimaActividad: new Date(item.createdDateTime)
          };
        }
        usuariosMap[u.displayName].registros++;
        usuariosMap[u.displayName].areas.add(getNombreArea(sitio.displayName));
        usuariosMap[u.displayName].formatos.add(nombre);
        const f = new Date(item.createdDateTime);
        if (f > usuariosMap[u.displayName].ultimaActividad)
          usuariosMap[u.displayName].ultimaActividad = f;
      });
    });

    const totalFormatos  = formatos.activos.length + formatos.inactivos.length + formatos.nunca.length;
    const totalRegistros = [...formatos.activos, ...formatos.inactivos, ...formatos.nunca].reduce((s,f) => s + f.totalItems, 0);
    const pctUso         = totalFormatos ? Math.round((formatos.activos.length / totalFormatos) * 100) : 0;

    areas.push({
      nombre:        getNombreArea(sitio.displayName),
      nombreOriginal: sitio.displayName,
      url:           sitio.webUrl,
      icon:          getIconoArea(sitio.displayName),
      color,
      formatos,
      totalFormatos,
      pctUso,
      totalRegistros,
      promRegistros: totalFormatos ? (totalRegistros / totalFormatos).toFixed(1) : '0',
    });
  });

  const usuarios = Object.values(usuariosMap)
    .sort((a,b) => b.registros - a.registros)
    .map(u => ({
      ...u,
      areas:          [...u.areas].join(', '),
      formatos:       u.formatos.size,
      ultimaActividad: u.ultimaActividad.toLocaleDateString('es-CO'),
    }));

  return { areas, usuarios };
}

// -----------------------------------------------
// DATOS DE EJEMPLO con areas reales PULPAFRUIT
// -----------------------------------------------
function cargarDatosEjemplo() {
  const areasEjemplo = [
    {
      nombre: 'Comercio Exterior', icon: '🌎', url: 'pulpafruit.sharepoint.com/sites/ComercioExterior',
      color: AREA_COLORS[0],
      formatos: {
        activos: [
          { codigo:'CE-FO-001', estado:'ACTIVA', totalItems:87,  frecuencia:'Semanal',   diasSinUso:5,   ultimaMod:'23/05/2026', ultimoUsuario:'Diana Ospina',    emailUltimo:'d.ospina@pulpafruit.com',   fechaCreacion:'10/01/2024' },
          { codigo:'CE-FO-002', estado:'ACTIVA', totalItems:54,  frecuencia:'Semanal',   diasSinUso:8,   ultimaMod:'20/05/2026', ultimoUsuario:'Diana Ospina',    emailUltimo:'d.ospina@pulpafruit.com',   fechaCreacion:'10/01/2024' },
          { codigo:'CE-FO-003', estado:'ACTIVA', totalItems:31,  frecuencia:'Mensual',   diasSinUso:22,  ultimaMod:'06/05/2026', ultimoUsuario:'Diana Ospina',    emailUltimo:'d.ospina@pulpafruit.com',   fechaCreacion:'15/02/2024' },
        ],
        inactivos: [
          { codigo:'CE-FO-004', estado:'INACTIVA', totalItems:9, frecuencia:'Esporadica', diasSinUso:142, ultimaMod:'07/01/2026', ultimoUsuario:'Diana Ospina', emailUltimo:'d.ospina@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'CE-FO-005', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/03/2024' },
        ]
      },
      totalFormatos:5, pctUso:60, totalRegistros:181, promRegistros:'36.2'
    },
    {
      nombre: 'Laboratorio', icon: '🔬', url: 'pulpafruit.sharepoint.com/sites/Laboratorio',
      color: AREA_COLORS[1],
      formatos: {
        activos: [
          { codigo:'LA-FO-001', estado:'ACTIVA', totalItems:215, frecuencia:'Diaria',    diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Marcela Rios',    emailUltimo:'m.rios@pulpafruit.com',    fechaCreacion:'05/01/2024' },
          { codigo:'LA-FO-002', estado:'ACTIVA', totalItems:143, frecuencia:'Diaria',    diasSinUso:1,  ultimaMod:'27/05/2026', ultimoUsuario:'Marcela Rios',    emailUltimo:'m.rios@pulpafruit.com',    fechaCreacion:'05/01/2024' },
          { codigo:'LA-AN-001', estado:'ACTIVA', totalItems:98,  frecuencia:'Diaria',    diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Juan Cardona',    emailUltimo:'j.cardona@pulpafruit.com', fechaCreacion:'10/01/2024' },
          { codigo:'LA-FO-003', estado:'ACTIVA', totalItems:67,  frecuencia:'Semanal',   diasSinUso:4,  ultimaMod:'24/05/2026', ultimoUsuario:'Marcela Rios',    emailUltimo:'m.rios@pulpafruit.com',    fechaCreacion:'15/01/2024' },
          { codigo:'LA-FO-004', estado:'ACTIVA', totalItems:45,  frecuencia:'Semanal',   diasSinUso:6,  ultimaMod:'22/05/2026', ultimoUsuario:'Juan Cardona',    emailUltimo:'j.cardona@pulpafruit.com', fechaCreacion:'20/01/2024' },
        ],
        inactivos: [
          { codigo:'LA-FO-005', estado:'INACTIVA', totalItems:18, frecuencia:'Esporadica', diasSinUso:112, ultimaMod:'07/02/2026', ultimoUsuario:'Juan Cardona', emailUltimo:'j.cardona@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: []
      },
      totalFormatos:6, pctUso:83, totalRegistros:586, promRegistros:'97.7'
    },
    {
      nombre: 'Materias Primas', icon: '🌿', url: 'pulpafruit.sharepoint.com/sites/MateriasPrimas',
      color: AREA_COLORS[2],
      formatos: {
        activos: [
          { codigo:'MP-FO-001', estado:'ACTIVA', totalItems:178, frecuencia:'Diaria',  diasSinUso:0, ultimaMod:'28/05/2026', ultimoUsuario:'Carlos Velez',  emailUltimo:'c.velez@pulpafruit.com',  fechaCreacion:'08/01/2024' },
          { codigo:'MP-FO-002', estado:'ACTIVA', totalItems:132, frecuencia:'Diaria',  diasSinUso:1, ultimaMod:'27/05/2026', ultimoUsuario:'Carlos Velez',  emailUltimo:'c.velez@pulpafruit.com',  fechaCreacion:'08/01/2024' },
          { codigo:'MP-CO-001', estado:'ACTIVA', totalItems:76,  frecuencia:'Semanal', diasSinUso:3, ultimaMod:'25/05/2026', ultimoUsuario:'Sandra Gil',    emailUltimo:'s.gil@pulpafruit.com',    fechaCreacion:'15/01/2024' },
        ],
        inactivos: [
          { codigo:'MP-FO-003', estado:'INACTIVA', totalItems:22, frecuencia:'Esporadica', diasSinUso:198, ultimaMod:'12/11/2025', ultimoUsuario:'Sandra Gil', emailUltimo:'s.gil@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'MP-FO-004', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/04/2024' },
        ]
      },
      totalFormatos:5, pctUso:60, totalRegistros:408, promRegistros:'81.6'
    },
    {
      nombre: 'I&D', icon: '💡', url: 'pulpafruit.sharepoint.com/sites/ID',
      color: AREA_COLORS[3],
      formatos: {
        activos: [
          { codigo:'ID-FO-001', estado:'ACTIVA', totalItems:63,  frecuencia:'Semanal', diasSinUso:7,  ultimaMod:'21/05/2026', ultimoUsuario:'Alejandro Muñoz', emailUltimo:'a.munoz@pulpafruit.com', fechaCreacion:'10/02/2024' },
          { codigo:'ID-IN-001', estado:'ACTIVA', totalItems:41,  frecuencia:'Mensual', diasSinUso:15, ultimaMod:'13/05/2026', ultimoUsuario:'Alejandro Muñoz', emailUltimo:'a.munoz@pulpafruit.com', fechaCreacion:'10/02/2024' },
        ],
        inactivos: [
          { codigo:'ID-FO-002', estado:'INACTIVA', totalItems:14, frecuencia:'Esporadica', diasSinUso:165, ultimaMod:'14/12/2025', ultimoUsuario:'Alejandro Muñoz', emailUltimo:'a.munoz@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'ID-FO-003', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/05/2024' },
        ]
      },
      totalFormatos:4, pctUso:50, totalRegistros:118, promRegistros:'29.5'
    },
    {
      nombre: 'Proyectos', icon: '📋', url: 'pulpafruit.sharepoint.com/sites/Proyectos',
      color: AREA_COLORS[4],
      formatos: {
        activos: [
          { codigo:'PR-FO-001', estado:'ACTIVA', totalItems:55,  frecuencia:'Semanal', diasSinUso:6,  ultimaMod:'22/05/2026', ultimoUsuario:'Felipe Castillo', emailUltimo:'f.castillo@pulpafruit.com', fechaCreacion:'01/03/2024' },
          { codigo:'PR-FO-002', estado:'ACTIVA', totalItems:33,  frecuencia:'Semanal', diasSinUso:10, ultimaMod:'18/05/2026', ultimoUsuario:'Felipe Castillo', emailUltimo:'f.castillo@pulpafruit.com', fechaCreacion:'01/03/2024' },
        ],
        inactivos: [
          { codigo:'PR-FO-003', estado:'INACTIVA', totalItems:8,  frecuencia:'Esporadica', diasSinUso:210, ultimaMod:'31/10/2025', ultimoUsuario:'Felipe Castillo', emailUltimo:'f.castillo@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'PR-FO-004', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', fechaCreacion:'01/06/2024' },
          { codigo:'PR-FO-005', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', fechaCreacion:'01/06/2024' },
        ]
      },
      totalFormatos:5, pctUso:40, totalRegistros:96, promRegistros:'19.2'
    },
    {
      nombre: 'Ambiental', icon: '♻️', url: 'pulpafruit.sharepoint.com/sites/Ambiental',
      color: AREA_COLORS[9],
      formatos: {
        activos: [
          { codigo:'AM-FO-001', estado:'ACTIVA', totalItems:74,  frecuencia:'Semanal', diasSinUso:4,  ultimaMod:'24/05/2026', ultimoUsuario:'Natalia Peña',   emailUltimo:'n.pena@pulpafruit.com',   fechaCreacion:'15/01/2024' },
          { codigo:'AM-FO-002', estado:'ACTIVA', totalItems:48,  frecuencia:'Mensual', diasSinUso:18, ultimaMod:'10/05/2026', ultimoUsuario:'Natalia Peña',   emailUltimo:'n.pena@pulpafruit.com',   fechaCreacion:'15/01/2024' },
          { codigo:'AM-RE-001', estado:'ACTIVA', totalItems:29,  frecuencia:'Mensual', diasSinUso:25, ultimaMod:'03/05/2026', ultimoUsuario:'Natalia Peña',   emailUltimo:'n.pena@pulpafruit.com',   fechaCreacion:'01/02/2024' },
        ],
        inactivos: [],
        nunca: [
          { codigo:'AM-FO-003', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/07/2024' },
        ]
      },
      totalFormatos:4, pctUso:75, totalRegistros:151, promRegistros:'37.8'
    },
    {
      nombre: 'Metrologia', icon: '📏', url: 'pulpafruit.sharepoint.com/sites/Metrologia',
      color: AREA_COLORS[7],
      formatos: {
        activos: [
          { codigo:'ME-FO-001', estado:'ACTIVA', totalItems:112, frecuencia:'Diaria',  diasSinUso:1,  ultimaMod:'27/05/2026', ultimoUsuario:'Roberto Acosta', emailUltimo:'r.acosta@pulpafruit.com', fechaCreacion:'08/01/2024' },
          { codigo:'ME-FO-002', estado:'ACTIVA', totalItems:89,  frecuencia:'Diaria',  diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Roberto Acosta', emailUltimo:'r.acosta@pulpafruit.com', fechaCreacion:'08/01/2024' },
          { codigo:'ME-CA-001', estado:'ACTIVA', totalItems:56,  frecuencia:'Semanal', diasSinUso:5,  ultimaMod:'23/05/2026', ultimoUsuario:'Roberto Acosta', emailUltimo:'r.acosta@pulpafruit.com', fechaCreacion:'15/01/2024' },
        ],
        inactivos: [
          { codigo:'ME-FO-003', estado:'INACTIVA', totalItems:11, frecuencia:'Esporadica', diasSinUso:189, ultimaMod:'21/11/2025', ultimoUsuario:'Roberto Acosta', emailUltimo:'r.acosta@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: []
      },
      totalFormatos:4, pctUso:75, totalRegistros:268, promRegistros:'67.0'
    },
    {
      nombre: 'SST', icon: '⛑️', url: 'pulpafruit.sharepoint.com/sites/SST',
      color: AREA_COLORS[5],
      formatos: {
        activos: [
          { codigo:'SS-FO-001', estado:'ACTIVA', totalItems:93,  frecuencia:'Semanal', diasSinUso:3,  ultimaMod:'25/05/2026', ultimoUsuario:'Paola Herrera',  emailUltimo:'p.herrera@pulpafruit.com', fechaCreacion:'10/01/2024' },
          { codigo:'SS-FO-002', estado:'ACTIVA', totalItems:67,  frecuencia:'Semanal', diasSinUso:7,  ultimaMod:'21/05/2026', ultimoUsuario:'Paola Herrera',  emailUltimo:'p.herrera@pulpafruit.com', fechaCreacion:'10/01/2024' },
          { codigo:'SS-IN-001', estado:'ACTIVA', totalItems:44,  frecuencia:'Mensual', diasSinUso:12, ultimaMod:'16/05/2026', ultimoUsuario:'Paola Herrera',  emailUltimo:'p.herrera@pulpafruit.com', fechaCreacion:'20/01/2024' },
        ],
        inactivos: [
          { codigo:'SS-FO-003', estado:'INACTIVA', totalItems:7,  frecuencia:'Esporadica', diasSinUso:230, ultimaMod:'10/10/2025', ultimoUsuario:'Paola Herrera', emailUltimo:'p.herrera@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'SS-FO-004', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/04/2024' },
        ]
      },
      totalFormatos:5, pctUso:60, totalRegistros:211, promRegistros:'42.2'
    },
    {
      nombre: 'Contable', icon: '💰', url: 'pulpafruit.sharepoint.com/sites/Contable',
      color: AREA_COLORS[3],
      formatos: {
        activos: [
          { codigo:'CO-FO-001', estado:'ACTIVA', totalItems:58,  frecuencia:'Semanal', diasSinUso:4,  ultimaMod:'24/05/2026', ultimoUsuario:'Gloria Ramirez', emailUltimo:'g.ramirez@pulpafruit.com', fechaCreacion:'08/01/2024' },
          { codigo:'CO-FO-002', estado:'ACTIVA', totalItems:41,  frecuencia:'Mensual', diasSinUso:20, ultimaMod:'08/05/2026', ultimoUsuario:'Gloria Ramirez', emailUltimo:'g.ramirez@pulpafruit.com', fechaCreacion:'08/01/2024' },
        ],
        inactivos: [
          { codigo:'CO-FO-003', estado:'INACTIVA', totalItems:15, frecuencia:'Esporadica', diasSinUso:175, ultimaMod:'04/12/2025', ultimoUsuario:'Gloria Ramirez', emailUltimo:'g.ramirez@pulpafruit.com', fechaCreacion:'01/01/2024' },
          { codigo:'CO-FO-004', estado:'INACTIVA', totalItems:6,  frecuencia:'Esporadica', diasSinUso:310, ultimaMod:'23/07/2025', ultimoUsuario:'Gloria Ramirez', emailUltimo:'g.ramirez@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: []
      },
      totalFormatos:4, pctUso:50, totalRegistros:120, promRegistros:'30.0'
    },
    {
      nombre: 'Recursos Humanos', icon: '👥', url: 'pulpafruit.sharepoint.com/sites/RecursosHumanos',
      color: AREA_COLORS[8],
      formatos: {
        activos: [
          { codigo:'RH-FO-001', estado:'ACTIVA', totalItems:102, frecuencia:'Diaria',  diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Ana Perez',      emailUltimo:'a.perez@pulpafruit.com',   fechaCreacion:'05/01/2024' },
          { codigo:'RH-FO-002', estado:'ACTIVA', totalItems:74,  frecuencia:'Semanal', diasSinUso:3,  ultimaMod:'25/05/2026', ultimoUsuario:'Ana Perez',      emailUltimo:'a.perez@pulpafruit.com',   fechaCreacion:'05/01/2024' },
          { codigo:'RH-IN-001', estado:'ACTIVA', totalItems:51,  frecuencia:'Mensual', diasSinUso:18, ultimaMod:'10/05/2026', ultimoUsuario:'Maria Torres',   emailUltimo:'m.torres@pulpafruit.com',  fechaCreacion:'15/01/2024' },
        ],
        inactivos: [
          { codigo:'RH-FO-003', estado:'INACTIVA', totalItems:13, frecuencia:'Esporadica', diasSinUso:238, ultimaMod:'02/10/2025', ultimoUsuario:'Maria Torres', emailUltimo:'m.torres@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'RH-FO-004', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/05/2024' },
        ]
      },
      totalFormatos:5, pctUso:60, totalRegistros:240, promRegistros:'48.0'
    },
    {
      nombre: 'Logistica', icon: '🚚', url: 'pulpafruit.sharepoint.com/sites/Logistica',
      color: AREA_COLORS[0],
      formatos: {
        activos: [
          { codigo:'LO-FO-004', estado:'ACTIVA', totalItems:187, frecuencia:'Diaria',  diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Carlos Ramirez', emailUltimo:'c.ramirez@pulpafruit.com', fechaCreacion:'05/01/2024' },
          { codigo:'LO-FO-060', estado:'ACTIVA', totalItems:143, frecuencia:'Diaria',  diasSinUso:1,  ultimaMod:'27/05/2026', ultimoUsuario:'Carlos Ramirez', emailUltimo:'c.ramirez@pulpafruit.com', fechaCreacion:'05/01/2024' },
          { codigo:'LO-FO-057', estado:'ACTIVA', totalItems:98,  frecuencia:'Diaria',  diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Pedro Vega',     emailUltimo:'p.vega@pulpafruit.com',   fechaCreacion:'10/01/2024' },
          { codigo:'LO-FO-041', estado:'ACTIVA', totalItems:55,  frecuencia:'Semanal', diasSinUso:7,  ultimaMod:'21/05/2026', ultimoUsuario:'Pedro Vega',     emailUltimo:'p.vega@pulpafruit.com',   fechaCreacion:'15/01/2024' },
        ],
        inactivos: [
          { codigo:'LO-FO-062', estado:'INACTIVA', totalItems:14, frecuencia:'Esporadica', diasSinUso:256, ultimaMod:'15/09/2025', ultimoUsuario:'Carlos Ramirez', emailUltimo:'c.ramirez@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: [
          { codigo:'LO-FO-008', estado:'NUNCA', totalItems:0, frecuencia:'Sin datos', diasSinUso:null, ultimaMod:'—', ultimoUsuario:'—', emailUltimo:'—', fechaCreacion:'01/03/2024' },
        ]
      },
      totalFormatos:6, pctUso:67, totalRegistros:497, promRegistros:'82.8'
    },
    {
      nombre: 'Limpieza y Desinfeccion', icon: '🧹', url: 'pulpafruit.sharepoint.com/sites/LimpiezaDesinfeccion',
      color: AREA_COLORS[10],
      formatos: {
        activos: [
          { codigo:'LD-FO-001', estado:'ACTIVA', totalItems:145, frecuencia:'Diaria',  diasSinUso:0,  ultimaMod:'28/05/2026', ultimoUsuario:'Jorge Morales',  emailUltimo:'j.morales@pulpafruit.com', fechaCreacion:'08/01/2024' },
          { codigo:'LD-FO-002', estado:'ACTIVA', totalItems:112, frecuencia:'Diaria',  diasSinUso:1,  ultimaMod:'27/05/2026', ultimoUsuario:'Jorge Morales',  emailUltimo:'j.morales@pulpafruit.com', fechaCreacion:'08/01/2024' },
          { codigo:'LD-PR-001', estado:'ACTIVA', totalItems:78,  frecuencia:'Semanal', diasSinUso:4,  ultimaMod:'24/05/2026', ultimoUsuario:'Jorge Morales',  emailUltimo:'j.morales@pulpafruit.com', fechaCreacion:'15/01/2024' },
          { codigo:'LD-FO-003', estado:'ACTIVA', totalItems:61,  frecuencia:'Semanal', diasSinUso:6,  ultimaMod:'22/05/2026', ultimoUsuario:'Laura Soto',     emailUltimo:'l.soto@pulpafruit.com',   fechaCreacion:'20/01/2024' },
        ],
        inactivos: [
          { codigo:'LD-FO-004', estado:'INACTIVA', totalItems:10, frecuencia:'Esporadica', diasSinUso:147, ultimaMod:'02/01/2026', ultimoUsuario:'Laura Soto', emailUltimo:'l.soto@pulpafruit.com', fechaCreacion:'01/01/2024' },
        ],
        nunca: []
      },
      totalFormatos:5, pctUso:80, totalRegistros:406, promRegistros:'81.2'
    },
  ];

  DATOS.areas = areasEjemplo;

  DATOS.usuarios = [
    { nombre:'Carlos Ramirez',  email:'c.ramirez@pulpafruit.com',  registros:344, areas:'Logistica',           formatos:4, ultimaActividad:'28/05/2026' },
    { nombre:'Jorge Morales',   email:'j.morales@pulpafruit.com',  registros:316, areas:'Limpieza y Desinfeccion', formatos:4, ultimaActividad:'28/05/2026' },
    { nombre:'Marcela Rios',    email:'m.rios@pulpafruit.com',     registros:283, areas:'Laboratorio',          formatos:4, ultimaActividad:'28/05/2026' },
    { nombre:'Carlos Velez',    email:'c.velez@pulpafruit.com',    registros:254, areas:'Materias Primas',       formatos:3, ultimaActividad:'28/05/2026' },
    { nombre:'Roberto Acosta',  email:'r.acosta@pulpafruit.com',   registros:201, areas:'Metrologia',           formatos:3, ultimaActividad:'28/05/2026' },
    { nombre:'Ana Perez',       email:'a.perez@pulpafruit.com',    registros:176, areas:'Recursos Humanos',      formatos:3, ultimaActividad:'28/05/2026' },
    { nombre:'Paola Herrera',   email:'p.herrera@pulpafruit.com',  registros:164, areas:'SST',                  formatos:3, ultimaActividad:'25/05/2026' },
    { nombre:'Diana Ospina',    email:'d.ospina@pulpafruit.com',   registros:131, areas:'Comercio Exterior',     formatos:3, ultimaActividad:'23/05/2026' },
    { nombre:'Natalia Peña',    email:'n.pena@pulpafruit.com',     registros:97,  areas:'Ambiental',            formatos:3, ultimaActividad:'24/05/2026' },
    { nombre:'Gloria Ramirez',  email:'g.ramirez@pulpafruit.com',  registros:89,  areas:'Contable',             formatos:2, ultimaActividad:'24/05/2026' },
    { nombre:'Juan Cardona',    email:'j.cardona@pulpafruit.com',  registros:87,  areas:'Laboratorio',          formatos:2, ultimaActividad:'24/05/2026' },
    { nombre:'Alejandro Muñoz', email:'a.munoz@pulpafruit.com',    registros:78,  areas:'I&D',                  formatos:2, ultimaActividad:'21/05/2026' },
  ];

  renderizarTodo();
  mostrarLoading(false);
}

// -----------------------------------------------
// RENDERIZADO
// -----------------------------------------------
function renderizarTodo() {
  actualizarMetricas();
  construirNavTabs();
  renderizarResumen();
  actualizarFecha();
}

function actualizarMetricas() {
  const totalAreas     = DATOS.areas.length;
  const totalFormatos  = DATOS.areas.reduce((s,a) => s + a.totalFormatos, 0);
  const totalActivos   = DATOS.areas.reduce((s,a) => s + a.formatos.activos.length, 0);
  const totalInactivos = DATOS.areas.reduce((s,a) => s + a.formatos.inactivos.length, 0);
  const totalNunca     = DATOS.areas.reduce((s,a) => s + a.formatos.nunca.length, 0);
  setText('m-areas',     totalAreas);
  setText('m-total',     totalFormatos);
  setText('m-activos',   totalActivos);
  setText('m-inactivos', totalInactivos);
  setText('m-nunca',     totalNunca);
  setText('footer-total', `${totalFormatos} formatos · ${totalActivos} activos · ${totalInactivos + totalNunca} sin actividad`);
}

function construirNavTabs() {
  const nav = document.getElementById('nav-tabs');
  if (!nav) return;
  nav.innerHTML = `
    <button class="nav-tab active" data-area="resumen" onclick="mostrarArea('resumen',this)">
      <span class="tab-dot" style="background:#3DAA35"></span>Resumen general
    </button>`;
  DATOS.areas.forEach(area => {
    const btn = document.createElement('button');
    btn.className = 'nav-tab';
    btn.dataset.area = area.nombre;
    btn.onclick = function() { mostrarArea(area.nombre, this); };
    btn.innerHTML = `<span class="tab-dot" style="background:${area.color.dot}"></span>${area.nombre}`;
    nav.appendChild(btn);
  });
}

function renderizarResumen() {
  const grid = document.getElementById('resumen-cards');
  grid.innerHTML = '';

  const table = document.createElement('div');
  table.className = 'resumen-table-wrap';
  table.innerHTML = `
    <table>
      <thead>
        <tr>
          <th style="width:180px">Area</th>
          <th style="width:60px">Form.</th>
          <th style="width:60px">Activos</th>
          <th style="width:70px">Inactivos</th>
          <th style="width:60px">Sin uso</th>
          <th style="width:70px">Registros</th>
          <th style="width:160px">% Uso</th>
          <th style="width:100px">Frecuencia</th>
          <th style="width:100px">Prom. registros</th>
          <th style="width:30px"></th>
        </tr>
      </thead>
      <tbody id="resumen-tbody"></tbody>
    </table>`;
  grid.appendChild(table);

  const tbody = document.getElementById('resumen-tbody');
  DATOS.areas.forEach(area => {
    const barColor = area.pctUso >= 70 ? '#3DAA35' : area.pctUso >= 40 ? '#EF9F27' : '#E24B4A';
    const pctColor = area.pctUso >= 70 ? 'var(--verde-dark)' : area.pctUso >= 40 ? 'var(--amber)' : 'var(--red)';
    const freq     = area.formatos.activos[0]?.frecuencia || 'Sin datos';
    const tr       = document.createElement('tr');
    tr.onclick     = () => mostrarArea(area.nombre, null);
    tr.innerHTML   = `
      <td>
        <div class="area-cell">
          <div class="area-cell-icon" style="background:${area.color.bg};color:${area.color.color}">${area.icon}</div>
          ${area.nombre}
        </div>
      </td>
      <td>${area.totalFormatos}</td>
      <td class="num-green">${area.formatos.activos.length}</td>
      <td class="num-amber">${area.formatos.inactivos.length}</td>
      <td class="num-red">${area.formatos.nunca.length}</td>
      <td style="font-weight:700">${area.totalRegistros}</td>
      <td>
        <div class="uso-bar-table">
          <div class="uso-bar-table-bg"><div class="uso-bar-table-fill" style="width:${area.pctUso}%;background:${barColor}"></div></div>
          <span class="uso-pct-table" style="color:${pctColor}">${area.pctUso}%</span>
        </div>
      </td>
      <td><span class="${freqClass(freq)}">${freq}</span></td>
      <td style="font-weight:700">${area.promRegistros}</td>
      <td class="arrow-cell">›</td>`;
    tbody.appendChild(tr);
  });

  const ul = document.getElementById('usuarios-globales');
  ul.innerHTML = '';
  DATOS.usuarios.slice(0,10).forEach((u,i) => {
    const av  = AVATAR_COLORS[i % AVATAR_COLORS.length];
    const ini = u.nombre.split(' ').map(n=>n[0]).join('').slice(0,2).toUpperCase();
    ul.innerHTML += `
      <div class="usuario-row">
        <div class="avatar" style="background:${av.bg};color:${av.color}">${ini}</div>
        <div class="user-info">
          <div class="user-name">${u.nombre}</div>
          <div class="user-detail">${u.email} · ${u.areas} · ${u.formatos} formatos · ${u.ultimaActividad}</div>
        </div>
        <div class="user-count">${u.registros} reg</div>
      </div>`;
  });
  setText('badge-fecha', `Generado: ${new Date().toLocaleDateString('es-CO')}`);
}

function renderizarArea(area) {
  document.getElementById('area-header-block').innerHTML = `
    <button class="btn-back" onclick="mostrarArea('resumen',null)" style="display:flex;align-items:center;gap:6px;padding:6px 12px;margin-bottom:12px;border-radius:20px;border:1.5px solid var(--border);background:var(--white);color:var(--muted);font-size:12px;font-weight:700;cursor:pointer;font-family:'Nunito Sans',sans-serif">
      ‹ Volver al resumen
    </button>` + `
    <div class="ah-left">
      <div class="ah-icon" style="background:${area.color.bg};color:${area.color.color}">${area.icon}</div>
      <div>
        <div class="ah-name">${area.nombre}</div>
        <div class="ah-url">${area.url}</div>
      </div>
    </div>
    <div class="ah-pills">
      <span class="pill pill-green">${area.formatos.activos.length} activos</span>
      <span class="pill pill-amber">${area.formatos.inactivos.length} inactivos</span>
      <span class="pill pill-red">${area.formatos.nunca.length} sin uso</span>
      <span class="pill pill-gray">${area.pctUso}% uso</span>
    </div>`;

  document.getElementById('area-stats-row').innerHTML = `
    <div class="area-stat"><div class="area-stat-val">${area.totalFormatos}</div><div class="area-stat-label">Total formatos</div></div>
    <div class="area-stat"><div class="area-stat-val" style="color:var(--verde)">${area.totalRegistros}</div><div class="area-stat-label">Total registros</div></div>
    <div class="area-stat"><div class="area-stat-val">${area.promRegistros}</div><div class="area-stat-label">Prom. registros</div></div>
    <div class="area-stat"><div class="area-stat-val" ${area.pctUso===0?'style="color:var(--red)"':''}>${area.pctUso}%</div><div class="area-stat-label">% uso activo</div></div>`;

  document.getElementById('area-activos-section').innerHTML = area.formatos.activos.length
    ? renderTablaFormatos(area.formatos.activos, 'Formatos activos', 'green', 'activa')
    : `<div class="alerta-inactiva">⚠ Esta area no tiene ningun formato activo en los ultimos ${CONFIG.diasInactividad} dias.</div>`;

  document.getElementById('area-inactivos-section').innerHTML = area.formatos.inactivos.length
    ? renderTablaFormatos(area.formatos.inactivos, `Inactivos — mas de ${CONFIG.diasInactividad} dias sin uso`, 'amber', 'inactiva')
    : '';

  document.getElementById('area-nunca-section').innerHTML = area.formatos.nunca.length
    ? renderTablaNunca(area.formatos.nunca) : '';

  const usuarios = DATOS.usuarios.filter(u => u.areas.includes(area.nombre));
  let uHtml = usuarios.length
    ? usuarios.map((u,i) => {
        const av  = AVATAR_COLORS[i % AVATAR_COLORS.length];
        const ini = u.nombre.split(' ').map(n=>n[0]).join('').slice(0,2).toUpperCase();
        return `<div class="usuario-row">
          <div class="avatar" style="background:${av.bg};color:${av.color}">${ini}</div>
          <div class="user-info">
            <div class="user-name">${u.nombre}</div>
            <div class="user-detail">${u.email} · ${u.formatos} formatos · ultima actividad ${u.ultimaActividad}</div>
          </div>
          <div class="user-count">${u.registros} reg</div>
        </div>`;
      }).join('')
    : '<div class="empty-state">Sin actividad de usuarios registrada</div>';

  document.getElementById('area-usuarios-section').innerHTML = `
    <div class="tabla-section">
      <div class="tabla-section-title gray">Usuarios activos en esta area</div>
      <div class="card">${uHtml}</div>
    </div>`;
}

function renderTablaFormatos(formatos, titulo, clase, tipo) {
  const maxItems = Math.max(...formatos.map(f => f.totalItems), 1);
  const POWERAPPS_BASE = 'https://apps.powerapps.com/play/e/5e486c6e-204c-ee38-94c6-2c68092acf16/a/b6eacb83-c58c-4324-b552-7fdb4db70dd6?tenantId=f16b5c8e-ad45-4fc4-a55b-0af26c456817&screenName=Sc_';
  const filas = formatos.map(f => {
    const pct  = Math.round((f.totalItems / maxItems) * 100);
    const barC = tipo === 'activa' ? '#3DAA35' : '#EF9F27';
    const dC   = f.diasSinUso === 0 ? 'dias-ok' : f.diasSinUso <= 30 ? 'dias-ok' : f.diasSinUso <= 90 ? 'dias-warn' : 'dias-danger';
    const powerAppsUrl = POWERAPPS_BASE + f.codigo;
    return `<tr class="${tipo === 'inactiva' ? 'row-inactiva' : ''}">
      <td><strong><a href="${powerAppsUrl}" target="_blank" style="color:inherit;text-decoration:none;border-bottom:1.5px dashed #3DAA35;cursor:pointer" title="Ver en Power Apps">${f.codigo} ↗</a></strong></td>
      <td><span class="badge badge-${tipo === 'activa' ? 'activa' : 'inactiva'}">${tipo === 'activa' ? 'Activa' : 'Inactiva'}</span></td>
      <td><div class="bar-mini-wrap"><div class="bar-mini-bg"><div class="bar-mini-fill" style="width:${pct}%;background:${barC}"></div></div>${f.totalItems}</div></td>
      <td><span class="${freqClass(f.frecuencia)}">${f.frecuencia}</span></td>
      <td>${f.ultimoUsuario}</td>
      <td style="font-size:11px;color:var(--muted)">${f.emailUltimo}</td>
      <td>${f.ultimaMod}</td>
      <td><span class="${dC}">${f.diasSinUso ?? '—'}</span></td>
    </tr>`;
  }).join('');
  return `<div class="tabla-section">
    <div class="tabla-section-title ${clase}">${titulo}</div>
    <div class="tabla-wrap"><table>
      <thead><tr>
        <th style="width:110px">Codigo</th><th style="width:72px">Estado</th>
        <th style="width:80px">Registros</th><th style="width:90px">Frecuencia</th>
        <th style="width:130px">Ultimo usuario</th><th>Email</th>
        <th style="width:100px">Ultima vez</th><th style="width:50px">Dias</th>
      </tr></thead>
      <tbody>${filas}</tbody>
    </table></div>
  </div>`;
}

function renderTablaNunca(formatos) {
  const filas = formatos.map(f => `
    <tr class="row-nunca">
      <td><strong>${f.codigo}</strong></td>
      <td><span class="badge badge-nunca">Sin uso</span></td>
      <td colspan="5" style="color:var(--hint)">Lista creada sin ningun registro ingresado</td>
      <td style="font-size:11px;color:var(--hint)">${f.fechaCreacion}</td>
    </tr>`).join('');
  return `<div class="tabla-section">
    <div class="tabla-section-title red">Nunca usados — sin ningun registro</div>
    <div class="tabla-wrap"><table>
      <thead><tr>
        <th style="width:110px">Codigo</th><th style="width:72px">Estado</th>
        <th colspan="5">Observacion</th><th style="width:100px">Fecha creacion</th>
      </tr></thead>
      <tbody>${filas}</tbody>
    </table></div>
  </div>`;
}

// -----------------------------------------------
// NAVEGACION
// -----------------------------------------------
function mostrarArea(nombre, btn) {
  document.querySelectorAll('.nav-tab').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  const vR = document.getElementById('vista-resumen');
  const vA = document.getElementById('vista-area');
  if (nombre === 'resumen') {
    vR.classList.add('active');
    vA.classList.remove('active');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  } else {
    const area = DATOS.areas.find(a => a.nombre === nombre);
    if (!area) return;
    vR.classList.remove('active');
    vA.classList.add('active');
    renderizarArea(area);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

// -----------------------------------------------
// UTILIDADES
// -----------------------------------------------
function setText(id, val) { const el = document.getElementById(id); if (el) el.textContent = val; }
function mostrarLoading(show) { const el = document.getElementById('loading'); if (el) el.classList.toggle('show', show); }
function actualizarFecha() {
  const now = new Date();
  setText('fecha-generacion', `Actualizado: ${now.toLocaleDateString('es-CO')} ${now.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'})}`);
}
function freqClass(f) {
  if (f==='Diaria')    return 'freq-d';
  if (f==='Semanal')   return 'freq-s';
  if (f==='Mensual')   return 'freq-m';
  if (f==='Esporadica') return 'freq-e';
  return 'freq-n';
}

// -----------------------------------------------
// INICIO
// -----------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  cargarDatosEjemplo();
});
