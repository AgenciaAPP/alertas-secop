import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

// ====================================================================================
// VARIABLES DE ENTORNO
// ====================================================================================
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;

const LIST_ID_TE_SERVIDORES = process.env.LIST_ID_TE_SERVIDORES; // f4eb9e0f-3223-41e8-b28e-12ef05e6c915
const LIST_ID_CONFIG_ALERTAS = process.env.LIST_ID_CONFIG_ALERTAS; // dd3181bf-4920-487b-81aa-33a8e0ab3882

const CRON_SECRET = process.env.CRON_SECRET; // protege el endpoint del cron
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD; // protege el panel de administración

// ====================================================================================
// MODO PRUEBA: si TEST_EMAIL_OVERRIDE tiene un valor, TODOS los correos se envían
// a esa dirección en vez de al supervisor real (el asunto deja claro el destinatario
// original). Déjala vacía/sin definir para que el envío sea real a cada supervisor.
// ====================================================================================
const TEST_EMAIL_OVERRIDE = process.env.TEST_EMAIL_OVERRIDE || '';

const NIT_AGENCIA_APP = '900623766';
const DIAS_ANTICIPACION = 30;
const DIAS_HABILES_RESPUESTA = 5;

// URL del logo institucional de Agencia APP (reemplazar por la URL real hospedada,
// ej. un enlace público de SharePoint/OneDrive con acceso directo a la imagen, o
// subir el PNG/SVG del manual de marca a un bucket público).
const LOGO_URL = process.env.LOGO_URL || 'https://via.placeholder.com/220x60?text=Agencia+APP';

// Paleta oficial del Manual de Identidad Gráfica 2026 - Agencia APP
const COLOR_PRIMARIO = '#1878ba'; // azul institucional
const COLOR_ACENTO = '#ffd500';   // amarillo
const COLOR_TEXTO = '#333333';
const COLOR_GRIS = '#6a6a6a';

// ====================================================================================
// AUTENTICACIÓN MICROSOFT GRAPH (mismo App Registration que STC)
// ====================================================================================
async function getMicrosoftGraphToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  try {
    const response = await axios.post(url, params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return response.data.access_token;
  } catch (error) {
    console.error('Error de autenticación con Azure AD:', error.response?.data || error.message);
    throw new Error('No se pudo adquirir el Token de Acceso de Microsoft.');
  }
}

// ====================================================================================
// MIDDLEWARE: PROTECCIÓN DEL PANEL ADMIN
// ====================================================================================
function requireAdminAuth(req, res, next) {
  const provided = req.headers['x-admin-password'];
  if (!ADMIN_PASSWORD || provided !== ADMIN_PASSWORD) {
    return res.status(401).json({ success: false, message: 'No autorizado.' });
  }
  next();
}

// ====================================================================================
// MIDDLEWARE: PROTECCIÓN DEL ENDPOINT DE CRON
// Vercel Cron envía automáticamente el header "Authorization: Bearer {CRON_SECRET}"
// cuando la variable de entorno CRON_SECRET está configurada en el proyecto.
// ====================================================================================
function requireCronAuth(req, res, next) {
  const authHeader = req.headers['authorization'];
  const permiteManual = req.query.secret && req.query.secret === CRON_SECRET;
  if (!CRON_SECRET || (authHeader !== `Bearer ${CRON_SECRET}` && !permiteManual)) {
    return res.status(401).json({ success: false, message: 'No autorizado.' });
  }
  next();
}

// ====================================================================================
// SHAREPOINT: LEER SUPERVISORES DESDE TE_Servidores (Título=cédula, NombreCompleto, Correo)
// ====================================================================================
async function obtenerSupervisores(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_TE_SERVIDORES}/items?expand=fields&$top=999`;
  const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });

  const mapa = {};
  for (const item of response.data.value) {
    const cedula = item.fields.Title ? String(item.fields.Title).trim() : '';
    if (!cedula) continue;
    mapa[cedula] = {
      cedula,
      nombre: item.fields.NombreCompleto || '',
      correo: item.fields.Correo || ''
    };
  }
  return mapa;
}

// ====================================================================================
// SHAREPOINT: LEER/ESCRIBIR SUPERVISORES (CRUD PARA EL PANEL ADMIN)
// ====================================================================================
async function listarSupervisoresConId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_TE_SERVIDORES}/items?expand=fields&$top=999`;
  const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
  return response.data.value.map(item => ({
    idSharePoint: item.id,
    cedula: item.fields.Title || '',
    nombre: item.fields.NombreCompleto || '',
    correo: item.fields.Correo || ''
  }));
}

async function crearSupervisor(token, { cedula, nombre, correo }) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_TE_SERVIDORES}/items`;
  const payload = { fields: { Title: String(cedula).trim(), NombreCompleto: nombre, Correo: correo } };
  await axios.post(url, payload, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
}

async function actualizarSupervisor(token, idSharePoint, { cedula, nombre, correo }) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_TE_SERVIDORES}/items/${idSharePoint}`;
  const payload = { fields: { Title: String(cedula).trim(), NombreCompleto: nombre, Correo: correo } };
  await axios.patch(url, payload, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
}

async function eliminarSupervisor(token, idSharePoint) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_TE_SERVIDORES}/items/${idSharePoint}`;
  await axios.delete(url, { headers: { 'Authorization': `Bearer ${token}` } });
}

// ====================================================================================
// SHAREPOINT: LEER/ESCRIBIR EL REMITENTE CONFIGURABLE (ConfigAlertasSecop)
// ====================================================================================
async function obtenerRemitenteConfigurado(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_CONFIG_ALERTAS}/items?expand=fields&$filter=fields/Title eq 'REMITENTE_ACTIVO'`;
  const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
  if (response.data.value.length === 0) return null;
  const item = response.data.value[0];
  return {
    idSharePoint: item.id,
    correoRemitente: item.fields.CorreoRemitente || ''
  };
}

async function actualizarRemitente(token, idSharePoint, nuevoCorreo) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_CONFIG_ALERTAS}/items/${idSharePoint}`;
  const payload = { fields: { CorreoRemitente: nuevoCorreo } };
  await axios.patch(url, payload, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
}

// ====================================================================================
// SECOP II: CONSULTAR CONTRATOS DE PRESTACIÓN DE SERVICIOS PRÓXIMOS A VENCER
// ====================================================================================
function formatoFechaSoQL(date) {
  return date.toISOString().split('T')[0]; // YYYY-MM-DD
}

async function obtenerContratosProximosAVencer() {
  const hoy = new Date();
  const limite = new Date();
  limite.setDate(hoy.getDate() + DIAS_ANTICIPACION);

  const desde = formatoFechaSoQL(hoy);
  const hasta = formatoFechaSoQL(limite);

  const whereClause = `fecha_de_fin_del_contrato between '${desde}T00:00:00.000' and '${hasta}T23:59:59.000'`;

  const url = 'https://www.datos.gov.co/resource/jbjy-vk9h.json';
  const params = {
    nit_entidad: NIT_AGENCIA_APP,
    tipo_de_contrato: 'Prestación de servicios',
    '$where': whereClause,
    '$limit': 1000
  };

  const response = await axios.get(url, {
    params,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });

  return response.data || [];
}

// ====================================================================================
// AGRUPAR CONTRATOS POR CÉDULA DE SUPERVISOR
// ====================================================================================
function agruparContratosPorSupervisor(contratos) {
  const grupos = {};
  for (const contrato of contratos) {
    const cedulaSupervisor = contrato.n_mero_de_documento_supervisor
      ? String(contrato.n_mero_de_documento_supervisor).trim()
      : '';
    if (!cedulaSupervisor) continue;

    if (!grupos[cedulaSupervisor]) grupos[cedulaSupervisor] = [];
    grupos[cedulaSupervisor].push(contrato);
  }
  return grupos;
}

// ====================================================================================
// FORMATEO
// ====================================================================================
function formatearFecha(fechaIso) {
  if (!fechaIso) return 'No definida';
  const fecha = new Date(fechaIso);
  return fecha.toLocaleDateString('es-CO', { year: 'numeric', month: 'long', day: 'numeric' });
}

// ====================================================================================
// PLANTILLA DEL CORREO (línea gráfica Agencia APP - Manual 2026)
// ====================================================================================
function construirFilasTabla(contratos) {
  return contratos.map(c => `
    <tr>
      <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; font-size:13px; color:${COLOR_TEXTO};">${c.referencia_del_contrato || 'No definido'}</td>
      <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; font-size:13px; color:${COLOR_TEXTO};">${c.proveedor_adjudicado || 'No definido'}</td>
      <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; font-size:13px; color:${COLOR_TEXTO}; line-height:1.5;">${c.objeto_del_contrato || 'No definido'}</td>
      <td style="padding:10px 12px; border-bottom:1px solid #e5e7eb; font-size:13px; color:${COLOR_TEXTO}; white-space:nowrap;"><strong>${formatearFecha(c.fecha_de_fin_del_contrato)}</strong></td>
    </tr>
  `).join('');
}

function construirCorreoHtml(nombreSupervisor, contratos) {
  const filas = construirFilasTabla(contratos);
  const cantidad = contratos.length;
  const plural = cantidad === 1 ? 'contrato' : 'contratos';

  return `
  <div style="font-family: 'Metropolis', 'Segoe UI', Arial, sans-serif; background-color:#f4f6f8; padding:32px 16px; margin:0;">
    <div style="max-width:680px; margin:0 auto; background-color:#ffffff; border-radius:10px; overflow:hidden; box-shadow:0 2px 10px rgba(0,0,0,0.06);">

      <!-- HEADER -->
      <div style="background-color:${COLOR_PRIMARIO}; padding:24px 32px; text-align:center;">
        <img src="${LOGO_URL}" alt="Agencia APP" style="max-height:48px;" />
      </div>

      <!-- CINTA DE ACENTO -->
      <div style="height:5px; background-color:${COLOR_ACENTO};"></div>

      <!-- CUERPO -->
      <div style="padding:32px;">
        <h2 style="color:${COLOR_PRIMARIO}; font-size:19px; margin:0 0 4px 0;">Hola, ${(nombreSupervisor || '').toUpperCase()}</h2>
        <p style="color:${COLOR_GRIS}; font-size:13px; margin:0 0 24px 0;">Alerta de vencimiento de contratos de prestación de servicios</p>

        <p style="font-size:14px; line-height:1.6; color:${COLOR_TEXTO};">
          Te informamos que tienes <strong>${cantidad} ${plural}</strong> de prestación de servicios bajo tu supervisión
          que vence${cantidad === 1 ? '' : 'n'} en los próximos <strong>${DIAS_ANTICIPACION} días</strong>:
        </p>

        <table style="width:100%; border-collapse:collapse; margin:20px 0; border:1px solid #e5e7eb; border-radius:6px; overflow:hidden;">
          <thead>
            <tr style="background-color:#f0f5f9;">
              <th style="padding:10px 12px; text-align:left; font-size:12px; color:${COLOR_PRIMARIO}; text-transform:uppercase; letter-spacing:0.3px;">Referencia</th>
              <th style="padding:10px 12px; text-align:left; font-size:12px; color:${COLOR_PRIMARIO}; text-transform:uppercase; letter-spacing:0.3px;">Contratista</th>
              <th style="padding:10px 12px; text-align:left; font-size:12px; color:${COLOR_PRIMARIO}; text-transform:uppercase; letter-spacing:0.3px;">Objeto</th>
              <th style="padding:10px 12px; text-align:left; font-size:12px; color:${COLOR_PRIMARIO}; text-transform:uppercase; letter-spacing:0.3px;">Vence</th>
            </tr>
          </thead>
          <tbody>
            ${filas}
          </tbody>
        </table>

        <div style="background-color:#fffbea; border-left:4px solid ${COLOR_ACENTO}; padding:14px 16px; border-radius:4px; margin:24px 0;">
          <p style="margin:0; font-size:13px; color:${COLOR_TEXTO}; line-height:1.6;">
            Se espera tu respuesta en los próximos <strong>${DIAS_HABILES_RESPUESTA} días hábiles</strong> para conocer
            las acciones a seguir respecto a estos contratos (prórroga, terminación, nueva contratación, etc.).
          </p>
        </div>

        <p style="font-size:13px; color:${COLOR_GRIS}; line-height:1.6;">
          Este es un mensaje automático generado a partir de la información pública registrada en SECOP II.
          Si tienes dudas sobre alguno de estos contratos, por favor contacta a la Dirección Técnica.
        </p>
      </div>

      <!-- FOOTER -->
      <div style="background-color:#f8f9fa; padding:18px 32px; text-align:center; border-top:1px solid #e5e7eb;">
        <p style="margin:0; font-size:11px; color:#94a3b8;">Alertas de Vencimiento de Contratos &bull; Dirección Técnica &bull; Agencia APP</p>
      </div>
    </div>
  </div>
  `;
}

// ====================================================================================
// ENVÍO DEL CORREO VÍA MICROSOFT GRAPH
// ====================================================================================
async function enviarCorreoAlerta(token, correoRemitente, correoSupervisor, nombreSupervisor, contratos) {
  if (!correoRemitente) {
    throw new Error('No hay un correo remitente configurado en ConfigAlertasSecop.');
  }
  if (!correoSupervisor) {
    console.warn(`Sin correo registrado para el supervisor "${nombreSupervisor}". No se envió alerta.`);
    return { enviado: false, motivo: 'sin_correo_supervisor' };
  }

  const url = `https://graph.microsoft.com/v1.0/users/${correoRemitente}/sendMail`;
  const cantidad = contratos.length;

  const modoPrueba = Boolean(TEST_EMAIL_OVERRIDE);
  const destinatarioFinal = modoPrueba ? TEST_EMAIL_OVERRIDE : correoSupervisor;
  const asunto = modoPrueba
    ? `🧪 [PRUEBA - destinatario real: ${correoSupervisor}] ${cantidad} contrato(s) de prestación de servicios próximos a vencer`
    : `⏰ Alerta: ${cantidad} contrato(s) de prestación de servicios próximos a vencer`;

  const mailPayload = {
    message: {
      subject: asunto,
      body: {
        contentType: 'HTML',
        content: construirCorreoHtml(nombreSupervisor, contratos)
      },
      toRecipients: [{ emailAddress: { address: destinatarioFinal } }]
    }
  };

  await axios.post(url, mailPayload, {
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
  });

  return { enviado: true };
}

// ====================================================================================
// RUTA: CRON DIARIO DE ALERTAS DE VENCIMIENTO
// ====================================================================================
app.get('/api/cron-alertas-vencimiento', requireCronAuth, async (req, res) => {
  try {
    const token = await getMicrosoftGraphToken();

    const remitenteConfig = await obtenerRemitenteConfigurado(token);
    const correoRemitente = remitenteConfig?.correoRemitente?.trim();

    if (!correoRemitente) {
      return res.status(200).json({
        success: false,
        message: 'No hay un correo remitente configurado en ConfigAlertasSecop (campo CorreoRemitente vacío). No se envió ninguna alerta.'
      });
    }

    const [contratos, supervisores] = await Promise.all([
      obtenerContratosProximosAVencer(),
      obtenerSupervisores(token)
    ]);

    const grupos = agruparContratosPorSupervisor(contratos);
    const resultados = [];

    for (const cedulaSupervisor of Object.keys(grupos)) {
      const supervisor = supervisores[cedulaSupervisor];
      const contratosDelSupervisor = grupos[cedulaSupervisor];

      if (!supervisor) {
        resultados.push({
          cedulaSupervisor,
          totalContratos: contratosDelSupervisor.length,
          enviado: false,
          motivo: 'supervisor_no_registrado_en_TE_Servidores'
        });
        continue;
      }

      try {
        const resultadoEnvio = await enviarCorreoAlerta(
          token,
          correoRemitente,
          supervisor.correo,
          supervisor.nombre,
          contratosDelSupervisor
        );
        resultados.push({
          cedulaSupervisor,
          nombreSupervisor: supervisor.nombre,
          totalContratos: contratosDelSupervisor.length,
          ...resultadoEnvio
        });
      } catch (errEnvio) {
        console.error(`Error enviando alerta a ${supervisor.correo}:`, errEnvio.response?.data || errEnvio.message);
        resultados.push({
          cedulaSupervisor,
          nombreSupervisor: supervisor.nombre,
          totalContratos: contratosDelSupervisor.length,
          enviado: false,
          motivo: 'error_envio'
        });
      }
    }

    return res.status(200).json({
      success: true,
      modoPrueba: Boolean(TEST_EMAIL_OVERRIDE),
      totalContratosEncontrados: contratos.length,
      totalSupervisoresNotificados: resultados.filter(r => r.enviado).length,
      detalle: resultados
    });
  } catch (error) {
    console.error('Error en cron de alertas de vencimiento:', error.response?.data || error.message);
    return res.status(500).json({ success: false, message: 'Error ejecutando el proceso de alertas.', detail: error.message });
  }
});

// ====================================================================================
// RUTA: EJECUCIÓN MANUAL (para pruebas desde el navegador con el secreto en query string)
// Ej: /api/cron-alertas-vencimiento?secret=TU_CRON_SECRET
// (usa la misma ruta y protección de arriba, ya contempla req.query.secret)
// ====================================================================================

// ====================================================================================
// RUTAS ADMIN: SUPERVISORES (CRUD sobre TE_Servidores)
// ====================================================================================
app.get('/api/admin/supervisores', requireAdminAuth, async (req, res) => {
  try {
    const token = await getMicrosoftGraphToken();
    const supervisores = await listarSupervisoresConId(token);
    res.json({ success: true, data: supervisores });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.message });
  }
});

app.post('/api/admin/supervisores', requireAdminAuth, async (req, res) => {
  const { cedula, nombre, correo } = req.body;
  if (!cedula || !correo) {
    return res.status(400).json({ success: false, message: 'Cédula y correo son obligatorios.' });
  }
  try {
    const token = await getMicrosoftGraphToken();
    await crearSupervisor(token, { cedula, nombre, correo });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.response?.data || error.message });
  }
});

app.patch('/api/admin/supervisores/:idSharePoint', requireAdminAuth, async (req, res) => {
  const { idSharePoint } = req.params;
  const { cedula, nombre, correo } = req.body;
  try {
    const token = await getMicrosoftGraphToken();
    await actualizarSupervisor(token, idSharePoint, { cedula, nombre, correo });
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.response?.data || error.message });
  }
});

app.delete('/api/admin/supervisores/:idSharePoint', requireAdminAuth, async (req, res) => {
  const { idSharePoint } = req.params;
  try {
    const token = await getMicrosoftGraphToken();
    await eliminarSupervisor(token, idSharePoint);
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.response?.data || error.message });
  }
});

// ====================================================================================
// RUTAS ADMIN: REMITENTE CONFIGURABLE (ConfigAlertasSecop)
// ====================================================================================
app.get('/api/admin/remitente', requireAdminAuth, async (req, res) => {
  try {
    const token = await getMicrosoftGraphToken();
    const config = await obtenerRemitenteConfigurado(token);
    res.json({ success: true, data: config });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.message });
  }
});

app.put('/api/admin/remitente', requireAdminAuth, async (req, res) => {
  const { correoRemitente } = req.body;
  if (!correoRemitente) {
    return res.status(400).json({ success: false, message: 'Falta el correo remitente.' });
  }
  try {
    const token = await getMicrosoftGraphToken();
    const config = await obtenerRemitenteConfigurado(token);
    if (!config) {
      return res.status(404).json({ success: false, message: 'No existe el ítem REMITENTE_ACTIVO en ConfigAlertasSecop.' });
    }
    await actualizarRemitente(token, config.idSharePoint, correoRemitente);
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.response?.data || error.message });
  }
});

// ====================================================================================
// RUTA: SALUD DEL SERVICIO
// ====================================================================================
app.get('/', (req, res) => {
  res.send('Servicio de Alertas de Vencimiento SECOP - Agencia APP, operando.');
});

export default app;
