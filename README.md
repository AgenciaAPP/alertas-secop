# Alertas de Vencimiento SECOP - Agencia APP

Automatización diaria que revisa los contratos de **prestación de servicios** de la
Agencia APP en **SECOP II** y envía un correo a cada supervisor con los contratos
que vencen en los próximos 30 días.

## ¿Cómo funciona?

1. Un **Vercel Cron** llama todos los días a las 8:00 a.m. (hora Bogotá) al endpoint
   `GET /api/cron-alertas-vencimiento`.
2. El endpoint consulta la API pública de SECOP II (dataset `jbjy-vk9h`) filtrando por
   `nit_entidad` de la Agencia APP, `tipo_de_contrato = "Prestación de servicios"` y
   `fecha_de_fin_del_contrato` entre hoy y hoy+30 días.
3. Agrupa los contratos encontrados por la cédula del supervisor
   (`n_mero_de_documento_supervisor`).
4. Cruza cada cédula con la lista de SharePoint `TE_Servidores` para obtener el
   nombre y correo del supervisor.
5. Envía un correo HTML (con la línea gráfica de Agencia APP) por supervisor, con
   una tabla de sus contratos y la solicitud de respuesta en 5 días hábiles.
6. El correo se envía vía Microsoft Graph, usando como remitente el valor guardado
   en la lista de SharePoint `ConfigAlertasSecop` (editable desde `/admin` sin
   redeploy).

## Despliegue en Vercel

1. Sube esta carpeta a un repositorio nuevo en GitHub (ej. `alertas-secop`).
2. Impórtalo en Vercel como un nuevo proyecto.
3. En **Settings → Environment Variables**, agrega todas las variables de
   `.env.example` con sus valores reales:
   - `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`: los mismos del App Registration de STC.
   - `SITE_ID`, `LIST_ID_TE_SERVIDORES`, `LIST_ID_CONFIG_ALERTAS`: ya vienen
     completados en `.env.example` con los valores confirmados para Agencia APP.
   - `CRON_SECRET`: genera un valor aleatorio (ver comentario en `.env.example`).
   - `ADMIN_PASSWORD`: la contraseña que quieras para entrar a `/admin`.
   - `LOGO_URL`: URL pública del logo de Agencia APP (súbelo a SharePoint con enlace
     "cualquiera con el enlace puede ver", o a cualquier hosting de imágenes).
4. Despliega. Vercel detectará automáticamente el cron definido en `vercel.json`.

## Panel de administración

Entra a `https://tu-proyecto.vercel.app/admin`, ingresa la contraseña
(`ADMIN_PASSWORD`) y podrás:

- Agregar, editar (próxima iteración) y eliminar supervisores.
- Cambiar el correo remitente sin tocar código ni redeploy.

## Requisito pendiente antes de producción

El buzón remitente (`direcciontecnica@app.gov.co` u otro que configures en `/admin`)
debe existir en Microsoft 365 y el App Registration debe tener permiso `Mail.Send`
para enviar en su nombre (igual que ya sucede con `notificacionesmop@app.gov.co` en
STC). Mientras el campo `CorreoRemitente` en `ConfigAlertasSecop` esté vacío, el
cron no enviará ningún correo (falla de forma segura, sin intentar enviar con un
remitente inválido).

## Prueba manual

Puedes disparar el proceso manualmente (sin esperar al cron) visitando:

```
https://tu-proyecto.vercel.app/api/cron-alertas-vencimiento?secret=TU_CRON_SECRET
```

Devuelve un JSON con el resumen de contratos encontrados y correos enviados.
