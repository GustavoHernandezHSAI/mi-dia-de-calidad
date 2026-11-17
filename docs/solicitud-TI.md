# Solicitud a TI — Lectura de calendario de Outlook para app de planificación personal

**Solicitante:** Coordinador Operativo de Procesos, Dirección de Calidad Corporativa
**Asunto:** Registro de una aplicación en Microsoft Entra ID con permiso de solo lectura de calendario

Buen día. Uso una aplicación web personal de planificación de trabajo (“Mi Día de Calidad”) y quisiera que lea
mi calendario de Outlook para detectar horarios ocupados y conflictos. Solicito su apoyo para registrarla
en Microsoft Entra ID del hospital con estos parámetros:

| Concepto | Valor |
|---|---|
| Tipo de cuenta admitida | Solo cuentas de este directorio organizativo (inquilino único) |
| Plataforma | **Aplicación de página única (SPA)** |
| URI de redirección | `https://__________.vercel.app/` (la dirección exacta de la app, con `/` final) |
| Permisos de Microsoft Graph (delegados) | `Calendars.Read`, `offline_access`, `openid` |
| Secreto de cliente | **No se requiere** (flujo de código de autorización con PKCE) |
| Consentimiento | Si la política del inquilino lo exige, consentimiento de administrador para `Calendars.Read` |

**Qué hace y qué no hace la aplicación**

- Lee únicamente **mis** eventos (asunto, inicio, fin, estado ocupado/libre) con mi propia sesión.
- **No** crea, modifica ni elimina eventos; **no** lee correos, contactos, Teams ni archivos.
- Los eventos se procesan en el navegador y **no se almacenan** en servidores externos.
- No maneja datos de pacientes ni del expediente clínico.
- El acceso se puede revocar en cualquier momento desde Entra ID o desde *Mi cuenta → Aplicaciones*.

**Datos que necesito de vuelta**

1. ID de aplicación (cliente)
2. ID de directorio (inquilino)

En una fase posterior podría solicitar `Calendars.ReadWrite` para crear eventos, siempre con confirmación expresa
en cada caso; por ahora solo se requiere lectura.

Gracias.
