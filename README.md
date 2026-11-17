# Mi Día de Calidad — MVP

Aplicación web personal para planear y dar seguimiento al trabajo del Coordinador Operativo de Procesos (Calidad).
Se usa desde el navegador en computadora y celular. No se compila: son archivos HTML, CSS y JavaScript.

## Contenido

```
index.html               Página principal
css/app.css              Estilos (claro / oscuro, móvil / escritorio)
js/config.js             ← ÚNICO archivo que se edita (claves de Supabase y Outlook)
js/app.js                Interfaz: Hoy, Semana, Tareas, Intramed, Cierre, Resumen, Ajustes
js/seguimiento.js        Seguimientos (etapas, actividades, tiempo diario) y Avance documental
js/flows.js              Flujos de documentos y formatos (etapas, cancha, rondas) — aquí se ajustan
js/xlsx.js               Lectura de Excel/CSV sin internet para importar la base documental
js/planner.js            Jornada, reuniones fijas, capacidad, conflictos, sugerencias, imprevistos
js/parser.js             Interpretación de texto libre (reglas locales, sin IA)
js/store.js              Datos: caché local, cola sin conexión y sincronización en tiempo real
js/outlook.js            Lectura del calendario de Outlook (OAuth PKCE, solo lectura)
js/summary.js            Cierre de jornada y resumen para la Directora
sw.js, manifest          Funcionamiento sin conexión y acceso directo en el celular
supabase/schema.sql      Base de datos, seguridad por usuario y tiempo real
docs/solicitud-TI.md     Texto para pedir a TI el registro de la app en Microsoft Entra ID
vercel.json              Encabezados de seguridad para la publicación
```

## 1. Probar sin instalar nada (modo local)

Si `js/config.js` no tiene datos de Supabase, la app abre en **modo local**: todo se guarda solo en ese navegador.
Sirve para probar. Para usarla en computadora y celular con los mismos datos, sigue los pasos 2 y 3.

## 2. Base de datos y cuenta: Supabase (gratis, ~10 min)

1. Entra a <https://supabase.com>, crea una cuenta y luego **New project**. Ponle nombre (p. ej. `mi-dia-de-calidad`),
   una contraseña de base de datos (guárdala) y la región más cercana disponible.
2. En el proyecto abre **SQL Editor → New query**, pega todo el contenido de `supabase/schema.sql` y presiona **Run**.
   Debe terminar sin errores (puedes ejecutarlo de nuevo sin perder datos).
3. **Authentication → Users → Add user → Create new user**: escribe tu correo y una contraseña, y marca
   **Auto Confirm User**. Esa será tu cuenta para entrar a la app.
4. **Authentication → (Sign In / Providers o Settings)**: desactiva **Allow new users to sign up**, para que nadie más pueda crear cuenta.
5. **Project Settings → API** (o **API Keys**): copia **Project URL** y la clave **anon / public**.
6. Abre `js/config.js` y pega esos dos valores en `SUPABASE_URL` y `SUPABASE_ANON_KEY`.
   La clave *anon* es pública por diseño; la protección real son las políticas de seguridad (RLS) del paso 2:
   cada usuario solo ve sus propios registros.
7. Cuando tengas la dirección de Vercel (paso 3), regresa a **Authentication → URL Configuration** y ponla en
   **Site URL** (necesario para el correo de “Olvidé mi contraseña”).

> Plan gratuito de Supabase: el proyecto se pausa tras **1 semana sin uso** (se reactiva desde el panel con
> “Resume project”) y **no incluye respaldos automáticos**. Usa **Ajustes → Descargar respaldo** cada semana.
> Si quieres respaldos diarios automáticos, el plan Pro cuesta desde USD 25/mes. Es decisión tuya; el MVP no lo requiere.

## 3. Publicar: Vercel (gratis, ~10 min)

1. Crea una cuenta en <https://github.com> → **New repository** → nombre `mi-dia-de-calidad` → **Private** → **Create**.
2. En el repositorio: **Add file → Upload files** y arrastra **el contenido** de esta carpeta
   (index.html, css, js, icons, supabase, docs, sw.js, manifest.webmanifest, vercel.json). **Commit changes**.
3. Entra a <https://vercel.com> → **Sign up with GitHub** → **Add New… → Project** → importa `mi-dia-de-calidad`
   → *Framework Preset*: **Other** → **Deploy**.
4. Vercel te da una dirección tipo `https://mi-dia-de-calidad.vercel.app`. Esa es tu app.
5. Para cambiar algo después (p. ej. `js/config.js`): edítalo en GitHub (ícono de lápiz) y guarda;
   Vercel publica la nueva versión sola en 1 minuto.

## 4. Uso en celular y computadora

- **Computadora:** abre la dirección en Edge o Chrome e inicia sesión. Atajos: `N` nueva tarea, `I` imprevisto.
  Puedes arrastrar tareas en la línea del día y entre días en la semana.
- **Celular:** abre la misma dirección en Safari (iPhone) o Chrome (Android) e inicia sesión.
  Opcional: menú **Compartir → Agregar a pantalla de inicio** para abrirla como app (sigue siendo el navegador).
  En el celular, toca una tarea y usa **Mover a** para reprogramar.
- **Sincronización:** los cambios aparecen en el otro dispositivo en segundos. Sin internet, la app sigue
  funcionando con lo último guardado y envía los cambios al recuperar la conexión (el indicador superior lo muestra).
- **Avisos:** en Ajustes puedes activar avisos del navegador 5 minutos antes de cada bloque y al final de la jornada.
  Funcionan mientras la app está abierta (en una pestaña o en la pantalla de inicio).

## 5. Outlook (cuando TI lo autorice)

La app funciona completa sin Outlook. Para activarlo:

1. Envía a TI el texto de `docs/solicitud-TI.md`.
2. TI te entregará dos datos: **ID de aplicación (cliente)** e **ID de directorio (inquilino)**.
3. Ponlos en `js/config.js` (`MS_CLIENT_ID` y `MS_TENANT_ID`) y guarda en GitHub.
4. En la app: **Ajustes → Outlook → Conectar Outlook** e inicia sesión con tu cuenta del hospital.

Qué hace: lee tus eventos (solo lectura) para marcar horarios ocupados y conflictos. No crea, cambia ni borra eventos.
Los eventos se leen en tu navegador y **no se guardan en la base de datos de Supabase**.

## Actualizar una instalación existente (versión 1.1)

1. En Supabase → SQL Editor, vuelve a ejecutar `supabase/schema.sql` completo. Agrega las tablas `projects` y `docs` sin tocar tus datos.
2. En GitHub sube los archivos nuevos y cambiados (puedes subir la carpeta completa encima). Vercel publica solo.
3. Al abrir la app, los asuntos de Intramed existentes se convierten solos: incidencia → reporte, requerimiento → necesidad nueva.

## Seguimientos y avance documental

- **Seguimientos:** trabajos grandes con etapas (lista que marcas), actividades (tareas vinculadas) y, si quieres,
  **tiempo diario reservado** (p. ej. 1 h de lunes a viernes) que se agenda solo hasta terminar o hasta una fecha.
  Plantilla incluida: *Formatos de expediente clínico electrónico*.
- **Avance documental:** cada documento o formato tiene flujo, etapa, ronda de revisión y “cancha” (Calidad,
  dueño de proceso, consultor, proveedor, Compras). Cuando un documento llega a tu cancha se crea la tarea
  (revisar, visto bueno, diagrama, cargar en Intramed); al avanzarlo se cierra la tarea anterior.
- **Importar tu Excel:** Avance documental → Importar Excel / CSV. Eliges qué columna es cada dato; si el código
  ya existe, se actualiza. **Exportar CSV** devuelve la base para abrirla en Excel.
- Los flujos (nombres y orden de etapas) están en `js/flows.js` por si tu proceso cambia.

## Seguridad y datos

- No capturar datos de pacientes ni información clínica: la app es para gestión de tareas.
- Acceso con usuario y contraseña; registro de nuevos usuarios desactivado.
- Cada registro está protegido por políticas RLS en la base de datos (solo su dueño lo lee o modifica).
- Los borrados son lógicos (se marcan como eliminados) para que se reflejen en todos los dispositivos.
- Respaldo manual en JSON desde Ajustes; se puede volver a importar.
