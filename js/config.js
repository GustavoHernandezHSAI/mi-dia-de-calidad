// ============================================================
//  CONFIGURACIÓN — es el único archivo que debes editar.
// ============================================================
//  Si dejas SUPABASE_URL vacío, la app funciona en "modo local":
//  los datos se guardan solo en este navegador (útil para probar).
// ============================================================

export const CONFIG = {
  // Supabase → Project Settings → API → "Project URL"
  SUPABASE_URL: 'ofmiyejmkxtrvallmste',

  // Supabase → Project Settings → API → "anon public" key
  // (es una clave pública; la seguridad la dan las políticas RLS del archivo schema.sql)
  SUPABASE_ANON_KEY: 'sb_publishable_KrOwE0fSaD_gAojXtTIyAA_MZv6Xjhc',

  // Outlook (opcional). Los entrega TI al registrar la app en Microsoft Entra ID.
  // Déjalos vacíos hasta tener la autorización.
  MS_CLIENT_ID: '',
  MS_TENANT_ID: '', // ID del directorio (tenant) del hospital

  // Zona horaria para leer el calendario de Outlook
  TIMEZONE: 'America/Mexico_City',
};
