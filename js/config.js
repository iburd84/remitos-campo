// Configuración de la app. Lo único que hay que completar al instalar.
// Estos datos NO son secretos (el tenant y el client ID de una app SPA son públicos).
// Los datos de remitos, fotos, clientes y usuarios viven en SharePoint, nunca en este repo.

export const CONFIG = {
  version: '1.7.0',

  // Microsoft Entra (registro de la app, tipo SPA)
  tenantId: 'ad9b3575-dfa1-4374-83ee-797ea46ee8da',
  clientId: '6accec4b-39ef-40c6-9bd8-f861b007bb38',
  scopes: ['User.Read', 'Sites.ReadWrite.All', 'Files.ReadWrite.All'],
  // Solo para la pantalla de Configuración: crear listas, columnas y biblioteca
  scopesConfiguracion: ['Sites.Manage.All'],

  // Sitio de SharePoint del piloto, ej. host 'alongersrl.sharepoint.com', ruta '/sites/RemitosCampo'
  sitio: {
    host: 'alongercomar.sharepoint.com',
    ruta: '/sites/RemitosCampo',
  },

  // Nombres de listas y biblioteca que crea la pantalla de Configuración
  listaRemitos: 'Remitos Campo',
  listaUsuarios: 'Usuarios Remitos',
  biblioteca: 'Remitos Escaneados',
  archivoConfig: '_config/clientes.json',
};

// Sin Entra configurado (o con ?demo=1) la app corre en modo demostración:
// guarda en el navegador y simula la planilla, para probar la interfaz.
export const MODO_DEMO =
  CONFIG.clientId.startsWith('COMPLETAR') || CONFIG.sitio.host.startsWith('COMPLETAR') ||
  new URLSearchParams(location.search).has('demo');
