function sincronizar() {
  var SB_URL = 'https://wnnwgcatjsbcqekktvmy.supabase.co';
  var SB_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndubndnY2F0anNiY3Fla2t0dm15Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODc2ODA5OTksImV4cCI6MjEwMzI1Njk5OX0.BKpREaJdjv4V9kMZVfCAw4I4GAEX2aUYeZQtULlVzmw';

  var url = SB_URL + '/rest/v1/meeting_notes?select=*,meeting_items(*)&estado=eq.finalizada&order=creado_en.desc';
  var options = {
    method: 'get',
    headers: {
      'apikey': SB_KEY,
      'Authorization': 'Bearer ' + SB_KEY
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notas') ||
              SpreadsheetApp.getActiveSpreadsheet().insertSheet('Notas');
  sheet.clear();

  var headers = ['Fecha', 'Vendedor', 'Nombre', 'Razón Social', 'Nombre Fantasía', 'SKU', 'Compra actualmente a', 'Volumen (kg/mes)', 'Precio ($/kg)', 'Monto mensual estimado', 'Nota', 'Importancia'];
  sheet.appendRow(headers);

  var impLabel = { alto: 'ALTO', medio: 'MEDIO', bajo: 'BAJO' };

  data.forEach(function(nota) {
    var fecha = new Date(nota.creado_en);
    var items = (nota.meeting_items && nota.meeting_items.length > 0) ? nota.meeting_items : [{}];
    items.forEach(function(item) {
      sheet.appendRow([
        fecha,
        nota.vendedor || '',
        nota.cliente || '',
        nota.razon_social || '',
        nota.nombre_fantasia || '',
        item.sku || '',
        item.proveedor_actual || '',
        item.volumen_kg || '',
        item.precio_clp || '',
        item.monto_mensual || '',
        nota.nota || '',
        nota.importancia ? impLabel[nota.importancia] : ''
      ]);
    });
  });

  sheet.autoResizeColumns(1, headers.length);
}
