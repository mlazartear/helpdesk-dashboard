# Helpdesk Dashboard

Dashboard de estadísticas de tickets del Helpdesk de Accusys — cantidad de casos
creados por mes y por cliente, a partir de la Bandeja de Equipo.

🔗 **Demo:** https://mlazartear.github.io/helpdesk-dashboard/

## Cómo actualizar los datos

1. Entrar a `https://helpdesk.accusys.com.ar/bandejas/equipo`, ajustar el rango de
   fechas deseado (Desde / Hasta) y exportar a Excel (botón verde).
2. Correr el script de conversión (requiere `pip3 install openpyxl`):

   ```bash
   python3 helpdesk_stats.py
   # o indicando el archivo manualmente:
   python3 helpdesk_stats.py --archivo /ruta/al/export.xlsx
   ```

   Esto genera/actualiza `helpdesk_data.json` en esta carpeta.

3. Ver el resultado localmente:

   ```bash
   python3 -m http.server 8090
   # abrir http://localhost:8090/
   ```

4. Subir los cambios a GitHub (se refleja automáticamente en GitHub Pages):

   ```bash
   git add helpdesk_data.json
   git commit -m "Actualizar datos del dashboard"
   git push
   ```

## Qué muestra

- KPIs: total de tickets, mes actual vs. anterior, cliente con más casos, promedio mensual
- Casos creados por mes, apilados por cliente
- Ranking de casos por cliente
- Distribución por tipo y por estado de ticket
- Tabla detallada: ingresos por mes × cliente

El cliente **Accusys** se excluye por defecto (tickets internos, no de clientes).
Para incluirlo: `python3 helpdesk_stats.py --incluir-accusys`.
