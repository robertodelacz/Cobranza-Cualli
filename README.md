# Cobranza Preventiva v2 · Financiera Cualli

Plataforma de avisos de pago (aviso 1 a 5 días hábiles y aviso 2 a 1 día hábil del pago). Sigue siendo Apps Script + Google Sheets (libro "Maestro"), pero ahora trabaja por **tandas**, con bitácora, reenvío, ajustes de saldo, avisos en Chat y roles.

## Cómo funciona (resumen)
1. **Hoy** dice si toca descargar el Rep1, con qué filtro de "Fecha Vencimiento" y qué avisos salen hoy. Un vencimiento en sábado, domingo o inhábil CNBV se paga y se avisa el siguiente día hábil (el correo dice esa fecha, no la anterior).
2. La coordinadora sube Rep1 y Rep9. La plataforma valida, muestra el rango real del archivo y pide confirmar antes de guardar. Cada carga queda como un *corte* con ID.
3. **Cola de envío**: cada cuota tiene un estado (Lista, Revisar, Bloqueada, Programada, Espera aviso 2, Completa, Vence hoy). Al abrir una fila se ve cómo se calculó el total.
4. Se envía por selección, en bloques; todo queda en **Bitácora** (la bitácora es la fuente de verdad, no se duplican avisos). Reenvío con motivo obligatorio.
5. **Ajustes de saldo** (pago aplicado, disposición, corrección) cambian el monto mientras siga vigente el mismo Rep9; si el monto es ≥ umbral y hay aprobadores, otra persona debe aprobar.
6. **Chat**: recordatorio de descarga, escalación si no hay reportes y resumen de cada envío.

## Instalación (una sola vez)
1. En el proyecto de Apps Script actual: copia **todos** los archivos de la raíz de este repositorio (los `.html` se llaman `index`, `styles`, `app_js`; reemplaza `appsscript.json`). Borra los `.gs` viejos que se repitan para que no haya funciones duplicadas.
2. Ejecuta **`inicializarV2`** una vez y acepta permisos. Crea las hojas Cortes, Ajustes, Calendario_Inhabiles, Usuarios y Cambios_Catalogo, agrega columnas a la bitácora, completa Config y pone la zona horaria America/Mexico_City.
3. Hoja **Usuarios** (correo, nombre, rol: COORDINADORA, GERENTE, AUDITORIA, ADMIN). Vacía = todos son administrador.
4. En **Configuración** de la plataforma: pega el webhook del espacio de Chat, "Enviar mensaje de prueba", y "Instalar recordatorios". Elimina los triggers viejos de la v1.
5. Completa el catálogo **Correos** con los correos reales de clientes y la hoja de **Tasas** (la sección Catálogos de la plataforma lista lo que falta).
6. Primera semana: `MODO_PRUEBA = TRUE` y `MODO_PRUEBA_DESTINO` con un buzón interno; los correos no llegan a clientes ni consumen avisos. Después ponlo en FALSE.
7. Implementar → Nueva versión de la aplicación web (ejecutar como tú, acceso dominio).

## Supuestos que conviene confirmar
- **Cadencia de tanda**: por defecto una tanda por día hábil (`TANDA_DIAS_HABILES=1`). Si la gerencia prefiere descargar cada 2 o 3 días hábiles, se cambia ese parámetro y la ventana se recalcula.
- **Moratorios**: se proyectan con fecha nominal (como el motor v1): capital vencido × tasa contrato × 2 ÷ 360 × días desde el corte del Rep9 al vencimiento. Parámetro `BASE_MORATORIOS` (NOMINAL / EFECTIVA). Validar con Finanzas.
- **Antigüedad de reportes**: el Rep9 debe ser de hoy y el Rep1 de la tanda en curso; si no, la cuota se bloquea (ajustable en Config).
- **Inhábiles**: 2026 = calendario CNBV verificado. 2027–2028 se calculan por regla; validar cuando la CNBV publique el suyo (se puede corregir en Configuración).
- **Tasas**: no hay reporte que las traiga; siguen siendo catálogo manual.
- **Línea en USD 101335 (JKLD)**: en el Rep9 de agosto del libro no aparece; si el reporte real la trae, saldrá completa. Revisar el archivo original.
- El envío es manual con selección; el envío automático existe (`ENVIO_AUTOMATICO`) pero está apagado.

## Pruebas incluidas (carpeta `pruebas/` (necesitan `maestro.json`, un volcado del libro, que no se versiona))
75 pruebas de Node contra los datos reales del libro (calendario, motor, flujo de envío, ajustes, Chat, roles) y una prueba de interfaz con Playwright. No se despliegan en Apps Script.
