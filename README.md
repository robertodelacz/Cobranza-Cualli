# Cobranza Preventiva v2 · Financiera Cualli

Plataforma de avisos de pago (aviso 1 a 5 días hábiles y aviso 2 a 1 día hábil del pago). Sigue siendo Apps Script + Google Sheets (libro "Maestro"), pero ahora trabaja por **tandas**, con bitácora, reenvío, ajustes de saldo, avisos en Chat y roles.

## Cómo funciona (resumen)
Nueve secciones en el menú lateral, en lenguaje de todos los días (diseño con el logo de Cualli, amarillo y grises):
1. **Hoy** es una guía de tres pasos: *Sube los reportes → Revisa lo pendiente → Envía los avisos*. Incluye la regla de días hábiles (a quién se avisa hoy), cifras de la tanda, las **dos tarjetas de carga** y un resumen de cartera. Un vencimiento en sábado, domingo o inhábil CNBV se paga y se avisa el siguiente día hábil (el correo dice esa fecha).
2. **Reportes**: **una tarjeta para el Rep1 y otra para el Rep9**, cada una con su zona de arrastre, el filtro que hay que usar al descargar (con botón de copiar las fechas), estado "al día / vencido", **Ver contenido** del archivo guardado (con búsqueda y descarga a Excel) e historial de todas las cargas. Sin casillas: el archivo se revisa solo y se guarda; solo se detiene, con un botón claro, si ve algo raro. Si se suelta un Rep9 en la tarjeta del Rep1 (o al revés), lo avisa y ofrece *Usarlo como Rep9*.
3. **Avisos**: *Listos para enviar* (ya marcados), *Revisa antes de enviar*, *No se pueden enviar todavía* (con el botón que lo resuelve), *Más adelante* y *Ya atendidos*. Al tocar una fila se abre **"la cinta"**: el cálculo del monto línea por línea, con la columna del reporte de donde sale cada número (Rep1 col. D, Rep9 col. K, P+Q, R+S+T+U…), la fórmula de los moratorios y el total que lleva el correo; también la vista previa del correo. Barra fija al fondo para enviar lo seleccionado.
4. **Hoja de trabajo**: réplica de la hoja de cálculo de la v1 dentro de la plataforma. Pestaña *Cuotas del Rep1* (columnas A–Q, las calculadas en amarillo, totales por moneda, filtros, búsqueda, descarga a Excel; tocar una fila abre su cinta), pestaña *Saldos vencidos del Rep9* (con las dos columnas calculadas: intereses vencidos con IVA y moratorios con IVA) y pestaña *Cómo se calcula* (de dónde sale cada columna y las reglas vigentes).
5. **Cartera**: todas las líneas del Rep9 con saldo total, vencido y próxima cuota; filtros y búsqueda. Al abrir un cliente: saldos, cuotas, avisos enviados, ajustes, edición de tasa, correos y STP, y la calculadora **"¿Cuánto debe al día…?"** (también en forma de cinta).
6. **Historial de envíos**: cada envío agrupado por día, con reenvío (motivo obligatorio). La bitácora es la fuente de verdad: no se duplican avisos.
7. **Ajustes de saldo** (pago que llegó, disposición, corrección): cambian el monto mientras siga vigente el mismo Rep9; si el monto es ≥ umbral y hay aprobadores, otra persona debe aprobar.
8. **Datos de clientes** (qué le falta a cada línea) y 9. **Configuración** (agrupada, con interruptores; incluye Chat, recordatorios, modo prueba y días inhábiles).
**Chat**: recordatorio de descarga, escalación si no hay reportes y resumen de cada envío.

## Instalación (una sola vez)
1. En el proyecto de Apps Script actual: copia **todos** los archivos de la raíz de este repositorio (los `.html` se llaman `index`, `styles`, `app_js`; reemplaza `appsscript.json`). Borra los `.gs` viejos que se repitan para que no haya funciones duplicadas.
2. Ejecuta **`inicializarV2`** una vez y acepta permisos. Crea las hojas Cortes, Ajustes, Calendario_Inhabiles, Usuarios y Cambios_Catalogo, agrega columnas a la bitácora, completa Config y pone la zona horaria America/Mexico_City.
3. Hoja **Usuarios** (correo, nombre, rol: COORDINADORA, GERENTE, AUDITORIA, ADMIN). Vacía = todos son administrador.
4. En **Configuración** de la plataforma: pega el webhook del espacio de Chat, "Enviar mensaje de prueba", y "Instalar recordatorios". Elimina los triggers viejos de la v1.
5. Completa el catálogo **Correos** con los correos reales de clientes y la hoja de **Tasas** (la sección Datos de clientes lista lo que falta).
6. Primera semana: `MODO_PRUEBA = TRUE` y `MODO_PRUEBA_DESTINO` con un buzón interno; los correos no llegan a clientes ni consumen avisos. Después ponlo en FALSE.
7. Implementar → Nueva versión de la aplicación web (ejecutar como tú, acceso dominio).

## Supuestos que conviene confirmar
- **Cadencia de tanda**: por defecto una tanda por día hábil (`TANDA_DIAS_HABILES=1`). Si la gerencia prefiere descargar cada 2 o 3 días hábiles, se cambia ese parámetro y la ventana se recalcula.
- **Moratorios**: se proyectan con fecha nominal (como el motor v1): capital vencido × tasa contrato × 2 ÷ 360 × días desde el corte del Rep9 al vencimiento. Parámetro `BASE_MORATORIOS` (NOMINAL / EFECTIVA). Validar con Finanzas.
- **Antigüedad de reportes**: el Rep9 debe ser de hoy y el Rep1 de la tanda en curso; si no, la cuota se bloquea (ajustable en Config).
- **Inhábiles**: 2026 = calendario CNBV verificado. 2027–2028 se calculan por regla; validar cuando la CNBV publique el suyo (se puede corregir en Configuración).
- **Tasas**: no hay reporte que las traiga; siguen siendo catálogo manual.
- **Línea en USD 101335 (JKLD)**: en el Rep9 de agosto del libro no aparece; si el reporte real la trae, saldrá completa. Revisar el archivo original.
- **Calculadora "¿Cuánto debe al día…?"**: es una estimación con la fórmula del motor (cuotas del Rep1 hasta esa fecha + vencido + moratorios desde el corte). Conviene contrastarla con la hoja "Calculadora saldos" de Sheets antes de usarla con clientes.
- **Logo**: la barra lateral carga `https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png` (el de otros desarrollos de Cualli); si no carga, muestra una "C". Cambia la URL en `index.html` si el logo vive en otro lado.
- **Hoja de trabajo**: las letras de columna (A–Q) son las de la plataforma; confirmar que coinciden con tu hoja de Sheets si quieres compararlas celda por celda.
- El envío es manual con selección; el envío automático existe (`ENVIO_AUTOMATICO`) pero está apagado.

## Pruebas incluidas (carpeta `pruebas/` (necesitan `maestro.json`, un volcado del libro, que no se versiona))
102 pruebas de Node contra los datos reales del libro (calendario, motor, flujo de envío, ajustes, Chat, roles, cartera, ficha de cliente y hoja de trabajo/traza del cálculo) y una prueba de interfaz con Playwright (43 verificaciones: dos cargas separadas y reporte en tarjeta equivocada, enviar, cinta, hoja de trabajo con sus totales, cartera, historial, ajustes, móvil). No se despliegan en Apps Script.
