# GymaholicMx — Export del proceso de picking/taller

Checklist de archivos incluidos en este paquete:

- /prompts/prompts.txt          — system prompt del trigger nocturno + few-shots/catálogo (sin keys ni nombres reales)
- /logs/conversations_logs.txt  — 9 conversaciones reales donde se procesó/validó un pedido (anonimizadas)
- /samples/sample_order_1.json  — MercadoLibre, pedido simple
- /samples/sample_order_2.json  — MercadoLibre, caso de venta duplicada (Paso 0.5)
- /samples/sample_order_3.json  — Shopify, KITFUNC desglosado
- /samples/sample_order_4.json  — MercadoLibre, SETBAZOTE (EMPLAYADO)
- /samples/sample_order_5.json  — TikTok (hipotético, no activo hoy) + escalación a revisión manual
- /scripts/send_script.txt      — pseudocódigo Python del flujo leer->armar->entregar (sin keys)
- /specs/taller_spec.txt        — mecanismo real de entrega (texto plano por Gmail/Telegram), sin API real hoy
- /docs/rules.md                — validaciones y reglas de negocio antes de mandar a taller
- /deploy/deploy.md             — trigger de Claude Code Remote, cron, dependencias reales (no hay servidor propio)

Notas importantes sobre el contenido

1. Ninguna credencial real está incluida (token de Telegram, API keys, contraseñas) — todo aparece como placeholder.
2. Los nombres/emails/teléfonos de clientes son ficticios o redactados — ver el aviso dentro de cada archivo.
3. Varias de las entradas documentan la realidad actual (no hay servidor/VPS propio ni integración API con taller). Ver `deploy/deploy.md` y `taller_spec.txt` para más detalles.
4. `conversations_logs.txt` solo trae 9 conversaciones; está explicado en su cabecera.

Instrucciones para añadir secrets (cuando lo necesites)
- Ve a Settings → Secrets and variables → Actions
- Añade los siguientes secrets (valores los pones tú, no subas credenciales en PR):
  - SMTP_HOST
  - SMTP_PORT
  - SMTP_USER
  - SMTP_PASS
  - EMAIL_TO
  - TELEGRAM_BOT_TOKEN
  - TELEGRAM_CHAT_ID
  - (Opcionales) SHOPIFY_ADMIN_TOKEN, SHOPIFY_STORE, ML_CLIENT_ID, ML_CLIENT_SECRET, ML_REFRESH_TOKEN

Si quieres que yo continúe pegando el script principal ahora, responde “siguiente” y te doy scripts/gymaholic_picking.py listo para pegar.
