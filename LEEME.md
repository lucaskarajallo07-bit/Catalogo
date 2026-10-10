# La Plaza Perfumería — tienda web

Sitio estático (HTML, CSS y JavaScript, sin servidor ni base de datos), pensado para celular y publicado en GitHub Pages. Todo lo que cambia (perfumes, precios, stock, fotos, envíos, cupones, banner) se administra desde el Excel.

## Qué hay en la carpeta

| Archivo | Para qué sirve |
| --- | --- |
| `index.html`, `styles.css`, `app.js` | La tienda. No hace falta tocarlos. |
| `Catalogo_La_Plaza.xlsx` | Tu panel de administración: productos, precios, stock, fotos, cupones, envíos, redes y banner. |
| `herramientas/excel_a_json.py` | Convierte el Excel en los datos de la tienda y genera las páginas para Google. |
| `data/catalogo.json` | Datos generados desde el Excel. No editar a mano. |
| `producto/`, `categoria/` | Una página por perfume y por categoría (para Google y para compartir). Se generan solas. |
| `img/` | Fotos de los perfumes. |

## Actualizar la tienda

1. Editá `Catalogo_La_Plaza.xlsx`.
2. En la carpeta de la tienda, corré:
   ```
   pip install openpyxl
   python herramientas/excel_a_json.py Catalogo_La_Plaza.xlsx
   ```
3. Subí al repo lo que cambió (`data/`, `producto/`, `categoria/`, y `img/` si agregaste fotos).

### Qué controla cada columna del Excel

- **Precio vacío**: esa presentación no se vende.
- **ml disponibles para decants** en 0: decants agotados (se ven, pero no se pueden comprar).
- **Botellas en stock** vacío: la botella aparece con «Consultar disponibilidad» y botón a WhatsApp.
- **Activo = No**: oculta el perfume.
- **Foto (archivo)**: una o varias, separadas por coma (`khamrah.jpg, khamrah-caja.jpg`). La primera es la principal.
- **Momento, Clima, Ocasión**: alimentan los filtros, el test y «Cuándo usarlo». Los completé a partir de tus descripciones; revisalos.
- **Intensidad y Duración**: vacíos por ahora. Cuando los cargues aparecen en la ficha, en los filtros y en el comparador.
- **Novedad / Más vendido**: muestran la etiqueta y la sección en el inicio. Marcalos solo si es real.
- **Oferta %**: descuento real; la tienda muestra el precio anterior tachado y la sección Ofertas.
- **Hoja Cupones**: código, tipo (porcentaje o monto fijo), valor, compra mínima y Activo. El campo de cupón aparece en el carrito solo si hay alguno activo. Los códigos quedan visibles para quien inspeccione la página: usalos para promos, no como claves.
- **Hoja Configuración**: WhatsApp, Instagram, TikTok, lugar de retiro, pagos, zonas de envío, URL del sitio y banner de promoción.

## Publicar en GitHub Pages

1. Subí todo el contenido de esta carpeta a la raíz del repositorio.
2. En GitHub: **Settings → Pages → Branch: main / root → Save**.
3. La tienda queda en `https://<tu-usuario>.github.io/<repo>/`.
4. Copiá esa dirección en «URL del sitio» del Excel y volvé a correr el script: se generan `sitemap.xml` y `robots.txt`, y los enlaces para compartir apuntan a la página de cada perfume.

## Cómo funciona el pedido

El cliente arma el carrito, completa sus datos y elige envío o retiro. Al confirmar, se abre WhatsApp con el pedido escrito (productos, presentaciones, cantidades, cupón, envío y total) dirigido a tu número. El cliente lo envía y vos confirmás. El sitio no guarda datos de clientes ni cobra en línea.

## Integraciones pendientes

| Qué | Estado | Qué hace falta |
| --- | --- | --- |
| Pago online | No integrado. Hoy se cobra en efectivo o transferencia por WhatsApp. | Contrato con una pasarela que opere en Paraguay (por ejemplo Bancard o Pagopar), sus credenciales y un servidor que firme los pagos: GitHub Pages no alcanza para eso. |
| Pedidos e inventario automáticos | El stock se actualiza a mano en el Excel. | Una base de datos y un servidor (o volver a tu MySQL con el bot) para descontar stock al vender. |
| Reseñas de clientes | Oculto. | Reseñas reales con permiso del cliente. |
| Más vendidos, novedades, ofertas, combos | Ocultos hasta que los marques en el Excel. | Marcar las columnas correspondientes. Los combos necesitan definir qué incluyen y su precio. |
| TikTok | Oculto. | Cargar el usuario en Configuración. |
| Analítica (visitas, carritos) | No integrada. | Una cuenta de Google Analytics u otra herramienta, con aviso de privacidad. |
