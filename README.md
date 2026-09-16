# El Palmar — Digital Catalog MVP

Catálogo digital de Distribuidora El Palmar construido con React + Vite.

## Incluye
- Catálogo de productos desde `public/data/products-list.json`.
- 3.764 productos del JSON descargado desde el endpoint.
- Imágenes de producto cuando el campo `imagen` está disponible.
- Búsqueda por producto, familia, categoría y código de barras.
- Filtros por categoría y familia.
- Carrito con cantidades editables y límite según stock.
- Descuento por volumen para hielo.
- Pedido por WhatsApp.
- Modal de despacho y medios de pago.
- Diseño responsive para desktop, tablet y celular.

## Ejecutar localmente

```bash
npm install
npm run dev
```

Para probarlo desde el celular en la misma Wi-Fi:

```bash
npm run dev -- --host
```

Luego abre en el celular la URL `Network` que muestre Vite.

## Build de producción

```bash
npm run build
npm run preview
```

## Actualizar productos

Reemplaza `public/data/products-list.json` por una nueva descarga del endpoint. La estructura esperada es la respuesta con un arreglo `productos` y campos como `id`, `nombre`, `precio`, `stock`, `codigoBarras`, `categoria`, `familia` e `imagen`.
