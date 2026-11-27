# Publicar en la Microsoft Store

La versión de la Microsoft Store es el mismo programa en un paquete `.msix`. Microsoft lo firma al publicarlo, así que el Control inteligente de aplicaciones de Windows 11 no lo bloquea, y la Tienda lo instala y lo actualiza. Publicar es gratis.

La compilación automática ([`.github/workflows/windows.yml`](../.github/workflows/windows.yml)) genera el `.msix` en cada versión y lo prueba en Windows: lo instala, abre el programa y lo desinstala. Lo que hay que hacer a mano es esto:

1. Crear la cuenta de desarrollador (una vez).
2. Reservar el nombre y copiar la identidad del paquete en GitHub (una vez).
3. Publicar una versión en GitHub y descargar el `.msix`.
4. Crear el envío en Partner Center.

Partner Center puede salir en español o en inglés: los nombres de cada sección van en los dos idiomas.

## 1. Cuenta de desarrollador

Entra en [storedeveloper.microsoft.com](https://storedeveloper.microsoft.com), pulsa «Get started for free» y elige el tipo de cuenta. Las dos son gratis:

- **Individual:** publicas a tu nombre. Se verifica con tu documento de identidad y un selfi. Microsoft la reserva para quien no publica en relación con su negocio o profesión.
- **Empresa (Company):** publicas a nombre de una empresa registrada. Se verifica con documentos oficiales de la empresa: de constitución, de un registro público o fiscales.

Después no se puede pasar de individual a empresa: habría que crear otra cuenta.

El nombre de editor que elijas es el que verá la gente en la Tienda (por ejemplo, «Mischka99»).

## 2. Nombre e identidad del paquete

1. En [Partner Center](https://partner.microsoft.com/dashboard), en Aplicaciones y juegos (Apps and games), crea un producto nuevo de tipo **Aplicación MSIX o PWA** (MSIX or PWA app).
2. Reserva el nombre **Centros educativos**. Si ya está cogido, reserva otro, por ejemplo «Centros educativos de España».
3. Abre la aplicación y ve a Administración de productos > Identidad del producto (Product management > Product identity). Copia estos valores en GitHub, en el repositorio: Settings > Secrets and variables > Actions > pestaña Variables > New repository variable.

| Partner Center | Variable de GitHub |
|---|---|
| `Package/Identity/Name` | `STORE_IDENTITY_NAME` |
| `Package/Identity/Publisher` (empieza por `CN=`) | `STORE_PUBLISHER` |
| `Package/Properties/PublisherDisplayName` | `STORE_PUBLISHER_DISPLAY_NAME` |
| El nombre reservado, solo si no es «Centros educativos» | `STORE_DISPLAY_NAME` |

Cópialos tal cual, con mayúsculas, espacios y signos: si no coinciden, Partner Center rechaza el paquete. Son variables, no secretos: van dentro del paquete y no son confidenciales.

Sin estas variables, la compilación usa una identidad de prueba que sirve para probar el paquete, pero no para subirlo.

## 3. El paquete

1. Publica una versión en GitHub, como siempre: Releases > Draft a new release, etiqueta `vX.Y.Z` (por ejemplo `v1.2.0`) y Publish release.
2. En Actions, abre la ejecución «Paquete de Windows» de esa versión y espera a que termine.
3. En el resumen de la ejecución aparece el apartado «Paquete de la Microsoft Store». Comprueba que la versión es la de la etiqueta y que la identidad es la de Partner Center, no la de prueba.
4. Abajo, en Artifacts, descarga **CentrosEducativos.msix**.

Usa siempre el `.msix` de una versión publicada. Los de los demás cambios llevan la versión `0.0.N` y no sirven para la Tienda.

El `.msix` va sin firmar a propósito: lo firma la Tienda. Windows no instala un `.msix` sin firmar, pero no hace falta probarlo en tu PC: la compilación ya lo instala y lo abre en Windows, firmado con un certificado de prueba.

## 4. Envío

En la página de la aplicación, pulsa «Iniciar el envío» (Start submission) y completa cada sección.

**Precios y disponibilidad (Pricing and availability)**

- Mercados: todos, o solo España.
- Visibilidad: pública.
- Precio: gratis.

**Propiedades (Properties)**

- Categoría: Productividad.
- Directiva de privacidad (privacy policy): `https://github.com/Mischka99/Centros-educativos-espana#privacidad`
- Sitio web: `https://github.com/Mischka99/Centros-educativos-espana`
- Contacto de soporte: `https://github.com/Mischka99/Centros-educativos-espana/issues`, o tu email.

**Clasificación por edades (Age ratings)**

Es el cuestionario de IARC. Es una aplicación, no un juego. No tiene violencia, compras, chat ni contenido de otros usuarios, no comparte la ubicación y no es un navegador: los enlaces se abren en el navegador del usuario.

**Paquetes (Packages)**

- Arrastra `CentrosEducativos.msix`.
- En la disponibilidad por familia de dispositivos debe quedar solo el escritorio (Windows Desktop).
- Es normal el aviso de que la capacidad restringida `runFullTrust` necesita aprobación: se justifica en «Opciones de envío».

**Descripciones de Store (Store listings)**, en español:

- **Descripción:**

  > Directorio de los 33.555 centros educativos no universitarios de España, con su teléfono, email, web y dirección. Los datos salen de fuentes oficiales: el Ministerio de Educación y los directorios de cada comunidad autónoma.
  >
  > Busca y filtra por comunidad, provincia, municipio, tipo de centro (escuela infantil, primaria, secundaria, FP, adultos, idiomas, música…), titularidad o si tiene email. Cada centro tiene su ficha con un mapa y botones para copiar el email o el teléfono, escribir un email o llamar.
  >
  > Marca los centros que te interesen y descárgalos a Excel, o descarga la lista filtrada entera.
  >
  > La lista va incluida, así que el programa abre al instante y funciona sin conexión. Mientras está abierto, se actualiza solo cada semana con los datos oficiales y avisa de los centros nuevos y de los emails nuevos.
  >
  > No tiene publicidad ni recoge datos tuyos. Es software libre (GNU GPL v3).

- **Novedades de esta versión:** «Primera versión en la Microsoft Store».
- **Características del producto:**
  - 33.555 centros educativos de España con teléfono, email, web y dirección
  - Datos oficiales del Ministerio de Educación y de las comunidades autónomas
  - Filtros por comunidad, provincia, municipio, tipo de centro, titularidad y email
  - Ficha de cada centro con el mapa del Instituto Geográfico Nacional
  - Descarga a Excel de los centros marcados o de la lista filtrada
  - Se actualiza sola cada semana y avisa de los centros y emails nuevos
  - Funciona sin conexión con la lista incluida
  - Sin publicidad ni recogida de datos
- **Capturas de pantalla:** [`docs/captura.png`](../docs/captura.png) (1400 × 880, el mínimo es 1366 × 768).
- **Términos de búsqueda:** colegios, institutos, escuelas, centros educativos, directorio de colegios, email de colegios, Ministerio de Educación.
- **Copyright:** «Copyright © 2026 Mischka99».
- **Términos de licencia adicionales:** «Software libre con licencia GNU GPL v3: https://www.gnu.org/licenses/gpl-3.0.html. Código fuente: https://github.com/Mischka99/Centros-educativos-espana».

**Opciones de envío (Submission options)**

En «Capacidades restringidas» (Restricted capabilities), la justificación de `runFullTrust`. El equipo de certificación lee en inglés:

> This is a Win32 desktop application (Python 3.13 and Tkinter) packaged as MSIX, so it needs runFullTrust to run. It doesn't install services or drivers, and it only writes its own data (in the package's app data folder) and the Excel files the user chooses to save. Links, emails and phone numbers open in the user's default apps only when the user clicks them.

En «Notas para la certificación» (Notes for certification):

> Free, open-source desktop app. No account or sign-in is needed. It ships with the full list of Spanish schools and works offline; with an internet connection it updates the list weekly from public government open data and shows maps from the Spanish National Geographic Institute (IGN). To test it: choose a region on the left, select a school to see its details and map, mark it and click "Descargar lista" to export to Excel. Source code: https://github.com/Mischka99/Centros-educativos-espana

Por último, pulsa «Enviar a la Store» (Submit to the Store). La certificación tarda hasta tres días laborables y Partner Center avisa por email. Cuando se publique, el enlace de la Tienda aparece en Identidad del producto: añádelo a la sección «Descargar y usar» del [README](../README.md).

## Actualizaciones

1. Publica la versión nueva en GitHub (con una etiqueta mayor que la anterior) y descarga su `.msix`, como en el paso 3.
2. En Partner Center, abre la aplicación y pulsa «Actualizar» (Update) junto al último envío.
3. En Paquetes, sube el `.msix` nuevo y quita el anterior. Cambia «Novedades de esta versión» y envíalo.

Quien tenga el programa instalado desde la Tienda recibe la actualización automáticamente. La identidad del paquete no cambia nunca: si cambiaras las variables de GitHub, Partner Center rechazaría el paquete.
