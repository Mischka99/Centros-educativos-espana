# Firma del instalador con SignPath

El instalador `CentrosEducativos.msi` se firma gratis con [SignPath Foundation](https://signpath.org), que firma proyectos de código abierto. La firma la hace SignPath.io a partir del .msi que compila GitHub Actions desde este repositorio ([`.github/workflows/windows.yml`](../.github/workflows/windows.yml)), y cada firma se aprueba a mano.

Con el .msi firmado, el Control inteligente de aplicaciones de Windows 11 deja de bloquearlo. Todo lo que va dentro ya está firmado: el Python oficial lo firman la Python Software Foundation y Microsoft, y el instalador no tiene ejecutables propios.

## 1. Antes de solicitarlo

- **Verificación en dos pasos** en GitHub: Settings > Password and authentication > Two-factor authentication. SignPath la exige a todo el equipo, también en su web.
- **Una versión publicada con el .msi**, aunque vaya sin firmar: SignPath exige que el proyecto ya se distribuya en la forma que se va a firmar. Al publicar una versión (sección 4), la compilación automática adjunta el .msi.
- **La «Code signing policy» y la política de privacidad** en la página del proyecto: ya están en el [README](../README.md#code-signing-policy).

## 2. Solicitud

Rellena el formulario de [signpath.org/apply](https://signpath.org/apply) con estos datos:

| Dato | Valor |
|---|---|
| Proyecto | Centros educativos |
| Repositorio | https://github.com/Mischka99/Centros-educativos-espana |
| Licencia | GPL-3.0 |
| Descripción | Programa para Windows que busca y filtra los centros educativos de España, con su teléfono, email y mapa, y los descarga a Excel. Usa datos públicos del Ministerio de Educación y de las comunidades autónomas. |
| Página de descarga | https://github.com/Mischka99/Centros-educativos-espana/releases |
| Qué se firma | `CentrosEducativos.msi` (instalador de Windows Installer, sin acciones personalizadas). Los ejecutables de Python que lleva dentro ya vienen firmados por la Python Software Foundation y Microsoft, y no se vuelven a firmar. |
| Cómo se compila | GitHub Actions, en runners de GitHub, desde `.github/workflows/windows.yml` |
| Code signing policy | https://github.com/Mischka99/Centros-educativos-espana#code-signing-policy |

SignPath revisa la solicitud y responde por email.

## 3. Cuando la aprueben

SignPath Foundation crea la organización y el proyecto en [app.signpath.io](https://app.signpath.io). Allí:

1. **Trusted build systems:** añade «GitHub.com» y enlázalo al proyecto.
2. **Artifact configuration:** deja esta como la predeterminada. La raíz es `zip-file` porque GitHub entrega el .msi dentro de un zip; los valores de `subject` y `author` son los que GitHub muestra en cada compilación, en el aviso «Metadatos del .msi para SignPath».
   ```xml
   <?xml version="1.0" encoding="utf-8"?>
   <artifact-configuration xmlns="http://signpath.io/artifact-configuration/v1">
     <zip-file>
       <msi-file path="CentrosEducativos.msi" subject="Centros educativos" author="Mischka99">
         <authenticode-sign/>
       </msi-file>
     </zip-file>
   </artifact-configuration>
   ```
3. **Signing policy:** usa la de publicación (normalmente `release-signing`), con aprobación manual y tú como aprobador.
4. **API token:** en tu perfil > API Tokens, crea uno con el rol «Submitter» en el proyecto. Cópialo, porque no se vuelve a mostrar.
5. **Aplicación de GitHub (opcional):** instala [github.com/apps/signpath](https://github.com/apps/signpath) en el repositorio. Solo sirve para los registros de auditoría.

Después, en GitHub, ve a Settings > Secrets and variables > Actions y añade:

| Tipo | Nombre | Valor |
|---|---|---|
| Secret | `SIGNPATH_API_TOKEN` | El token del paso 4 |
| Variable | `SIGNPATH_ORGANIZATION_ID` | El identificador de la organización (está en la dirección de app.signpath.io) |
| Variable | `SIGNPATH_PROJECT_SLUG` | El identificador del proyecto |
| Variable | `SIGNPATH_SIGNING_POLICY_SLUG` | El de la política de publicación (`release-signing`) |

Mientras `SIGNPATH_ORGANIZATION_ID` esté vacía, las versiones se publican con el .msi sin firmar.

## 4. Publicar una versión

1. En GitHub, ve a Releases > Draft a new release.
2. Crea la etiqueta `vX.Y.Z` (por ejemplo `v1.2.0`) sobre `main`, pon el título y la descripción, y pulsa «Publish release». No hace falta adjuntar nada.
3. La compilación automática monta el zip y el .msi, los prueba y envía el .msi a SignPath.
4. SignPath te avisa por email: revisa la solicitud y apruébala. La compilación espera hasta 5 horas.
5. Al aprobarla, el .msi firmado y el zip aparecen en la versión.

Para comprobar la firma, en Windows haz clic derecho en el .msi > Propiedades > Firmas digitales. El firmante es «SignPath Foundation».
