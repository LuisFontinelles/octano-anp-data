# Política de Privacidad — Octano

**Última actualización: 5 de julio de 2026**

Octano es una aplicación iOS gratuita, mantenida por un desarrollador independiente, que ayuda a los conductores a elegir dónde repostar usando datos públicos oficiales. Esta política explica, conforme a la Ley General de Protección de Datos de Brasil (Ley n.º 13.709/2018 — LGPD), qué datos se tratan y cómo.

**Responsable:** Luis Fontinelles — contacto: **[fontinelles.com/contact](https://fontinelles.com/contact)**

## El resumen honesto

Octano **no tiene cuenta, no tiene inicio de sesión, no usa publicidad y no vende datos**. El contenido que creas (precios, votos) permanece en tu dispositivo. Dos cosas salen del dispositivo: tu **ubicación aproximada**, enviada a Google para buscar estaciones cercanas, y **datos de uso anónimos**, enviados a Firebase (Google) para entender cómo se usa la app y mejorarla. Las compras opcionales de apoyo las procesa **Apple** — nunca vemos los datos de tu tarjeta.

## 1. Datos que trata la app

### 1.1 Ubicación (con tu permiso)
- **Para qué:** encontrar estaciones cercanas, ordenar la lista por distancia y trazar rutas.
- **Cómo:** procesada en el dispositivo. Se envía **solo** a la API de Google Places (búsqueda de estaciones) y a Apple (MKDirections, cálculo de ruta), que actúan como encargados de esas funciones.
- **Base legal:** consentimiento (lo autorizas en iOS y puedes revocarlo en cualquier momento en Ajustes → Privacidad → Localización).
- La app **no** almacena historial de ubicación en ningún servidor del desarrollador.

### 1.2 Datos de uso y diagnóstico (analytics)
- **Qué:** eventos de navegación (pantallas visitadas, tiempo en cada pantalla, toques en botones como buscar, iniciar ruta, abrir estación), modelo del dispositivo y versión de iOS, región aproximada (a partir de la IP) y un **identificador seudónimo de la app** generado por Firebase. También recopilamos informes de **fallos (crash)** para corregir errores.
- **Para qué:** entender cómo se usa la app, priorizar mejoras y corregir problemas. **No** lo usamos para publicidad y **no** recopilamos el identificador de publicidad (IDFA).
- **Quién lo procesa:** **Firebase (Google LLC)** — Analytics, Crashlytics, Performance. Los datos se tratan bajo la política de Google.
- **Base legal:** interés legítimo en la mejora del servicio; los datos son seudonimizados y agregados.
- **Notificaciones:** si aceptas recibir notificaciones, se usa un token de Firebase Messaging solo para entregarlas.

### 1.3 Compras en la app (contribuciones opcionales)
- La app ofrece **contribuciones únicas** ("propinas") para apoyar el proyecto. No desbloquean funciones.
- El **pago lo procesa Apple** (App Store). El desarrollador recibe solo la confirmación de la transacción — **nunca** el número de tarjeta ni datos financieros.

### 1.4 Datos que la app NO recopila
Nombre, correo, teléfono, contactos, fotos, datos de tarjeta e identificador de publicidad (IDFA). No hay inicio de sesión ni cuenta.

## 2. Servicios de terceros

| Servicio | Qué recibe | Política |
|---|---|---|
| Google Maps / Places (Google LLC) | Ubicación aproximada en las búsquedas; telemetría propia del SDK de mapas | [policies.google.com/privacy](https://policies.google.com/privacy) |
| Firebase (Google LLC) | Eventos de uso, diagnóstico/fallos y un identificador seudónimo de la app | [firebase.google.com/support/privacy](https://firebase.google.com/support/privacy) |
| Apple (App Store / StoreKit) | Procesamiento de las compras de apoyo | [apple.com/privacy](https://www.apple.com/privacy/) |
| Apple (MapKit/MKDirections) | Coordenadas de origen/destino para el cálculo de rutas | [apple.com/privacy](https://www.apple.com/privacy/) |

## 3. Datos públicos mostrados por la app

La app reproduce bases **públicas y oficiales** de la ANP (fiscalización, registro de revendedores y precios), de los **Procons estatales** y del sistema nacional **SINDEC/Senacon** (empresas sancionadas y reclamaciones fundamentadas), del **IPEM-SP** (surtidores certificados) y contenido de Google (reseñas, calificaciones), siempre con indicación de fuente y fecha.

- Los datos de **empresas** (CNPJ, razón social) se muestran tal como constan en las bases públicas.
- Los **CPF** (identificación fiscal de persona física) de revendedores individuales, cuando están presentes en las bases, son **enmascarados** por la app y el pipeline de datos — el documento completo no se redistribuye.
- Las bases de defensa del consumidor contienen empresas (personas jurídicas), recopiladas de portales de transparencia y datos abiertos (Ley de Acceso a la Información).
- Las reseñas mostradas son de autoría de usuarios de Google y siguen bajo responsabilidad de sus autores y de la plataforma de origen.

## 4. Tus derechos (art. 18 de la LGPD)

Puedes solicitar confirmación del tratamiento, acceso, corrección o eliminación de datos por nuestro canal de contacto: **[fontinelles.com/contact](https://fontinelles.com/contact)**. En la práctica: **desinstalar la app elimina los datos que creaste en el dispositivo**; el permiso de ubicación es revocable en los Ajustes de iOS; y puedes desactivar la personalización de anuncios/rastreo de iOS en cualquier momento.

Si eres **titular de datos presentes en las bases públicas** mostradas (p. ej., un revendedor persona física), puedes contactarnos para revisar la exhibición — y, respecto a las bases de origen, ejercer tus derechos ante la ANP, el Procon o el IPEM.

## 5. Niños y adolescentes

La app está destinada a conductores habilitados y no está dirigida a menores de edad.

## 6. Cambios

Esta política puede actualizarse; la versión vigente estará siempre en esta dirección, con la fecha en la parte superior.
