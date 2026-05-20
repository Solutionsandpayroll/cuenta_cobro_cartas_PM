# Generador de Cartas — Punto Medical

Aplicación web interna desarrollada para **Punto Medical** que automatiza la generación de **Cartas Cuenta de Cobro** en formato Word (`.docx`) y PDF, a partir de datos de empleados y tarifas ReteICA cargados desde archivos Excel.

---

## Características

- **Carga de Excel**: importa la base de empleados y la tabla de tarifas ReteICA desde archivos `.xlsx`/`.xls`
- **Búsqueda de empleado**: campo con autocompletado y datos editables en tiempo real (NIT, nombre, teléfono, dirección)
- **Editor de información**: formulario para fecha, tipo de documento, concepto, valor, retenciones (ICA y Fuente)
- **Cálculo automático de retenciones**:
  - `@@VALORICA@@` → tarifa/1000 × valor bruto
  - `@@VALORFUENTE@@` → porcentaje × valor bruto
  - `@@TOTALPAGAR@@` → valor bruto − ICA − Fuente
- **Valor en letras**: convierte el valor numérico a palabras en español (formato colombiano)
- **Bloque condicional ICA**: si no aplica ReteICA, el párrafo completo se elimina del documento
- **Generación Word** (`.docx`): reemplaza marcadores en la plantilla con los datos del formulario
- **Generación PDF**: renderiza el `.docx` con `docx-preview`, captura con `html2canvas` y ensambla con `jsPDF`
- **Consecutivo en la nube**: el número de carta se sincroniza con JSONBin.io — funciona en cualquier computador
- **Animaciones de entrada** en todos los elementos de la interfaz
- **100% responsive**

---

## Stack Tecnológico

| Librería | Uso |
|----------|-----|
| React 18 + Vite 5 | Framework principal |
| SheetJS (`xlsx`) | Lectura de archivos Excel |
| PizZip | Manipulación del `.docx` (descomprimir/recomprimir) |
| `docx-preview` | Renderizado del Word en DOM para captura PDF |
| `html2canvas` | Captura del DOM como imagen |
| `jsPDF` | Generación del archivo PDF |
| JSONBin.io | Persistencia del consecutivo en la nube |

---

## Estructura del Proyecto

```
├── public/
│   ├── Carta cuenta de cobro.docx   ← Plantilla Word con marcadores @@
│   └── Base ReteICA.xlsx            ← Tabla de tarifas (se carga automáticamente)
├── src/
│   ├── App.jsx                      ← Toda la lógica y UI de la aplicación
│   └── App.css                      ← Estilos y animaciones
├── .env                             ← Variables de entorno (no se sube a Git)
├── .env.example                     ← Plantilla de variables de entorno
└── .gitignore
```

---

## Marcadores en la Plantilla Word

La plantilla `public/Carta cuenta de cobro.docx` debe contener estos marcadores:

| Marcador | Valor generado |
|----------|---------------|
| `@@FECHA@@` | Fecha en formato "21 de abril de 2026" |
| `@@NOMBRECOMPLETO@@` | Nombre del empleado |
| `@@NUMERODOCUMENTO@@` | NIT / Cédula |
| `@@TELEFONO@@` | Teléfono |
| `@@DIRECCION@@` | Dirección |
| `@@DOCUMENTO@@` | Tipo de documento |
| `@@CONCEPTO@@` | Concepto de la carta |
| `@@VALOR@@` | Valor en formato COP (ej: `$ 1.500.000`) |
| `@@VALORENLETRAS@@` | Valor en letras (ej: `UN MILLÓN QUINIENTOS MIL PESOS M/CTE`) |
| `@@VALORRETEF@@` | Retención en la Fuente o `No Aplica` |
| `@@VALORRETEICA@@` | Tarifa ICA (ej: `5/1000`) o `No Aplica` |
| `@@CODIGOACTIVIDADECONOMICA@@` | Código CIIU |
| `@@VALORICA@@` | Valor calculado de ICA o `-` |
| `@@VALORFUENTE@@` | Valor calculado de Fuente o `-` |
| `@@TOTALPAGAR@@` | Total a pagar (bruto − ICA − Fuente) |
| `@@CONSECUTIVO@@` | Número de carta |
| `@@INICIOICA@@` ... `@@FINICA@@` | Bloque eliminado si ICA = No aplica |

---

## Configuración Local

### 1. Clonar e instalar

```bash
git clone https://github.com/Solutionsandpayroll/cuenta_cobro_cartas_PM.git
cd cuenta_cobro_cartas_PM
npm install
```

### 2. Variables de entorno

Crear el archivo `.env` en la raíz del proyecto (ver `.env.example`):

```env
VITE_JSONBIN_BIN_ID=tu_bin_id_aqui
VITE_JSONBIN_API_KEY_B64=base64_de_tu_api_key
```

> **¿Cómo obtener el base64 de la API Key?**
> La API Key de JSONBin suele contener caracteres `$` que rompen el parseo de `.env`.
> Convertila a base64 con Python:
> ```python
> import base64
> k = 'TU_API_KEY_AQUI'
> print(base64.b64encode(k.encode()).decode())
> ```
> El resultado es el valor que va en `VITE_JSONBIN_API_KEY_B64`.

### 3. Iniciar en desarrollo

```bash
npm run dev
```

---

## Despliegue en Vercel

1. Conectar el repositorio en [vercel.com](https://vercel.com)
2. En **Settings → Environment Variables** agregar:
   - `VITE_JSONBIN_BIN_ID` → el ID del Bin
   - `VITE_JSONBIN_API_KEY_B64` → el base64 de la API Key (igual que en el `.env` local)
3. Hacer **Redeploy** para que tome las variables

> El consecutivo se sincroniza con JSONBin en tiempo real — funciona igual en local y en Vercel.

---

## Configuración de JSONBin.io

1. Crear cuenta gratuita en [jsonbin.io](https://jsonbin.io)
2. Crear un nuevo **Bin** con contenido inicial:
   ```json
   { "consecutivo": 1 }
   ```
3. Copiar el **Bin ID** de la URL
4. En **Account → API Keys** copiar la **X-Master-Key**

---

## Scripts

```bash
npm run dev      # Servidor de desarrollo (http://localhost:5173)
npm run build    # Build de producción
npm run preview  # Preview del build
```

---

Desarrollado por **Solutions & Payroll** para uso interno de Punto Medical.


Template base reutilizable para proyectos React con el diseño corporativo de Solutions & Payroll.

## ✨ Características Incluidas

- ✅ **Header corporativo** con logo y bienvenida
- ✅ **Diseño profesional** con colores y estilos de S&P
- ✅ **Sección de ayuda colapsable** (opcional)
- ✅ **Sistema de cards** con animaciones suaves
- ✅ **Footer** corporativo
- ✅ **100% responsive** para móviles y desktop
- ✅ **Animaciones** de entrada elegantes
- ✅ **Variables CSS** fáciles de personalizar

## 🚀 Cómo Usar Este Template

### Opción 1: Copiar para Nuevo Proyecto

```bash
# 1. Copiar la carpeta completa
cp -r syp-react-template mi-nuevo-proyecto

# 2. Entrar al nuevo proyecto
cd mi-nuevo-proyecto

# 3. Instalar dependencias
npm install

# 4. Iniciar desarrollo
npm run dev
```

### Opción 2: Clonar y Modificar

```bash
# 1. Copiar todo el contenido
Copy-Item -Path "syp-react-template" -Destination "nuevo-proyecto" -Recurse

# 2. Cambiar nombre en package.json
# Edita la línea: "name": "tu-nombre-proyecto"

# 3. Instalar y ejecutar
cd nuevo-proyecto
npm install
npm run dev
```

## 📝 Estructura del Template

```
syp-react-template/
├── public/
│   └── Logo syp.png          # Logo corporativo S&P
├── src/
│   ├── App.jsx               # Componente principal (limpio)
│   ├── App.css               # Estilos completos
│   ├── index.css             # Estilos globales
│   └── main.jsx              # Entry point
├── index.html                # HTML base con favicon
├── package.json              # Dependencias mínimas
└── vite.config.js            # Configuración Vite
```

## 🎯 Personalización Rápida

### 1. Cambiar Título de la App

Edita `src/App.jsx` línea ~20:
```jsx
<p className="subtitle">Tu Nuevo Título</p>
```

### 2. Modificar Mensaje de Bienvenida

Edita `src/App.jsx` línea ~30:
```jsx
<span>Bienvenido, Tu Usuario</span>
```

### 3. Personalizar Colores

Edita `src/App.css`, variables CSS al inicio:
```css
:root {
  --primary: #2563eb;        /* Azul principal */
  --primary-dark: #1e40af;   /* Azul oscuro */
  /* ... más colores */
}
```

### 4. Agregar tu Lógica

En `src/App.jsx`, dentro del `<div className="card-body">`:
- Agrega tus estados con `useState`
- Crea tus funciones
- Añade tus componentes de formulario

## 📦 Agregar Dependencias

Según lo que necesites para tu proyecto:

```bash
# Para procesar archivos Excel
npm install xlsx exceljs file-saver

# Para formularios
npm install react-hook-form

# Para hacer requests
npm install axios

# Para routing
npm install react-router-dom

# etc...
```

## 🎨 Componentes Disponibles

### Sección de Ayuda Colapsable

Si no la necesitas, puedes eliminar todo el bloque:
```jsx
<div className="help-section">
  {/* ... */}
</div>
```

### Form Groups

```jsx
<div className="form-group">
  <label className="label">
    {/* Icono SVG */}
    Tu Label
  </label>
  <input className="select-input" />
</div>
```

### Botones

```jsx
<button className="btn-primary">
  {/* Icono SVG */}
  Texto del Botón
</button>
```

## 🌈 Estilos Predefinidos

Clases disponibles en `App.css`:
- `.card` - Contenedor con sombra
- `.form-section` - Espaciado de formularios
- `.form-group` - Grupo de campo
- `.label` - Label con icono
- `.select-input` - Input/Select estilizado
- `.btn-primary` - Botón principal
- `.btn-remove` - Botón eliminar
- `.drop-zone` - Zona drag & drop
- `.modal-overlay` - Overlay de modal
- `.help-section` - Sección colapsable

## 💡 Tips

1. **Mantén limpio el App.jsx** - Crea componentes separados si crece mucho
2. **Usa las variables CSS** - No modifiques los colores directamente
3. **Los SVG están inline** - Puedes cambiarlos fácilmente o usar íconos de librerías
4. **Las animaciones ya están configuradas** - Se activarán automáticamente

## 📚 Recursos

- [Documentación React](https://react.dev/)
- [Documentación Vite](https://vitejs.dev/)
- [Iconos SVG](https://feathericons.com/)
- [Colores](https://tailwindcss.com/docs/customizing-colors)

## 🔒 No Subir a Git

Si inicias Git en tu nuevo proyecto, asegúrate de tener `.gitignore`:
```
node_modules
dist
.env
```

## 📄 Licencia

© 2026 Solutions & Payroll. Template de uso interno.

---

**¡Listo para crear tu próximo proyecto!** 🚀
