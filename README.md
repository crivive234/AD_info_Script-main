# 🗂️ AD Info Tool

Herramienta interactiva de consola en PowerShell para consultar, auditar y gestionar objetos en **Active Directory**: usuarios, equipos, GPOs y objetos eliminados, todo desde un menú visual con salida en color.

---

## 📋 Requisitos

| Requisito | Detalle |
|-----------|---------|
| **PowerShell** | 5.1 o superior |
| **Módulo ActiveDirectory** | RSAT – Active Directory PowerShell |
| **Permisos** | Cuenta con acceso de lectura al dominio (y permisos de escritura para operaciones de gestión) |
| **Conectividad** | Acceso al Domain Controller |

### Instalar el módulo ActiveDirectory (RSAT)

**Opción 1 – Servidor Windows (PowerShell como Administrador):**
```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

**Opción 2 – Windows 10/11:**
```
Configuración → Aplicaciones → Características opcionales → RSAT: Active Directory Domain Services
```

---

## 🚀 Cómo ejecutar

```powershell
.\AD-Info.ps1
```

> **Nota:** Si el sistema bloquea la ejecución de scripts, usa:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```

---

## 🗃️ Menú principal

Al iniciar el script se muestra un menú numerado con todas las opciones disponibles:

```
  +================================================+
  |        AD INFO TOOL  -  Active Directory       |
  +================================================+

  --- USUARIOS ------------------------------------
  [1]  Ver info detallada de un usuario
  [2]  Buscar usuarios  (nombre / depto / OU / grupo / descripcion / ...)
  [6]  Comparar dos usuarios  (atributos y grupos)
  [8]  Desbloquear usuario
  [13] Habilitar / Deshabilitar usuario

  --- EQUIPOS -------------------------------------
  [3]  Ver info detallada de un equipo
  [4]  Buscar equipos
  [7]  Comparar dos equipos  (OU, atributos y grupos)
  [11] Auditoria de hostname  (creacion / modificaciones / baja)

  --- GPO / POLITICAS -----------------------------
  [9]  Consultar GPOs  (por grupo, usuario o equipo)
  [12] Comparar GPOs  (usuario vs usuario  /  equipo vs equipo)

  --- AUDITORIA / REPORTES ------------------------
  [5]  Reportes de auditoria  (bloqueados / inactivos / privilegios / ...)
  [10] AD Recycle Bin  (ver objetos eliminados)

  [0]  Salir
```

---

## 🔍 Descripción de funciones

### 👤 Módulo de Usuarios

#### `[1]` Ver info detallada de un usuario
Muestra un perfil completo del usuario. Acepta como entrada el **SamAccountName**, **email** o **nombre completo**.

Información que muestra:
- Datos generales: nombre, email, teléfono, cargo, departamento, manager
- Estado de la cuenta: activa/deshabilitada, bloqueada, intentos de logon fallidos, expiración
- Contraseña: último cambio, si nunca expira, si está expirada
- Actividad: último inicio de sesión, fecha de creación, modificación y OU donde está ubicado
- Grupos: lista completa de membresías

---

#### `[2]` Buscar usuarios
Búsqueda flexible de usuarios con múltiples filtros:

| Opción | Filtro |
|--------|--------|
| 1 | Por nombre o apellido |
| 2 | Por departamento |
| 3 | Por OU (requiere el DistinguishedName completo, ej: `OU=Ventas,DC=empresa,DC=com`) |
| 4 | Usuarios bloqueados |
| 5 | Usuarios deshabilitados |
| 6 | Contraseña expirada |
| 7 | Sin inicio de sesión en X días |
| 8 | Por texto en la descripción |
| 9 | Miembros de un grupo (búsqueda recursiva) |

Los resultados se muestran en tabla con: nombre, usuario, email, departamento, estado, bloqueo, contraseña expirada y último logon.

---

#### `[6]` Comparar dos usuarios
Compara lado a lado todos los atributos principales de dos cuentas de usuario, resaltando visualmente las diferencias. Incluye comparación de grupos de membresía.

---

#### `[8]` Desbloquear usuario
Muestra el estado actual de bloqueo del usuario (intentos fallidos, último intento) y ofrece tres opciones:
- **S** → Aplicar el desbloqueo real
- **W** → Simulación (`-WhatIf`), sin cambios reales
- **N** → Cancelar

Tras el desbloqueo, verifica que la operación fue exitosa.

---

#### `[13]` Habilitar / Deshabilitar usuario
Permite cambiar el estado de una cuenta de usuario. Muestra el estado actual y solicita confirmación antes de aplicar cualquier cambio.

---

### 💻 Módulo de Equipos

#### `[3]` Ver info detallada de un equipo
Muestra el perfil completo de una cuenta de equipo en AD:
- Datos del equipo: nombre, DNS hostname, IPv4, descripción
- Sistema operativo y versión
- Estado: habilitado/deshabilitado, último contacto con AD, fechas de creación y modificación
- Administrador asignado (`ManagedBy`) y ubicación en OU
- Grupos de membresía

---

#### `[4]` Buscar equipos
Búsqueda de equipos con los siguientes filtros:

| Opción | Filtro |
|--------|--------|
| 1 | Por nombre del equipo |
| 2 | Por sistema operativo (ej: `Windows 10`, `Server 2019`) |
| 3 | Por OU (DistinguishedName) |
| 4 | Equipos deshabilitados |
| 5 | Equipos inactivos (sin contacto en X días) |

---

#### `[7]` Comparar dos equipos
Compara atributos y grupos de dos equipos entre sí, destacando diferencias en OU, sistema operativo, estado y membresías de grupos.

---

#### `[11]` Auditoría de hostname
Consulta el **registro de seguridad del Domain Controller** en busca de eventos relacionados con el equipo:

| Evento | Descripción |
|--------|-------------|
| `4741` | Equipo dado de alta en el dominio (join) |
| `4742` | Cuenta de equipo modificada |
| `4743` | Equipo eliminado del dominio |

Muestra quién realizó cada acción y en qué fecha, y verifica si el equipo aún existe en AD.

---

### 📜 Módulo de GPOs

#### `[9]` Consultar GPOs
Consulta las políticas de grupo aplicables a un objeto del dominio. Permite buscar GPOs vinculadas a:
- Un usuario específico
- Un equipo específico
- Un grupo de seguridad

---

#### `[12]` Comparar GPOs
Compara las GPOs aplicadas a dos objetos del mismo tipo (usuario vs usuario o equipo vs equipo), identificando políticas en común y políticas exclusivas de cada uno.

---

### 📊 Módulo de Auditoría y Reportes

#### `[5]` Reportes de auditoría
Genera reportes sobre el estado del directorio. Tipos de reporte disponibles:
- Usuarios bloqueados
- Usuarios inactivos (sin inicio de sesión reciente)
- Cuentas con contraseña expirada
- Cuentas con privilegios elevados
- Equipos inactivos

Los resultados se pueden **exportar a CSV** indicando la carpeta de destino.

---

#### `[10]` AD Recycle Bin – Objetos eliminados
Permite buscar objetos eliminados en la **Papelera de reciclaje de Active Directory**:

**Modo 1 – Búsqueda por tipo y nombre:**
- Busca usuarios o equipos eliminados por nombre (parcial o completo)
- Muestra nombre, SID, fecha de eliminación, quién lo eliminó y OU de origen
- Permite exportar los resultados a CSV

**Modo 2 – Búsqueda por ObjectSID:**
- Localiza un objeto específico mediante su SID completo
- Formato esperado: `S-1-5-21-XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXX`

> ⚠️ La Recycle Bin debe estar habilitada en el dominio. Si no lo está, los objetos eliminados no serán recuperables por este método.

---

## 📁 Exportación de datos

Las funciones de búsqueda, reporte y Recycle Bin ofrecen la opción de exportar resultados a **CSV (UTF-8)**. Al confirmar la exportación, el script solicita la carpeta de destino y genera el archivo con un nombre que incluye la fecha y hora:

```
C:\Reportes\Usuarios_Bloqueados_20250401_1430.csv
```

---

## ⚠️ Consideraciones de seguridad

- Ejecutar con una cuenta con los **permisos mínimos necesarios** para cada operación.
- Las operaciones de escritura (desbloqueo, habilitar/deshabilitar) solicitan confirmación explícita antes de aplicarse.
- El modo **simulación (`WhatIf`)** disponible en el desbloqueo de usuarios permite revisar el impacto antes de confirmar.
- La auditoría de hostname requiere acceso al **registro de seguridad del DC**, lo cual generalmente exige privilegios administrativos en el controlador de dominio.

---

## 🛠️ Solución de problemas comunes

| Problema | Solución |
|----------|----------|
| `El módulo ActiveDirectory no está instalado` | Instalar RSAT (ver sección de Requisitos) |
| `No se encontró el usuario / equipo` | Verificar que el SamAccountName o nombre sea correcto |
| `Formato de DN inválido` | El DistinguishedName debe comenzar con `OU=`, `CN=` o `DC=` y contener `DC=` |
| `Error al leer el registro del DC` | Se requieren permisos de administrador sobre el Domain Controller |
| `No se encontraron objetos en la Recycle Bin` | Verificar que la Recycle Bin esté habilitada o que el SID sea correcto |

---

## 📄 Licencia

Libre para uso interno y personal. Revisar las políticas de tu organización antes de ejecutar scripts con privilegios de dominio.
