#Requires -Module ActiveDirectory
<#
.SYNOPSIS
    AD Info Tool - Consulta, comparacion y gestion de Active Directory
.DESCRIPTION
    Usuarios, equipos, comparaciones visuales, busquedas avanzadas y desbloqueo.
.NOTES
    Requiere el modulo ActiveDirectory (RSAT) instalado.
    Ejecutar con privilegios suficientes en el dominio.
#>

# ============================================================
#  CONFIGURACION Y UTILIDADES
# ============================================================

$Host.UI.RawUI.WindowTitle = "AD Info Tool"

function Write-Header {
    param([string]$Titulo)
    Write-Host ""
    Write-Host "  +============================================================+" -ForegroundColor Cyan
    Write-Host ("  |  " + $Titulo.ToUpper().PadRight(58) + "|") -ForegroundColor White
    Write-Host "  +============================================================+" -ForegroundColor Cyan
    Write-Host ""
}

function Write-SubHeader {
    param([string]$Titulo)
    Write-Host ""
    Write-Host "  +-- $Titulo " -NoNewline -ForegroundColor Yellow
    Write-Host ("-" * ([Math]::Max(1, 54 - $Titulo.Length))) -NoNewline -ForegroundColor DarkGray
    Write-Host "+" -ForegroundColor Yellow
    Write-Host ""
}

function Write-Campo {
    param([string]$Etiqueta, [string]$Valor, [ConsoleColor]$Color = "Gray")
    $etiq = $Etiqueta.PadRight(26)
    Write-Host "  $etiq : " -NoNewline -ForegroundColor DarkCyan
    Write-Host $Valor -ForegroundColor $Color
}

function Write-Separador {
    Write-Host "  " -NoNewline
    Write-Host ("-" * 60) -ForegroundColor DarkGray
}

function Get-OUdesdeDN {
    param([string]$DN)
    if ([string]::IsNullOrEmpty($DN)) { return "N/A" }
    $partes = $DN -split ","
    ($partes | Where-Object { $_ -notmatch "^CN=" } | ForEach-Object { ($_ -split "=")[1] }) -join " > "
}

function Pause-Pantalla {
    Write-Host ""
    $null = Read-Host "  Presiona Enter para continuar"
}

# Imprime un badge de estado con color
function Write-Badge {
    param([string]$Texto, [string]$Tipo = "info")
    switch ($Tipo) {
        "ok"      { Write-Host " [ OK ] $Texto " -ForegroundColor Black -BackgroundColor Green  }
        "warn"    { Write-Host " [ ! ] $Texto "  -ForegroundColor Black -BackgroundColor Yellow }
        "error"   { Write-Host " [ X ] $Texto "  -ForegroundColor White -BackgroundColor Red    }
        "info"    { Write-Host " [ i ] $Texto "  -ForegroundColor White -BackgroundColor DarkCyan }
        default   { Write-Host $Texto }
    }
}

# Devuelve el PDC Emulator del dominio actual, o $null si no se pudo determinar.
# El estado de bloqueo (LockedOut) solo es 100% confiable si se lee y se escribe
# contra el PDC Emulator: es el unico DC que recibe replicacion urgente e
# inmediata de intentos fallidos y bloqueos. Consultar/aplicar contra cualquier
# otro DC puede devolver un estado desactualizado, sobre todo en entornos
# multi-sitio con varios DCs.
function Get-PDCEmulator {
    try {
        (Get-ADDomain -ErrorAction Stop).PDCEmulator
    } catch {
        $null
    }
}

# Helper: flujo compartido de "preguntar si exportar -> pedir carpeta -> crear si
# no existe -> exportar con timestamp". $Data ya debe venir con las columnas
# finales seleccionadas (Select-Object) por quien llama. Devuelve $true si
# exporto, $false si el usuario declino o no habia datos.
function Export-DatosCSV {
    param(
        $Data,
        [string]$NombreArchivoBase,
        [string]$Prompt
    )
    if (-not $Data -or @($Data).Count -eq 0) {
        Write-Host "  No hay datos para exportar." -ForegroundColor Yellow
        return $false
    }
    $exp = Read-Host $Prompt
    if ($exp -notmatch "^[sS]$") { return $false }

    $carpeta = Read-Host "  Carpeta destino (ej: C:\Reportes)"
    if (-not (Test-Path $carpeta)) {
        New-Item -ItemType Directory -Path $carpeta -Force | Out-Null
        Write-Host "  Carpeta creada: $carpeta" -ForegroundColor DarkGray
    }
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmm")
    $archivo   = Join-Path $carpeta "${NombreArchivoBase}_$timestamp.csv"
    $Data | Export-Csv -Path $archivo -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  OK - CSV guardado en:" -ForegroundColor Green
    Write-Host "  $archivo" -ForegroundColor Cyan
    return $true
}

# Helper: separador de una tabla de comparacion de 2 columnas (atributo | A | B)
function Write-SepComparacion {
    param([int]$LabelW = 22, [int]$ColW = 20)
    Write-Host ("  +" + ("-" * ($LabelW+2)) + "+" + ("-" * ($ColW+2)) + "+" + ("-" * ($ColW+2)) + "+") -ForegroundColor DarkGray
}

# Helper: encabezado de una tabla de comparacion de 2 columnas
function Write-EncabezadoComparacion {
    param([string]$NombreA, [string]$NombreB, [string]$ColorA = "Cyan", [string]$ColorB = "Magenta", [int]$LabelW = 22, [int]$ColW = 20)
    Write-SepComparacion -LabelW $LabelW -ColW $ColW
    Write-Host ("  | " + "ATRIBUTO".PadRight($LabelW) + " | ") -NoNewline -ForegroundColor DarkGray
    Write-Host $NombreA.PadRight($ColW) -NoNewline -ForegroundColor $ColorA
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
    Write-Host $NombreB.PadRight($ColW) -NoNewline -ForegroundColor $ColorB
    Write-Host " |" -ForegroundColor DarkGray
    Write-SepComparacion -LabelW $LabelW -ColW $ColW
}

# Helper: una fila de tabla de comparacion, resaltando en amarillo si A y B difieren.
# Usado por Compare-Usuarios y Compare-Equipos para no repetir esta funcion interna.
function Write-FilaComparacion {
    param([string]$Label, [string]$ValA, [string]$ValB, [string]$ColorA = "Cyan", [string]$ColorB = "Magenta", [int]$LabelW = 22, [int]$ColW = 20)
    $igual  = ($ValA -eq $ValB)
    $colFil = if ($igual) {"Gray"} else {"Yellow"}
    $icono  = if ($igual) {"  "} else {">>"}
    $v1 = if ($ValA.Length -gt $ColW) {$ValA.Substring(0,$ColW-3)+"..."} else {$ValA.PadRight($ColW)}
    $v2 = if ($ValB.Length -gt $ColW) {$ValB.Substring(0,$ColW-3)+"..."} else {$ValB.PadRight($ColW)}
    Write-Host ("  |$icono" + $Label.PadRight($LabelW) + " | ") -NoNewline -ForegroundColor $colFil
    Write-Host $v1 -NoNewline -ForegroundColor $(if ($igual) {"Gray"} else {$ColorA})
    Write-Host " | " -NoNewline -ForegroundColor DarkGray
    Write-Host $v2 -NoNewline -ForegroundColor $(if ($igual) {"Gray"} else {$ColorB})
    Write-Host " |" -ForegroundColor DarkGray
}

# Helper: compara dos listas de MemberOf (CN extraido), muestra el resumen y el
# detalle (comunes / solo en A / solo en B). Devuelve las tres listas para que
# quien llama pueda exportarlas si lo necesita. Usado por Compare-Usuarios y
# Compare-Equipos para no repetir esta logica.
function Compare-YMostrarGrupos {
    param(
        [string[]]$MemberOfA, [string[]]$MemberOfB,
        [string]$NombreA, [string]$NombreB,
        [string]$ColorA = "Cyan", [string]$ColorB = "Magenta",
        [string]$Sustantivo = "elementos"
    )
    $grupos1 = @($MemberOfA | ForEach-Object {($_ -split ",")[0] -replace "^CN=",""} | Sort-Object)
    $grupos2 = @($MemberOfB | ForEach-Object {($_ -split ",")[0] -replace "^CN=",""} | Sort-Object)
    if ($grupos1.Count -eq 0) { $grupos1 = @("__EMPTY__") }
    if ($grupos2.Count -eq 0) { $grupos2 = @("__EMPTY__") }

    $diff    = Compare-Object -ReferenceObject $grupos1 -DifferenceObject $grupos2 -IncludeEqual
    $soloA   = @($diff | Where-Object {$_.SideIndicator -eq "<="} | Where-Object {$_.InputObject -ne "__EMPTY__"})
    $soloB   = @($diff | Where-Object {$_.SideIndicator -eq "=>"} | Where-Object {$_.InputObject -ne "__EMPTY__"})
    $comunes = @($diff | Where-Object {$_.SideIndicator -eq "=="} | Where-Object {$_.InputObject -ne "__EMPTY__"})

    Write-Host ""
    Write-Host "  +---------------------------+-------+" -ForegroundColor DarkGray
    Write-Host "  | Categoria                 | Total |" -ForegroundColor DarkGray
    Write-Host "  +---------------------------+-------+" -ForegroundColor DarkGray
    Write-Host ("  | Grupos en comun           |  " + "$($comunes.Count)".PadLeft(4) + " |") -ForegroundColor Green
    Write-Host ("  | Solo en " + $NombreA.PadRight(16) + "  |  " + "$($soloA.Count)".PadLeft(4) + " |") -ForegroundColor $ColorA
    Write-Host ("  | Solo en " + $NombreB.PadRight(16) + "  |  " + "$($soloB.Count)".PadLeft(4) + " |") -ForegroundColor $ColorB
    Write-Host "  +---------------------------+-------+" -ForegroundColor DarkGray
    Write-Host ""

    if ($comunes.Count -gt 0) {
        Write-Host "  [=] EN COMUN:" -ForegroundColor Green
        $comunes | Sort-Object InputObject | ForEach-Object { Write-Host "      (=) $($_.InputObject)" -ForegroundColor Green }
        Write-Host ""
    }
    if ($soloA.Count -gt 0) {
        Write-Host ("  [A] SOLO " + $NombreA.ToUpper() + " -- le faltan a " + $NombreB + ":") -ForegroundColor $ColorA
        $soloA | Sort-Object InputObject | ForEach-Object { Write-Host "      (+) $($_.InputObject)" -ForegroundColor $ColorA }
        Write-Host ""
    }
    if ($soloB.Count -gt 0) {
        Write-Host ("  [B] SOLO " + $NombreB.ToUpper() + " -- le faltan a " + $NombreA + ":") -ForegroundColor $ColorB
        $soloB | Sort-Object InputObject | ForEach-Object { Write-Host "      (+) $($_.InputObject)" -ForegroundColor $ColorB }
        Write-Host ""
    }
    if ($soloA.Count -eq 0 -and $soloB.Count -eq 0) {
        Write-Host "  Ambos $Sustantivo tienen exactamente los mismos grupos." -ForegroundColor Green
        Write-Host ""
    }

    [PSCustomObject]@{ Comunes = $comunes; SoloA = $soloA; SoloB = $soloB }
}

# ============================================================
#  MODULO USUARIOS
# ============================================================

function Get-InfoUsuario {
    Write-Header "INFORMACION DE USUARIO"
    $identidad = Read-Host "  Usuario (SamAccountName, email o nombre)"
    try {
        $props = @(
            "Name","SamAccountName","EmailAddress","Enabled","LockedOut",
            "PasswordLastSet","PasswordNeverExpires","PasswordExpired",
            "LastLogonDate","Created","Modified","Description",
            "Department","Title","Manager","MemberOf",
            "DistinguishedName","TelephoneNumber","OfficePhone",
            "AccountExpirationDate","BadLogonCount"
        )
        $u = Get-ADUser -Identity $identidad -Properties $props -ErrorAction Stop

        Write-SubHeader "Datos Generales"
        Write-Campo "Nombre Completo"      $u.Name                                                        "White"
        Write-Campo "Usuario (SAM)"        $u.SamAccountName                                              "White"
        Write-Campo "Email"                $(if ($u.EmailAddress)    {$u.EmailAddress}    else {"No asignado"})
        Write-Campo "Telefono"             $(if ($u.TelephoneNumber) {$u.TelephoneNumber} else {"No asignado"})
        Write-Campo "Cargo"                $(if ($u.Title)           {$u.Title}           else {"No asignado"})
        Write-Campo "Departamento"         $(if ($u.Department)      {$u.Department}      else {"No asignado"})
        Write-Campo "Descripcion"          $(if ($u.Description)     {$u.Description}     else {"Sin descripcion"})
        if ($u.Manager) {
            $mgr = (Get-ADUser -Identity $u.Manager -ErrorAction SilentlyContinue).Name
            Write-Campo "Manager"          $(if ($mgr) {$mgr} else {$u.Manager})
        } else { Write-Campo "Manager" "No asignado" }

        Write-SubHeader "Estado de la Cuenta"
        $estCuenta  = if ($u.Enabled)   {"ACTIVA"}     else {"DESHABILITADA"}
        $colCuenta  = if ($u.Enabled)   {"Green"}      else {"Red"}
        Write-Campo "Estado"               $estCuenta $colCuenta
        $estBloqueo = if ($u.LockedOut) {"BLOQUEADA"}  else {"Desbloqueada"}
        $colBloqueo = if ($u.LockedOut) {"Red"}        else {"Green"}
        Write-Campo "Bloqueo"              $estBloqueo $colBloqueo
        Write-Campo "Intentos fallidos"    "$($u.BadLogonCount)"
        $expCuenta = if ($u.AccountExpirationDate) {$u.AccountExpirationDate.ToString("dd/MM/yyyy")} else {"No expira"}
        Write-Campo "Expiracion cuenta"    $expCuenta

        Write-SubHeader "Contrasena"
        $ultCambio = if ($u.PasswordLastSet) {$u.PasswordLastSet.ToString("dd/MM/yyyy HH:mm")} else {"Nunca"}
        Write-Campo "Ultimo cambio"        $ultCambio
        Write-Campo "Nunca expira"         $(if ($u.PasswordNeverExpires) {"Si"} else {"No"})
        $passExp = if ($u.PasswordExpired) {"SI"} else {"No"}
        $colPass = if ($u.PasswordExpired) {"Red"} else {"Green"}
        Write-Campo "Contrasena expirada"  $passExp $colPass

        Write-SubHeader "Actividad y Ubicacion"
        $ultLogon = if ($u.LastLogonDate) {$u.LastLogonDate.ToString("dd/MM/yyyy HH:mm")} else {"Nunca / No registrado"}
        Write-Campo "Ultimo inicio sesion" $ultLogon
        Write-Campo "Creado"               $u.Created.ToString("dd/MM/yyyy")
        Write-Campo "Modificado"           $u.Modified.ToString("dd/MM/yyyy")
        Write-Campo "Ubicacion (OU)"       (Get-OUdesdeDN $u.DistinguishedName)

        Write-SubHeader "Grupos de Membresia ($($u.MemberOf.Count))"
        if ($u.MemberOf -and $u.MemberOf.Count -gt 0) {
            $u.MemberOf | ForEach-Object {
                $gn = ($_ -split ",")[0] -replace "^CN=",""
                Write-Host "    (*) $gn" -ForegroundColor Magenta
            }
        } else { Write-Host "    Sin grupos asignados" -ForegroundColor DarkGray }

    } catch {
        Write-Host ""
        Write-Host "  ERROR: No se encontro el usuario '$identidad'" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkRed
    }
    Pause-Pantalla
}

function Search-Usuarios {
    Write-Header "BUSQUEDA DE USUARIOS"
    Write-Host "  Opciones de busqueda:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] Por nombre o apellido"
    Write-Host "  [2] Por departamento"
    Write-Host "  [3] Por OU (unidad organizativa)"
    Write-Host "  [4] Usuarios bloqueados"
    Write-Host "  [5] Usuarios deshabilitados"
    Write-Host "  [6] Contrasena expirada"
    Write-Host "  [7] Sin inicio de sesion reciente"
    Write-Host "  [8] Por descripcion"
    Write-Host "  [9] Miembros de un grupo"
    Write-Host ""
    $opc = Read-Host "  Opcion"

    $props = @("Name","SamAccountName","EmailAddress","Enabled","LockedOut",
               "LastLogonDate","Department","PasswordExpired","Description")
    try {
        $lista = $null
        switch ($opc) {
            "1" {
                $q = Read-Host "  Nombre a buscar"
                $lista = Get-ADUser -Filter "Name -like '*$q*'" -Properties $props | Sort-Object Name
            }
            "2" {
                $q = Read-Host "  Departamento"
                $lista = Get-ADUser -Filter "Department -like '*$q*'" -Properties $props | Sort-Object Name
            }
            "3" {
                $q = Read-Host "  DistinguishedName de la OU (ej: OU=Ventas,DC=empresa,DC=com)"
                if (-not ($q -match "^(OU|CN|DC)=.+,DC=")) {
                    Write-Host "  Formato de DN invalido. Debe comenzar con OU=, CN= o DC= y contener DC=." -ForegroundColor Red
                    Pause-Pantalla; return
                }
                $lista = Get-ADUser -Filter * -SearchBase $q -Properties $props | Sort-Object Name
            }
            "4" {
                $pdc  = Get-PDCEmulator
                $srvP = if ($pdc) { @{ Server = $pdc } } else { @{} }
                $lista = Search-ADAccount @srvP -LockedOut -UsersOnly | Get-ADUser @srvP -Properties $props | Sort-Object Name
            }
            "5" {
                $lista = Get-ADUser -Filter "Enabled -eq '$false'" -Properties $props | Sort-Object Name
            }
            "6" {
                $lista = Search-ADAccount -PasswordExpired -UsersOnly | Get-ADUser -Properties $props | Sort-Object Name
            }
            "7" {
                $dias = Read-Host "  Dias sin inicio de sesion (ej: 90)"
                $fecha = (Get-Date).AddDays(-[int]$dias)
                $fechaStr = $fecha.ToFileTime()
                $lista = Get-ADUser -Filter "Enabled -eq '$true' -and (LastLogonTimestamp -lt $fechaStr -or LastLogonTimestamp -notlike '*')" -Properties $props |
                    Sort-Object LastLogonDate
            }
            "8" {
                $q = Read-Host "  Texto en descripcion a buscar"
                $lista = Get-ADUser -Filter "Description -like '*$q*'" -Properties $props | Sort-Object Name
            }
            "9" {
                $grp = Read-Host "  Nombre del grupo"
                try {
                    $miembros = Get-ADGroupMember -Identity $grp -Recursive |
                        Where-Object { $_.objectClass -eq "user" }
                    if ($miembros) {
                        $lista = $miembros | ForEach-Object {
                            Get-ADUser -Identity $_.SamAccountName -Properties $props
                        } | Sort-Object Name
                    } else {
                        $lista = @()
                    }
                } catch {
                    Write-Host ""
                    Write-Host "  ERROR: No se encontro el grupo '$grp'" -ForegroundColor Red
                    Pause-Pantalla
                    return
                }
            }
            default {
                Write-Host "  Opcion no valida." -ForegroundColor Red
                Pause-Pantalla
                return
            }
        }

        if (-not $lista -or @($lista).Count -eq 0) {
            Write-Host ""
            Write-Host "  No se encontraron usuarios." -ForegroundColor Yellow
        } else {
            Write-Host ""
            Write-Host "  Total: $(@($lista).Count) usuarios encontrados" -ForegroundColor Green
            Write-Host ""
            Write-Separador
            $lista | Select-Object `
                @{N="Nombre";       E={$_.Name}},
                @{N="Usuario";      E={$_.SamAccountName}},
                @{N="Email";        E={if($_.EmailAddress){$_.EmailAddress}else{"-"}}},
                @{N="Depto";        E={if($_.Department){$_.Department}else{"-"}}},
                @{N="Estado";       E={if($_.Enabled){"Activa"}else{"Deshabilitada"}}},
                @{N="Bloqueado";    E={if($_.LockedOut){"SI"}else{"No"}}},
                @{N="Pass.Exp.";    E={if($_.PasswordExpired){"SI"}else{"No"}}},
                @{N="Ultimo logon"; E={if($_.LastLogonDate){$_.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"}}} |
            Format-Table -AutoSize
        }
    } catch {
        Write-Host ""
        Write-Host "  ERROR en la busqueda: $($_.Exception.Message)" -ForegroundColor Red
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO DESBLOQUEO DE USUARIOS
# ============================================================

function Unlock-Usuario {
    Write-Header "DESBLOQUEO DE USUARIO"

    $pdc  = Get-PDCEmulator
    $srvP = if ($pdc) { @{ Server = $pdc } } else { @{} }
    if ($pdc) {
        Write-Host "  (Verificando y aplicando contra el PDC Emulator: $pdc)" -ForegroundColor DarkGray
    } else {
        Write-Host "  [!] No se pudo determinar el PDC Emulator; se usara el DC por defecto." -ForegroundColor Yellow
        Write-Host "      El estado de bloqueo podria no ser el mas reciente." -ForegroundColor Yellow
    }

    $identidad = Read-Host "  Usuario a desbloquear (SamAccountName)"

    try {
        $u = Get-ADUser @srvP -Identity $identidad -Properties LockedOut, BadLogonCount, LastBadPasswordAttempt -ErrorAction Stop

        Write-Host ""
        Write-Separador
        Write-Host "  Usuario   : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $u.Name -ForegroundColor White
        Write-Host "  SAM       : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $u.SamAccountName -ForegroundColor White
        Write-Host "  Bloqueado : " -NoNewline -ForegroundColor DarkCyan
        if ($u.LockedOut) {
            Write-Host "SI" -ForegroundColor Red
        } else {
            Write-Host "No (la cuenta no esta bloqueada actualmente)" -ForegroundColor Yellow
        }
        Write-Host "  Intentos fallidos : " -NoNewline -ForegroundColor DarkCyan
        Write-Host "$($u.BadLogonCount)" -ForegroundColor White
        if ($u.LastBadPasswordAttempt) {
            Write-Host "  Ultimo intento    : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $u.LastBadPasswordAttempt.ToString("dd/MM/yyyy HH:mm") -ForegroundColor White
        }
        Write-Separador
        Write-Host ""

        if (-not $u.LockedOut) {
            Write-Host "  La cuenta no esta bloqueada. No se requiere accion." -ForegroundColor Yellow
            Pause-Pantalla
            return
        }

        $conf = Read-Host "  Confirmar desbloqueo de '$($u.SamAccountName)'? (S=Confirmar / W=Simular / N=Cancelar)"
        if ($conf -match "^[wW]$") {
            Unlock-ADAccount @srvP -Identity $identidad -WhatIf
            Write-Host ""
            Write-Host "  [SIMULACION] El comando se ejecutaria sobre '$($u.SamAccountName)'." -ForegroundColor Yellow
            Write-Host "  Vuelve a ejecutar y confirma con S para aplicar el cambio." -ForegroundColor DarkGray
        } elseif ($conf -match "^[sS]$") {
            Unlock-ADAccount @srvP -Identity $identidad -ErrorAction Stop
            Write-Host ""
            Write-Host "  OK - Cuenta desbloqueada exitosamente." -ForegroundColor Green

            # Verificar resultado
            $check = Get-ADUser @srvP -Identity $identidad -Properties LockedOut
            $estado = if ($check.LockedOut) {"SIGUE BLOQUEADA (verifica permisos)"} else {"Desbloqueada correctamente"}
            $col    = if ($check.LockedOut) {"Red"} else {"Green"}
            Write-Host "  Verificacion : $estado" -ForegroundColor $col
        } else {
            Write-Host ""
            Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
        }

    } catch {
        Write-Host ""
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO HABILITACION / DESHABILITACION DE USUARIOS
# ============================================================

function Enable-Usuario {
    Write-Header "HABILITAR / DESHABILITAR USUARIO"

    $identidad = Read-Host "  Usuario (SamAccountName)"

    try {
        $u = Get-ADUser -Identity $identidad -Properties Enabled, Description, Department -ErrorAction Stop

        Write-Host ""
        Write-Separador
        Write-Host "  Usuario   : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $u.Name -ForegroundColor White
        Write-Host "  SAM       : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $u.SamAccountName -ForegroundColor White
        Write-Host "  Estado    : " -NoNewline -ForegroundColor DarkCyan
        if ($u.Enabled) {
            Write-Host "ACTIVA" -ForegroundColor Green
        } else {
            Write-Host "DESHABILITADA" -ForegroundColor Red
        }
        Write-Separador
        Write-Host ""

        if ($u.Enabled) {
            Write-Host "  La cuenta esta actualmente HABILITADA." -ForegroundColor Green
            Write-Host "  [1] Deshabilitar la cuenta"
            Write-Host "  [W] Simular deshabilitacion (WhatIf)"
            Write-Host "  [N] Cancelar"
        } else {
            Write-Host "  La cuenta esta actualmente DESHABILITADA." -ForegroundColor Red
            Write-Host "  [1] Habilitar la cuenta"
            Write-Host "  [W] Simular habilitacion (WhatIf)"
            Write-Host "  [N] Cancelar"
        }
        Write-Host ""

        $accion = Read-Host "  Opcion"

        if ($accion -match "^[wW]$") {
            if ($u.Enabled) {
                Disable-ADAccount -Identity $identidad -WhatIf
                Write-Host ""
                Write-Host "  [SIMULACION] La cuenta '$($u.SamAccountName)' seria DESHABILITADA." -ForegroundColor Yellow
            } else {
                Enable-ADAccount -Identity $identidad -WhatIf
                Write-Host ""
                Write-Host "  [SIMULACION] La cuenta '$($u.SamAccountName)' seria HABILITADA." -ForegroundColor Yellow
            }
            Write-Host "  Ninguna cuenta fue modificada." -ForegroundColor DarkGray

        } elseif ($accion -match "^[1]$") {
            if ($u.Enabled) {
                $conf = Read-Host "  Confirmar DESHABILITACION de '$($u.SamAccountName)'? (S/N)"
                if ($conf -match "^[sS]$") {
                    Disable-ADAccount -Identity $identidad -ErrorAction Stop
                    Write-Host ""
                    Write-Host "  OK - Cuenta deshabilitada exitosamente." -ForegroundColor Green
                    $check = Get-ADUser -Identity $identidad -Properties Enabled
                    $estado = if (-not $check.Enabled) {"Deshabilitada correctamente"} else {"SIGUE ACTIVA (verifica permisos)"}
                    $col    = if (-not $check.Enabled) {"Green"} else {"Red"}
                    Write-Host "  Verificacion : $estado" -ForegroundColor $col
                } else {
                    Write-Host ""
                    Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
                }
            } else {
                $conf = Read-Host "  Confirmar HABILITACION de '$($u.SamAccountName)'? (S/N)"
                if ($conf -match "^[sS]$") {
                    Enable-ADAccount -Identity $identidad -ErrorAction Stop
                    Write-Host ""
                    Write-Host "  OK - Cuenta habilitada exitosamente." -ForegroundColor Green
                    $check = Get-ADUser -Identity $identidad -Properties Enabled
                    $estado = if ($check.Enabled) {"Habilitada correctamente"} else {"SIGUE DESHABILITADA (verifica permisos)"}
                    $col    = if ($check.Enabled) {"Green"} else {"Red"}
                    Write-Host "  Verificacion : $estado" -ForegroundColor $col
                } else {
                    Write-Host ""
                    Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
                }
            }
        } else {
            Write-Host ""
            Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
        }

    } catch {
        Write-Host ""
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO EQUIPOS
# ============================================================

function Get-InfoEquipo {
    Write-Header "INFORMACION DE EQUIPO"
    $nombre = Read-Host "  Nombre del equipo (hostname)"
    try {
        $props = @(
            "Name","DNSHostName","Enabled","OperatingSystem","OperatingSystemVersion",
            "LastLogonDate","Created","Modified","Description","MemberOf",
            "DistinguishedName","IPv4Address","ManagedBy"
        )
        $e = Get-ADComputer -Identity $nombre -Properties $props -ErrorAction Stop

        Write-SubHeader "Datos del Equipo"
        Write-Campo "Nombre"               $e.Name                                                           "White"
        Write-Campo "DNS Hostname"         $(if ($e.DNSHostName)  {$e.DNSHostName}  else {"No registrado"})
        Write-Campo "IPv4"                 $(if ($e.IPv4Address)  {$e.IPv4Address}  else {"No disponible"})
        Write-Campo "Descripcion"          $(if ($e.Description)  {$e.Description}  else {"Sin descripcion"})

        Write-SubHeader "Sistema Operativo"
        Write-Campo "SO"                   $(if ($e.OperatingSystem)        {$e.OperatingSystem}        else {"No registrado"})
        Write-Campo "Version"              $(if ($e.OperatingSystemVersion) {$e.OperatingSystemVersion} else {"No registrado"})

        Write-SubHeader "Estado"
        $estEq = if ($e.Enabled) {"Habilitado"} else {"DESHABILITADO"}
        $colEq = if ($e.Enabled) {"Green"}      else {"Red"}
        Write-Campo "Estado"               $estEq $colEq
        $ultLogon = if ($e.LastLogonDate) {$e.LastLogonDate.ToString("dd/MM/yyyy HH:mm")} else {"Nunca / No registrado"}
        Write-Campo "Ultimo contacto AD"   $ultLogon
        Write-Campo "Creado en AD"         $e.Created.ToString("dd/MM/yyyy")
        Write-Campo "Modificado"           $e.Modified.ToString("dd/MM/yyyy")
        if ($e.ManagedBy) {
            $adm = (Get-ADUser -Identity $e.ManagedBy -ErrorAction SilentlyContinue).Name
            Write-Campo "Administrador"    $(if ($adm) {$adm} else {$e.ManagedBy})
        } else { Write-Campo "Administrador" "No asignado" }

        Write-SubHeader "Ubicacion (OU)"
        Write-Campo "OU Path"              (Get-OUdesdeDN $e.DistinguishedName)

        Write-SubHeader "Grupos de Membresia ($($e.MemberOf.Count))"
        if ($e.MemberOf -and $e.MemberOf.Count -gt 0) {
            $e.MemberOf | ForEach-Object {
                $gn = ($_ -split ",")[0] -replace "^CN=",""
                Write-Host "    (*) $gn" -ForegroundColor Magenta
            }
        } else { Write-Host "    Sin grupos asignados" -ForegroundColor DarkGray }

    } catch {
        Write-Host ""
        Write-Host "  ERROR: No se encontro el equipo '$nombre'" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkRed
    }
    Pause-Pantalla
}

function Search-Equipos {
    Write-Header "BUSQUEDA DE EQUIPOS"
    Write-Host "  Opciones:" -ForegroundColor Yellow
    Write-Host "  [1] Por nombre del equipo"
    Write-Host "  [2] Por sistema operativo"
    Write-Host "  [3] Por OU (unidad organizativa)"
    Write-Host "  [4] Equipos deshabilitados"
    Write-Host "  [5] Equipos inactivos (sin contacto reciente)"
    Write-Host ""
    $opc = Read-Host "  Opcion"
    $props = @("Name","DNSHostName","Enabled","OperatingSystem","LastLogonDate","DistinguishedName")
    try {
        $lista = $null
        switch ($opc) {
            "1" {
                $q = Read-Host "  Nombre a buscar"
                $lista = Get-ADComputer -Filter "Name -like '*$q*'" -Properties $props | Sort-Object Name
            }
            "2" {
                $q = Read-Host "  SO a buscar (ej: Windows 10, Server 2019)"
                $lista = Get-ADComputer -Filter "OperatingSystem -like '*$q*'" -Properties $props | Sort-Object Name
            }
            "3" {
                $q = Read-Host "  DistinguishedName de la OU"
                if (-not ($q -match "^(OU|CN|DC)=.+,DC=")) {
                    Write-Host "  Formato de DN invalido. Debe comenzar con OU=, CN= o DC= y contener DC=." -ForegroundColor Red
                    Pause-Pantalla; return
                }
                $lista = Get-ADComputer -Filter * -SearchBase $q -Properties $props | Sort-Object Name
            }
            "4" {
                $lista = Get-ADComputer -Filter "Enabled -eq '$false'" -Properties $props | Sort-Object Name
            }
            "5" {
                $dias = Read-Host "  Dias sin contacto (ej: 90)"
                $fecha = (Get-Date).AddDays(-[int]$dias)
                $fechaStr = $fecha.ToFileTime()
                $lista = Get-ADComputer -Filter "Enabled -eq '$true' -and (LastLogonTimestamp -lt $fechaStr -or LastLogonTimestamp -notlike '*')" -Properties $props |
                    Sort-Object LastLogonDate
            }
            default {
                Write-Host "  Opcion no valida." -ForegroundColor Red
                Pause-Pantalla
                return
            }
        }
        if (-not $lista -or @($lista).Count -eq 0) {
            Write-Host ""
            Write-Host "  No se encontraron equipos." -ForegroundColor Yellow
        } else {
            Write-Host ""
            Write-Host "  Total: $(@($lista).Count) equipos" -ForegroundColor Green
            Write-Host ""
            $lista | Select-Object `
                @{N="Nombre";       E={$_.Name}},
                @{N="DNS";          E={if($_.DNSHostName){$_.DNSHostName}else{"-"}}},
                @{N="Sistema Op.";  E={if($_.OperatingSystem){$_.OperatingSystem}else{"-"}}},
                @{N="Estado";       E={if($_.Enabled){"Habilitado"}else{"Deshabilitado"}}},
                @{N="Ultimo logon"; E={if($_.LastLogonDate){$_.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"}}} |
            Format-Table -AutoSize
        }
    } catch {
        Write-Host ""
        Write-Host "  ERROR en la busqueda: $($_.Exception.Message)" -ForegroundColor Red
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO AUDITORIA DE EQUIPO - HISTORIAL DE EVENTOS
# ============================================================

# Helper: parsea eventos 4741 / 4742 / 4743 via XML con fallback a Properties[]
function Parse-CompEvent {
    param($ev)
    try {
        $xml   = [xml]$ev.ToXml()
        $xdata = $xml.Event.EventData.Data
        $comp  = ($xdata | Where-Object { $_.Name -eq "TargetComputerName" }).'#text'
        $actor = ($xdata | Where-Object { $_.Name -eq "SubjectUserName"   }).'#text'
        $dom   = ($xdata | Where-Object { $_.Name -eq "SubjectDomainName" }).'#text'
        if (-not $comp)  { $comp  = "$($ev.Properties[1].Value)" }
        if (-not $actor) { $actor = "$($ev.Properties[4].Value)" }
        [PSCustomObject]@{
            Equipo = ($comp -replace '\$$','').Trim()
            Actor  = if ($actor -and $dom) { "$dom\$actor" } elseif ($actor) { $actor } else { "Desconocido" }
            Fecha  = $ev.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
        }
    } catch { $null }
}

# Helper: verifica si una subcategoria de auditoria esta registrando Success en un DC.
# Requiere WinRM habilitado en el DC (PS Remoting) - es un mecanismo distinto al
# que usa Get-WinEvent -ComputerName (que va por RPC), asi que puede fallar aunque
# la consulta de eventos si funcione. El fallo aqui no bloquea el resto del flujo.
# Generica: la usan las auditorias de equipos, OU y usuarios (cada una con su propia
# subcategoria: "Computer Account Management", "Directory Service Changes",
# "User Account Management").
function Test-AuditPolicySubcategory {
    param([string]$DC, [string]$Subcategory)
    try {
        $raw = Invoke-Command -ComputerName $DC -ScriptBlock {
                    param($sub)
                    auditpol /get /subcategory:"$sub" /r
                } -ArgumentList $Subcategory -ErrorAction Stop
        $csv  = $raw | ConvertFrom-Csv
        $incl = $csv.'Inclusion Setting'
        $estado = if ($incl -like "*Success*") { "OK" } else { "OFF" }
        [PSCustomObject]@{ DC = $DC; Estado = $estado; Detalle = $incl }
    } catch {
        [PSCustomObject]@{ DC = $DC; Estado = "ERROR"; Detalle = $_.Exception.Message }
    }
}

# Helper: descubre los DCs del dominio, pregunta cuales consultar y, opcionalmente,
# verifica en cada uno la subcategoria de auditoria indicada. Devuelve la lista de
# DCs a consultar. Compartida por las tres auditorias basadas en el log de Seguridad
# (equipos, OU, usuarios) para no repetir esta logica en cada una.
function Get-DCsParaAuditoria {
    param([string]$Subcategory)

    Write-Host ""
    Write-Host "  Descubriendo Domain Controllers..." -ForegroundColor DarkGray
    try {
        $todosDCs = (Get-ADDomainController -Filter *).HostName
        Write-Host "  DCs encontrados: $($todosDCs -join ', ')" -ForegroundColor DarkGray
    } catch {
        Write-Host "  ERROR obteniendo lista de DCs: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Se continuara consultando unicamente el DC local." -ForegroundColor Yellow
        $todosDCs = @()
    }
    $dcInput = Read-Host "  DC especifico a consultar (Enter = TODOS los DCs listados arriba)"
    $dcsAConsultar = if ([string]::IsNullOrWhiteSpace($dcInput)) {
        if ($todosDCs.Count -gt 0) { $todosDCs } else { @($null) }
    } else { @($dcInput) }

    Write-Host ""
    $verificarPol = Read-Host "  Verificar politica '$Subcategory' via WinRM? (S/N, Enter = N)"
    if ($verificarPol -match "^[sS]$") {
        Write-Host "  Verificando politica en cada DC..." -ForegroundColor DarkGray
        $estadosPolitica = foreach ($dc in $dcsAConsultar) {
            if ($dc) { Test-AuditPolicySubcategory -DC $dc -Subcategory $Subcategory }
        }
        if ($estadosPolitica) {
            Write-SubHeader "Estado de auditoria por DC"
            foreach ($e in $estadosPolitica) {
                switch ($e.Estado) {
                    "OK"    { Write-Badge "$($e.DC)  ->  $($e.Detalle)" "ok" }
                    "OFF"   { Write-Badge "$($e.DC)  ->  $($e.Detalle)" "warn" }
                    "ERROR" { Write-Badge "$($e.DC)  ->  ERROR WinRM: $($e.Detalle)" "error" }
                }
            }
            $dcsSinAuditoria = @($estadosPolitica | Where-Object { $_.Estado -eq "OFF" })
            if ($dcsSinAuditoria.Count -gt 0) {
                Write-Host ""
                Write-Host "  [!] En los DCs marcados OFF, si el cambio se proceso alli, los" -ForegroundColor Yellow
                Write-Host "  |   eventos correspondientes NO existen y no se podran recuperar." -ForegroundColor Yellow
            }
        } else {
            Write-Host "  No se pudo verificar (sin DC remoto valido para WinRM)." -ForegroundColor DarkGray
        }
    } else {
        Write-Host "  Verificacion de politica omitida - se pasa directo a la consulta de eventos." -ForegroundColor DarkGray
    }

    return $dcsAConsultar
}

# Helper: intenta determinar quien creo un objeto de AD (usuario o equipo),
# probando en cascada: ms-DS-CreatorSID -> propietario del objeto (ACL) ->
# primera ACE de creacion. $AdObj debe traer las propiedades DistinguishedName
# y "ms-DS-CreatorSID" ya cargadas. Generica: la usan las auditorias de
# equipos y de usuarios para no repetir esta logica en cada una.
function Resolve-CreadorObjetoAD {
    param($AdObj)

    $creadorNombre  = $null
    $creadorFuente  = $null
    $creadorEsGrupo = $false

    # METODO 1: ms-DS-CreatorSID — almacena el SID del creador en ciertos escenarios
    $creadorSID = $AdObj."ms-DS-CreatorSID"
    if ($creadorSID) {
        $sidStr = $null
        try { $sidStr = $creadorSID.Value } catch {}
        if (-not $sidStr) { try { $sidStr = (New-Object System.Security.Principal.SecurityIdentifier($creadorSID, 0)).Value } catch {} }

        if ($sidStr) {
            # Intentar resolver como usuario primero
            try {
                $objCreador = Get-ADUser -Filter "objectSid -eq '$sidStr'" `
                                -Properties SamAccountName, Name -ErrorAction Stop
                if ($objCreador) {
                    $creadorNombre = "$($objCreador.SamAccountName)  ($($objCreador.Name))"
                    $creadorFuente = "ms-DS-CreatorSID (usuario AD)"
                }
            } catch {}

            # Si no es usuario, buscar como cualquier objeto (puede ser cuenta de servicio/grupo)
            if (-not $creadorNombre) {
                try {
                    $objCreador = Get-ADObject -Filter "objectSid -eq '$sidStr'" `
                                    -Properties SamAccountName, Name, ObjectClass -ErrorAction Stop
                    if ($objCreador) {
                        $creadorNombre = "$($objCreador.SamAccountName)  ($($objCreador.Name))"
                        $creadorFuente = "ms-DS-CreatorSID ($($objCreador.ObjectClass))"
                        if ($objCreador.ObjectClass -eq "group") { $creadorEsGrupo = $true }
                    }
                } catch {}
            }

            # Fallback: traduccion SID -> NTAccount
            if (-not $creadorNombre) {
                try {
                    $ntAcct = (New-Object System.Security.Principal.SecurityIdentifier($sidStr)).Translate([System.Security.Principal.NTAccount]).Value
                    $creadorNombre = $ntAcct
                    $creadorFuente = "ms-DS-CreatorSID (traduccion SID)"
                } catch {
                    $creadorNombre = $sidStr
                    $creadorFuente = "ms-DS-CreatorSID (SID crudo)"
                }
            }
        }
    }

    # METODO 2: Propietario del objeto via Get-Acl
    if (-not $creadorNombre) {
        try {
            $acl = Get-Acl -Path ("AD:\" + $AdObj.DistinguishedName) -ErrorAction Stop
            $ownerStr = $acl.Owner
            if ($ownerStr -match "^S-1-") {
                try {
                    $ownerStr = (New-Object System.Security.Principal.SecurityIdentifier($ownerStr)).Translate([System.Security.Principal.NTAccount]).Value
                } catch {}
            }
            if ($ownerStr) {
                $creadorNombre = $ownerStr
                $creadorFuente = "Propietario del objeto (nTSecurityDescriptor)"
                if ($ownerStr -match "Enterprise Admins|Domain Admins|Admins del dominio") {
                    $creadorEsGrupo = $true
                }
            }
        } catch {}
    }

    # METODO 3: Primer ACE explicita de tipo Allow sobre el objeto
    if (-not $creadorNombre) {
        try {
            $acl2 = Get-Acl -Path ("AD:\" + $AdObj.DistinguishedName) -ErrorAction Stop
            $aceCreador = $acl2.Access |
                Where-Object { $_.ActiveDirectoryRights -match "GenericAll|CreateChild" `
                               -and $_.AccessControlType -eq "Allow" `
                               -and $_.IdentityReference -notmatch "SYSTEM|Domain Admins|Enterprise|Administrators" } |
                Select-Object -First 1
            if ($aceCreador) {
                $creadorNombre = $aceCreador.IdentityReference.Value
                $creadorFuente = "ACL del objeto (primer ACE de creacion)"
            }
        } catch {}
    }

    [PSCustomObject]@{
        Nombre  = $creadorNombre
        Fuente  = $creadorFuente
        EsGrupo = $creadorEsGrupo
    }
}

function Get-AuditoriaEquipo {
    Write-Header "AUDITORIA DE EQUIPO - HISTORIAL DE EVENTOS"

    Write-Host "  Consulta los eventos de creacion, modificacion y eliminacion" -ForegroundColor DarkGray
    Write-Host "  de una cuenta de equipo en el Registro de Seguridad del DC." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Eventos auditados:" -ForegroundColor DarkGray
    Write-Host "    4741  ->  Equipo dado de alta en el dominio (join)" -ForegroundColor DarkGray
    Write-Host "    4742  ->  Cuenta de equipo modificada" -ForegroundColor DarkGray
    Write-Host "    4743  ->  Equipo eliminado del dominio" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Estos eventos son LOCALES a cada DC (no se replican) y requieren" -ForegroundColor DarkGray
    Write-Host "  que la auditoria 'Computer Account Management' este activa." -ForegroundColor DarkGray
    Write-Host ""

    $hostname = (Read-Host "  Nombre del equipo a auditar (hostname)").Trim()
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        Write-Host "  Nombre no puede estar vacio." -ForegroundColor Red
        Pause-Pantalla; return
    }

    # ── Verificar objeto en AD ───────────────────────────────────────────
    Write-Host ""
    Write-Host "  Verificando objeto en AD..." -ForegroundColor DarkGray
    $adExiste = $false
    $adComp   = $null
    try {
        $adComp   = Get-ADComputer -Identity $hostname `
                        -Properties Created,Modified,DistinguishedName,OperatingSystem,Enabled,`
                                    "ms-DS-CreatorSID","ManagedBy","Description" `
                        -ErrorAction Stop
        $adExiste = $true

        Write-SubHeader "Estado actual en AD"
        Write-Campo "Nombre"       $adComp.Name  "White"
        $estEq = if ($adComp.Enabled) {"Habilitado"} else {"DESHABILITADO"}
        $colEq = if ($adComp.Enabled) {"Green"}      else {"Red"}
        Write-Campo "Estado"       $estEq $colEq
        Write-Campo "SO"           $(if ($adComp.OperatingSystem) {$adComp.OperatingSystem} else {"No registrado"})
        Write-Campo "Descripcion"  $(if ($adComp.Description)     {$adComp.Description}     else {"-"})
        Write-Campo "Creado en AD" $adComp.Created.ToString("dd/MM/yyyy HH:mm:ss")  "Cyan"
        Write-Campo "Modificado"   $adComp.Modified.ToString("dd/MM/yyyy HH:mm:ss") "Cyan"
        Write-Campo "OU"           (Get-OUdesdeDN $adComp.DistinguishedName)

        # ── Resolver creador (metodo compartido con auditoria de usuarios) ──
        Write-Host ""
        $creador        = Resolve-CreadorObjetoAD -AdObj $adComp
        $creadorNombre  = $creador.Nombre
        $creadorFuente  = $creador.Fuente
        $creadorEsGrupo = $creador.EsGrupo

        # ── Mostrar resultado ────────────────────────────────────────────
        Write-Host "  +-- Creador del objeto en AD " -NoNewline -ForegroundColor Green
        Write-Host ("-" * 30) -ForegroundColor DarkGray
        if ($creadorNombre) {
            $colCreador = if ($creadorEsGrupo) {"Yellow"} else {"Green"}
            Write-Campo "Unido al dominio por" $creadorNombre $colCreador
            Write-Host "  | Fuente: $creadorFuente" -ForegroundColor DarkGray
            Write-Host "  | Dato permanente en AD, no depende de la rotacion del log." -ForegroundColor DarkGray
            if ($creadorEsGrupo) {
                Write-Host ""
                Write-Host "  [!] El resultado muestra un grupo ($creadorNombre)." -ForegroundColor Yellow
                Write-Host "  | Windows asigna el grupo como propietario cuando un miembro de" -ForegroundColor DarkGray
                Write-Host "  | ese grupo realiza el join. Para identificar al usuario individual:" -ForegroundColor DarkGray
                Write-Host "  | - Consulta los eventos 4741 en el DC (opcion de eventos abajo)" -ForegroundColor DarkGray
                Write-Host "  | - O usa la opcion [11] para revisar el log de seguridad del DC." -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  No fue posible determinar el creador con los permisos actuales." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Para obtener este dato necesitas uno de los siguientes:" -ForegroundColor DarkGray
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
            Write-Host "  | 1. Ser miembro de Domain Admins o Account Operators" -ForegroundColor DarkGray
            Write-Host "  | 2. Tener delegado 'Read ms-DS-CreatorSID' en la OU del equipo" -ForegroundColor DarkGray
            Write-Host "  | 3. Revisar el propietario manualmente en ADUC:" -ForegroundColor DarkGray
            Write-Host "  |    Propiedades del equipo -> Seguridad -> Avanzado -> Propietario" -ForegroundColor DarkGray
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        }
        Write-Host ""
    } catch {
        Write-Host ""
        Write-Host "  '$hostname' NO existe actualmente en AD." -ForegroundColor Yellow
        Write-Host "  Puede haber sido eliminado -> consulta la opcion [10] AD Recycle Bin" -ForegroundColor DarkGray
    }

    # ── Descubrir DCs del dominio ────────────────────────────────────────
    # Los eventos de seguridad son LOCALES a cada DC: no se replican entre
    # ellos. Si el join/modificacion se proceso en un DC distinto al que
    # consultes, obtendras cero resultados aunque todo lo demas este bien.
    Write-Host ""
    Write-Host "  La consulta de eventos de seguridad puede tardar varios minutos" -ForegroundColor DarkGray
    Write-Host "  dependiendo del tamano del log y de cuantos DCs se consulten." -ForegroundColor DarkGray
    Write-Host ""
    $consultarEventos = Read-Host "  Consultar eventos de seguridad (4741/4742/4743)? (S/N)"

    $evCreate    = @()
    $evModify    = @()
    $evDelete    = @()
    $dcsConError = @()

    if ($consultarEventos -match "^[sS]$") {

        $dcsAConsultar = Get-DCsParaAuditoria -Subcategory "Computer Account Management"

        # ── Consultar 4741/4742/4743 en cada DC ──────────────────────────
        Write-Host ""
        Write-Host "  Consultando eventos de seguridad..." -ForegroundColor DarkGray
        foreach ($dc in $dcsAConsultar) {
            $dcLabel = if ($dc) { $dc } else { "DC local" }
            Write-Host "    -> $dcLabel ..." -ForegroundColor DarkGray -NoNewline

            $wevP = @{}
            if ($dc) { $wevP['ComputerName'] = $dc }

            $encontradosEnDC = 0
            $huboError       = $false

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4741} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-CompEvent $_ } |
                        Where-Object { $_ -and $_.Equipo -like $hostname })
                    $evCreate += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4742} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-CompEvent $_ } |
                        Where-Object { $_ -and $_.Equipo -like $hostname })
                    $evModify += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4743} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-CompEvent $_ } |
                        Where-Object { $_ -and $_.Equipo -like $hostname })
                    $evDelete += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            if ($huboError) {
                Write-Host "  ERROR" -ForegroundColor Red
                $dcsConError += $dcLabel
            } else {
                $col = if ($encontradosEnDC -gt 0) { "Green" } else { "DarkGray" }
                Write-Host "  $encontradosEnDC evento(s)" -ForegroundColor $col
            }
        }

        $evModify = @($evModify | Sort-Object Fecha -Descending)

        if ($dcsConError.Count -gt 0) {
            Write-Host ""
            Write-Host "  [!] No se pudo consultar el log de Seguridad en estos DCs:" -ForegroundColor Yellow
            foreach ($e in $dcsConError) { Write-Host "  | $e" -ForegroundColor DarkGray }
            Write-Host ""
            Write-Host "  Requisitos para leer los eventos:" -ForegroundColor Yellow
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
            $req = @(
                "1. GPO: 'Audit Computer Account Management' activo en el DC",
                "2. Permisos de lectura sobre el log de Seguridad del DC",
                "3. En DC remotos: firewall abierto (puerto 445 / RPC dinamico)",
                "4. Los eventos se depuran segun rotacion configurada del log"
            )
            foreach ($r in $req) { Write-Host "  | $r" -ForegroundColor DarkGray }
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        }

    } else {
        Write-Host "  Consulta de eventos omitida." -ForegroundColor DarkGray
    }

    # ── Seccion: Creacion ────────────────────────────────────────────────
    Write-SubHeader "Creacion / Join al dominio  (Event 4741)"
    if ($evCreate.Count -gt 0) {
        $evCreate | Select-Object `
            @{N="Fecha de alta";   E={$_.Fecha}},
            @{N="Dado de alta por"; E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de creacion en el log para '$hostname'." -ForegroundColor Yellow
        if ($adExiste) {
            Write-Host ("  Referencia AD: objeto creado el {0}" -f $adComp.Created.ToString("dd/MM/yyyy HH:mm:ss")) -ForegroundColor DarkGray
        }
        Write-Host "  (El evento 4741 pudo no haberse generado, haberse depurado por" -ForegroundColor DarkGray
        Write-Host "  rotacion del log, o haber ocurrido en un DC no consultado)" -ForegroundColor DarkGray
    }

    # ── Seccion: Modificaciones ──────────────────────────────────────────
    Write-SubHeader "Modificaciones  (Event 4742)"
    if ($evModify.Count -gt 0) {
        Write-Host ("  {0} modificacion(es) encontrada(s)  -  orden: mas reciente primero" -f $evModify.Count) -ForegroundColor Cyan
        Write-Host ""
        $evModify | Select-Object `
            @{N="Fecha de modificacion"; E={$_.Fecha}},
            @{N="Modificado por";        E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de modificacion en el log para '$hostname'." -ForegroundColor Yellow
    }

    # ── Seccion: Eliminacion ─────────────────────────────────────────────
    Write-SubHeader "Eliminacion del dominio  (Event 4743)"
    if ($evDelete.Count -gt 0) {
        $evDelete | Select-Object `
            @{N="Fecha de eliminacion"; E={$_.Fecha}},
            @{N="Eliminado por";        E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de eliminacion en el log para '$hostname'." -ForegroundColor Yellow
        if (-not $adExiste) {
            Write-Host "  El equipo no existe en AD: usa la opcion [10] AD Recycle Bin para" -ForegroundColor DarkGray
            Write-Host "  buscar el objeto eliminado y ver mas detalles." -ForegroundColor DarkGray
        }
    }

    if ($consultarEventos -match "^[sS]$" -and $evCreate.Count -eq 0 -and $evModify.Count -eq 0 -and $evDelete.Count -eq 0) {
        Write-Host ""
        Write-Host "  Posibles causas de no encontrar ningun evento:" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        Write-Host "  | 1. La auditoria 'Computer Account Management' no estaba" -ForegroundColor DarkGray
        Write-Host "  |    activa en el DC que proceso el cambio (verifica arriba)" -ForegroundColor DarkGray
        Write-Host "  | 2. El cambio se proceso en un DC que no quedo incluido en" -ForegroundColor DarkGray
        Write-Host "  |    la consulta (revisa la lista de DCs consultados arriba)" -ForegroundColor DarkGray
        Write-Host "  | 3. El log ya roto ese evento por tamano/retencion" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
    }

    # ── Exportar ─────────────────────────────────────────────────────────
    $todosEventos = @()
    foreach ($ev in $evCreate) { $todosEventos += [PSCustomObject]@{Tipo="Creacion";    Equipo=$ev.Equipo; Actor=$ev.Actor; Fecha=$ev.Fecha} }
    foreach ($ev in $evModify) { $todosEventos += [PSCustomObject]@{Tipo="Modificacion";Equipo=$ev.Equipo; Actor=$ev.Actor; Fecha=$ev.Fecha} }
    foreach ($ev in $evDelete) { $todosEventos += [PSCustomObject]@{Tipo="Eliminacion"; Equipo=$ev.Equipo; Actor=$ev.Actor; Fecha=$ev.Fecha} }

    Write-Separador
    Export-DatosCSV -Data $todosEventos -NombreArchivoBase "Auditoria_${hostname}" `
        -Prompt "  Exportar auditoria de '$hostname' a CSV? (S/N)" | Out-Null

    Pause-Pantalla
}

# ============================================================
#  MODULO AUDITORIA DE OU - QUIEN LA MOVIO (Evento 5139)
# ============================================================

# Helper: verifica si "Directory Service Changes" esta auditando Success en un DC.
# Requiere WinRM habilitado en el DC (PS Remoting) - es un mecanismo distinto al
# que usa Get-WinEvent -ComputerName (que va por RPC), asi que puede fallar aunque
# la consulta de eventos si funcione. El fallo aqui no bloquea el resto del flujo.
# Helper: parsea el evento 5139 (objeto de directorio movido) via XML
function Parse-OUMoveEvent {
    param($ev, [string]$DCOrigen)
    try {
        $xml   = [xml]$ev.ToXml()
        $xdata = $xml.Event.EventData.Data
        $actor    = ($xdata | Where-Object { $_.Name -eq "SubjectUserName"   }).'#text'
        $dom      = ($xdata | Where-Object { $_.Name -eq "SubjectDomainName" }).'#text'
        $objDN    = ($xdata | Where-Object { $_.Name -eq "ObjectDN"         }).'#text'
        $objClase = ($xdata | Where-Object { $_.Name -eq "ObjectClass"      }).'#text'
        $oldTree  = ($xdata | Where-Object { $_.Name -eq "OldTree"          }).'#text'
        $newTree  = ($xdata | Where-Object { $_.Name -eq "NewTree"          }).'#text'

        [PSCustomObject]@{
            Fecha       = $ev.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
            Actor       = if ($actor -and $dom) { "$dom\$actor" } elseif ($actor) { $actor } else { "Desconocido" }
            ObjectDN    = $objDN
            ObjectClass = $objClase
            Desde       = $oldTree
            Hacia       = $newTree
            DC          = $DCOrigen
        }
    } catch { $null }
}

function Get-AuditoriaOU {
    Write-Header "AUDITORIA DE OU - QUIEN LA MOVIO"

    Write-Host "  Busca el evento 5139 (objeto de directorio movido) en el log de" -ForegroundColor DarkGray
    Write-Host "  Seguridad de los DCs. Requiere que 'Audit Directory Service" -ForegroundColor DarkGray
    Write-Host "  Changes' haya estado activo ANTES del movimiento; si no lo estaba" -ForegroundColor DarkGray
    Write-Host "  en el DC que proceso el cambio, el evento no existe y no se puede" -ForegroundColor DarkGray
    Write-Host "  recuperar de forma retroactiva." -ForegroundColor DarkGray
    Write-Host ""

    $termino = (Read-Host "  Nombre o fragmento de la OU a buscar (ej: Ventas)").Trim()
    if ([string]::IsNullOrWhiteSpace($termino)) {
        Write-Host "  El termino de busqueda no puede estar vacio." -ForegroundColor Red
        Pause-Pantalla; return
    }

    $diasInput = Read-Host "  Buscar en los ultimos cuantos dias? (Enter = 30)"
    $dias = 30
    if ($diasInput -match '^\d+$') { $dias = [int]$diasInput }
    $desde = (Get-Date).AddDays(-$dias)

    $dcsAConsultar = Get-DCsParaAuditoria -Subcategory "Directory Service Changes"

    # ── Consultar evento 5139 en cada DC ─────────────────────────────────
    Write-Host ""
    Write-Host "  Consultando evento 5139 desde $($desde.ToString('dd/MM/yyyy')) ..." -ForegroundColor DarkGray
    $resultados  = @()
    $dcsConError = @()

    foreach ($dc in $dcsAConsultar) {
        Write-Host "    -> $dc ..." -ForegroundColor DarkGray -NoNewline
        try {
            $raw = Get-WinEvent -ComputerName $dc -FilterHashtable @{
                        LogName   = 'Security'
                        Id        = 5139
                        StartTime = $desde
                    } -ErrorAction Stop

            $parseados = @($raw | ForEach-Object { Parse-OUMoveEvent -ev $_ -DCOrigen $dc })
            $coincidencias = @($parseados | Where-Object {
                $_ -and (
                    ($_.ObjectDN -and $_.ObjectDN -like "*$termino*") -or
                    ($_.Desde    -and $_.Desde    -like "*$termino*") -or
                    ($_.Hacia    -and $_.Hacia    -like "*$termino*")
                )
            })
            $resultados += $coincidencias
            $col = if ($coincidencias.Count -gt 0) { "Green" } else { "DarkGray" }
            Write-Host "  $($coincidencias.Count) coincidencia(s)" -ForegroundColor $col
        } catch {
            if ($_.Exception.Message -match "No events were found") {
                Write-Host "  sin eventos 5139 en el rango" -ForegroundColor DarkGray
            } else {
                Write-Host "  ERROR" -ForegroundColor Red
                $dcsConError += "$dc  ->  $($_.Exception.Message)"
            }
        }
    }

    # ── Resultados ────────────────────────────────────────────────────────
    Write-SubHeader "Resultados - movimientos que coinciden con '$termino'"
    if ($resultados.Count -gt 0) {
        $resultados | Sort-Object Fecha -Descending | Select-Object `
            @{N="Fecha";      E={$_.Fecha}},
            @{N="Movido por"; E={$_.Actor}},
            @{N="Desde";      E={$_.Desde}},
            @{N="Hacia";      E={$_.Hacia}},
            @{N="DC (log)";   E={$_.DC}} |
        Format-Table -AutoSize -Wrap
    } else {
        Write-Host "  No se encontraron eventos 5139 que coincidan con '$termino'." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Posibles causas:" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        Write-Host "  | 1. La auditoria 'Directory Service Changes' no estaba" -ForegroundColor DarkGray
        Write-Host "  |    activa en el DC que proceso el movimiento (ver arriba)" -ForegroundColor DarkGray
        Write-Host "  | 2. El movimiento ocurrio fuera del rango de $dias dia(s)" -ForegroundColor DarkGray
        Write-Host "  | 3. El termino de busqueda no coincide con el DN real" -ForegroundColor DarkGray
        Write-Host "  |    (revisa mayusculas/typos, o prueba con otro fragmento)" -ForegroundColor DarkGray
        Write-Host "  | 4. El log ya roto ese evento por tamano/retencion" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
    }

    if ($dcsConError.Count -gt 0) {
        Write-Host ""
        Write-Host "  [!] No se pudo consultar el log en estos DCs:" -ForegroundColor Yellow
        foreach ($e in $dcsConError) { Write-Host "  | $e" -ForegroundColor DarkGray }
    }

    # ── Exportar ─────────────────────────────────────────────────────────
    $datosExport = @($resultados | Sort-Object Fecha -Descending |
        Select-Object Fecha, Actor, ObjectDN, ObjectClass, Desde, Hacia, DC)
    Write-Separador
    Export-DatosCSV -Data $datosExport -NombreArchivoBase "AuditoriaOU_${termino}" `
        -Prompt "  Exportar $($resultados.Count) resultado(s) a CSV? (S/N)" | Out-Null

    Pause-Pantalla
}

# Helper: parsea eventos 4720/4738/4726 (cuenta de usuario creada/modificada/eliminada)
function Parse-UserEvent {
    param($ev)
    try {
        $xml   = [xml]$ev.ToXml()
        $xdata = $xml.Event.EventData.Data
        $usr   = ($xdata | Where-Object { $_.Name -eq "TargetUserName"    }).'#text'
        $actor = ($xdata | Where-Object { $_.Name -eq "SubjectUserName"   }).'#text'
        $dom   = ($xdata | Where-Object { $_.Name -eq "SubjectDomainName" }).'#text'
        [PSCustomObject]@{
            Usuario = $usr
            Actor   = if ($actor -and $dom) { "$dom\$actor" } elseif ($actor) { $actor } else { "Desconocido" }
            Fecha   = $ev.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
        }
    } catch { $null }
}

function Get-AuditoriaUsuario {
    Write-Header "AUDITORIA DE USUARIO - HISTORIAL DE EVENTOS"

    Write-Host "  Consulta los eventos de creacion, modificacion y eliminacion" -ForegroundColor DarkGray
    Write-Host "  de una cuenta de usuario en el Registro de Seguridad del DC." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Eventos auditados:" -ForegroundColor DarkGray
    Write-Host "    4720  ->  Cuenta de usuario creada" -ForegroundColor DarkGray
    Write-Host "    4738  ->  Cuenta de usuario modificada" -ForegroundColor DarkGray
    Write-Host "    4726  ->  Cuenta de usuario eliminada" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Estos eventos son LOCALES a cada DC (no se replican) y requieren" -ForegroundColor DarkGray
    Write-Host "  que la auditoria 'User Account Management' este activa." -ForegroundColor DarkGray
    Write-Host ""

    $samAccountName = (Read-Host "  SamAccountName del usuario a auditar").Trim()
    if ([string]::IsNullOrWhiteSpace($samAccountName)) {
        Write-Host "  Nombre no puede estar vacio." -ForegroundColor Red
        Pause-Pantalla; return
    }

    # ── Verificar objeto en AD ───────────────────────────────────────────
    Write-Host ""
    Write-Host "  Verificando objeto en AD..." -ForegroundColor DarkGray
    $adExiste = $false
    $adUsr    = $null
    try {
        $adUsr    = Get-ADUser -Identity $samAccountName `
                        -Properties Created,Modified,DistinguishedName,Enabled,`
                                    "ms-DS-CreatorSID","Description","UserPrincipalName" `
                        -ErrorAction Stop
        $adExiste = $true

        Write-SubHeader "Estado actual en AD"
        Write-Campo "Nombre"       $adUsr.Name  "White"
        $estUsr = if ($adUsr.Enabled) {"Habilitado"} else {"DESHABILITADO"}
        $colUsr = if ($adUsr.Enabled) {"Green"}      else {"Red"}
        Write-Campo "Estado"       $estUsr $colUsr
        Write-Campo "UPN"          $(if ($adUsr.UserPrincipalName) {$adUsr.UserPrincipalName} else {"-"})
        Write-Campo "Descripcion"  $(if ($adUsr.Description)       {$adUsr.Description}       else {"-"})
        Write-Campo "Creado en AD" $adUsr.Created.ToString("dd/MM/yyyy HH:mm:ss")  "Cyan"
        Write-Campo "Modificado"   $adUsr.Modified.ToString("dd/MM/yyyy HH:mm:ss") "Cyan"
        Write-Campo "OU"           (Get-OUdesdeDN $adUsr.DistinguishedName)

        # ── Resolver creador (metodo compartido con auditoria de equipos) ──
        Write-Host ""
        $creador        = Resolve-CreadorObjetoAD -AdObj $adUsr
        $creadorNombre  = $creador.Nombre
        $creadorFuente  = $creador.Fuente
        $creadorEsGrupo = $creador.EsGrupo

        # ── Mostrar resultado ────────────────────────────────────────────
        Write-Host "  +-- Creador del objeto en AD " -NoNewline -ForegroundColor Green
        Write-Host ("-" * 30) -ForegroundColor DarkGray
        if ($creadorNombre) {
            $colCreador = if ($creadorEsGrupo) {"Yellow"} else {"Green"}
            Write-Campo "Creado por" $creadorNombre $colCreador
            Write-Host "  | Fuente: $creadorFuente" -ForegroundColor DarkGray
            Write-Host "  | Dato permanente en AD, no depende de la rotacion del log." -ForegroundColor DarkGray
            if ($creadorEsGrupo) {
                Write-Host ""
                Write-Host "  [!] El resultado muestra un grupo ($creadorNombre)." -ForegroundColor Yellow
                Write-Host "  | Windows asigna el grupo como propietario cuando un miembro de" -ForegroundColor DarkGray
                Write-Host "  | ese grupo creo la cuenta. Para identificar al usuario individual:" -ForegroundColor DarkGray
                Write-Host "  | - Consulta el evento 4720 en el DC (opcion de eventos abajo)" -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  No fue posible determinar el creador con los permisos actuales." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Para obtener este dato necesitas uno de los siguientes:" -ForegroundColor DarkGray
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
            Write-Host "  | 1. Ser miembro de Domain Admins o Account Operators" -ForegroundColor DarkGray
            Write-Host "  | 2. Tener delegado 'Read ms-DS-CreatorSID' en la OU del usuario" -ForegroundColor DarkGray
            Write-Host "  | 3. Revisar el propietario manualmente en ADUC:" -ForegroundColor DarkGray
            Write-Host "  |    Propiedades del usuario -> Seguridad -> Avanzado -> Propietario" -ForegroundColor DarkGray
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        }
        Write-Host ""
    } catch {
        Write-Host ""
        Write-Host "  '$samAccountName' NO existe actualmente en AD." -ForegroundColor Yellow
        Write-Host "  Puede haber sido eliminado -> consulta la opcion [10] AD Recycle Bin" -ForegroundColor DarkGray
    }

    # ── Descubrir DCs y consultar eventos ────────────────────────────────
    # Los eventos de seguridad son LOCALES a cada DC: no se replican entre
    # ellos. Si la creacion/modificacion se proceso en un DC distinto al que
    # consultes, obtendras cero resultados aunque todo lo demas este bien.
    Write-Host ""
    Write-Host "  La consulta de eventos de seguridad puede tardar varios minutos" -ForegroundColor DarkGray
    Write-Host "  dependiendo del tamano del log y de cuantos DCs se consulten." -ForegroundColor DarkGray
    Write-Host ""
    $consultarEventos = Read-Host "  Consultar eventos de seguridad (4720/4738/4726)? (S/N)"

    $evCreate    = @()
    $evModify    = @()
    $evDelete    = @()
    $dcsConError = @()

    if ($consultarEventos -match "^[sS]$") {

        $dcsAConsultar = Get-DCsParaAuditoria -Subcategory "User Account Management"

        # ── Consultar 4720/4738/4726 en cada DC ──────────────────────────
        Write-Host ""
        Write-Host "  Consultando eventos de seguridad..." -ForegroundColor DarkGray
        foreach ($dc in $dcsAConsultar) {
            $dcLabel = if ($dc) { $dc } else { "DC local" }
            Write-Host "    -> $dcLabel ..." -ForegroundColor DarkGray -NoNewline

            $wevP = @{}
            if ($dc) { $wevP['ComputerName'] = $dc }

            $encontradosEnDC = 0
            $huboError       = $false

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4720} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-UserEvent $_ } |
                        Where-Object { $_ -and $_.Usuario -like $samAccountName })
                    $evCreate += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4738} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-UserEvent $_ } |
                        Where-Object { $_ -and $_.Usuario -like $samAccountName })
                    $evModify += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            try {
                $raw = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4726} @wevP -ErrorAction Stop
                if ($raw) {
                    $parsed = @($raw | ForEach-Object { Parse-UserEvent $_ } |
                        Where-Object { $_ -and $_.Usuario -like $samAccountName })
                    $evDelete += $parsed
                    $encontradosEnDC += $parsed.Count
                }
            } catch {
                if ($_.Exception.Message -notmatch "No events|no se encontr") { $huboError = $true }
            }

            if ($huboError) {
                Write-Host "  ERROR" -ForegroundColor Red
                $dcsConError += $dcLabel
            } else {
                $col = if ($encontradosEnDC -gt 0) { "Green" } else { "DarkGray" }
                Write-Host "  $encontradosEnDC evento(s)" -ForegroundColor $col
            }
        }

        $evModify = @($evModify | Sort-Object Fecha -Descending)

        if ($dcsConError.Count -gt 0) {
            Write-Host ""
            Write-Host "  [!] No se pudo consultar el log de Seguridad en estos DCs:" -ForegroundColor Yellow
            foreach ($e in $dcsConError) { Write-Host "  | $e" -ForegroundColor DarkGray }
            Write-Host ""
            Write-Host "  Requisitos para leer los eventos:" -ForegroundColor Yellow
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
            $req = @(
                "1. GPO: 'Audit User Account Management' activo en el DC",
                "2. Permisos de lectura sobre el log de Seguridad del DC",
                "3. En DC remotos: firewall abierto (puerto 445 / RPC dinamico)",
                "4. Los eventos se depuran segun rotacion configurada del log"
            )
            foreach ($r in $req) { Write-Host "  | $r" -ForegroundColor DarkGray }
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        }

    } else {
        Write-Host "  Consulta de eventos omitida." -ForegroundColor DarkGray
    }

    # ── Seccion: Creacion ────────────────────────────────────────────────
    Write-SubHeader "Creacion de la cuenta  (Event 4720)"
    if ($evCreate.Count -gt 0) {
        $evCreate | Select-Object `
            @{N="Fecha de alta"; E={$_.Fecha}},
            @{N="Creado por";    E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de creacion en el log para '$samAccountName'." -ForegroundColor Yellow
        if ($adExiste) {
            Write-Host ("  Referencia AD: objeto creado el {0}" -f $adUsr.Created.ToString("dd/MM/yyyy HH:mm:ss")) -ForegroundColor DarkGray
        }
        Write-Host "  (El evento 4720 pudo no haberse generado, haberse depurado por" -ForegroundColor DarkGray
        Write-Host "  rotacion del log, o haber ocurrido en un DC no consultado)" -ForegroundColor DarkGray
    }

    # ── Seccion: Modificaciones ──────────────────────────────────────────
    Write-SubHeader "Modificaciones  (Event 4738)"
    if ($evModify.Count -gt 0) {
        Write-Host ("  {0} modificacion(es) encontrada(s)  -  orden: mas reciente primero" -f $evModify.Count) -ForegroundColor Cyan
        Write-Host ""
        $evModify | Select-Object `
            @{N="Fecha de modificacion"; E={$_.Fecha}},
            @{N="Modificado por";        E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de modificacion en el log para '$samAccountName'." -ForegroundColor Yellow
    }

    # ── Seccion: Eliminacion ─────────────────────────────────────────────
    Write-SubHeader "Eliminacion de la cuenta  (Event 4726)"
    if ($evDelete.Count -gt 0) {
        $evDelete | Select-Object `
            @{N="Fecha de eliminacion"; E={$_.Fecha}},
            @{N="Eliminado por";        E={$_.Actor}} |
        Format-Table -AutoSize
    } else {
        Write-Host "  Sin registros de eliminacion en el log para '$samAccountName'." -ForegroundColor Yellow
        if (-not $adExiste) {
            Write-Host "  El usuario no existe en AD: usa la opcion [10] AD Recycle Bin para" -ForegroundColor DarkGray
            Write-Host "  buscar el objeto eliminado y ver mas detalles." -ForegroundColor DarkGray
        }
    }

    if ($consultarEventos -match "^[sS]$" -and $evCreate.Count -eq 0 -and $evModify.Count -eq 0 -and $evDelete.Count -eq 0) {
        Write-Host ""
        Write-Host "  Posibles causas de no encontrar ningun evento:" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
        Write-Host "  | 1. La auditoria 'User Account Management' no estaba" -ForegroundColor DarkGray
        Write-Host "  |    activa en el DC que proceso el cambio (verifica arriba)" -ForegroundColor DarkGray
        Write-Host "  | 2. El cambio se proceso en un DC que no quedo incluido en" -ForegroundColor DarkGray
        Write-Host "  |    la consulta (revisa la lista de DCs consultados arriba)" -ForegroundColor DarkGray
        Write-Host "  | 3. El log ya roto ese evento por tamano/retencion" -ForegroundColor DarkGray
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkGray
    }

    # ── Exportar ─────────────────────────────────────────────────────────
    $todosEventos = @()
    foreach ($ev in $evCreate) { $todosEventos += [PSCustomObject]@{Tipo="Creacion";     Usuario=$ev.Usuario; Actor=$ev.Actor; Fecha=$ev.Fecha} }
    foreach ($ev in $evModify) { $todosEventos += [PSCustomObject]@{Tipo="Modificacion"; Usuario=$ev.Usuario; Actor=$ev.Actor; Fecha=$ev.Fecha} }
    foreach ($ev in $evDelete) { $todosEventos += [PSCustomObject]@{Tipo="Eliminacion";  Usuario=$ev.Usuario; Actor=$ev.Actor; Fecha=$ev.Fecha} }

    Write-Separador
    Export-DatosCSV -Data $todosEventos -NombreArchivoBase "AuditoriaUsuario_${samAccountName}" `
        -Prompt "  Exportar auditoria de '$samAccountName' a CSV? (S/N)" | Out-Null

    Pause-Pantalla
}

# ============================================================
#  MODULO EXPORTACION - AUDITORIA
# ============================================================

# Columnas estandar de usuario para todos los reportes
function Get-ColsUsuario {
    param($u)
    [PSCustomObject]@{
        Nombre           = $u.Name
        Usuario          = $u.SamAccountName
        Email            = if ($u.EmailAddress)  { $u.EmailAddress }  else { "-" }
        Departamento     = if ($u.Department)     { $u.Department }     else { "-" }
        Cargo            = if ($u.Title)          { $u.Title }          else { "-" }
        Estado           = if ($u.Enabled)        { "Activa" }          else { "Deshabilitada" }
        Bloqueado        = if ($u.LockedOut)       { "SI" }             else { "No" }
        PassExpirada     = if ($u.PasswordExpired) { "SI" }             else { "No" }
        PassNuncaExpira  = if ($u.PasswordNeverExpires) { "SI" }        else { "No" }
        AdminCount       = $u.adminCount
        UltimoLogon      = if ($u.LastLogonDate)  { $u.LastLogonDate.ToString("dd/MM/yyyy HH:mm") } else { "Nunca" }
        FechaCreacion    = if ($u.Created)         { $u.Created.ToString("dd/MM/yyyy") }            else { "-" }
        OU               = Get-OUdesdeDN $u.DistinguishedName
        DN               = $u.DistinguishedName
    }
}

# Columnas estandar de equipo para todos los reportes
function Get-ColsEquipo {
    param($e)
    [PSCustomObject]@{
        Nombre          = $e.Name
        DNS             = if ($e.DNSHostName)           { $e.DNSHostName }           else { "-" }
        SistemaOperativo= if ($e.OperatingSystem)       { $e.OperatingSystem }       else { "-" }
        Version         = if ($e.OperatingSystemVersion){ $e.OperatingSystemVersion } else { "-" }
        Estado          = if ($e.Enabled)               { "Habilitado" }             else { "Deshabilitado" }
        UltimoLogon     = if ($e.LastLogonDate)         { $e.LastLogonDate.ToString("dd/MM/yyyy") } else { "Nunca" }
        FechaCreacion   = if ($e.Created)               { $e.Created.ToString("dd/MM/yyyy") }       else { "-" }
        PasswordLastSet = if ($e.PasswordLastSet)       { $e.PasswordLastSet.ToString("dd/MM/yyyy") } else { "-" }
        OU              = Get-OUdesdeDN $e.DistinguishedName
        DN              = $e.DistinguishedName
    }
}

function Export-Reporte {
    Write-Header "AUDITORIA - REPORTES DE SEGURIDAD"

    Write-Host "  Se mostraran los resultados en pantalla." -ForegroundColor DarkGray
    Write-Host "  Al finalizar podras exportar a CSV si lo deseas." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  --- USUARIOS -----------------------------------------" -ForegroundColor Yellow
    Write-Host "  [1]  Usuarios bloqueados"
    Write-Host "  [2]  Usuarios deshabilitados"
    Write-Host "  [3]  Password expirado"
    Write-Host "  [4]  Password que NUNCA expira  (cuentas de servicio u omisiones)"
    Write-Host "  [5]  Usuarios inactivos >90 dias"
    Write-Host "  [6]  Cuentas con privilegios elevados  (adminCount = 1)"
    Write-Host ""
    Write-Host "  --- EQUIPOS ------------------------------------------" -ForegroundColor Yellow
    Write-Host "  [7]  Equipos sin OU  (contenedor Computers por defecto)"
    Write-Host ""
    Write-Host "  --- GRUPOS -------------------------------------------" -ForegroundColor Yellow
    Write-Host "  [8]  Grupos vacios"
    Write-Host ""

    $opc = Read-Host "  Opcion"

    $propsU = @("Name","SamAccountName","EmailAddress","Department","Title","Enabled",
                "LockedOut","PasswordExpired","PasswordNeverExpires","PasswordLastSet",
                "LastLogonDate","Created","DistinguishedName","adminCount","Description")
    $propsC = @("Name","DNSHostName","OperatingSystem","OperatingSystemVersion",
                "Enabled","LastLogonDate","Created","DistinguishedName","PasswordLastSet")

    $fecha90 = (Get-Date).AddDays(-90)

    # Titulo descriptivo para mostrar en pantalla
    $titulos = @{
        "1" = "USUARIOS BLOQUEADOS"
        "2" = "USUARIOS DESHABILITADOS"
        "3" = "USUARIOS CON PASSWORD EXPIRADO"
        "4" = "USUARIOS CON PASSWORD QUE NUNCA EXPIRA"
        "5" = "USUARIOS INACTIVOS (>90 DIAS SIN LOGON)"
        "6" = "CUENTAS CON PRIVILEGIOS ELEVADOS (adminCount=1)"
        "7" = "EQUIPOS SIN OU (contenedor Computers por defecto)"
        "8" = "GRUPOS VACIOS"
    }

    try {
        $datos   = $null
        $tipoObj = "usuario"

        switch ($opc) {

            # ── USUARIOS ──────────────────────────────────────────────────
            "1" {
                $pdc  = Get-PDCEmulator
                $srvP = if ($pdc) { @{ Server = $pdc } } else { @{} }
                if ($pdc) {
                    Write-Host "  Buscando usuarios bloqueados (contra el PDC Emulator: $pdc)..." -ForegroundColor DarkGray
                } else {
                    Write-Host "  [!] No se pudo determinar el PDC Emulator; usando el DC por defecto." -ForegroundColor Yellow
                    Write-Host "      El estado de bloqueo podria no ser el mas reciente en entornos multi-DC." -ForegroundColor Yellow
                    Write-Host "  Buscando usuarios bloqueados..." -ForegroundColor DarkGray
                }
                $datos = Search-ADAccount @srvP -LockedOut -UsersOnly |
                    Get-ADUser @srvP -Properties $propsU | Sort-Object Name |
                    ForEach-Object { Get-ColsUsuario $_ }
            }
            "2" {
                Write-Host "  Buscando usuarios deshabilitados..." -ForegroundColor DarkGray
                $datos = Get-ADUser -Filter "Enabled -eq '$false'" -Properties $propsU |
                    Sort-Object Name | ForEach-Object { Get-ColsUsuario $_ }
            }
            "3" {
                Write-Host "  Buscando usuarios con password expirado..." -ForegroundColor DarkGray
                $datos = Search-ADAccount -PasswordExpired -UsersOnly |
                    Get-ADUser -Properties $propsU | Sort-Object Name |
                    ForEach-Object { Get-ColsUsuario $_ }
            }
            "4" {
                Write-Host "  Buscando usuarios con password que nunca expira..." -ForegroundColor DarkGray
                $datos = Get-ADUser -Filter "PasswordNeverExpires -eq '$true' -and Enabled -eq '$true'" -Properties $propsU |
                    Sort-Object Name | ForEach-Object { Get-ColsUsuario $_ }
            }
            "5" {
                Write-Host "  Buscando usuarios inactivos (>90 dias)..." -ForegroundColor DarkGray
                $fechaStr = $fecha90.ToFileTime()
                $datos = Get-ADUser -Filter "Enabled -eq '$true' -and (LastLogonTimestamp -lt $fechaStr -or LastLogonTimestamp -notlike '*')" -Properties $propsU |
                    Sort-Object LastLogonDate | ForEach-Object { Get-ColsUsuario $_ }
            }
            "6" {
                Write-Host "  Buscando cuentas con adminCount=1..." -ForegroundColor DarkGray
                $datos = Get-ADUser -Filter "adminCount -eq 1" -Properties $propsU |
                    Sort-Object Name | ForEach-Object { Get-ColsUsuario $_ }
            }

            # ── EQUIPOS ───────────────────────────────────────────────────
            "7" {
                $tipoObj = "equipo"
                Write-Host "  Buscando equipos en el contenedor Computers (sin OU)..." -ForegroundColor DarkGray
                $datos = Get-ADComputer -Filter * -Properties $propsC |
                    Where-Object { $_.DistinguishedName -like "*CN=Computers,DC=*" } |
                    Sort-Object Name | ForEach-Object { Get-ColsEquipo $_ }
            }

            # ── GRUPOS ────────────────────────────────────────────────────
            "8" {
                $tipoObj = "grupo"
                Write-Host "  Buscando grupos vacios (puede tardar)..." -ForegroundColor DarkGray
                $datos = Get-ADGroup -Filter * -Properties Members, Description, Created, DistinguishedName |
                    Where-Object { $_.Members.Count -eq 0 } | Sort-Object Name |
                    ForEach-Object {
                        [PSCustomObject]@{
                            Nombre        = $_.Name
                            SamAccount    = $_.SamAccountName
                            Tipo          = $_.GroupScope
                            Categoria     = $_.GroupCategory
                            Descripcion   = if ($_.Description) { $_.Description } else { "-" }
                            FechaCreacion = if ($_.Created) { $_.Created.ToString("dd/MM/yyyy") } else { "-" }
                            OU            = Get-OUdesdeDN $_.DistinguishedName
                            DN            = $_.DistinguishedName
                        }
                    }
            }

            default {
                Write-Host "  Opcion no valida." -ForegroundColor Red
                Pause-Pantalla; return
            }
        }

        # ── Mostrar resultados en pantalla ────────────────────────────────
        $titulo = $titulos[$opc]
        $total  = if ($datos) { @($datos).Count } else { 0 }

        Write-Host ""
        Write-Host "  +============================================================+" -ForegroundColor Cyan
        Write-Host ("  |  $titulo".PadRight(63) + "|") -ForegroundColor White
        Write-Host ("  |  Total encontrados: $total".PadRight(63) + "|") -ForegroundColor $(if ($total -gt 0) {"Yellow"} else {"Green"})
        Write-Host "  +============================================================+" -ForegroundColor Cyan

        if ($total -eq 0) {
            Write-Host ""
            Write-Host "  No se encontraron resultados para este filtro." -ForegroundColor Green
            Pause-Pantalla; return
        }

        Write-Host ""

        # Vista en tabla segun tipo de objeto
        switch ($tipoObj) {
            "usuario" {
                $datos | Select-Object Nombre, Usuario, Departamento, Estado, Bloqueado,
                    PassExpirada, PassNuncaExpira, AdminCount, UltimoLogon, FechaCreacion, OU |
                Format-Table -AutoSize -Wrap
            }
            "equipo" {
                $datos | Select-Object Nombre, DNS, SistemaOperativo, Estado,
                    UltimoLogon, FechaCreacion, OU |
                Format-Table -AutoSize -Wrap
            }
            "grupo" {
                $datos | Select-Object Nombre, Tipo, Categoria, Descripcion, FechaCreacion, OU |
                Format-Table -AutoSize -Wrap
            }
        }

        # ── Desbloqueo interactivo (solo reporte de bloqueados) ──────────────
        if ($opc -eq "1" -and $total -gt 0) {
            Write-Separador
            Write-Host ""
            Write-Host "  ACCION DE DESBLOQUEO" -ForegroundColor Yellow
            Write-Host "  [A] Desbloquear TODOS los usuarios del reporte"
            Write-Host "  [S] Seleccionar usuarios a desbloquear"
            Write-Host "  [W] Simular (WhatIf) sin aplicar cambios"
            Write-Host "  [N] Omitir"
            Write-Host ""
            $opcDesbloqueo = Read-Host "  Opcion"

            if ($opcDesbloqueo -match "^[wW]$") {
                Write-Host ""
                Write-Host "  [SIMULACION] Los siguientes usuarios serian desbloqueados:" -ForegroundColor Yellow
                foreach ($usr in $datos) {
                    Unlock-ADAccount @srvP -Identity $usr.Usuario -WhatIf
                }
                Write-Host ""
                Write-Host "  Ninguna cuenta fue modificada. Repite y confirma con A o S para aplicar." -ForegroundColor DarkGray

            } elseif ($opcDesbloqueo -match "^[aA]$") {
                Write-Host ""
                $conf = Read-Host "  Confirmar desbloqueo de $total usuarios? (S/N)"
                if ($conf -match "^[sS]$") {
                    Write-Host ""
                    # Re-consultar en tiempo real para obtener el estado actual de bloqueo
                    # (evita problemas de cache del reporte original)
                    Write-Host "  Verificando estado actual de bloqueo en AD..." -ForegroundColor DarkGray
                    $exitosos = 0; $fallidos = 0; $yaLibres = 0
                    foreach ($usr in $datos) {
                        try {
                            $estadoActual = Get-ADUser @srvP -Identity $usr.Usuario -Properties LockedOut -ErrorAction Stop
                            if ($estadoActual.LockedOut) {
                                Unlock-ADAccount @srvP -Identity $usr.Usuario -ErrorAction Stop
                                Write-Host ("  [ OK ] {0,-25} {1}" -f $usr.Usuario, $usr.Nombre) -ForegroundColor Green
                                $exitosos++
                            } else {
                                Write-Host ("  [SKIP] {0,-25} ya desbloqueada" -f $usr.Usuario) -ForegroundColor DarkGray
                                $yaLibres++
                            }
                        } catch {
                            Write-Host ("  [ ERR] {0,-25} {1}" -f $usr.Usuario, $_.Exception.Message) -ForegroundColor Red
                            $fallidos++
                        }
                    }
                    Write-Host ""
                    Write-Host ("  Resultado: {0} desbloqueadas, {1} ya estaban libres, {2} con error." -f $exitosos, $yaLibres, $fallidos) -ForegroundColor Cyan
                } else {
                    Write-Host ""
                    Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
                }

            } elseif ($opcDesbloqueo -match "^[sS]$") {
                Write-Host ""
                $arrDatos = @($datos)
                for ($i = 0; $i -lt $arrDatos.Count; $i++) {
                    $usr = $arrDatos[$i]
                    Write-Host ("  [{0,2}] {1,-25} {2}" -f ($i + 1), $usr.Usuario, $usr.Nombre) -ForegroundColor White
                }
                Write-Host ""
                $seleccion = Read-Host "  Numeros a desbloquear separados por coma (ej: 1,3,5)"
                $indices = $seleccion -split "," |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { $_ -match "^\d+$" }
                Write-Host ""
                $exitosos = 0; $fallidos = 0; $yaLibres = 0
                foreach ($idx in $indices) {
                    $num = [int]$idx - 1
                    if ($num -ge 0 -and $num -lt $arrDatos.Count) {
                        $usr = $arrDatos[$num]
                        try {
                            $estadoActual = Get-ADUser @srvP -Identity $usr.Usuario -Properties LockedOut -ErrorAction Stop
                            if ($estadoActual.LockedOut) {
                                Unlock-ADAccount @srvP -Identity $usr.Usuario -ErrorAction Stop
                                Write-Host ("  [ OK ] {0,-25} {1}" -f $usr.Usuario, $usr.Nombre) -ForegroundColor Green
                                $exitosos++
                            } else {
                                Write-Host ("  [SKIP] {0,-25} ya desbloqueada" -f $usr.Usuario) -ForegroundColor DarkGray
                                $yaLibres++
                            }
                        } catch {
                            Write-Host ("  [ ERR] {0,-25} {1}" -f $usr.Usuario, $_.Exception.Message) -ForegroundColor Red
                            $fallidos++
                        }
                    } else {
                        Write-Host "  Numero $idx fuera de rango, ignorado." -ForegroundColor Yellow
                    }
                }
                Write-Host ""
                Write-Host ("  Resultado: {0} desbloqueadas, {1} ya estaban libres, {2} con error." -f $exitosos, $yaLibres, $fallidos) -ForegroundColor Cyan
            } else {
                Write-Host ""
                Write-Host "  Sin cambios. Continuando..." -ForegroundColor DarkGray
            }
        }

        # ── Habilitar interactivo (solo reporte de deshabilitados) ────────────
        if ($opc -eq "2" -and $total -gt 0) {
            Write-Separador
            Write-Host ""
            Write-Host "  ACCION DE HABILITACION" -ForegroundColor Yellow
            Write-Host "  [A] Habilitar TODOS los usuarios del reporte"
            Write-Host "  [S] Seleccionar usuarios a habilitar"
            Write-Host "  [W] Simular (WhatIf) sin aplicar cambios"
            Write-Host "  [N] Omitir"
            Write-Host ""
            $opcHabilitar = Read-Host "  Opcion"

            if ($opcHabilitar -match "^[wW]$") {
                Write-Host ""
                Write-Host "  [SIMULACION] Los siguientes usuarios serian habilitados:" -ForegroundColor Yellow
                foreach ($usr in $datos) {
                    Enable-ADAccount -Identity $usr.Usuario -WhatIf
                }
                Write-Host ""
                Write-Host "  Ninguna cuenta fue modificada. Repite y confirma con A o S para aplicar." -ForegroundColor DarkGray

            } elseif ($opcHabilitar -match "^[aA]$") {
                Write-Host ""
                $conf = Read-Host "  Confirmar habilitacion de $total usuarios? (S/N)"
                if ($conf -match "^[sS]$") {
                    Write-Host ""
                    $exitosos = 0; $fallidos = 0
                    foreach ($usr in $datos) {
                        try {
                            Enable-ADAccount -Identity $usr.Usuario -ErrorAction Stop
                            Write-Host ("  [ OK ] {0,-25} {1}" -f $usr.Usuario, $usr.Nombre) -ForegroundColor Green
                            $exitosos++
                        } catch {
                            Write-Host ("  [ ERR] {0,-25} {1}" -f $usr.Usuario, $_.Exception.Message) -ForegroundColor Red
                            $fallidos++
                        }
                    }
                    Write-Host ""
                    Write-Host ("  Resultado: {0} habilitados correctamente, {1} con error." -f $exitosos, $fallidos) -ForegroundColor Cyan
                } else {
                    Write-Host ""
                    Write-Host "  Operacion cancelada." -ForegroundColor DarkGray
                }

            } elseif ($opcHabilitar -match "^[sS]$") {
                Write-Host ""
                $arrDatos = @($datos)
                for ($i = 0; $i -lt $arrDatos.Count; $i++) {
                    $usr = $arrDatos[$i]
                    Write-Host ("  [{0,2}] {1,-25} {2}" -f ($i + 1), $usr.Usuario, $usr.Nombre) -ForegroundColor White
                }
                Write-Host ""
                $seleccion = Read-Host "  Numeros a habilitar separados por coma (ej: 1,3,5)"
                $indices = $seleccion -split "," |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { $_ -match "^\d+$" }
                Write-Host ""
                $exitosos = 0; $fallidos = 0
                foreach ($idx in $indices) {
                    $num = [int]$idx - 1
                    if ($num -ge 0 -and $num -lt $arrDatos.Count) {
                        $usr = $arrDatos[$num]
                        try {
                            Enable-ADAccount -Identity $usr.Usuario -ErrorAction Stop
                            Write-Host ("  [ OK ] {0,-25} {1}" -f $usr.Usuario, $usr.Nombre) -ForegroundColor Green
                            $exitosos++
                        } catch {
                            Write-Host ("  [ ERR] {0,-25} {1}" -f $usr.Usuario, $_.Exception.Message) -ForegroundColor Red
                            $fallidos++
                        }
                    } else {
                        Write-Host "  Numero $idx fuera de rango, ignorado." -ForegroundColor Yellow
                    }
                }
                Write-Host ""
                Write-Host ("  Resultado: {0} habilitados correctamente, {1} con error." -f $exitosos, $fallidos) -ForegroundColor Cyan
            } else {
                Write-Host ""
                Write-Host "  Sin cambios. Continuando..." -ForegroundColor DarkGray
            }            Write-Host ""
        }

        # ── Preguntar si exportar ──────────────────────────────────────────
        $nombres = @{
            "1"="Usuarios_Bloqueados"; "2"="Usuarios_Deshabilitados"
            "3"="Usuarios_PasswordExpirado"; "4"="Usuarios_PasswordNuncaExpira"
            "5"="Usuarios_Inactivos90dias"; "6"="Usuarios_PrivilegiosElevados"
            "7"="Equipos_SinOU"; "8"="Grupos_Vacios"
        }
        Write-Separador
        Export-DatosCSV -Data $datos -NombreArchivoBase $nombres[$opc] `
            -Prompt "  Exportar estos $total resultados a CSV? (S/N)" | Out-Null

    } catch {
        Write-Host ""
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO COMPARACION DE USUARIOS
# ============================================================

function Compare-Usuarios {
    Write-Header "COMPARACION ENTRE DOS USUARIOS"

    $u1sam = Read-Host "  Usuario A (SamAccountName)"
    $u2sam = Read-Host "  Usuario B (SamAccountName)"

    $props = @(
        "Name","SamAccountName","EmailAddress","Department","Title",
        "Enabled","LockedOut","PasswordNeverExpires","PasswordExpired",
        "LastLogonDate","DistinguishedName","MemberOf","Manager","Description"
    )
    try {
        $u1 = Get-ADUser -Identity $u1sam -Properties $props -ErrorAction Stop
        $u2 = Get-ADUser -Identity $u2sam -Properties $props -ErrorAction Stop
    } catch {
        Write-Host ""
        Write-Host "  ERROR: No se pudo obtener uno de los usuarios." -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkRed
        Pause-Pantalla
        return
    }

    $colA = "Cyan"
    $colB = "Magenta"

    function Get-MgrNombre {
        param($mgr)
        if (-not $mgr) { return "No asignado" }
        $n = (Get-ADUser -Identity $mgr -ErrorAction SilentlyContinue).Name
        if ($n) { return $n } else { return $mgr }
    }

    # ----------------------------------------------------------
    #  Tabla de atributos
    # ----------------------------------------------------------
    Write-SubHeader "Comparacion de Atributos"
    Write-EncabezadoComparacion -NombreA $u1.SamAccountName -NombreB $u2.SamAccountName

    Write-FilaComparacion "Nombre completo"   $u1.Name                                               $u2.Name
    Write-FilaComparacion "Email"             $(if($u1.EmailAddress){$u1.EmailAddress}else{"-"})      $(if($u2.EmailAddress){$u2.EmailAddress}else{"-"})
    Write-FilaComparacion "Departamento"      $(if($u1.Department){$u1.Department}else{"-"})          $(if($u2.Department){$u2.Department}else{"-"})
    Write-FilaComparacion "Cargo"             $(if($u1.Title){$u1.Title}else{"-"})                    $(if($u2.Title){$u2.Title}else{"-"})
    Write-FilaComparacion "Manager"           (Get-MgrNombre $u1.Manager)                             (Get-MgrNombre $u2.Manager)
    Write-FilaComparacion "Cuenta activa"     $(if($u1.Enabled){"SI"}else{"NO"})                      $(if($u2.Enabled){"SI"}else{"NO"})
    Write-FilaComparacion "Bloqueado"         $(if($u1.LockedOut){"SI"}else{"No"})                    $(if($u2.LockedOut){"SI"}else{"No"})
    Write-FilaComparacion "Pass expirada"     $(if($u1.PasswordExpired){"SI"}else{"No"})              $(if($u2.PasswordExpired){"SI"}else{"No"})
    Write-FilaComparacion "Pass nunca expira" $(if($u1.PasswordNeverExpires){"SI"}else{"No"})         $(if($u2.PasswordNeverExpires){"SI"}else{"No"})
    Write-FilaComparacion "Ultimo logon"      $(if($u1.LastLogonDate){$u1.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"}) $(if($u2.LastLogonDate){$u2.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"})
    Write-FilaComparacion "OU"                (Get-OUdesdeDN $u1.DistinguishedName)                   (Get-OUdesdeDN $u2.DistinguishedName)

    Write-SepComparacion
    Write-Host ""
    Write-Host "  >> = Diferencia encontrada" -ForegroundColor Yellow
    Write-Host ("  " + $u1.SamAccountName + " = columna izquierda  |  " + $u2.SamAccountName + " = columna derecha") -ForegroundColor DarkGray

    # ----------------------------------------------------------
    #  Comparacion de grupos
    # ----------------------------------------------------------
    Write-SubHeader "Comparacion de Grupos"
    $diffGrupos = Compare-YMostrarGrupos -MemberOfA $u1.MemberOf -MemberOfB $u2.MemberOf `
        -NombreA $u1.SamAccountName -NombreB $u2.SamAccountName -Sustantivo "usuarios"
    $comunes = $diffGrupos.Comunes; $soloA = $diffGrupos.SoloA; $soloB = $diffGrupos.SoloB

    # ----------------------------------------------------------
    #  Exportar
    # ----------------------------------------------------------
    $exp = Read-Host "  Exportar comparacion a CSV? (S/N)"
    if ($exp -match "^[sS]$") {
        $ruta = Read-Host "  Ruta (ej: C:\Reportes\comp_usuarios.csv)"
        try {
            $filas = @()
            foreach ($g in $comunes) {
                $filas += [PSCustomObject]@{Grupo=$g.InputObject; Estado="Comun"; UsuarioA=$u1.SamAccountName; UsuarioB=$u2.SamAccountName}
            }
            foreach ($g in $soloA) {
                $filas += [PSCustomObject]@{Grupo=$g.InputObject; Estado="Solo en $($u1.SamAccountName)"; UsuarioA=$u1.SamAccountName; UsuarioB=$u2.SamAccountName}
            }
            foreach ($g in $soloB) {
                $filas += [PSCustomObject]@{Grupo=$g.InputObject; Estado="Solo en $($u2.SamAccountName)"; UsuarioA=$u1.SamAccountName; UsuarioB=$u2.SamAccountName}
            }
            $filas | Export-Csv -Path $ruta -NoTypeInformation -Encoding UTF8
            Write-Host "  OK - Exportado en: $ruta" -ForegroundColor Green
        } catch {
            Write-Host "  ERROR al exportar: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Pause-Pantalla
}

# ============================================================
#  MODULO COMPARACION DE EQUIPOS
# ============================================================

function Compare-Equipos {
    Write-Header "COMPARACION ENTRE DOS EQUIPOS"

    $e1nom = Read-Host "  Equipo A (hostname)"
    $e2nom = Read-Host "  Equipo B (hostname)"

    $props = @(
        "Name","DNSHostName","Enabled","OperatingSystem","OperatingSystemVersion",
        "LastLogonDate","Created","DistinguishedName","MemberOf",
        "IPv4Address","Description","ManagedBy"
    )
    try {
        $e1 = Get-ADComputer -Identity $e1nom -Properties $props -ErrorAction Stop
        $e2 = Get-ADComputer -Identity $e2nom -Properties $props -ErrorAction Stop
    } catch {
        Write-Host ""
        Write-Host "  ERROR: No se pudo obtener uno de los equipos." -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkRed
        Pause-Pantalla
        return
    }

    $colA  = "Cyan"
    $colB  = "Magenta"

    # ----------------------------------------------------------
    #  Tabla de atributos
    # ----------------------------------------------------------
    Write-SubHeader "Comparacion de Atributos"
    Write-EncabezadoComparacion -NombreA $e1.Name -NombreB $e2.Name

    $ou1 = Get-OUdesdeDN $e1.DistinguishedName
    $ou2 = Get-OUdesdeDN $e2.DistinguishedName

    Write-FilaComparacion "Hostname"          $e1.Name                                                            $e2.Name
    Write-FilaComparacion "DNS"               $(if($e1.DNSHostName){$e1.DNSHostName}else{"-"})                   $(if($e2.DNSHostName){$e2.DNSHostName}else{"-"})
    Write-FilaComparacion "IPv4"              $(if($e1.IPv4Address){$e1.IPv4Address}else{"-"})                    $(if($e2.IPv4Address){$e2.IPv4Address}else{"-"})
    Write-FilaComparacion "SO"                $(if($e1.OperatingSystem){$e1.OperatingSystem}else{"-"})            $(if($e2.OperatingSystem){$e2.OperatingSystem}else{"-"})
    Write-FilaComparacion "Version SO"        $(if($e1.OperatingSystemVersion){$e1.OperatingSystemVersion}else{"-"}) $(if($e2.OperatingSystemVersion){$e2.OperatingSystemVersion}else{"-"})
    Write-FilaComparacion "Habilitado"        $(if($e1.Enabled){"SI"}else{"NO"})                                  $(if($e2.Enabled){"SI"}else{"NO"})
    Write-FilaComparacion "Ultimo logon"      $(if($e1.LastLogonDate){$e1.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"}) $(if($e2.LastLogonDate){$e2.LastLogonDate.ToString("dd/MM/yyyy")}else{"Nunca"})
    Write-FilaComparacion "Creado"            $e1.Created.ToString("dd/MM/yyyy")                                  $e2.Created.ToString("dd/MM/yyyy")
    Write-FilaComparacion "OU"                $ou1                                                                 $ou2

    Write-SepComparacion
    Write-Host ""
    Write-Host "  >> = Diferencia encontrada" -ForegroundColor Yellow
    Write-Host ("  " + $e1.Name + " = columna izquierda  |  " + $e2.Name + " = columna derecha") -ForegroundColor DarkGray

    # ----------------------------------------------------------
    #  Ubicacion OU desglosada nivel a nivel
    # ----------------------------------------------------------
    Write-SubHeader "Ubicacion en el Directorio (OU)"

    $mismoObjeto = ($e1.DistinguishedName -eq $e2.DistinguishedName)
    $mismaOU     = ($ou1 -eq $ou2)

    if ($mismoObjeto) {
        Write-Host "  [!] ATENCION: Ambos nombres apuntan al MISMO objeto en AD." -ForegroundColor Red
        Write-Host "      DN: $($e1.DistinguishedName)" -ForegroundColor DarkGray
    } elseif ($mismaOU) {
        Write-Host "  [OK] Ambos equipos estan en la MISMA OU:" -ForegroundColor Green
        Write-Host "       $ou1" -ForegroundColor Green
    } else {
        Write-Host "  [!!] Los equipos estan en OUs DIFERENTES" -ForegroundColor Yellow
        Write-Host ""
        Write-Host ("  " + $e1.Name.PadRight(20) + " : ") -NoNewline -ForegroundColor $colA
        Write-Host $ou1 -ForegroundColor $colA
        Write-Host ("  " + $e2.Name.PadRight(20) + " : ") -NoNewline -ForegroundColor $colB
        Write-Host $ou2 -ForegroundColor $colB

        # Desglose nivel a nivel
        Write-Host ""
        Write-Host "  Desglose por nivel de OU (desde la raiz):" -ForegroundColor White
        Write-Host ""
        $partes1 = @($e1.DistinguishedName -split "," | Where-Object {$_ -notmatch "^CN="})
        $partes2 = @($e2.DistinguishedName -split "," | Where-Object {$_ -notmatch "^CN="})
        $maxLen  = [Math]::Max($partes1.Count, $partes2.Count)

        $sepNivel = "  +---------+" + ("-" * 24) + "+" + ("-" * 24) + "+"
        Write-Host $sepNivel -ForegroundColor DarkGray
        Write-Host ("  | Nivel   | " + $e1.Name.PadRight(22) + " | " + $e2.Name.PadRight(22) + " |") -ForegroundColor DarkGray
        Write-Host $sepNivel -ForegroundColor DarkGray

        for ($i = $maxLen - 1; $i -ge 0; $i--) {
            $p1    = if ($i -lt $partes1.Count) {($partes1[$i] -split "=")[1]} else {"(no existe)"}
            $p2    = if ($i -lt $partes2.Count) {($partes2[$i] -split "=")[1]} else {"(no existe)"}
            $igual = ($p1 -eq $p2)
            $col   = if ($igual) {"Gray"} else {"Yellow"}
            $icono = if ($igual) {" OK "} else {"DIFF"}
            $nivel = "Nivel $($maxLen - $i)"
            Write-Host ("  | " + $icono.PadRight(7) + " | ") -NoNewline -ForegroundColor $col
            Write-Host $p1.PadRight(22) -NoNewline -ForegroundColor $(if($igual){"Gray"}else{$colA})
            Write-Host " | " -NoNewline -ForegroundColor DarkGray
            Write-Host $p2.PadRight(22) -NoNewline -ForegroundColor $(if($igual){"Gray"}else{$colB})
            Write-Host " |" -ForegroundColor DarkGray
        }
        Write-Host $sepNivel -ForegroundColor DarkGray
    }

    # ----------------------------------------------------------
    #  Comparacion de grupos
    # ----------------------------------------------------------
    Write-SubHeader "Comparacion de Grupos"
    $diffGrupos = Compare-YMostrarGrupos -MemberOfA $e1.MemberOf -MemberOfB $e2.MemberOf `
        -NombreA $e1.Name -NombreB $e2.Name -Sustantivo "equipos"
    $comunes = $diffGrupos.Comunes; $soloA = $diffGrupos.SoloA; $soloB = $diffGrupos.SoloB

    # ----------------------------------------------------------
    #  Exportar
    # ----------------------------------------------------------
    $exp = Read-Host "  Exportar comparacion a CSV? (S/N)"
    if ($exp -match "^[sS]$") {
        $ruta = Read-Host "  Ruta (ej: C:\Reportes\comp_equipos.csv)"
        try {
            $filas = @()
            $atribs = @(
                @{C="Hostname";    A=$e1.Name;                   B=$e2.Name},
                @{C="DNS";         A=$e1.DNSHostName;            B=$e2.DNSHostName},
                @{C="IPv4";        A=$e1.IPv4Address;            B=$e2.IPv4Address},
                @{C="SO";          A=$e1.OperatingSystem;        B=$e2.OperatingSystem},
                @{C="VersionSO";   A=$e1.OperatingSystemVersion; B=$e2.OperatingSystemVersion},
                @{C="Habilitado";  A=$e1.Enabled;                B=$e2.Enabled},
                @{C="UltimoLogon"; A=$e1.LastLogonDate;          B=$e2.LastLogonDate},
                @{C="OU";          A=$ou1;                       B=$ou2}
            )
            foreach ($a in $atribs) {
                $filas += [PSCustomObject]@{
                    Tipo=    "Atributo"
                    Elemento=$a.C
                    Estado=  if ($a.A -eq $a.B) {"Igual"} else {"Diferente"}
                    EquipoA= $e1.Name; ValorA=$a.A
                    EquipoB= $e2.Name; ValorB=$a.B
                }
            }
            foreach ($g in $comunes) {
                $filas += [PSCustomObject]@{Tipo="Grupo";Elemento=$g.InputObject;Estado="Comun";EquipoA=$e1.Name;ValorA="SI";EquipoB=$e2.Name;ValorB="SI"}
            }
            foreach ($g in $soloA) {
                $filas += [PSCustomObject]@{Tipo="Grupo";Elemento=$g.InputObject;Estado="Solo en $($e1.Name)";EquipoA=$e1.Name;ValorA="SI";EquipoB=$e2.Name;ValorB="NO"}
            }
            foreach ($g in $soloB) {
                $filas += [PSCustomObject]@{Tipo="Grupo";Elemento=$g.InputObject;Estado="Solo en $($e2.Name)";EquipoA=$e1.Name;ValorA="NO";EquipoB=$e2.Name;ValorB="SI"}
            }
            $filas | Export-Csv -Path $ruta -NoTypeInformation -Encoding UTF8
            Write-Host "  OK - Exportado en: $ruta" -ForegroundColor Green
        } catch {
            Write-Host "  ERROR al exportar: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Pause-Pantalla
}


# ============================================================
#  MODULO GPO - Politicas de Grupo
# ============================================================

# Verifica si el modulo GroupPolicy esta disponible
function Test-GPOModule {
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        Write-Host ""
        Write-Host "  ERROR: El modulo GroupPolicy no esta disponible." -ForegroundColor Red
        Write-Host "  Requiere GPMC (Group Policy Management Console) instalado." -ForegroundColor Yellow
        Write-Host "  Instalalo con: Add-WindowsFeature GPMC  o via RSAT en Windows 10/11" -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    Import-Module GroupPolicy -ErrorAction SilentlyContinue
    return $true
}

# Parsea el XML de una GPO y devuelve un resumen de sus configuraciones
function Get-ResumenGPO {
    param([xml]$xmlGPO)

    $resumen = @{
        Seguridad      = @()
        Scripts        = @()
        Registro       = @()
        Unidades       = @()
        Impresoras     = @()
        Software       = @()
        Carpetas       = @()
        Otros          = @()
    }

    # Namespaces helpers
    $ns = New-Object System.Xml.XmlNamespaceManager($xmlGPO.NameTable)
    $ns.AddNamespace("gp",  "http://www.microsoft.com/GroupPolicy/Settings")

    function Get-Nodos { param($nodo, $xpath) try { $nodo.SelectNodes($xpath, $ns) } catch { $null } }

    foreach ($scope in @("Computer","User")) {
        $scopeNode = $xmlGPO.GPO.$scope
        if (-not $scopeNode) { continue }

        $enabled = $scopeNode.Enabled
        if ($enabled -eq "false") { continue }

        foreach ($ext in $scopeNode.ExtensionData.Extension) {
            if (-not $ext) { continue }
            $tipo = $ext.LocalName

            switch -Wildcard ($tipo) {
                # ── Seguridad ─────────────────────────────────────────────
                "SecuritySettings" {
                    # Politica de contrasenas
                    $pp = $ext.Account | Where-Object { $_.Type -eq "PasswordPolicies" }
                    if ($pp) {
                        foreach ($s in $pp.SettingBoolean + $pp.SettingNumber) {
                            if ($s.Name) {
                                $resumen.Seguridad += "[Contrasena] $($s.Name) = $($s.SettingNumber)$($s.SettingBoolean)"
                            }
                        }
                    }
                    # Bloqueo de cuenta
                    $lp = $ext.Account | Where-Object { $_.Type -eq "LockoutPolicies" }
                    if ($lp) {
                        foreach ($s in $lp.SettingNumber) {
                            if ($s.Name) {
                                $resumen.Seguridad += "[Bloqueo] $($s.Name) = $($s.SettingNumber)"
                            }
                        }
                    }
                    # Derechos de usuario (User Rights)
                    foreach ($ur in $ext.UserRightsAssignment) {
                        if ($ur.Name) {
                            $who = ($ur.Member | ForEach-Object { $_.Name.'#text' }) -join ", "
                            $resumen.Seguridad += "[Derecho] $($ur.Name) => $who"
                        }
                    }
                    # Grupos restringidos
                    foreach ($rg in $ext.RestrictedGroups) {
                        if ($rg.GroupName) {
                            $members = ($rg.Member | ForEach-Object { $_.Name.'#text' }) -join ", "
                            $resumen.Seguridad += "[Grupo restringido] $($rg.GroupName.'#text') => Miembros: $members"
                        }
                    }
                    # Auditoria
                    foreach ($ap in $ext.AuditSetting) {
                        if ($ap.SubcategoryName) {
                            $resumen.Seguridad += "[Auditoria] $($ap.SubcategoryName) = $($ap.SettingValue)"
                        }
                    }
                    # Opciones de seguridad
                    foreach ($so in $ext.SecurityOptions) {
                        if ($so.KeyName -and $so.SettingString) {
                            $resumen.Seguridad += "[Opcion] $($so.KeyName) = $($so.SettingString)"
                        }
                    }
                }

                # ── Scripts ───────────────────────────────────────────────
                "Scripts*" {
                    foreach ($sc in $ext.Script) {
                        $tipo2 = if ($sc.Type) { $sc.Type } else { "Script" }
                        $cmd   = if ($sc.Command) { $sc.Command } else { "(sin comando)" }
                        $resumen.Scripts += "[$tipo2] $cmd"
                    }
                }

                # ── Registro ──────────────────────────────────────────────
                "Registry*" {
                    foreach ($rk in $ext.Registry) {
                        $key = $rk.Properties.key
                        $val = $rk.Properties.value
                        $dat = $rk.Properties.data
                        if ($key) { $resumen.Registro += "$key\$val = $dat" }
                    }
                }

                # ── Unidades de red ───────────────────────────────────────
                "DriveMapSettings" {
                    foreach ($drv in $ext.DriveMap) {
                        $letra  = $drv.Properties.letter
                        $ruta   = $drv.Properties.path
                        $label  = $drv.Properties.label
                        $accion = $drv.Properties.action
                        $resumen.Unidades += "[$accion] Letra:$letra  Ruta:$ruta  Etiqueta:$label"
                    }
                }

                # ── Impresoras ────────────────────────────────────────────
                "PrinterSettings" {
                    foreach ($p in $ext.Printer) {
                        $ruta   = $p.Properties.path
                        $accion = $p.Properties.action
                        $deflt  = $p.Properties.default
                        $resumen.Impresoras += "[$accion] $ruta  Predeterminada:$deflt"
                    }
                }

                # ── Software ──────────────────────────────────────────────
                "AppMgmtSetting" {
                    foreach ($app in $ext.MsiApplication) {
                        $resumen.Software += "$($app.Name) v$($app.ProductVersion)"
                    }
                }
                "SoftwareInstallationSettings" {
                    foreach ($app in $ext.Application) {
                        $resumen.Software += "$($app.Name) - $($app.Path)"
                    }
                }

                # ── Redireccion de carpetas ───────────────────────────────
                "FolderRedirectionSettings" {
                    foreach ($fr in $ext.Folder) {
                        $resumen.Carpetas += "$($fr.Id) => $($fr.Location.DestinationPath)"
                    }
                }

                # ── Resto ─────────────────────────────────────────────────
                default {
                    if ($tipo -and $tipo -notin @("","#comment")) {
                        $resumen.Otros += "Configuracion detectada: $tipo ($scope)"
                    }
                }
            }
        }
    }
    return $resumen
}

# Imprime una seccion del resumen GPO
function Write-SeccionGPO {
    param([string]$Titulo, [array]$Items, [ConsoleColor]$Color = "Gray", [string]$Icono = "(*)  ")
    if (-not $Items -or $Items.Count -eq 0) { return }
    Write-Host ""
    Write-Host ("  +-- " + $Titulo + " (" + $Items.Count + ") " + ("-" * [Math]::Max(1,48-$Titulo.Length)) + "+") -ForegroundColor Yellow
    foreach ($item in $Items) {
        Write-Host "  $Icono" -NoNewline -ForegroundColor DarkGray
        Write-Host $item -ForegroundColor $Color
    }
}

# Helper: obtiene el reporte XML de una GPO, lo parsea y muestra sus
# configuraciones por categoria. Usado por Get-GPOsPorGrupo y Get-GPOsDeObjeto
# para no repetir este bloque (Get-GPOReport + Get-ResumenGPO + 8x Write-SeccionGPO).
function Write-ConfiguracionesGPO {
    param($Gpo)
    try {
        $xmlStr = Get-GPOReport -Guid $Gpo.Id -ReportType Xml -ErrorAction Stop
        $xml    = [xml]$xmlStr
        $res    = Get-ResumenGPO $xml

        $totalCfg = ($res.Seguridad.Count + $res.Scripts.Count + $res.Registro.Count +
                     $res.Unidades.Count + $res.Impresoras.Count + $res.Software.Count +
                     $res.Carpetas.Count + $res.Otros.Count)

        if ($totalCfg -gt 0) {
            Write-SeccionGPO "Seguridad y Derechos"     $res.Seguridad   "White"    "(SEC)"
            Write-SeccionGPO "Scripts"                  $res.Scripts     "Cyan"     "(SCR)"
            Write-SeccionGPO "Registro"                 $res.Registro    "Gray"     "(REG)"
            Write-SeccionGPO "Unidades de red"          $res.Unidades    "Yellow"   "(DRV)"
            Write-SeccionGPO "Impresoras"               $res.Impresoras  "Magenta"  "(PRT)"
            Write-SeccionGPO "Software"                 $res.Software    "Green"    "(APP)"
            Write-SeccionGPO "Redireccion de carpetas"  $res.Carpetas    "DarkCyan" "(FLD)"
            Write-SeccionGPO "Otras configuraciones"    $res.Otros       "DarkGray" "( ? )"
        } else {
            Write-Host "  (No se detectaron configuraciones parseables - usa GPMC para detalles completos)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  No se pudo parsear la GPO: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
}

# ── FUNCION 1: Info detallada de una GPO ─────────────────────────────────────
function Get-InfoGPO {
    Write-Header "INFORMACION DE GPO"
    if (-not (Test-GPOModule)) { Pause-Pantalla; return }

    $nombre = Read-Host "  Nombre de la GPO (o GUID entre llaves)"

    try {
        $gpo = Get-GPO -Name $nombre -ErrorAction Stop
    } catch {
        try { $gpo = Get-GPO -Guid $nombre -ErrorAction Stop } catch {
            Write-Host "  ERROR: GPO no encontrada." -ForegroundColor Red
            Pause-Pantalla; return
        }
    }

    Write-SubHeader "Datos de la GPO"
    Write-Campo "Nombre"              $gpo.DisplayName                "White"
    Write-Campo "GUID"                $gpo.Id.ToString()
    Write-Campo "Dominio"             $gpo.DomainName
    Write-Campo "Estado"              $gpo.GpoStatus.ToString()
    Write-Campo "Creada"              $gpo.CreationTime.ToString("dd/MM/yyyy HH:mm")
    Write-Campo "Modificada"          $gpo.ModificationTime.ToString("dd/MM/yyyy HH:mm")
    Write-Campo "Conf. Equipo"        $(if($gpo.Computer.Enabled){"Habilitada"}else{"Deshabilitada"})
    Write-Campo "Conf. Usuario"       $(if($gpo.User.Enabled){"Habilitada"}else{"Deshabilitada"})

    # OUs donde esta vinculada
    Write-SubHeader "OUs donde esta vinculada"
    try {
        $links = Get-ADOrganizationalUnit -Filter * -Properties gpLink |
            Where-Object { $_.gpLink -like "*$($gpo.Id)*" }
        if ($links) {
            foreach ($ou in $links) {
                $ouPath = Get-OUdesdeDN $ou.DistinguishedName
                Write-Host "  --> $ouPath" -ForegroundColor Cyan
                Write-Host "      DN: $($ou.DistinguishedName)" -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  No se encontraron OUs vinculadas (puede ser a nivel de dominio/sitio)." -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  No se pudo obtener los vinculos: $($_.Exception.Message)" -ForegroundColor DarkGray
    }

    # Filtrado de seguridad (quien puede aplicarla)
    Write-SubHeader "Filtrado de Seguridad (quien puede aplicarla)"
    try {
        $perms = Get-GPPermission -Guid $gpo.Id -All |
            Where-Object { $_.Permission -eq "GpoApply" }
        if ($perms) {
            foreach ($p in $perms) {
                $tipo = $p.Trustee.SidType.ToString()
                Write-Host ("  [" + $tipo.PadRight(8) + "] ") -NoNewline -ForegroundColor DarkCyan
                Write-Host $p.Trustee.Name -ForegroundColor Green
            }
        } else {
            Write-Host "  Sin entradas de filtrado (aplica a todos los usuarios autenticados)." -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "  No se pudo leer filtrado de seguridad." -ForegroundColor DarkGray
    }


    Pause-Pantalla
}

# ── FUNCION 2: GPOs que aplican a un Grupo ───────────────────────────────────
function Get-GPOsPorGrupo {
    Write-Header "GPOs QUE APLICAN A UN GRUPO"
    if (-not (Test-GPOModule)) { Pause-Pantalla; return }

    $grpNombre = Read-Host "  Nombre del grupo"

    try {
        $grupo = Get-ADGroup -Identity $grpNombre -ErrorAction Stop
    } catch {
        Write-Host "  ERROR: Grupo no encontrado." -ForegroundColor Red
        Pause-Pantalla; return
    }

    Write-Host ""
    Write-Host "  Buscando GPOs con filtrado de seguridad para: " -NoNewline -ForegroundColor DarkCyan
    Write-Host $grupo.Name -ForegroundColor Cyan
    Write-Host "  (Esto puede tardar unos segundos...)" -ForegroundColor DarkGray
    Write-Host ""

    try {
        $todasGPOs   = Get-GPO -All -ErrorAction Stop
        $gposDelGrupo = @()

        foreach ($gpo in $todasGPOs) {
            try {
                $perms = Get-GPPermission -Guid $gpo.Id -All -ErrorAction SilentlyContinue |
                    Where-Object { $_.Permission -eq "GpoApply" -and
                                   $_.Trustee.Name -like "*$($grupo.SamAccountName)*" }
                if ($perms) {
                    $gposDelGrupo += $gpo
                }
            } catch { continue }
        }

        if ($gposDelGrupo.Count -eq 0) {
            Write-Host "  No se encontraron GPOs con este grupo en el filtrado de seguridad." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  NOTA: El grupo puede recibir GPOs por pertenecer a una OU o sitio" -ForegroundColor DarkGray
            Write-Host "  sin estar explicitamente en el filtrado de seguridad." -ForegroundColor DarkGray
            Pause-Pantalla; return
        }

        Write-Host "  Se encontraron $($gposDelGrupo.Count) GPO(s) con este grupo:" -ForegroundColor Green
        Write-Host ""

        foreach ($gpo in $gposDelGrupo) {
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor Cyan
            Write-Host ("  |  GPO: " + $gpo.DisplayName.PadRight(51) + "|") -ForegroundColor White
            Write-Host "  +----------------------------------------------------------+" -ForegroundColor Cyan

            Write-Host "  Estado    : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $gpo.GpoStatus.ToString() -ForegroundColor Gray
            Write-Host "  Modificada: " -NoNewline -ForegroundColor DarkCyan
            Write-Host $gpo.ModificationTime.ToString("dd/MM/yyyy HH:mm") -ForegroundColor Gray
            Write-Host "  GUID      : " -NoNewline -ForegroundColor DarkCyan
            Write-Host $gpo.Id.ToString() -ForegroundColor DarkGray

            # OUs vinculadas
            $links = Get-ADOrganizationalUnit -Filter * -Properties gpLink |
                Where-Object { $_.gpLink -like "*$($gpo.Id)*" }
            if ($links) {
                Write-Host "  Vinculada a:" -ForegroundColor DarkCyan
                foreach ($ou in $links) {
                    Write-Host ("      --> " + (Get-OUdesdeDN $ou.DistinguishedName)) -ForegroundColor Yellow
                }
            }

            # Configuraciones de la GPO
            Write-ConfiguracionesGPO -Gpo $gpo
            Write-Host ""
        }

    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }

    Pause-Pantalla
}

# ── FUNCION 3: GPOs que aplican a un usuario o equipo (RSoP) ─────────────────
function Get-GPOsDeObjeto {
    Write-Header "GPOs QUE APLICAN A UN USUARIO O EQUIPO"
    if (-not (Test-GPOModule)) { Pause-Pantalla; return }

    Write-Host "  Buscar GPOs para:" -ForegroundColor Yellow
    Write-Host "  [1] Usuario"
    Write-Host "  [2] Equipo"
    Write-Host ""
    $tipo = Read-Host "  Tipo"

    if ($tipo -eq "1") {
        $identidad = Read-Host "  SamAccountName del usuario"
        try {
            $obj   = Get-ADUser -Identity $identidad -Properties DistinguishedName -ErrorAction Stop
            $ouDN  = ($obj.DistinguishedName -split ",",2)[1]
            $tipoTxt = "usuario"
        } catch {
            Write-Host "  ERROR: Usuario no encontrado." -ForegroundColor Red
            Pause-Pantalla; return
        }
    } elseif ($tipo -eq "2") {
        $identidad = Read-Host "  Hostname del equipo"
        try {
            $obj   = Get-ADComputer -Identity $identidad -Properties DistinguishedName -ErrorAction Stop
            $ouDN  = ($obj.DistinguishedName -split ",",2)[1]
            $tipoTxt = "equipo"
        } catch {
            Write-Host "  ERROR: Equipo no encontrado." -ForegroundColor Red
            Pause-Pantalla; return
        }
    } else {
        Write-Host "  Opcion no valida." -ForegroundColor Red
        Pause-Pantalla; return
    }

    Write-Host ""
    Write-Host "  Objeto  : " -NoNewline -ForegroundColor DarkCyan
    Write-Host $obj.Name -ForegroundColor White
    Write-Host "  OU      : " -NoNewline -ForegroundColor DarkCyan
    Write-Host (Get-OUdesdeDN $obj.DistinguishedName) -ForegroundColor Cyan
    Write-Host ""

    # Recopilar todas las OUs en la jerarquia desde el objeto hasta la raiz
    $ouChain = @()
    $current = $ouDN
    while ($current -match "^(OU|DC)=") {
        $ouChain += $current
        $current  = ($current -split ",",2)[1]
    }

    # Buscar GPOs vinculadas a cada nivel
    Write-Host "  GPOs encontradas en la jerarquia de OUs:" -ForegroundColor Yellow
    Write-Host ""

    $gposEncontradas = @()

    foreach ($nivel in $ouChain) {
        try {
            $adObj = Get-ADObject -Identity $nivel -Properties gpLink -ErrorAction SilentlyContinue
            if ($adObj -and $adObj.gpLink) {
                # Extraer GUIDs de gpLink
                $guids = [regex]::Matches($adObj.gpLink, '\{([A-Fa-f0-9\-]+)\}') |
                    ForEach-Object { $_.Groups[1].Value }

                foreach ($guid in $guids) {
                    try {
                        $gpo = Get-GPO -Guid $guid -ErrorAction SilentlyContinue
                        if ($gpo) {
                            $gposEncontradas += [PSCustomObject]@{
                                GPO   = $gpo
                                Nivel = Get-OUdesdeDN $nivel
                            }
                        }
                    } catch { continue }
                }
            }
        } catch { continue }
    }

    if ($gposEncontradas.Count -eq 0) {
        Write-Host "  No se encontraron GPOs vinculadas en la jerarquia de OUs." -ForegroundColor Yellow
        Pause-Pantalla; return
    }

    Write-Host "  Total de GPOs que aplican: $($gposEncontradas.Count)" -ForegroundColor Green
    Write-Host ""

    foreach ($entrada in $gposEncontradas) {
        $gpo = $entrada.GPO
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host ("  |  GPO: " + $gpo.DisplayName.PadRight(51) + "|") -ForegroundColor White
        Write-Host "  +----------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  Vinculada en : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $entrada.Nivel -ForegroundColor Yellow
        Write-Host "  Estado       : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $gpo.GpoStatus -ForegroundColor Gray
        Write-Host "  Modificada   : " -NoNewline -ForegroundColor DarkCyan
        Write-Host $gpo.ModificationTime.ToString("dd/MM/yyyy HH:mm") -ForegroundColor Gray

        # Filtrado de seguridad
        try {
            $applyPerms = Get-GPPermission -Guid $gpo.Id -All -ErrorAction SilentlyContinue |
                Where-Object { $_.Permission -eq "GpoApply" }
            if ($applyPerms) {
                $quienAplica = ($applyPerms | ForEach-Object { $_.Trustee.Name }) -join ", "
                Write-Host "  Aplica a     : " -NoNewline -ForegroundColor DarkCyan
                Write-Host $quienAplica -ForegroundColor Green
            }
        } catch {}

        # Configuraciones
        Write-ConfiguracionesGPO -Gpo $gpo
        Write-Host ""
    }

    Pause-Pantalla
}

# ── FUNCION 4: Comparar GPOs entre dos usuarios o dos equipos ────────────────
function Compare-GPO {
    Write-Header "COMPARADOR DE GPO"
    if (-not (Test-GPOModule)) { Pause-Pantalla; return }

    Write-Host "  Tipo de comparacion:" -ForegroundColor Yellow
    Write-Host "  [1]  Comparar GPOs entre dos USUARIOS"
    Write-Host "  [2]  Comparar GPOs entre dos EQUIPOS"
    Write-Host ""
    $tipoComp = Read-Host "  Opcion"

    if ($tipoComp -notin @("1","2")) {
        Write-Host "  Opcion no valida." -ForegroundColor Red
        Pause-Pantalla; return
    }

    $esUsuario = ($tipoComp -eq "1")
    $labelA    = if ($esUsuario) { "Usuario A (SamAccountName)" } else { "Equipo A (hostname)" }
    $labelB    = if ($esUsuario) { "Usuario B (SamAccountName)" } else { "Equipo B (hostname)" }
    $tituloH   = if ($esUsuario) { "COMPARADOR DE GPO: USUARIO vs USUARIO" } else { "COMPARADOR DE GPO: EQUIPO vs EQUIPO" }

    Write-Host ""
    Write-Header $tituloH
    $identA = (Read-Host "  $labelA").Trim()
    $identB = (Read-Host "  $labelB").Trim()

    try {
        if ($esUsuario) {
            $objA = Get-ADUser     -Identity $identA -Properties DistinguishedName,Name -ErrorAction Stop
            $objB = Get-ADUser     -Identity $identB -Properties DistinguishedName,Name -ErrorAction Stop
        } else {
            $objA = Get-ADComputer -Identity $identA -Properties DistinguishedName,Name -ErrorAction Stop
            $objB = Get-ADComputer -Identity $identB -Properties DistinguishedName,Name -ErrorAction Stop
        }
    } catch {
        Write-Host "  ERROR: No se pudo obtener uno de los objetos." -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkRed
        Pause-Pantalla; return
    }

    $colA = "Cyan"
    $colB = "Magenta"

    Write-Host ""
    Write-Host ("  [A] " + $objA.Name.PadRight(30)) -NoNewline -ForegroundColor $colA
    Write-Host "OU: $(Get-OUdesdeDN $objA.DistinguishedName)" -ForegroundColor DarkGray
    Write-Host ("  [B] " + $objB.Name.PadRight(30)) -NoNewline -ForegroundColor $colB
    Write-Host "OU: $(Get-OUdesdeDN $objB.DistinguishedName)" -ForegroundColor DarkGray
    Write-Host ""

    # ── Helper: recorre jerarquia de OUs y recolecta GPOs vinculadas ─────
    function Get-GPOsJerarquia {
        param([string]$DN)
        $lista     = [System.Collections.Generic.List[PSCustomObject]]::new()
        $vistos    = @{}
        $current   = ($DN -split ",",2)[1]
        while ($current -match "^(OU|DC)=") {
            try {
                $adObj = Get-ADObject -Identity $current -Properties gpLink -ErrorAction SilentlyContinue
                if ($adObj -and $adObj.gpLink) {
                    $guids = [regex]::Matches($adObj.gpLink, '\{([A-Fa-f0-9\-]+)\}') |
                        ForEach-Object { $_.Groups[1].Value }
                    foreach ($guid in $guids) {
                        if (-not $vistos.ContainsKey($guid)) {
                            $vistos[$guid] = $true
                            try {
                                $gpo = Get-GPO -Guid $guid -ErrorAction SilentlyContinue
                                if ($gpo) {
                                    $lista.Add([PSCustomObject]@{
                                        Nombre = $gpo.DisplayName
                                        GUID   = $guid
                                        Nivel  = Get-OUdesdeDN $current
                                        Estado = $gpo.GpoStatus.ToString()
                                    })
                                }
                            } catch { continue }
                        }
                    }
                }
            } catch {}
            $current = ($current -split ",",2)[1]
        }
        return $lista.ToArray()
    }

    Write-Host "  Recopilando GPOs de [A] $($objA.Name)..." -ForegroundColor DarkGray
    $gposA = @(Get-GPOsJerarquia -DN $objA.DistinguishedName)
    Write-Host "  Recopilando GPOs de [B] $($objB.Name)..." -ForegroundColor DarkGray
    $gposB = @(Get-GPOsJerarquia -DN $objB.DistinguishedName)

    $nombresA = @($gposA | ForEach-Object { $_.Nombre } | Sort-Object)
    $nombresB = @($gposB | ForEach-Object { $_.Nombre } | Sort-Object)

    if ($nombresA.Count -eq 0) { $nombresA = @("__EMPTY__") }
    if ($nombresB.Count -eq 0) { $nombresB = @("__EMPTY__") }

    $diff    = Compare-Object -ReferenceObject $nombresA -DifferenceObject $nombresB -IncludeEqual
    $comunes = @($diff | Where-Object { $_.SideIndicator -eq "==" -and $_.InputObject -ne "__EMPTY__" })
    $soloA   = @($diff | Where-Object { $_.SideIndicator -eq "<=" -and $_.InputObject -ne "__EMPTY__" })
    $soloB   = @($diff | Where-Object { $_.SideIndicator -eq "=>" -and $_.InputObject -ne "__EMPTY__" })

    # ── Resumen ──────────────────────────────────────────────────────────
    Write-SubHeader "Resumen de comparacion"
    $tipoLabel = if ($esUsuario) { "usuario" } else { "equipo" }
    Write-Host "  +-------------------------------+---------+" -ForegroundColor DarkGray
    Write-Host "  | GPOs                          |  Total  |" -ForegroundColor DarkGray
    Write-Host "  +-------------------------------+---------+" -ForegroundColor DarkGray
    Write-Host ("  | GPOs de [A] " + $objA.Name.PadRight(18) + " | " + "$($gposA.Count)".PadLeft(7) + " |") -ForegroundColor $colA
    Write-Host ("  | GPOs de [B] " + $objB.Name.PadRight(18) + " | " + "$($gposB.Count)".PadLeft(7) + " |") -ForegroundColor $colB
    Write-Host ("  | En comun                      | " + "$($comunes.Count)".PadLeft(7) + " |") -ForegroundColor Green
    Write-Host ("  | Solo en [A]                   | " + "$($soloA.Count)".PadLeft(7) + " |") -ForegroundColor $colA
    Write-Host ("  | Solo en [B]                   | " + "$($soloB.Count)".PadLeft(7) + " |") -ForegroundColor $colB
    Write-Host "  +-------------------------------+---------+" -ForegroundColor DarkGray

    # ── GPOs en comun ────────────────────────────────────────────────────
    if ($comunes.Count -gt 0) {
        Write-SubHeader "GPOs en COMUN (ambos las reciben)"
        foreach ($g in $comunes) {
            $infoA = $gposA | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
            $infoB = $gposB | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
            Write-Host "  [=] " -NoNewline -ForegroundColor Green
            Write-Host $g.InputObject -ForegroundColor White
            if ($infoA) { Write-Host ("      [A] vinculada en: " + $infoA.Nivel) -ForegroundColor $colA }
            if ($infoB) { Write-Host ("      [B] vinculada en: " + $infoB.Nivel) -ForegroundColor $colB }
        }
    }

    # ── Solo A ───────────────────────────────────────────────────────────
    if ($soloA.Count -gt 0) {
        Write-SubHeader "GPOs SOLO en [A]  $($objA.Name)  (no aplican a [B])"
        foreach ($g in $soloA) {
            $info = $gposA | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
            Write-Host "  [A] " -NoNewline -ForegroundColor $colA
            Write-Host $g.InputObject -ForegroundColor White
            if ($info) { Write-Host ("      Vinculada en: " + $info.Nivel) -ForegroundColor $colA }
        }
    }

    # ── Solo B ───────────────────────────────────────────────────────────
    if ($soloB.Count -gt 0) {
        Write-SubHeader "GPOs SOLO en [B]  $($objB.Name)  (no aplican a [A])"
        foreach ($g in $soloB) {
            $info = $gposB | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
            Write-Host "  [B] " -NoNewline -ForegroundColor $colB
            Write-Host $g.InputObject -ForegroundColor White
            if ($info) { Write-Host ("      Vinculada en: " + $info.Nivel) -ForegroundColor $colB }
        }
    }

    if ($soloA.Count -eq 0 -and $soloB.Count -eq 0 -and $comunes.Count -gt 0) {
        Write-Host ""
        Write-Host "  Ambos objetos reciben exactamente las mismas GPOs." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  [=] Comun   [A] Solo en A   [B] Solo en B" -ForegroundColor DarkGray

    # ── Exportar ─────────────────────────────────────────────────────────
    Write-Separador
    $exp = Read-Host "  Exportar comparacion a CSV? (S/N)"
    if ($exp -match "^[sS]$") {
        $ruta = Read-Host "  Ruta del archivo (ej: C:\Reportes\comp_gpo.csv)"
        try {
            $filas = @()
            $labelTipo = if ($esUsuario) { "Usuario" } else { "Equipo" }
            foreach ($g in $comunes) {
                $infoA = $gposA | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
                $infoB = $gposB | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
                $filas += [PSCustomObject]@{
                    GPO      = $g.InputObject
                    Estado   = "Comun"
                    NivelA   = if ($infoA) { $infoA.Nivel } else { "-" }
                    NivelB   = if ($infoB) { $infoB.Nivel } else { "-" }
                    "${labelTipo}A" = $objA.Name
                    "${labelTipo}B" = $objB.Name
                }
            }
            foreach ($g in $soloA) {
                $info = $gposA | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
                $filas += [PSCustomObject]@{
                    GPO      = $g.InputObject
                    Estado   = "Solo en A"
                    NivelA   = if ($info) { $info.Nivel } else { "-" }
                    NivelB   = "-"
                    "${labelTipo}A" = $objA.Name
                    "${labelTipo}B" = $objB.Name
                }
            }
            foreach ($g in $soloB) {
                $info = $gposB | Where-Object { $_.Nombre -eq $g.InputObject } | Select-Object -First 1
                $filas += [PSCustomObject]@{
                    GPO      = $g.InputObject
                    Estado   = "Solo en B"
                    NivelA   = "-"
                    NivelB   = if ($info) { $info.Nivel } else { "-" }
                    "${labelTipo}A" = $objA.Name
                    "${labelTipo}B" = $objB.Name
                }
            }
            $filas | Export-Csv -Path $ruta -NoTypeInformation -Encoding UTF8
            Write-Host "  OK - Exportado en: $ruta" -ForegroundColor Green
        } catch {
            Write-Host "  ERROR al exportar: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Pause-Pantalla
}

# ── MENU GPO ─────────────────────────────────────────────────────────────────
function Show-MenuGPO {
    Write-Header "MODULO GPO"
    Write-Host "  [1]  Ver detalles de una GPO  (configuraciones, vinculos, filtrado)"
    Write-Host "  [2]  GPOs que aplican a un grupo  (que permisos da ese grupo)"
    Write-Host "  [3]  GPOs que aplican a un usuario o equipo  (jerarquia de OUs)"
    Write-Host "  [4]  Comparar GPOs  (usuario vs usuario  /  equipo vs equipo)"
    Write-Host "  [0]  Volver al menu principal"
    Write-Host ""
    $opc = Read-Host "  Opcion"
    switch ($opc) {
        "1" { Get-InfoGPO        }
        "2" { Get-GPOsPorGrupo   }
        "3" { Get-GPOsDeObjeto   }
        "4" { Compare-GPO        }
        "0" { return }
        default { Write-Host "  Opcion no valida." -ForegroundColor Red; Start-Sleep 1 }
    }
}
# ============================================================
#  MODULO AD RECYCLE BIN - OBJETOS ELIMINADOS
# ============================================================

function Get-ObjetosEliminados {
    Write-Header "AD RECYCLE BIN - OBJETOS ELIMINADOS"

    # ── Resolver el dominio (con reintento manual de DC si ADWS no responde) ──
    Write-Host ""
    $dominioDN = $null
    $srvAD     = @{}
    try {
        $dominioDN = (Get-ADDomain -ErrorAction Stop).DistinguishedName
    } catch {
        Write-Host "  ERROR: no se pudo contactar Active Directory Web Services (ADWS)." -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Esto NO significa que el dominio este caido: significa que el DC" -ForegroundColor Yellow
        Write-Host "  elegido automaticamente no respondio en el puerto de ADWS (TCP 9389)." -ForegroundColor Yellow
        Write-Host "  Puede ser ese DC puntual con el servicio ADWS detenido, o un" -ForegroundColor Yellow
        Write-Host "  problema de red/firewall hacia el en este momento." -ForegroundColor Yellow
        Write-Host ""
        $dcManual = (Read-Host "  Reintentar especificando un DC? (nombre o IP, Enter = cancelar)").Trim()
        if ($dcManual) {
            try {
                $dominioDN = (Get-ADDomain -Server $dcManual -ErrorAction Stop).DistinguishedName
                $srvAD = @{ Server = $dcManual }
                Write-Host "  OK - conectado contra $dcManual" -ForegroundColor Green
            } catch {
                Write-Host "  ERROR tambien contra $dcManual : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        if (-not $dominioDN) { Pause-Pantalla; return }
    }
    $deletedBase = "CN=Deleted Objects,$dominioDN"

    # ── Verificar estado de la Recycle Bin ───────────────────────────────
    Write-Host ""
    $recycleActiva = $false
    try {
        $rb = Get-ADOptionalFeature @srvAD -Filter * -ErrorAction Stop |
              Where-Object { $_.Name -like "*Recycle*" }
        if ($rb -and $rb.EnabledScopes.Count -gt 0) {
            $recycleActiva = $true
            Write-Host "  [OK] Papelera de Reciclaje de AD: ACTIVA" -ForegroundColor Green
        } else {
            Write-Host "  [!]  Papelera de Reciclaje de AD: NO HABILITADA" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Sin ella los objetos borrados pierden atributos y no se pueden" -ForegroundColor DarkGray
            Write-Host "  restaurar completamente. Para habilitarla ejecuta en el DC:" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Enable-ADOptionalFeature 'Recycle Bin Feature' \" -ForegroundColor Cyan
            Write-Host "    -Scope ForestOrConfigurationSet \" -ForegroundColor Cyan
            Write-Host "    -Target (Get-ADForest).Name" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  NOTA: Requiere nivel funcional 2008 R2 o superior. La accion es irreversible." -ForegroundColor DarkGray
            Write-Host ""
            $continuar = Read-Host "  Continuar de todas formas buscando objetos eliminados? (S/N)"
            if ($continuar -notmatch "^[sS]$") { Pause-Pantalla; return }
        }
    } catch {
        Write-Host "  (No se pudo verificar la Recycle Bin: $($_.Exception.Message))" -ForegroundColor DarkGray
    }

    # ── Menu de modo de busqueda ─────────────────────────────────────────
    # Nota: IPv4Address no es un atributo almacenado en AD (es resolucion DNS dinamica),
    # por eso fue removido de $propsEliminado para evitar el error de parametro invalido.
    do {
    Write-Host ""
    Write-Host "  Modo de busqueda:" -ForegroundColor Yellow
    Write-Host "  [1]  Listar todos los eliminados  (por tipo)"
    Write-Host "  [2]  Buscar por ObjectSID  (ej: S-1-5-21-...-XXXXX)"
    Write-Host "  [0]  Volver al menu principal"
    Write-Host ""
    $modoOpc = Read-Host "  Modo"

    if ($modoOpc -eq "0") { return }

    if ($modoOpc -notin @("1","2")) {
        Write-Host "  Opcion no valida." -ForegroundColor Red
        continue
    }

    # IPv4Address removido - no es atributo real de AD (causa error "nombre del parametro: IPv4Address")
    # Es una resolucion DNS dinamica que no se puede recuperar de objetos eliminados

    # ── Propiedades a recuperar ──────────────────────────────────────────
    $propsEliminado = @(
        "Name","ObjectClass","ObjectSID","isDeleted","whenChanged","whenCreated",
        "lastKnownParent","msDS-LastKnownRDN","DistinguishedName","ObjectGUID",
        "Description","OperatingSystem","OperatingSystemVersion",
        "SamAccountName","UserPrincipalName","mail",
        "DisplayName","Department","Title","Manager",
        "DNSHostName","ManagedBy",
        "Created","Modified","Enabled"
    )

    # ── Funcion interna: mostrar detalle de un objeto eliminado ──────────
    function Show-DetalleEliminado {
        param($obj, [bool]$recycleActiva)

        # Resolver nombre original
        $nombreOriginal = $obj."msDS-LastKnownRDN"
        if (-not $nombreOriginal) {
            $nombreOriginal = $obj.Name -replace '\\\nDEL:.*','' -replace '\nDEL:.*',''
            $nombreOriginal = ($nombreOriginal -split "`n")[0].Trim()
        }
        if (-not $nombreOriginal) { $nombreOriginal = $obj.Name }

        $ouOrigen = if ($obj.lastKnownParent) { Get-OUdesdeDN $obj.lastKnownParent } else { "Desconocida" }
        $clase    = $obj.ObjectClass

        Write-Host ""
        Write-Host "  +============================================================+" -ForegroundColor Cyan
        Write-Host ("  |  DETALLE: " + $nombreOriginal.ToUpper().PadRight(51) + "|") -ForegroundColor White
        Write-Host "  +============================================================+" -ForegroundColor Cyan
        Write-Host ""

        Write-SubHeader "Identificacion"
        Write-Campo "Nombre original"      $nombreOriginal                                                         "White"
        Write-Campo "Tipo de objeto"       $clase                                                                  "White"
        Write-Campo "ObjectSID"            $(if ($obj.ObjectSID)   { $obj.ObjectSID.Value }   else { "No disponible" }) "Yellow"
        Write-Campo "ObjectGUID"           $(if ($obj.ObjectGUID)  { $obj.ObjectGUID.ToString() } else { "No disponible" }) "DarkGray"
        Write-Campo "SamAccountName"       $(if ($obj.SamAccountName) { $obj.SamAccountName } else { "-" })

        Write-SubHeader "Fechas"
        Write-Campo "Creado en AD"         $(if ($obj.whenCreated) { $obj.whenCreated.ToString("dd/MM/yyyy HH:mm:ss") } else { "Desconocida" })
        Write-Campo "Fecha de borrado"     $(if ($obj.whenChanged) { $obj.whenChanged.ToString("dd/MM/yyyy HH:mm:ss") } else { "Desconocida" }) "Red"

        Write-SubHeader "Ubicacion"
        Write-Campo "OU de Origen"         $ouOrigen
        Write-Campo "DN en Recycle Bin"    $obj.DistinguishedName   "DarkGray"

        if ($clase -eq "computer") {
            Write-SubHeader "Datos del Equipo"
            Write-Campo "SO"               $(if ($obj.OperatingSystem)        { $obj.OperatingSystem }        else { "-" })
            Write-Campo "Version SO"       $(if ($obj.OperatingSystemVersion) { $obj.OperatingSystemVersion } else { "-" })
            Write-Campo "DNS Hostname"     $(if ($obj.DNSHostName)            { $obj.DNSHostName }            else { "-" })
            Write-Campo "Descripcion"      $(if ($obj.Description)            { $obj.Description }            else { "-" })
        }
        if ($clase -eq "user") {
            Write-SubHeader "Datos del Usuario"
            Write-Campo "Nombre visible"   $(if ($obj.DisplayName)    { $obj.DisplayName }    else { "-" })
            Write-Campo "Email"            $(if ($obj.mail)           { $obj.mail }           else { "-" })
            Write-Campo "Departamento"     $(if ($obj.Department)     { $obj.Department }     else { "-" })
            Write-Campo "Cargo"            $(if ($obj.Title)          { $obj.Title }          else { "-" })
            Write-Campo "Descripcion"      $(if ($obj.Description)    { $obj.Description }    else { "-" })
        }

        Write-Host ""
        if ($recycleActiva) {
            Write-Badge "  Objeto restaurable (Recycle Bin activa)" "ok"
            Write-Host ""
            $confRestore = Read-Host "  Restaurar este objeto ahora? (S=Confirmar / W=Simular / N=Omitir)"
            if ($confRestore -match "^[wW]$") {
                Write-Host ""
                Write-Host "  [SIMULACION] Se restauraria:" -ForegroundColor Yellow
                Write-Host "  Restore-ADObject -Identity '$($obj.ObjectGUID)'" -ForegroundColor Cyan
                Write-Host "  Ningun objeto fue modificado." -ForegroundColor DarkGray
            } elseif ($confRestore -match "^[sS]$") {
                try {
                    Restore-ADObject @srvAD -Identity $obj.ObjectGUID -ErrorAction Stop
                    Write-Host ""
                    Write-Host "  OK - '$nombreOriginal' restaurado exitosamente." -ForegroundColor Green
                    try {
                        $verif = Get-ADObject @srvAD -Identity $obj.ObjectGUID -Properties DistinguishedName -ErrorAction Stop
                        Write-Host "  Verificacion: activo en $(Get-OUdesdeDN $verif.DistinguishedName)" -ForegroundColor Green
                    } catch {
                        Write-Host "  [!] No se pudo verificar de inmediato (puede tardar en replicar)." -ForegroundColor Yellow
                        Write-Host "      Revisa manualmente con Get-ADObject -Identity '$($obj.ObjectGUID)'." -ForegroundColor DarkGray
                    }
                } catch {
                    Write-Host ""
                    Write-Host "  ERROR al restaurar: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "  Verifica que el objeto no haya superado el tiempo de retencion" -ForegroundColor DarkGray
                    Write-Host "  (deletedObjectLifetime) y que tengas permisos suficientes." -ForegroundColor DarkGray
                }
            } else {
                Write-Host ""
                Write-Host "  Sin cambios." -ForegroundColor DarkGray
            }
        } else {
            Write-Badge "  Recycle Bin inactiva: atributos incompletos, restauracion limitada" "warn"
        }
        Write-Host ""
    }

    # ════════════════════════════════════════════════════════════════════
    #  MODO 1 — Listar todos los eliminados por tipo (con busqueda)
    # ════════════════════════════════════════════════════════════════════
    if ($modoOpc -eq "1") {

        Write-Host ""
        Write-Host "  Tipo de objeto a listar:" -ForegroundColor Yellow
        Write-Host "  [1]  Equipos"
        Write-Host "  [2]  Usuarios"
        Write-Host "  [3]  Grupos"
        Write-Host ""
        $tipoOpc = Read-Host "  Tipo"

        switch ($tipoOpc) {
            "1" { $filtroClase = "computer"; $tipoTxt = "EQUIPOS"  }
            "2" { $filtroClase = "user";     $tipoTxt = "USUARIOS" }
            "3" { $filtroClase = "group";    $tipoTxt = "GRUPOS"   }
            default {
                Write-Host "  Opcion no valida." -ForegroundColor Red
                continue
            }
        }

        Write-Host ""
        Write-Host "  Consultando eventos del DC para saber quien borro los objetos:" -ForegroundColor DarkGray
        $consultarDC = Read-Host "  Consultar eventos del DC? (S/N)"
        $dcNombre    = $null
        if ($consultarDC -match "^[sS]$") {
            $dcNombre = Read-Host "  Nombre o IP del Domain Controller"
        }

        Write-Host ""
        Write-Host "  Buscando $tipoTxt eliminados..." -ForegroundColor DarkGray

        try {
            $todosBorrados = Get-ADObject @srvAD `
                -Filter * `
                -IncludeDeletedObjects `
                -SearchBase $deletedBase `
                -Properties $propsEliminado `
                -ErrorAction Stop

            $eliminados = $todosBorrados |
                Where-Object { $_.isDeleted -eq $true -and $_.ObjectClass -eq $filtroClase } |
                Sort-Object whenChanged -Descending

            if (-not $eliminados -or @($eliminados).Count -eq 0) {
                Write-Host ""
                Write-Host "  No se encontraron $tipoTxt eliminados." -ForegroundColor Yellow
                if (-not $recycleActiva) {
                    Write-Host "  La Recycle Bin no estaba activa al momento del borrado." -ForegroundColor DarkGray
                }
                continue
            }

            # Cargar eventos del DC si se pidio
            $eventosSeguridad = @{}
            if ($dcNombre) {
                Write-Host "  Consultando eventos de seguridad en $dcNombre..." -ForegroundColor DarkGray
                $eventIds = switch ($filtroClase) {
                    "computer" { @(4743) }
                    "user"     { @(4726) }
                    "group"    { @(4730,4734,4758,4763) }
                }
                try {
                    $eventos = Get-WinEvent -ComputerName $dcNombre `
                        -FilterHashtable @{ LogName='Security'; Id=$eventIds } `
                        -ErrorAction Stop
                    foreach ($ev in $eventos) {
                        try {
                            $datosEvt = ([xml]$ev.ToXml()).Event.EventData.Data
                            $objName  = ($datosEvt | Where-Object { $_.Name -eq "TargetUserName" }).'#text'
                            $actor    = ($datosEvt | Where-Object { $_.Name -eq "SubjectUserName" }).'#text'
                            $dom      = ($datosEvt | Where-Object { $_.Name -eq "SubjectDomainName" }).'#text'
                            if ($objName -and -not $eventosSeguridad.ContainsKey($objName)) {
                                $eventosSeguridad[$objName] = [PSCustomObject]@{
                                    BorradoPor  = if ($actor -and $dom) { "$dom\$actor" } elseif ($actor) { $actor } else { "Desconocido" }
                                    FechaEvento = $ev.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
                                }
                            }
                        } catch { continue }
                    }
                    Write-Host "  $($eventosSeguridad.Count) eventos de borrado encontrados." -ForegroundColor DarkGray
                } catch {
                    Write-Host "  No se pudieron leer los eventos del DC: $($_.Exception.Message)" -ForegroundColor Yellow
                    $dcNombre = $null
                }
            }

            # Construir tabla de resultados
            $resultados = @()
            foreach ($objItem in $eliminados) {
                $nombreItem = $objItem."msDS-LastKnownRDN"
                if (-not $nombreItem) {
                    $nombreItem = ($objItem.Name -split "`n")[0].Trim() -replace '\\\nDEL:.*','' -replace '\nDEL:.*',''
                }
                if (-not $nombreItem) { $nombreItem = $objItem.Name }

                $borradoPor = if ($dcNombre -and $eventosSeguridad.ContainsKey($nombreItem)) {
                    $eventosSeguridad[$nombreItem].BorradoPor
                } elseif ($dcNombre) { "No encontrado en eventos" } else { "Sin datos del DC" }

                $sidStr = if ($objItem.ObjectSID) { $objItem.ObjectSID.Value } else { "-" }

                $resultados += [PSCustomObject]@{
                    Nombre           = $nombreItem
                    SID              = $sidStr
                    FechaEliminacion = if ($objItem.whenChanged) { $objItem.whenChanged.ToString("dd/MM/yyyy HH:mm") } else { "?" }
                    BorradoPor       = $borradoPor
                    OUOrigen         = if ($objItem.lastKnownParent) { Get-OUdesdeDN $objItem.lastKnownParent } else { "Desconocida" }
                    Clase            = $objItem.ObjectClass
                    ObjectGUID       = $objItem.ObjectGUID.ToString()
                    _AdObj           = $objItem   # referencia interna para detalle
                }
            }

            $totalLista = @($resultados).Count

            # ── Mostrar tabla completa ──────────────────────────────────────
            Write-Host ""
            Write-Host "  +============================================================+" -ForegroundColor Cyan
            Write-Host ("  |  $tipoTxt ELIMINADOS - Total: $totalLista".PadRight(63) + "|") -ForegroundColor White
            Write-Host "  +============================================================+" -ForegroundColor Cyan
            Write-Host ""

            # Numerar para facilitar la seleccion posterior
            $i = 1
            foreach ($r in $resultados) {
                $idx = "[$i]".PadRight(5)
                Write-Host ("  " + $idx) -NoNewline -ForegroundColor DarkCyan
                Write-Host ($r.Nombre.PadRight(30)) -NoNewline -ForegroundColor White
                Write-Host ($r.FechaEliminacion.PadRight(18)) -NoNewline -ForegroundColor Yellow
                Write-Host ($r.OUOrigen) -ForegroundColor DarkGray
                $i++
            }

            # ── Busqueda en los resultados listados ─────────────────────────
            Write-Host ""
            Write-Host "  Puedes buscar dentro de los resultados anteriores." -ForegroundColor DarkGray
            Write-Host "  Ingresa un texto para filtrar (nombre / OU) o un numero [#] para ver detalle." -ForegroundColor DarkGray
            Write-Host "  Deja vacio y presiona Enter para saltar a exportar." -ForegroundColor DarkGray

            do {
                Write-Host ""
                $busqInput = (Read-Host "  Buscar / Ver detalle (texto, #numero o Enter para salir)").Trim()

                if ([string]::IsNullOrWhiteSpace($busqInput)) { break }

                # Si el usuario escribe un numero, mostrar detalle de ese item
                if ($busqInput -match '^\d+$') {
                    $numSel = [int]$busqInput - 1
                    if ($numSel -ge 0 -and $numSel -lt $resultados.Count) {
                        Show-DetalleEliminado -obj $resultados[$numSel]._AdObj -recycleActiva $recycleActiva
                    } else {
                        Write-Host "  Numero fuera de rango. El listado tiene $totalLista elementos." -ForegroundColor Yellow
                    }
                } else {
                    # Filtrar por texto en nombre o OU
                    $filtrados = $resultados | Where-Object {
                        $_.Nombre  -like "*$busqInput*" -or
                        $_.OUOrigen -like "*$busqInput*" -or
                        $_.SID     -like "*$busqInput*"
                    }

                    if (-not $filtrados -or @($filtrados).Count -eq 0) {
                        Write-Host "  Sin coincidencias para '$busqInput'." -ForegroundColor Yellow
                    } else {
                        $arrFilt = @($filtrados)
                        Write-Host ""
                        Write-Host ("  Coincidencias: " + $arrFilt.Count) -ForegroundColor Green
                        Write-Host ""
                        $fi = 1
                        foreach ($r in $arrFilt) {
                            $idxF = "[$fi]".PadRight(5)
                            Write-Host ("  " + $idxF) -NoNewline -ForegroundColor DarkCyan
                            Write-Host ($r.Nombre.PadRight(30)) -NoNewline -ForegroundColor White
                            Write-Host ($r.FechaEliminacion.PadRight(18)) -NoNewline -ForegroundColor Yellow
                            Write-Host ($r.OUOrigen) -ForegroundColor DarkGray
                            $fi++
                        }
                        Write-Host ""
                        $selDet = (Read-Host "  Numero para ver detalle (o Enter para seguir buscando)").Trim()
                        if ($selDet -match '^\d+$') {
                            $numDet = [int]$selDet - 1
                            if ($numDet -ge 0 -and $numDet -lt $arrFilt.Count) {
                                Show-DetalleEliminado -obj $arrFilt[$numDet]._AdObj -recycleActiva $recycleActiva
                            } else {
                                Write-Host "  Numero fuera de rango." -ForegroundColor Yellow
                            }
                        }
                    }
                }
            } while ($true)

            # ── Exportar ────────────────────────────────────────────────────
            $datosExport = @($resultados | Select-Object Nombre, SID, FechaEliminacion, BorradoPor, OUOrigen, Clase, ObjectGUID)
            Write-Separador
            Export-DatosCSV -Data $datosExport -NombreArchivoBase "${tipoTxt}_Eliminados" `
                -Prompt "  Exportar $totalLista resultados a CSV? (S/N)" | Out-Null

        } catch {
            Write-Host ""
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }

    } # fin modoOpc = 1

    # ════════════════════════════════════════════════════════════════════
    #  MODO 2 — Buscar por ObjectSID
    # ════════════════════════════════════════════════════════════════════
    elseif ($modoOpc -eq "2") {

        Write-Host ""
        Write-Host "  Ingresa el SID completo, por ejemplo:" -ForegroundColor DarkGray
        Write-Host "  S-1-5-21-433858523-2103985357-1423778804-122575" -ForegroundColor DarkGray
        Write-Host ""
        $sidInput = (Read-Host "  ObjectSID").Trim()

        if ([string]::IsNullOrWhiteSpace($sidInput)) {
            Write-Host "  El SID no puede estar vacio." -ForegroundColor Red
            continue
        }

        # Validar formato basico del SID
        if ($sidInput -notmatch '^S-\d+-\d+(-\d+)+$') {
            Write-Host "  Formato de SID invalido." -ForegroundColor Red
            Write-Host "  Formato esperado: S-1-5-21-XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXX" -ForegroundColor Yellow
            continue
        }

        Write-Host ""
        Write-Host "  Buscando SID en la Recycle Bin..." -ForegroundColor DarkGray

        try {
            $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sidInput)

            $todos = Get-ADObject @srvAD `
                -Filter * `
                -IncludeDeletedObjects `
                -SearchBase $deletedBase `
                -Properties $propsEliminado `
                -ErrorAction Stop |
                Where-Object { $_.isDeleted -eq $true }

            $encontrado = @($todos | Where-Object {
                $_.ObjectSID -and $_.ObjectSID.Value -eq $sidObj.Value
            })

            if ($encontrado.Count -eq 0) {
                Write-Host ""
                Write-Host "  No se encontro ningun objeto con ese SID en la Recycle Bin." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Posibles causas:" -ForegroundColor DarkGray
                Write-Host "  - La Recycle Bin no estaba activa cuando se borro el objeto" -ForegroundColor DarkGray
                Write-Host "  - El objeto fue eliminado definitivamente (tombstone expirado)" -ForegroundColor DarkGray
                Write-Host "  - El SID ingresado tiene un error tipografico" -ForegroundColor DarkGray
            } else {
                Write-Host "  $($encontrado.Count) objeto(s) encontrado(s) con ese SID." -ForegroundColor Green
                foreach ($objSID in $encontrado) {
                    Show-DetalleEliminado -obj $objSID -recycleActiva $recycleActiva
                    Write-Separador
                }
            }
        } catch {
            Write-Host ""
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            if ($_.Exception.Message -like "*SID*" -or $_.Exception.Message -like "*format*") {
                Write-Host "  Verifica que el SID tenga el formato correcto: S-1-5-21-...-XXXXX" -ForegroundColor Yellow
            }
        }

    } # fin modoOpc = 2

    Write-Host ""
    $otraBusqueda = Read-Host "  Realizar otra busqueda en la Recycle Bin? (S/N)"
    } while ($otraBusqueda -match "^[sS]$")
}

# ============================================================
#  MENU PRINCIPAL
# ============================================================

function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "  +================================================+" -ForegroundColor Cyan
    Write-Host "  |        AD INFO TOOL  -  Active Directory       |" -ForegroundColor Cyan
    Write-Host "  +================================================+" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  --- USUARIOS ------------------------------------" -ForegroundColor Yellow
    Write-Host "  [1]  Ver info detallada de un usuario"
    Write-Host "  [2]  Buscar usuarios  (nombre / depto / OU / grupo / descripcion / ...)"
    Write-Host "  [6]  Comparar dos usuarios  (atributos y grupos)"
    Write-Host "  [8]  Desbloquear usuario"
    Write-Host "  [13] Habilitar / Deshabilitar usuario"
    Write-Host "  [15] Auditoria de usuario  (creacion / modificaciones / baja)"
    Write-Host ""
    Write-Host "  --- EQUIPOS -------------------------------------" -ForegroundColor Yellow
    Write-Host "  [3]  Ver info detallada de un equipo"
    Write-Host "  [4]  Buscar equipos"
    Write-Host "  [7]  Comparar dos equipos  (OU, atributos y grupos)"
    Write-Host "  [11] Auditoria de hostname  (creacion / modificaciones / baja)"
    Write-Host ""
    Write-Host "  --- GPO / POLITICAS -----------------------------" -ForegroundColor Yellow
    Write-Host "  [9]  Consultar GPOs  (por grupo, usuario o equipo)"
    Write-Host "  [12] Comparar GPOs  (usuario vs usuario  /  equipo vs equipo)"
    Write-Host ""
    Write-Host "  --- AUDITORIA / REPORTES ------------------------" -ForegroundColor Yellow
    Write-Host "  [5]  Reportes de auditoria  (bloqueados / inactivos / privilegios / ...)"
    Write-Host "  [10] AD Recycle Bin  (ver objetos eliminados)"
    Write-Host "  [14] Auditoria de movimiento de OU  (quien la movio, desde/hacia donde)"
    Write-Host ""
    Write-Host "  [0]  Salir"
    Write-Host ""
}

# ============================================================
#  VERIFICACION DEL MODULO AD
# ============================================================

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host ""
    Write-Host "  ERROR: El modulo ActiveDirectory no esta instalado." -ForegroundColor Red
    Write-Host "  Instalalo con: Install-WindowsFeature RSAT-AD-PowerShell" -ForegroundColor Yellow
    Write-Host "  O desde: Configuracion > Caracteristicas opcionales > RSAT" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

# ============================================================
#  LOOP PRINCIPAL
# ============================================================

do {
    Show-Menu
    $sel = Read-Host "  Selecciona una opcion"

    switch ($sel) {
        "1"  { Get-InfoUsuario           }
        "2"  { Search-Usuarios           }
        "3"  { Get-InfoEquipo            }
        "4"  { Search-Equipos            }
        "5"  { Export-Reporte            }
        "6"  { Compare-Usuarios          }
        "7"  { Compare-Equipos           }
        "8"  { Unlock-Usuario            }
        "9"  { Show-MenuGPO              }
        "12" { Compare-GPO              }
        "10" { Get-ObjetosEliminados     }
        "11" { Get-AuditoriaEquipo       }
        "13" { Enable-Usuario            }
        "14" { Get-AuditoriaOU           }
        "15" { Get-AuditoriaUsuario      }
        "0" {
            Write-Host ""
            Write-Host "  Hasta luego." -ForegroundColor Cyan
            Write-Host ""
        }
        default {
            Write-Host ""
            Write-Host "  Opcion no valida. Intenta de nuevo." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
} while ($sel -ne "0")
