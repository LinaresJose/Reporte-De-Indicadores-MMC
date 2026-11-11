# LYNX — Sistema de Indicadores de Ventas

> Aplicación Python que replica y automatiza el reporte de indicadores de ventas
> originalmente implementado en un archivo Excel con macros VBA (`.xlsm`),
> importando directamente los 6 archivos exportados desde Odoo ERP.

---

## 📋 Tabla de contenidos

- [Descripción](#descripción)
- [Requisitos](#requisitos)
- [Instalación](#instalación)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Uso](#uso)
- [Mantenimiento mensual](#mantenimiento-mensual)
- [Archivos fuente de Odoo](#archivos-fuente-de-odoo)
- [Reportes generados](#reportes-generados)
- [Lógica de KPIs](#lógica-de-kpis)
- [Pruebas](#pruebas)

---

## Descripción

**LYNX** genera automáticamente:

| Reporte | Descripción |
|---------|-------------|
| `Reporte_General_MES.xlsx` | Hoja General (34 cols) + 8 hojas regionales |
| `PDF_INDICADORES/` | Un PDF A4 por vendedor con los 4 KPIs |
| `REGIONALES/` | Un `.xlsx` limpio por región |
| `manifest_indicadores_vendedores.csv` | Listado de PDFs generados |
| `manifest_indicadores_regionales.csv` | Listado de Excels regionales |

---

## Requisitos

- **Python 3.11+** (probado con 3.14.3)
- Windows 10/11 (también funciona en Linux/macOS)

### Dependencias Python

```
pandas>=2.0
openpyxl>=3.1
fpdf2>=2.8
pillow>=10.0
python-dateutil>=2.9
```

---

## Instalación

```powershell
# 1. Clonar el repositorio
git clone https://github.com/LinaresJose/Reporte-De-Indicadores-MMC.git
cd Reporte-De-Indicadores-MMC

# 2. (Opcional) Crear entorno virtual
py -m venv venv
.\venv\Scripts\Activate.ps1

# 3. Instalar dependencias
py -m pip install -r requirements.txt

# 4. Ejecutar
py main.py
```

---

## Estructura del proyecto

```
lynx_reportes/
├── main.py                     # Punto de entrada — configura logging y lanza la GUI
├── config.py                   # Constantes globales, cabeceras, regiones, KPIs
├── requirements.txt
│
├── core/                       # Lógica de negocio pura (sin dependencias GUI)
│   ├── etl.py                  # Lectura y limpieza de los 6 archivos Odoo
│   ├── engine.py               # Cálculo de KPIs, rankings, totales
│   ├── models.py               # Dataclasses: VendorRecord, BudgetRecord, etc.
│   └── exceptions.py           # Jerarquía de excepciones LYNX
│
├── reports/                    # Generadores de reportes
│   ├── excel_writer.py         # Workbook consolidado + hojas regionales
│   ├── pdf_writer.py           # PDFs individuales por vendedor
│   └── manifest_writer.py      # CSVs de manifiestos
│
├── maintenance/                # Tareas de mantenimiento mensual
│   ├── budget_loader.py        # Cargar archivo de Presupuestos del gerente
│   └── cobranza_loader.py      # Cargar INDICADOR_COBRANZA cierre de mes
│
├── gui/                        # Interfaz gráfica Tkinter
│   ├── app.py                  # Ventana principal con 3 pestañas
│   └── widgets.py              # FilePickerRow, LogTextArea, SectionFrame
│
├── data/                       # Archivos de configuración (no incluidos en repo)
│   ├── zonas.xlsx              # Tabla maestra: Nombre, Zona, Región
│   ├── presupuestos.xlsx       # Presupuestos guardados automáticamente
│   └── config.json             # Preferencias de la última sesión
│
├── output/                     # Salida generada (ignorada en .gitignore)
│   ├── Reporte_General_*.xlsx
│   ├── PDF_INDICADORES/
│   ├── REGIONALES/
│   └── lynx.log
│
└── tests/
    ├── test_etl.py
    └── test_engine.py
```

---

## Uso

### 1. Ejecutar la aplicación

```powershell
py main.py
```

### 2. Pestaña "Importar Datos"

Seleccione los 6 archivos exportados desde Odoo (ver [sección abajo](#archivos-fuente-de-odoo)).

### 3. Pestaña "Generar Reportes"

| Botón | Resultado |
|-------|-----------|
| 📊 Reporte General + Regionales | Excel consolidado + 8 Excels regionales |
| 📄 Generar PDFs Individuales | Un PDF A4 por vendedor |
| 🚀 Generar TODO | Todo lo anterior + manifiestos CSV |

---

## Mantenimiento mensual

### A. Actualizar Presupuestos (metas del mes)

> El gerente entrega un Excel con columnas: `Nombre`, `VENTAS`, `ACTIVACIÓN`, `SKU`, `Region`.

1. Ir a pestaña **Mantenimiento**.
2. Sección **"Presupuestos"** → seleccionar el archivo del gerente → clic **"Cargar y guardar Presupuestos"**.
3. El sistema guarda los datos en `data/presupuestos.xlsx` automáticamente.

### B. Cambiar el mes del reporte

1. Pestaña **Mantenimiento** → sección **"Mes del reporte"**.
2. Seleccionar mes en el desplegable y año en el spinner.
3. Clic **"Aplicar mes"**.

### C. Cobranza cierre de mes (INDICADOR_COBRANZA) vs. Mes en curso

> **Mes en curso (Automático desde Odoo):** Mientras transcurre el mes, el sistema calcula automáticamente el `% Sin Vencer` a partir del archivo de Odoo `cuentas_por_cobrar_cierre.xlsx` usando la política de crédito comercial (32 días estándar, 45 días clientes VIP). No requiere intervención manual y la columna nunca queda vacía.
>
> **Cierre de mes (Definitivo):** Al finalizar el mes, el gerente entrega el archivo oficial `INDICADOR_COBRANZA_MES_AÑO.xlsx`.
> Este archivo actualiza la cobranza con los datos auditados de la hoja **"RESUMEN"**.

1. Pestaña **Mantenimiento** → sección **"Cobranza Cierre (fin de mes)"**.
2. Seleccionar el archivo `INDICADOR_COBRANZA_*.xlsx`.
3. Clic **"Cargar Cobranza Cierre"**.
4. El sistema guarda automáticamente una copia persistente en `data/cobranza_cierre.xlsx` y activa el modo de cierre definitivo.
5. Si desea volver al cálculo dinámico desde Odoo para el nuevo mes, use el botón **"Modo Automático (Odoo)"**.

---

## Archivos fuente de Odoo

| Archivo | Descripción | Columnas clave |
|---------|-------------|----------------|
| `facturado_cierre_.xlsx` | Facturas emitidas del período | Comercial, Subtotal, NIF, Ref. interna |
| `Pedido_recibido_cierre.xlsx` | Pedidos de venta | Comercial, Referencia, Base imponible |
| `devoluciones_cierre.xlsx` | Notas de crédito / devoluciones | Comercial, Subtotal |
| `cuentas_por_cobrar_cierre.xlsx` | Saldo por cobrar | Comercial, Fecha de entrega, Importe |
| `Pedido_por_facturar_cierre_.xlsx` | Pedidos pendientes de facturar | Comercial, Base imponible |
| `backorder_cierre.xlsx` | Backorders pendientes | Comercial, Backorder Value |

---

## Reportes generados

### Excel Consolidado — hoja General (34 columnas)

| Col | Campo |
|-----|-------|
| A | Nombre |
| B | Zona |
| C | Región |
| D | Cant. Pedidos |
| E | Monto Pedidos |
| F | Devoluciones |
| G | Total Facturado |
| H | Ventas Netas |
| I | Monto por Procesar |
| J | BackOrder |
| K | Presup. Ventas |
| L | Diferencia Ventas |
| M | Alcance Ventas |
| N | Puntos Ventas (30) |
| O | Clientes Activados |
| P | Presup. Activación |
| Q | Diferencia Activación |
| R | Alcance Activación |
| S | Puntos Activación (10) |
| T | SKUs Vendidos |
| U | Presup. SKUs |
| V | Diferencia SKUs |
| W | Alcance SKUs |
| X | Puntos SKUs (10) |
| Y | % Sin Vencer |
| Z | Puntos Cobranza (50) |
| AA | Total Indicadores |
| AB | Incentivo |
| AC | Lugar Ventas |
| AD | Lugar Activación |
| AE | Lugar SKUs |
| AF | Lugar Indicadores |
| AG | Rifa |
| AH | Participación |

---

## Lógica de KPIs

| KPI | Puntaje máx. | Fórmula |
|-----|-------------|---------|
| Ventas | 30 | `MIN(30, 30 × ventas_netas / presup_ventas)` |
| Cobranza | 50 | `50 × (sin_vencer / (sin_vencer + vencido))` |
| Activación | 10 | `MIN(10, 10 × clientes_activados / presup_activacion)` |
| SKUs | 10 | `MIN(10, 10 × skus_vendidos / presup_skus)` |
| **Total** | **100** | Suma de los 4 anteriores |

### Bonos

| Bono | Condición |
|------|-----------|
| **Incentivo** (1 % comisión extra) | `alcance_ventas ≥ 89.45 %` **Y** `pct_sin_vencer ≥ 89.45 %` |
| **Rifa** | `alcance_activacion ≥ 90 %` **Y** `alcance_skus ≥ 90 %` |

### Corte horario

El sistema usa **hoy** como fecha de corte si la hora actual es ≥ 16:30, de lo contrario usa **ayer**.
Esto determina cuántos días hábiles han transcurrido en el mes.

---

## Pruebas

```powershell
# Instalar pytest si no está instalado
py -m pip install pytest

# Ejecutar todas las pruebas
py -m pytest tests/ -v

# Sólo pruebas del motor de KPIs
py -m pytest tests/test_engine.py -v

# Sólo pruebas de ETL (requiere archivos fuente)
py -m pytest tests/test_etl.py -v
```

---

## Configuración de zonas (`data/zonas.xlsx`)

Cree un archivo Excel con las siguientes columnas en la hoja principal:

| Nombre | Zona | Región |
|--------|------|--------|
| MARIA GARCIA | ZC-01 | CAPITAL |
| PEDRO LOPEZ | ZL-02 | LLANOS |
| … | … | … |

Solo los vendedores presentes en este archivo aparecerán en el reporte.

---

## Licencia

Uso interno — MMC Venezuela. Todos los derechos reservados.
