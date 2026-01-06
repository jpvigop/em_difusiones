# Script de Difusión WhatsApp - Filtrado de Contactos

Script para filtrar contactos de campañas de WhatsApp, eliminando duplicados, socios/exsocios, y clientes que ya compraron recientemente.

## Requisitos

- Python 3.10+
- Librerías: `pandas`, `openpyxl`

```bash
pip install pandas openpyxl
```

## Configuración

Editar las variables al inicio del archivo `script.py`:

```python
# Archivos de entrada
CANDIDATOS_FILE = "archivos/enviar.xlsx"    # Contactos a enviar
EXCLUIR_FILE = "archivos/excluir.xlsx"      # Teléfonos a excluir

# Nombre de la campaña (nombre del archivo de salida)
CAMPAIGN_NAME = "nombre_campana"

# Filtros
MIN_SALIDAS = 1          # Mínimo de compras
MAX_SALIDAS = 2          # Máximo de compras
MAX_DESCUENTO_PERMITIDO = 10  # Excluir clientes con descuento > 10%
```

## Archivos de Entrada

### 1. Archivo de Candidatos (`archivos/enviar.xlsx`)

**¿De dónde se saca?**
> [COMPLETAR: Describir de qué reporte o sección del sistema RPCRM se exporta este archivo]
> 
> Ejemplo: "Desde Estadística de Compras y Ventas → Filtrar por producto X → Rango de fechas de 3 a 6 meses → Exportar a Excel"

**Columnas requeridas:**
- `Nombre` - Nombre del cliente
- `Celular` - Número(s) de teléfono (puede tener varios separados por espacio)
- `Salidas` - Cantidad de compras
- `Ventas` - Monto de ventas (opcional)

### 2. Archivo de Exclusión (`archivos/excluir.xlsx`)

**¿De dónde se saca?**
> [COMPLETAR: Describir de qué reportes se sacan los teléfonos a excluir]
> 
> Este archivo debe contener los teléfonos de:
> - [ ] Socios activos: [COMPLETAR: de dónde se exporta]
> - [ ] Ex-socios: [COMPLETAR: de dónde se exporta]
> - [ ] Clientes que compraron en el período reciente (ej: últimos 3 meses): [COMPLETAR: de dónde se exporta]
> - [ ] Clientes que recibieron difusión reciente (últimos 30 días): [COMPLETAR: de dónde se exporta]

**Formato:** Puede ser cualquier estructura - el script extrae automáticamente todos los números de teléfono de cualquier celda.

## Uso

1. Colocar archivos en la carpeta `archivos/`
2. Configurar las variables en `script.py`
3. Ejecutar:

```bash
python script.py
```

## Archivo de Salida

Se genera un archivo Excel en `output/{CAMPAIGN_NAME}.xlsx` con dos hojas:

- **Envios**: Contactos filtrados listos para enviar
- **Excluidos**: Contactos excluidos por descuento >10% (con motivo)

### Columnas de salida (hoja Envios):
| Columna | Descripción |
|---------|-------------|
| Código/N° | ID del cliente |
| Nombre | Nombre completo |
| Nombre limpio | Primera palabra del nombre (capitalizada) |
| Salidas | Cantidad de compras |
| Ventas | Monto total |
| Telefono | Número normalizado (formato 598XXXXXXXX) |

## Lógica de Filtrado

El script aplica los filtros en este orden:

1. **Filtro de Descuento**: Excluye clientes con >10% de descuento en el nombre
2. **Filtro de Salidas/Ventas**: Solo mantiene clientes con Salidas entre MIN y MAX
3. **Filtro de Teléfono**: Excluye si algún teléfono está en la lista de exclusión
4. **Filtro sin teléfono**: Excluye registros sin teléfono válido
5. **Deduplicación**: Elimina duplicados por número de teléfono

## Normalización de Teléfonos Uruguay

El script normaliza automáticamente los formatos:
- `099123456` → `59899123456`
- `99123456` → `59899123456`
- `+598 99 123 456` → `59899123456`
- `099 123 456` → `59899123456`
- `098702834(MADRE)` → extrae `098702834`

## Notas

- Los teléfonos múltiples en una celda (separados por espacio) se procesan correctamente
- Si un cliente tiene múltiples teléfonos y UNO está en la lista de exclusión, se excluye al cliente completo
- El archivo de exclusión puede tener cualquier formato - el script busca patrones numéricos en todas las celdas

