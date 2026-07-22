# Especificaciones Técnicas y Documentación SDD: FACE-Sync V2 (MVP PNUD)

**Cliente / Organización:** RED PROCOSI — Programa PNUD (Bolivia)  
**Aplicación:** FACE-Sync V2 (SaaS Web)  
**Propósito:** Automatizar la extracción del Mayor Presupuestario de VISUAL e inyectarlo en la plantilla oficial FACE (`SR_Transaction_Details`), con soporte para exención tributaria (87%/13%), lectura de celda K22 e identificación de `Account` y `Activity ID`.

---

## 1. Reglas de Negocio e Inyección de Datos

### A. Presupuesto Total del Proyecto
* **Celda Objetivo:** `SR_Transaction_Details!K22` (Fila 22, Columna 11).
* Si la plantilla contiene el valor en `K22`, la aplicación lo toma de manera prioritaria como el techo presupuestario oficial para los KPIs y el % de absorción.

### B. Cruce Compuesto de Importación (`Budget Line No`)
$$\text{Clave de Cruce} = (\text{Activity ID}, \text{UNDP Budget Account})$$
* **Activity ID:** Extraído de la glosa (`ACT: 23` o `/ 23 /`).
* **Account:** Código de 5 dígitos extraído del encabezado de la partida en VISUAL (ej. `71610`).
* **Matriz de Respaldo:** El sistema incluye por defecto la matriz local `data/FACE MALARIA_pestaña Budget.csv` con 77 claves únicas de cruce.

### C. Mapeo en la Plantilla FACE (`SR_Transaction_Details` desde Fila 22)

| Columna Excel | Campo en FACE | Tipo de Dato / Fórmula |
| :---: | :--- | :--- |
| **B (2)** | `Serial Number` | Contador secuencial (`1, 2, 3...`) |
| **C (3)** | `Date` | Fecha en formato nativo Excel (`yyyy-mm-dd`) |
| **D (4)** | `Voucher No.` | Comprobante contable de VISUAL |
| **E (5)** | `Payment / Ref` | Número de Cheque o Transferencia (`CH-xxxx`) |
| **F (6)** | `Payee / Vendor` | Nombre del Proveedor |
| **G (7)** | `Description` | Detalle o concepto de la transacción |
| **H (8)** | `Budget Line No.` | Calculado mediante la combinación `Activity ID` + `Account` |
| **I (9)** | `Account` | Código de cuenta de 5 dígitos (ej. `71610`) |
| **N (14)** | `Payment` | Importe Total en Bs |
| **Q (17)** | `Expenses` | **87%** del Importe Total |
| **R (18)** | `VAT tax` | **13%** del Importe Total (IVA/Importaciones) |

---

## 2. Estructura de Archivos

```text
├── assets/Logo_PROCOSI_OK_1.png
├── data/FACE MALARIA_pestaña Budget.csv
├── app.py
├── lector_mayor.py
├── requirements.txt
└── README.md