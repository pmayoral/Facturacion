# Sintaxis Markdown para Texto Libre

El campo "Texto Libre" utiliza una sintaxis Markdown simplificada para guardar formato en Google Sheets.

## 📝 Formatos de Texto

### Negrita
```
**texto en negrita**
```
Ejemplo: **Importante**

### Cursiva
```
*texto en cursiva*
```
Ejemplo: *énfasis*

### Subrayado
```
__texto subrayado__
```
Ejemplo: __destacado__

## 📏 Tamaños de Fuente

### Muy Grande
```
## Título principal
```
Ejemplo: ## AVISO IMPORTANTE

### Grande
```
# Subtítulo
```
Ejemplo: # Condiciones de pago

### Normal
```
Texto sin modificadores
```
Ejemplo: Este es texto normal

### Pequeño
```
texto pequeño~
```
Ejemplo: Nota adicional~

### Muy Pequeño
```
texto muy pequeño~~
```
Ejemplo: Texto legal~~

## 📋 Listas

```
- Primer elemento
- Segundo elemento
- Tercer elemento
```

Ejemplo:
- Incluye IVA
- Pago a 30 días
- Transferencia bancaria

## 💡 Ejemplos Completos

### Ejemplo 1: Condiciones de pago
```
## CONDICIONES DE PAGO
**Forma de pago:** Transferencia bancaria
**Plazo:** *30 días desde fecha de factura*
**IBAN:** __ES00 1234 5678 9012 3456 7890__

Gracias por su confianza~
```

### Ejemplo 2: Información bancaria
```
# Datos bancarios
- Banco: BBVA
- IBAN: ES00 1234 5678 9012 3456 7890
- Titular: Mi Empresa S.L.

*Incluya número de factura en concepto*
```

### Ejemplo 3: Avisos legales
```
## AVISO IMPORTANTE
Esta factura debe ser **abonada** antes del __31/12/2024__

Condiciones generales disponibles en www.miempresa.com~~
```

## ⚙️ Funcionamiento Técnico

- **Al guardar**: El HTML del editor se convierte automáticamente a Markdown
- **En Google Sheets**: Se guarda como texto plano con sintaxis Markdown (legible)
- **Al cargar**: El Markdown se convierte de vuelta a HTML para el editor
- **En la factura impresa**: Se renderiza con todo el formato aplicado

## 🔍 Ventajas

✅ **Legible** - Incluso en Google Sheets sin procesar
✅ **Portable** - Se puede editar manualmente en Sheets si es necesario
✅ **Estándar** - Usa sintaxis Markdown ampliamente conocida
✅ **Reconstruible** - Se puede convertir de vuelta a HTML sin pérdida

