# 🧾 Automatización de Documentación para Administración de Contratos en el Servicio Público del Ecuador

Este proyecto automatiza la extracción de información clave desde órdenes de compra en PDF y facilita su integración con documentos Word vinculados a un archivo maestro de Excel, cumpliendo con los procesos administrativos del sector público en Ecuador.

---

## 🚀 Funcionalidades

- 📄 Extrae automáticamente los siguientes campos desde PDFs:
  - Número de orden
  - Objeto de contratación
  - Proveedor
  - Plazo de ejecución
  - Fecha de suscripción
  - Fecha límite de entrega
  - Fecha de la orden de compra
  - Número de certificación presupuestaria
  - Valor sin IVA (desde bloques tipo `UNIDAD ... $ ...`)
- 📋 Imprime los resultados listos para copiar en Excel o insertar automáticamente.
- 🌐 Soporte de fechas con nombres de mes en español (`locale`).
- 🤝 Compatible con flujos Word–Excel existentes vía vínculos o combinación de correspondencia.

---

## 🧠 Requisitos

- Python 3.8 o superior
- Librerías: PyMuPDF y openpyxl

---

## 📁 Estructura del Proyecto

AUTOMATIZATION-OF-CONTRACT-ADMINISTRATION-DOCUMENTATION-IN-ECUADOR-PUBLIC-SERVICE/

│

├──> extractor.py           # Script principal de extracción

├──> README.md              # Este archivo

├──> datos.xlsx             # Excel maestro con los datos extraídos

├──> plantilla.docx         # Plantilla Word vinculada al Excel

└──> ejemplos/              # PDFs de ejemplo

---

## ⚙️ ¿Cómo usar?

1. Coloca un archivo PDF en la carpeta ejemplos/.
2. Ejecuta el script de python.
3. Ingresa la fecha de suscripción cuando se solicite.
4. Los datos se imprimirán en consola y pueden agregarse automáticamente a datos.xlsx.
5. Abre tu documento Word vinculado al Excel para ver la información reflejada automáticamente.
