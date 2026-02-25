# Excel VBA Automation Toolkit: Pricing & Logistics 📊🚀

## 📋 Descripción
Este repositorio contiene un conjunto de herramientas avanzadas desarrolladas en **VBA (Visual Basic for Applications)** para optimizar la cadena de suministro de datos en el sector Car Rental. El sistema automatiza la validación, transformación y exportación de tarifas masivas hacia plataformas corporativas y sistemas de terceros como Crossborder Xpress.

## 🛡️ Nota sobre Confidencialidad y Ética
El código contenido en este repositorio ha sido **anonimizado y sanitizado**. Los nombres de socios comerciales, tipos de tarifas específicas y rutas de servidores locales han sido reemplazados por etiquetas genéricas (`TYPE_A`, `PARTNER_01`, etc.) para proteger la propiedad intelectual de la empresa de origen, manteniendo intacta la arquitectura lógica y la funcionalidad técnica del software.

## 🛠️ Herramientas Incluidas

### 1. Rule Engine (`Rule_Engine.vba`)
- **Función:** Valida y asigna reglas de negocio dinámicas a cada tarifa antes de su carga al sistema central.
- **Capacidades:** - Verificación preventiva de campos obligatorios (Location, Effective Date/Time).
  - Uso de diccionarios de datos (`Scripting.Dictionary`) para garantizar la integridad de valores únicos.
  - Clasificación automática de registros según la longitud y prefijo de los códigos de locación.

### 2. Rate Generator (`Rate_Generator.vba`)
- **Función:** Pipeline de procesamiento que transforma datos crudos en archivos de carga masiva (CSV).
- **Capacidades:** - Filtrado inteligente por marca (Hertz, Dollar, Thrifty, Firefly).
  - Normalización de precisión numérica (redondeo a 2 decimales en columnas financieras).
  - Lógica multimoneda automática (USD/MXN) basada en la parametrización de la tarifa.

### 3. CBX Processor (`CBX_Processor.vba`)
- **Función:** Módulo de exportación para la plataforma Crossborder Xpress.
- **Capacidades:** - Validación masiva de celdas para prevenir valores negativos o nulos.
  - Fragmentación automática de datos en múltiples archivos CSV según el tipo de servicio (CBX, DCBX, TCBX).
  - Limpieza automática de metadatos y columnas de control antes de la exportación final.

## ⚙️ Habilidades Técnicas Demostradas
* **Automatización de Procesos (RPA Lite):** Reducción de tiempos de carga de horas a segundos.
* **Manejo de Errores (Error Handling):** Implementación de mensajes críticos y salidas controladas para evitar corrupción de datos.
* **Data Wrangling en Excel:** Limpieza y estructuración de datos para asegurar interoperabilidad entre sistemas.

## 📈 Impacto de Negocio
- **Eliminación de Errores Manuales:** Se eliminó el riesgo de rechazo por parte del sistema receptor mediante validaciones previas al commit.
- **Escalabilidad Operativa:** Capacidad para procesar cientos de combinaciones de tarifas y locaciones con un solo clic.

---
**Desarrollado por:** [Leonardo Morales](https://github.com/LeoMorales-Dev)
