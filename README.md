# Risk Analysis Agent

Repositorio asociado a la prueba de concepto desarrollada para automatizar el análisis de riesgos de proveedores SaaS mediante n8n, RAG, Qdrant, Ollama y agentes de inteligencia artificial.

El objetivo del proyecto es reducir la carga manual asociada a la revisión de cuestionarios de seguridad, automatizando la ingesta de información, el análisis de respuestas, la generación de follow-ups, la consolidación de riesgos y la creación de un informe final en formato DOCX.

---

## 1. Descripción general

Este repositorio contiene los principales artefactos técnicos utilizados durante el desarrollo del prototipo:

* Workflows exportados de n8n.
* Scripts Python utilizados para ingesta, RAG, scoring y generación de informes.
* Catálogo de riesgos transformado a formato JSON.
* Artefactos JSON generados durante el caso de prueba.
* Ejemplo de follow-up automático.
* Ejemplo de informe final generado.
* Plantillas Word utilizadas para el informe.
* Configuración Docker para levantar el entorno local.
* Documentación técnica de las limitaciones encontradas.

El flujo completo permite procesar un cuestionario de proveedor, consultar un catálogo de riesgos mediante RAG, evaluar las respuestas con un agente, generar aclaraciones cuando la información es insuficiente y producir un informe final en formato Word.

---

## 2. Arquitectura del entorno

La prueba de concepto se ejecuta en un entorno local basado en Docker Compose. El entorno está formado por tres servicios principales:

| Servicio | Contenedor   | Función                                                                                 |
| -------- | ------------ | --------------------------------------------------------------------------------------- |
| n8n      | `tfm_n8n`    | Orquestación del flujo, ejecución de nodos, agentes, scripts y generación de artefactos |
| Qdrant   | `tfm_qdrant` | Base de datos vectorial utilizada para la recuperación RAG                              |
| Ollama   | `tfm_ollama` | Generación local de embeddings mediante el modelo `nomic-embed-text`                    |

La configuración principal se encuentra en:

```text
docker-compose.yml
```

El servicio de n8n monta la carpeta local `workspace/` dentro del contenedor como `/data`. Esto permite que los scripts, casos, plantillas y resultados sean accesibles desde los nodos de n8n.

---

## 3. Estructura del repositorio

La estructura principal del repositorio es la siguiente:

```text
risk-analysis-agent/
├── Workflows/
│   ├── Chat Agente AARR.json
│   ├── GROQ copy.json
│   ├── Indexar documentos RAG en Qdrant.json
│   ├── Intake correos AARR GROQ copy.json
│   └── RAG FINAL.json
│
├── n8n_custom/
│   └── Dockerfile
│
├── qdrant_storage/
│
├── workspace/
│   ├── cases/
│   │   ├── Codigo_del_proyecto/
│   │   └── Codigo_del_proyecto_NOFOLLOWUP_OK/
│   │
│   ├── chat_state/
│   ├── rag_docs/
│   ├── rag_manual/
│   ├── scripts/
│   └── templates/
│
├── docker-compose.yml
└── .gitignore
```

Además, el entorno espera la existencia de algunas carpetas locales que pueden no estar incluidas en el repositorio por tamaño o sensibilidad:

```text
n8n_data/
ollama_data/
```

Estas carpetas se explican más adelante.

---

## 4. Carpetas principales del sistema

### 4.1. `Workflows/`

Contiene los workflows exportados desde n8n en formato JSON. Estos ficheros permiten importar el flujo en otra instancia de n8n y revisar la configuración de nodos, conexiones, expresiones, prompts y lógica del proceso.

Los workflows incluidos son:

```text
Workflows/
├── Chat Agente AARR.json
├── GROQ copy.json
├── Indexar documentos RAG en Qdrant.json
├── Intake correos AARR GROQ copy.json
└── RAG FINAL.json
```

Los workflows más relevantes para la prueba de concepto final eson:

```text
Workflows/Indexar documentos RAG en Qdrant.json
Workflows/GROQ copy.json
```

Estos workflows representan los principales flujos del análisis, integrando ingesta, recuperación RAG, scoring, follow-up, consolidación de resultados y generación del informe.

Los workflows:

```text
Workflows/Indexar documentos RAG en Qdrant.json
Workflows/RAG FINAL.json
```

se utilizan para preparar la base vectorial en Qdrant.

El workflow `Chat Agente AARR.json` corresponde al chat que te pregunta el TIER del caso antes de comenzar el análisis del mismo.

---

### 4.2. `n8n_custom/`

Contiene la configuración personalizada de la imagen de n8n.

```text
n8n_custom/
└── Dockerfile
```

Esta imagen permite adaptar el contenedor de n8n al entorno del proyecto, instalando dependencias necesarias para ejecutar scripts Python, manipular documentos y acceder a los ficheros montados en `/data`.

---

### 4.3. `workspace/`

Es la carpeta principal de trabajo del prototipo. En el contenedor de n8n se monta como:

```text
/data
```

Por tanto, cualquier ruta utilizada dentro de los nodos de n8n suele partir de `/data`.

Su estructura principal es:

```text
workspace/
├── cases/
├── chat_state/
├── rag_docs/
├── rag_manual/
├── scripts/
└── templates/
```

---

### 4.4. `workspace/cases/`

Contiene los casos analizados por el sistema. Cada caso dispone de una carpeta propia. En el caso de prueba principal se utilizó:

```text
workspace/cases/Codigo_del_proyecto/
```

La estructura interna de un caso es:

```text
workspace/cases/Codigo_del_proyecto/
├── 00_state/
├── 01_originales/
├── 02_json/
├── 03_analysis/
├── 04_followups/
├── 04_report/
└── 05_report/
```

Esta estructura permite separar las fuentes originales, los artefactos intermedios y las salidas finales.

---

## 5. Estructura de carpetas de un caso

Cada caso analizado genera una estructura similar a la siguiente:

```text
Codigo_del_proyecto/
├── 00_state/
│   └── tier_started.lock
│
├── 01_originales/
│   ├── cuestionario.xlsx
│   └── triaje_email.txt
│
├── 02_json/
│   ├── case_meta.json
│   ├── questionnaire.json
│   ├── triaje.json
│   ├── parse_summary_questionnaire.json
│   └── parse_summary_triaje.json
│
├── 03_analysis/
│   ├── ai_batches/
│   ├── scoring_batches/
│   ├── followup_reanalysis_batches/
│   ├── followup_reanalysis_results/
│   ├── scoring_preliminary.json
│   ├── followups_pending.json
│   ├── scoring_final.json
│   ├── risk_analysis.json
│   ├── report_data.json
│   └── case_status.json
│
├── 04_followups/
│   └── round_001/
│       ├── followup_email_sent.txt
│       ├── provider_response_raw.txt
│       ├── provider_response_clean.txt
│       └── provider_response_parsed.json
│
├── 04_report/
│   ├── assets/
│   ├── report_pack.json
│   ├── report_narrative_ai.json
│   ├── report_render_data.json
│   └── google_drive_upload_pack.json
│
└── 05_report/
    ├── Informe_Ciberseguridad_Codigo_del_proyecto.docx
    ├── report_generation_status.json
    ├── _sanitized_template_Codigo_del_proyecto.docx
    └── evidencias/
```

---

## 6. Artefactos JSON generados

Durante el flujo se generan múltiples artefactos JSON. Estos ficheros permiten mantener trazabilidad entre la información de entrada, las decisiones del agente, los follow-ups, los riesgos finales y el informe generado.

Los principales artefactos son:

| Fichero                         | Ubicación      | Finalidad                                                                     |
| ------------------------------- | -------------- | ----------------------------------------------------------------------------- |
| `questionnaire.json`            | `02_json/`     | Contiene las preguntas, respuestas y explicaciones extraídas del cuestionario |
| `triaje.json`                   | `02_json/`     | Contiene el contexto del caso: criticidad, datos tratados, GDPR, etc.   |
| `case_meta.json`                | `02_json/`     | Metadatos generales y TIER del caso                                                  |
| `scoring_preliminary.json`      | `03_analysis/` | Resultado preliminar del análisis de preguntas                                |
| `followups_pending.json`        | `03_analysis/` | Preguntas que requieren aclaración del proveedor                              |
| `scoring_final.json`            | `03_analysis/` | Resultado final consolidado tras follow-up                                    |
| `risk_analysis.json`            | `03_analysis/` | Riesgos agrupados, severidades y mitigaciones                                 |
| `report_pack.json`              | `04_report/`   | Paquete de datos estructurados para el informe                                |
| `report_narrative_ai.json`      | `04_report/`   | Narrativa generada para el informe                                            |
| `report_render_data.json`       | `04_report/`   | Entrada final para renderizar el informe DOCX                                 |
| `google_drive_upload_pack.json` | `04_report/`   | Paquete de subida del informe y evidencias a Google Drive                     |
| `report_generation_status.json` | `05_report/`   | Estado de generación del informe final                                        |

---

## 7. Reglas de scoring y normalización

El sistema utiliza un modelo de scoring basado en tres valores:

| Score | Significado                                          | Acción                        |
| ----- | ---------------------------------------------------- | ----------------------------- |
| `1`   | Incumplimiento claro                                 | Genera hallazgo y riesgo      |
| `2`   | Información insuficiente, parcial o ambigua          | Genera follow-up al proveedor |
| `3`   | Cumplimiento claro, no aplica o pregunta informativa | No genera hallazgo            |

Durante el análisis inicial, el agente puede asignar `score = 2` cuando la respuesta del proveedor no es suficiente para cerrar la evaluación.

Tras el follow-up, el sistema evita mantener preguntas en estado pendiente. En esa fase, una pregunta debe resolverse como:

| Resultado tras follow-up | Significado                                                       |
| ------------------------ | ----------------------------------------------------------------- |
| `final_score = 3`        | La aclaración resuelve la duda                                    |
| `final_score = 1`        | La aclaración sigue siendo insuficiente o confirma incumplimiento |

Además, el sistema utiliza normalizadores deterministas para evitar que el modelo de lenguaje tome decisiones críticas sin control. Estos normalizadores se encargan de consolidar:

* Identificador de riesgo.
* Nombre del riesgo.
* Controles afectados.
* Severidad.
* Generación de hallazgos.
* Necesidad de follow-up.
* Plazos de mitigación.
* Agrupación de riesgos.
* Revisión humana.

Las mitigaciones no son inventadas por el modelo. Proceden del catálogo de riesgos y se aplican de forma controlada.

---

## 8. Catálogo de riesgos y RAG

El catálogo de riesgos se encuentra en:

```text
workspace/rag_manual/
├── catalog_questions.json
└── catalog_questions_v2.json
```

Este catálogo contiene la relación entre preguntas, controles, riesgos, severidades, mitigaciones y reglas específicas.

Para que el agente pueda consultar el catálogo durante el análisis, las preguntas se indexan en Qdrant mediante embeddings. El script principal de indexación es:

```text
workspace/scripts/index_manual_qdrant.py
```

El modelo de embeddings utilizado es:

```text
nomic-embed-text
```

Este modelo se ejecuta a través de Ollama.

El comando de preparación del modelo es:

```bash
docker exec -it tfm_ollama ollama pull nomic-embed-text
```

La colección final utilizada para el RAG del agente es:

```text
tfm_rag_manual_agent
```

Un ejemplo de comando de indexación sería:

```bash
docker exec -it tfm_n8n python /data/scripts/index_manual_qdrant.py \
  --catalog-json /data/rag_manual/catalog_questions_v2.json \
  --collection tfm_rag_manual_agent \
  --qdrant-url http://qdrant:6333 \
  --ollama-url http://ollama:11434 \
  --embedding-model nomic-embed-text \
  --recreate
```

Cada ficha del catálogo se indexa con un identificador exacto de pregunta, por ejemplo:

```text
EXACT_QUESTION_ID_Q_020
```

Esto permite que el agente consulte la ficha concreta de cada pregunta y reduce el riesgo de mezclar información entre controles distintos.

---

## 9. Prompts principales utilizados

Los prompts principales se encuentran integrados en los workflows exportados de n8n, especialmente en los nodos de agente y de generación narrativa.

Los prompts cubren principalmente:

* Scoring inicial de preguntas.
* Uso obligatorio de RAG.
* Validación de `question_id`.
* Generación de follow-ups.
* Reanálisis tras respuesta del proveedor.
* Generación narrativa para el informe.
* Restricciones para evitar que el modelo invente riesgos o mitigaciones.

La lógica principal del prompt obliga al agente a consultar la herramienta RAG antes de evaluar cada pregunta. El agente debe utilizar la ficha del catálogo correspondiente y no debe copiar riesgos, mitigaciones o controles de otras preguntas.


---

## 10. Configuración de nodos n8n

La configuración de nodos de n8n se documenta mediante los workflows exportados en JSON:

```text
Workflows/
```

Estos ficheros pueden importarse en cualquier instancia de n8n para visualizar el flujo completo, sus nodos, conexiones, expresiones y configuración.

Para importar un workflow:

1. Abrir n8n.
2. Crear un workflow nuevo.
3. Abrir el menú de los tres puntos.
4. Seleccionar `Import from File`.
5. Elegir el fichero `.json` correspondiente.
6. Configurar credenciales, URLs y variables locales.


---

## 11. Scripts principales

Los scripts principales se encuentran en:

```text
workspace/scripts/
```

Los más relevantes son:

| Script                             | Función                                                                          |
| ---------------------------------- | -------------------------------------------------------------------------------- |
| `parse_ingesta_n8n.py`             | Procesa el cuestionario y el triaje, generando artefactos JSON normalizados      |
| `index_manual_qdrant.py`           | Indexa el catálogo manual de riesgos en Qdrant                                   |
| `build_question_rag_context.py`    | Recupera contexto RAG por pregunta y genera batches para el agente               |
| `build_question_rag_context_v2.py` | Versión mejorada del generador de contexto y batches                             |
| `rag_qdrant.py`                    | Utilidades de indexación, consulta, health-check y gestión de colecciones Qdrant |
| `generate_risk_report_docx.py`     | Genera el informe final DOCX a partir de `report_render_data.json`               |

El script de generación del informe lee:

```text
/data/cases/<case_id>/04_report/report_render_data.json
```

y genera:

```text
/data/cases/<case_id>/05_report/Informe_Ciberseguridad_<case_id>.docx
```

---

## 12. Ejemplo de informe generado

El repositorio incluye un ejemplo de informe generado para el caso anonimizado:

```text
workspace/cases/Codigo_del_proyecto/05_report/Informe_Ciberseguridad_Codigo_del_proyecto.docx
```

Este informe se genera automáticamente a partir de:

```text
workspace/cases/Codigo_del_proyecto/04_report/report_render_data.json
```

El proceso de generación utiliza la plantilla:

```text
workspace/templates/NuevaPlantillaInforme_MARKERS.docx
```

Durante la generación se realizan tareas como:

* Sustitución de marcadores.
* Creación de tablas.
* Inserción de resultados de análisis.
* Aplicación de colores por severidad.
* Incorporación de deuda tecnológica.
* Generación de documento DOCX final.

Esta parte cubre el contenido previsto en el **Anexo F. Ejemplo de informe generado**.

---

## 13. Ejemplo de follow-up

El caso de prueba incluye una ronda de follow-up ubicada en:

```text
workspace/cases/Codigo_del_proyecto/04_followups/round_001/
```

Los ficheros principales son:

| Fichero                         | Descripción                                     |
| ------------------------------- | ----------------------------------------------- |
| `followup_email_sent.txt`       | Correo de aclaración generado para el proveedor |
| `provider_response_raw.txt`     | Respuesta original recibida                     |
| `provider_response_clean.txt`   | Respuesta limpiada y normalizada                |
| `provider_response_parsed.json` | Respuesta parseada en formato JSON              |

Después de recibir la respuesta, el sistema genera batches de reanálisis:

```text
workspace/cases/Codigo_del_proyecto/03_analysis/followup_reanalysis_batches/
```

y resultados de reanálisis:

```text
workspace/cases/Codigo_del_proyecto/03_analysis/followup_reanalysis_results/
```

---

## 14. Resultados del caso de prueba

El caso principal del repositorio es:

```text
Codigo_del_proyecto
```

Se trata de un caso anonimizado utilizado para validar el flujo completo.

Los resultados principales fueron:

| Indicador                             | Resultado              |
| ------------------------------------- | ---------------------- |
| Preguntas analizadas                  | 56                     |
| Batches iniciales                     | 53                     |
| Preguntas con follow-up               | 18                     |
| Preguntas resueltas tras follow-up    | 14                     |
| Preguntas no resueltas tras follow-up | 4                      |
| Hallazgos finales                     | 22                     |
| Riesgos agrupados                     | 7                      |
| Riesgos altos                         | 3                      |
| Riesgos medios                        | 3                      |
| Riesgos bajos                         | 1                      |
| Riesgo global                         | Alto                   |
| Informe final                         | Generado correctamente |

El estado final del caso puede consultarse en:

```text
workspace/cases/Codigo_del_proyecto/03_analysis/case_status.json
```

El estado de generación del informe puede consultarse en:

```text
workspace/cases/Codigo_del_proyecto/05_report/report_generation_status.json
```

---

## 15. Limitaciones de proveedores LLM y cuotas

Durante el desarrollo se identificaron varias limitaciones asociadas al uso de proveedores de modelos de lenguaje.

### 15.1. Groq

Se detectaron limitaciones relacionadas con:

* Tamaño de contexto.
* Límites de tokens.
* Cuotas de uso.
* Dificultad para procesar prompts muy largos.
* Necesidad de reducir el tamaño de los batches.

Por este motivo, se ajustó el diseño del flujo para dividir las preguntas en batches pequeños y reducir la dependencia de prompts excesivamente largos.

### 15.2. Google AI Studio / Gemini

El uso de modelos externos permite mejorar el rendimiento del análisis, pero también introduce dependencia de:

* Cuotas de uso.
* Límites de peticiones.
* Disponibilidad del proveedor.
* Posibles cambios en modelos y APIs.
* Coste o restricciones del servicio.

### 15.3. Ollama local

Ollama se utilizó principalmente para generar embeddings locales con:

```text
nomic-embed-text
```

Sin embargo, el uso de modelos locales sin GPU puede ser lento para tareas de razonamiento complejas. Por este motivo, en la prueba de concepto se utiliza Ollama principalmente como soporte local para embeddings, mientras que el razonamiento principal puede apoyarse en modelos externos.


---

## 16. Carpetas no incluidas o generadas localmente


### 16.1. `n8n_data/`

Esta carpeta corresponde al volumen persistente de n8n:

```text
./n8n_data:/home/node/.n8n
```

Contiene datos internos de n8n, como configuración local, base de datos, credenciales y ejecuciones. Por seguridad, no debe compartirse en GitHub.

Si no existe, se genera automáticamente al levantar el entorno con Docker Compose.

### 16.2. `ollama_data/`

Esta carpeta corresponde al volumen persistente de Ollama:

```text
./ollama_data:/root/.ollama
```

Aquí se almacenan los modelos descargados por Ollama. En este proyecto es especialmente relevante porque debe contener el modelo de embeddings:

```text
nomic-embed-text
```

Si la carpeta no existe, se genera al levantar Ollama. Después debe descargarse el modelo con:

```bash
docker exec -it tfm_ollama ollama pull nomic-embed-text
```

### 16.3. `qdrant_storage/`

Esta carpeta corresponde al volumen persistente de Qdrant:

```text
./qdrant_storage:/qdrant/storage
```

Contiene las colecciones vectoriales generadas durante la indexación. Puede regenerarse ejecutando el script de indexación del catálogo, por lo que no es imprescindible para comprender el proyecto.

---

## 17. Puesta en marcha básica

### 17.1. Levantar los contenedores

Desde la raíz del repositorio:

```bash
docker compose up -d --build
```

Esto levanta:

* n8n en `http://localhost:5678`
* Qdrant en `http://localhost:6333`
* Ollama en `http://localhost:11434`

### 17.2. Descargar el modelo de embeddings

```bash
docker exec -it tfm_ollama ollama pull nomic-embed-text
```

### 17.3. Indexar el catálogo en Qdrant

```bash
docker exec -it tfm_n8n python /data/scripts/index_manual_qdrant.py \
  --catalog-json /data/rag_manual/catalog_questions_v2.json \
  --collection tfm_rag_manual_agent \
  --qdrant-url http://qdrant:6333 \
  --ollama-url http://ollama:11434 \
  --embedding-model nomic-embed-text \
  --recreate
```

### 17.4. Importar los workflows en n8n

Desde la interfaz de n8n:

1. Abrir `http://localhost:5678`.
2. Crear o abrir un workflow.
3. Seleccionar el menú de los tres puntos.
4. Elegir `Import from File`.
5. Importar el workflow deseado desde la carpeta `Workflows/`.
6. Configurar credenciales y variables necesarias.

---

## 18. Estado del prototipo

El prototipo demuestra la viabilidad funcional de automatizar gran parte del análisis de riesgos de proveedores SaaS.

El sistema permite:

* Ingerir información de cuestionarios y triajes.
* Normalizar datos a JSON.
* Consultar un catálogo de riesgos mediante RAG.
* Evaluar respuestas con un agente.
* Generar follow-ups automáticos.
* Reanalizar respuestas del proveedor.
* Agrupar riesgos.
* Generar medidas mitigantes.
* Crear un informe final en DOCX.
* Preparar la subida del informe a Google Drive.

El prototipo no elimina la revisión humana. El informe final debe ser revisado por un analista antes de su uso o envío.

---

## 19. Próximas mejoras

Algunas mejoras futuras serían:

* Evaluación contra un dataset de referencia.
* Dashboard de métricas y resultados.
* Gestión multi-caso.
* Integración productiva con SharePoint o Microsoft Graph.
* Mejora de control de permisos y auditoría.
* Extracción de prompts a ficheros independientes.
* Uso de plantillas documentales más robustas.
* Mejora del catálogo de riesgos.
* Incorporación de un agente previo para recoger instrucciones especiales del analista.
* Seguimiento de deuda tecnológica y estado de mitigación.

---

## 20. Aviso

Este repositorio forma parte de una prueba de concepto académica. No debe considerarse una solución productiva sin revisar aspectos de seguridad, control de errores, gestión de credenciales, auditoría, privacidad, escalabilidad y gobierno del modelo.
