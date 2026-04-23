# 🎓 Asistente MESOB - Local y Gratis

Asistente de IA que responde preguntas sobre documentación MESOB sin costos de API.

## 🚀 Instalación Rápida

### 1. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 2. Asegurar que Ollama está corriendo
Abre terminal y ejecuta:
```bash
ollama serve
```

En otra terminal/ventana, descarga el modelo Mistral:
```bash
ollama pull mistral
```

### 3. Procesar los PDFs (una sola vez)
```bash
python process_pdfs.py
```

Esto:
- Extrae texto de todos los PDFs
- Crea embeddings locales
- Almacena en base de datos ChromaDB

### 4. Ejecutar la app
```bash
streamlit run app.py
```

La app abrirá en: http://localhost:8501

## ✨ Características

- 💬 Chat interactivo sobre documentación MESOB
- 🔍 Búsqueda semántica en PDFs
- 🚀 Totalmente local (sin internet necesario)
- 💰 Sin costos (0€)
- ⚡ Respuestas en streaming

## 📁 Estructura

```
├── app.py                 # App principal Streamlit
├── process_pdfs.py        # Procesa PDFs y crea embeddings
├── requirements.txt       # Dependencias
└── mesob_db/             # Base de datos vectorial (se crea automáticamente)
```

## 🤔 Solución de problemas

**"Ollama not found"**
- Instala Ollama desde: https://ollama.ai
- Asegúrate que está corriendo (`ollama serve`)

**"Model not found"**
- Ejecuta: `ollama pull mistral`

**"No documents found"**
- Verifica la ruta de PDFs en `process_pdms.py`
- Asegúrate que están en la carpeta correcta

## 📊 Rendimiento

- Modelo: Mistral (7B parámetros)
- Tiempo de respuesta: 5-15 segundos
- Requisitos: ~8GB RAM, GPU opcional

¡Listo! 🎉
