# 📅 Generador de Horarios Semanales / Anuales

Aplicación web para generar horarios de actividades estudiantiles de forma aleatoria,
respetando restricciones por sexo y evitando repetición dentro de la semana.

## 🚀 Publicar en Streamlit Cloud (gratis, 5 minutos)

1. Sube este proyecto a un repositorio de GitHub
2. Ve a **[share.streamlit.io](https://share.streamlit.io)** e inicia sesión con GitHub
3. Haz clic en **"New app"**
4. Selecciona tu repositorio y pon `app_web.py` como archivo principal
5. Haz clic en **"Deploy"** — en 1-2 minutos tendrás tu URL pública

## 💻 Correr localmente

```bash
pip install -r requirements.txt
streamlit run app_web.py
```

## 📁 Archivos del proyecto

| Archivo | Descripción |
|---|---|
| `app_web.py` | App web (Streamlit) — para publicar en internet |
| `generador_horario.py` | App de escritorio (Tkinter) — para usar en Mac/PC |
| `Ejemplo_Actividades_Alumnos.xlsx` | Excel de ejemplo para probar |
| `requirements.txt` | Dependencias Python |
| `crear_ejecutable_windows.bat` | Crea un .exe para Windows |

## 📋 Formato del Excel de entrada

**Hoja "Actividades":**
| Nombre actividad | Sexo |
|---|---|
| Barrer el patio | Indiferente |
| Limpiar baño hombres | Hombre |
| Limpiar baño mujeres | Mujer |

**Hoja "Alumnos":**
| Nombre Alumnos | Sexo |
|---|---|
| Juan Pérez | Hombre |
| María García | Mujer |

## ⚙️ Reglas de asignación

- Actividades **Hombre** → solo alumnos masculinos
- Actividades **Mujer** → solo alumnos femeninos  
- Actividades **Indiferente** → cualquier alumno
- Cada alumno: máximo una actividad por día
- Sin repetición actividad-alumno dentro de la semana (en lo posible)
- Semana: lunes a sábado
