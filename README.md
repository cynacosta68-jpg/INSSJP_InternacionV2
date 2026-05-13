# 🏥 Procesador PAMI — Internación

Sistema de procesamiento, valorización y generación de templates para prácticas de internación INSSJP (PAMI).

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
streamlit run app_pami.py
```

Se abre en `http://localhost:8501`.

## Archivos requeridos

### 1. Archivo crudo PAMI (factura ARCA)
Excel con hoja que contenga: `PRESTADOR`, `BATE`, `PROFESIONAL ACTUANTE`, `C_PRACTICA`, `PRACTICA`, `MONTO`, `APELLIDO Y NOMBRE`, `NRO. BENEFICIO/GP`, `F. DE PRACTICA`, `N. DE OP`, etc.

### 2. Base de Datos (parámetros)
Excel con tres solapas:
- **DIRECTA** — atribución directa por código o profesional
- **NORMAL** — atribución normal por BATE + Profesional + Ayudante + Anatomía
- **VALORES** — valores para completar prácticas con MONTO = 0

## Flujo

1. **Cargar** archivos en la barra lateral
2. **Editar** prácticas: marcar Ayudante, Cuenta Ayudante, Anatomía, Matrícula con filtros por BATE/Modalidad/Monto
3. **Valorizar** según reglas DIRECTA y NORMAL con apertura por cuenta
4. **Exportar** Excel valorizado + Templates evweb (.zip) + link a evweb

## Reglas de atribución

### DIRECTA
- Códigos: 800001, 801001–801005, 801008, 801009, 816005
- Profesionales: RIOS PART, DE SABATO (cualquier variante)
- RIOS PART / RUFFINI → `%HP` × MONTO, Cuenta P
- DE SABATO → `%HC` × MONTO, Cuenta HC
- Resto por código → `%HC` × MONTO, Cuenta HC

### NORMAL
Busca BATE + PROFESIONAL + Ayudante + Anatomía → apertura en sub-filas:
- `%HC` → Cuenta HC
- `%HP` → Cuenta P
- `%Ay` → Cuenta Ayudante (usuario)
- `%Ap` → Cuenta AP
- `%G` → Cuenta G

## Templates evweb
Genera archivos Excel con formato `ImportacionesEvweb` (1500 filas máx. por template) listos para importar en [evweb](https://cmsc.evweb.com.ar).
