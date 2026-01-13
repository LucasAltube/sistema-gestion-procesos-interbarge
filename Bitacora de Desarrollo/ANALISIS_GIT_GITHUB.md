# Análisis Git / GitHub (Interbarge)

## Estado actual
- El repositorio Git existe localmente en `Sistema de Gestion por Procesos/`.
- No hay remoto configurado (no hay GitHub conectado desde aquí).
- En el working tree hay dos tipos de contenido:
  1) Material “de sistema de gestión” (procedimientos, registros, planillas) en carpetas como `Comercial/`, `RRHH/`, etc.
  2) Archivos “de desarrollo/gestión del trabajo” (bitácora, extracciones, propuestas) en `Bitacora de Desarrollo/`.

## Recomendación de versionado
- Versionar (recomendado):
  - `Bitacora de Desarrollo/` (bitácora + análisis + extractos).
  - Archivos de trabajo propios (como la propuesta `FC-XX-01-01 Indicadores Procesos v4.xlsx`) si el objetivo es iterar y auditar cambios.
- Evaluar antes de versionar:
  - `Comercial/`, `RRHH/`, `SCH/`, etc. (pueden incluir información sensible, datos personales, o documentos “controlados” fuera de Git).

## Estrategia de commits
- Evitar commits gigantes mezclando documentos fuente + propuestas.
- Sugerencia de separación:
  - Commit 1: Estructura + bitácora (`Bitacora de Desarrollo/*`, `.gitignore`).
  - Commit 2: Propuesta de indicadores (`FC-XX-01-01 Indicadores Procesos v4.xlsx` + extractos si aplica).
  - Commit 3+: Ajustes de indicadores por iteración (cambios en v4).

## Convención de mensajes (simple)
- `chore:` para estructura/administrativo.
- `docs:` para bitácoras y documentación interna.
- `feat:` para nuevas propuestas/plantillas.
- `fix:` para correcciones puntuales.

## GitHub (cuando decidas usarlo)
- Pasos típicos (manuales):
  - `git remote add origin <url>`
  - `git push -u origin main`
- Recomendación: repositorio privado si se suben documentos internos.

## Riesgos y cuidado
- Excel/Word cambian binariamente: Git no muestra diffs útiles.
  - Mitigación: mantener extractos CSV/Markdown (como `indicadores_v3_extraido.csv`) para trazabilidad.
- Datos personales (RRHH) y comerciales: revisar antes de subir a GitHub.
