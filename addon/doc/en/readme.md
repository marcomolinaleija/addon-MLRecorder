# MLRecorder para NVDA

MLRecorder agrega comandos globales de NVDA para grabar:

- Audio del proceso enfocado.
- Audio del sistema (lo que se escucha en los altavoces).
- Micrófono.
- Mezcla del proceso enfocado + micrófono.
- Mezcla del sistema + micrófono.

## Gestos de entrada

Los comandos de MLRecorder **no tienen atajos asignados por defecto** para evitar conflictos con otras aplicaciones. Asígnalos desde:

**NVDA → Preferencias → Gestos de entrada → MLRecorder**

Los comandos disponibles son:

- Alternar grabación del proceso enfocado.
- Alternar grabación del sistema (escritorio).
- Alternar grabación de micrófono.
- Alternar grabación mixta (proceso enfocado + micrófono).
- Alternar grabación mixta (sistema + micrófono).
- Pausar o reanudar la grabación activa.
- Detener la grabación activa.
- Detener todas las grabaciones.
- Reportar estado de grabaciones.
- Abrir carpeta de grabaciones.
- Activar capa de comandos.

## Capa de comandos

La capa de comandos permite acceder a las funciones con teclas cortas. Actívala con el gesto que hayas asignado (se oirá un pitido) y luego pulsa:

- **p**: Alternar grabación de proceso.
- **Alt+p**: Alternar grabación mixta (Proceso + Mic).
- **s**: Alternar grabación de sistema.
- **Alt+s**: Alternar grabación mixta (Sistema + Mic).
- **m**: Alternar grabación de micrófono.
- **d**: Detener grabación activa.
- **Alt+d**: Detener todas las grabaciones.
- **i**: Reportar estado.
- **o**: Abrir carpeta de grabaciones.

## Pausa y reanudación

Durante una grabación activa, el comando de pausa alterna entre pausado y grabando sin cerrar el archivo de salida. El comando de estado informa si la sesión está activa o pausada.

## Estado de grabaciones

El comando de estado anuncia, para cada sesión activa:

- Tipo de sesión (proceso, sistema, micrófono o mezcla) y nombre.
- Estado: activo o pausado.
- Tiempo neto grabado y megabytes escritos.

Ejemplo: *"Proceso activo: Spotify, 1m 23s, 8.4 MB"*

## Notificación de sesión terminada

Si el proceso grabado se cierra inesperadamente o ocurre un error de dispositivo, NVDA anuncia el evento por voz y limpia la sesión automáticamente.

## Configuración

Abre **NVDA → Preferencias → Opciones → MLRecorder** para ajustar:

- Formato de salida (WAV, MP3, FLAC, Opus).
- Saltar silencios.
- Volumen de proceso (%).
- Dispositivo de micrófono.

## Carpeta de grabaciones

Las grabaciones se guardan en: `%userprofile%\Documents\NVDA_MLRecorder`
