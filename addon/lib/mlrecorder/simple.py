import atexit
from dataclasses import dataclass
from pathlib import Path
from threading import RLock
from typing import List, Optional, Union

from .core import (
    FORMAT_FLAC,
    FORMAT_MP3,
    FORMAT_OPUS,
    FORMAT_WAV,
    InputDeviceInfo,
    MLRecorder,
    MLRecorderError,
    ProcessInfo,
)


FormatInput = Union[int, str]

_runtime_lock = RLock()
_runtime: Optional[MLRecorder] = None


def _normalize_format(fmt: FormatInput) -> int:
    if isinstance(fmt, int):
        return fmt

    value = fmt.strip().lower()
    if value == "wav":
        return FORMAT_WAV
    if value == "mp3":
        return FORMAT_MP3
    if value == "opus":
        return FORMAT_OPUS
    if value == "flac":
        return FORMAT_FLAC
    raise ValueError(f"Formato no soportado: {fmt}")


def initialize(dll_path: Optional[str] = None) -> MLRecorder:
    global _runtime
    with _runtime_lock:
        if _runtime is None:
            _runtime = MLRecorder(dll_path=dll_path)
            _runtime.initialize()
            atexit.register(shutdown)
            return _runtime

        if dll_path is not None:
            requested = Path(dll_path).expanduser().resolve()
            if requested != _runtime.dll_path:
                raise MLRecorderError(
                    f"Runtime ya inicializado con {_runtime.dll_path}, no con {requested}"
                )
        return _runtime


def shutdown() -> None:
    global _runtime
    with _runtime_lock:
        if _runtime is not None:
            _runtime.shutdown()
            _runtime = None


def runtime() -> MLRecorder:
    return initialize()


def list_processes(dll_path: Optional[str] = None) -> List[ProcessInfo]:
    return initialize(dll_path=dll_path).list_processes(only_active_audio=False)


def list_active_processes(dll_path: Optional[str] = None) -> List[ProcessInfo]:
    return initialize(dll_path=dll_path).list_processes(only_active_audio=True)


def list_microphones(dll_path: Optional[str] = None) -> List[InputDeviceInfo]:
    return initialize(dll_path=dll_path).list_input_devices()


def _resolve_microphone_device_id(recorder: MLRecorder, input_device_id: Optional[str]) -> str:
    devices = recorder.list_input_devices()
    if not devices:
        raise MLRecorderError("No hay dispositivos de entrada disponibles")

    if input_device_id:
        for device in devices:
            if device.device_id == input_device_id:
                return device.device_id
        raise MLRecorderError(f"No se encontró el dispositivo de entrada: {input_device_id}")

    for device in devices:
        if device.is_default:
            return device.device_id
    return devices[0].device_id


def start_recorder(
    pid: int,
    output_dir: str = "recordings",
    fmt: FormatInput = "wav",
    bitrate: int = 0,
    skip_silence: bool = False,
    strict_process_isolation: bool = True,
    dll_path: Optional[str] = None,
) -> "RecorderSession":
    recorder = initialize(dll_path=dll_path)
    output_dir_resolved = str(Path(output_dir).expanduser().resolve())
    recorder.start_capture_to_directory(
        process_id=pid,
        output_dir=output_dir_resolved,
        fmt=_normalize_format(fmt),
        bitrate=bitrate,
        skip_silence=skip_silence,
        monitor_only=False,
        strict_process_isolation=strict_process_isolation,
    )
    return RecorderSession(process_id=pid)


def start_recorder_to_file(
    pid: int,
    output_file: str,
    fmt: FormatInput = "wav",
    bitrate: int = 0,
    skip_silence: bool = False,
    strict_process_isolation: bool = True,
    dll_path: Optional[str] = None,
) -> "RecorderSession":
    recorder = initialize(dll_path=dll_path)
    output_file_resolved = str(Path(output_file).expanduser().resolve())
    recorder.start_capture_to_file(
        process_id=pid,
        output_file=output_file_resolved,
        fmt=_normalize_format(fmt),
        bitrate=bitrate,
        skip_silence=skip_silence,
        monitor_only=False,
        strict_process_isolation=strict_process_isolation,
    )
    return RecorderSession(process_id=pid)


def stop_recorder(pid: int) -> None:
    initialize().stop_capture(pid)


def start_microphone_recorder(
    output_dir: str = "recordings",
    fmt: FormatInput = "wav",
    bitrate: int = 0,
    skip_silence: bool = False,
    input_device_id: Optional[str] = None,
    dll_path: Optional[str] = None,
) -> "MicrophoneSession":
    recorder = initialize(dll_path=dll_path)
    resolved_device_id = _resolve_microphone_device_id(recorder, input_device_id)
    output_dir_resolved = str(Path(output_dir).expanduser().resolve())
    recorder.start_microphone_capture_to_directory(
        output_dir=output_dir_resolved,
        input_device_id=resolved_device_id,
        fmt=_normalize_format(fmt),
        bitrate=bitrate,
        skip_silence=skip_silence,
        monitor_only=False,
    )
    return MicrophoneSession(device_id=resolved_device_id)


def stop_microphone(input_device_id: Optional[str] = None) -> None:
    initialize().stop_microphone_capture(input_device_id=input_device_id)


def stop_all_microphones() -> None:
    initialize().stop_all_microphone_captures()


def start_mixed_recorder(
    pid: int,
    output_dir: str = "recordings",
    fmt: FormatInput = "wav",
    bitrate: int = 0,
    skip_silence: bool = False,
    strict_process_isolation: bool = True,
    include_microphone: bool = True,
    input_device_id: Optional[str] = None,
    base_name: str = "Mixed",
    dll_path: Optional[str] = None,
) -> "MixedRecorderSession":
    recorder = initialize(dll_path=dll_path)
    output_dir_resolved = str(Path(output_dir).expanduser().resolve())
    fmt_value = _normalize_format(fmt)

    resolved_device_id: Optional[str] = None
    process_started = False
    microphone_started = False
    mixed_enabled = False

    try:
        recorder.start_capture_to_directory(
            process_id=pid,
            output_dir="",
            fmt=fmt_value,
            bitrate=bitrate,
            skip_silence=skip_silence,
            monitor_only=True,
            strict_process_isolation=strict_process_isolation,
        )
        process_started = True

        if include_microphone:
            resolved_device_id = _resolve_microphone_device_id(recorder, input_device_id)
            recorder.start_microphone_capture_to_directory(
                output_dir="",
                input_device_id=resolved_device_id,
                fmt=fmt_value,
                bitrate=bitrate,
                skip_silence=skip_silence,
                monitor_only=True,
            )
            microphone_started = True

        recorder.enable_mixed_recording_to_directory(
            output_dir=output_dir_resolved,
            fmt=fmt_value,
            bitrate=bitrate,
            base_name=base_name,
        )
        mixed_enabled = True
    except Exception:
        if mixed_enabled:
            recorder.disable_mixed_recording()
        if microphone_started and resolved_device_id:
            try:
                recorder.stop_microphone_capture(resolved_device_id)
            except Exception:
                pass
        if process_started:
            try:
                recorder.stop_capture(pid)
            except Exception:
                pass
        raise

    return MixedRecorderSession(process_id=pid, microphone_device_id=resolved_device_id)


def stop_all() -> None:
    recorder = initialize()
    recorder.disable_mixed_recording()
    recorder.stop_all_microphone_captures()
    recorder.stop_all_captures()


def is_recording(pid: int) -> bool:
    return initialize().is_capturing(pid)


def set_volume(pid: int, volume_0_to_1: float) -> None:
    initialize().set_capture_volume(pid, volume_0_to_1)


def active_session_count() -> int:
    return initialize().active_session_count()


@dataclass
class RecorderSession:
    process_id: int

    def stop(self) -> None:
        stop_recorder(self.process_id)

    def is_recording(self) -> bool:
        return is_recording(self.process_id)

    def set_volume(self, volume_0_to_1: float) -> None:
        set_volume(self.process_id, volume_0_to_1)

    def __enter__(self) -> "RecorderSession":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            self.stop()
        except Exception:
            pass


@dataclass
class MicrophoneSession:
    device_id: str

    def stop(self) -> None:
        stop_microphone(self.device_id)

    def __enter__(self) -> "MicrophoneSession":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            self.stop()
        except Exception:
            pass


@dataclass
class MixedRecorderSession:
    process_id: int
    microphone_device_id: Optional[str] = None

    def stop(self) -> None:
        recorder = initialize()
        recorder.disable_mixed_recording()
        if self.microphone_device_id:
            try:
                recorder.stop_microphone_capture(self.microphone_device_id)
            except Exception:
                pass
        try:
            recorder.stop_capture(self.process_id)
        except Exception:
            pass

    def is_mixing(self) -> bool:
        return initialize().is_mixed_recording_active()

    def __enter__(self) -> "MixedRecorderSession":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.stop()
