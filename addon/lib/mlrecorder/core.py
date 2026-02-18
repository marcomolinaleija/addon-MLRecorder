import ctypes
import os
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional


FORMAT_WAV = 0
FORMAT_MP3 = 1
FORMAT_OPUS = 2
FORMAT_FLAC = 3

RESULT_OK = 0


@dataclass
class ProcessInfo:
    process_id: int
    process_name: str
    window_title: str
    has_active_audio: bool


@dataclass
class InputDeviceInfo:
    device_id: str
    friendly_name: str
    is_default: bool


class _CProcessInfo(ctypes.Structure):
    _fields_ = [
        ("process_id", ctypes.c_uint32),
        ("process_name_utf8", ctypes.c_char_p),
        ("window_title_utf8", ctypes.c_char_p),
        ("has_active_audio", ctypes.c_int),
    ]


_ProcessCallback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.POINTER(_CProcessInfo), ctypes.c_void_p)


class _CInputDeviceInfo(ctypes.Structure):
    _fields_ = [
        ("device_id_utf8", ctypes.c_char_p),
        ("friendly_name_utf8", ctypes.c_char_p),
        ("is_default", ctypes.c_int),
    ]


_InputDeviceCallback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.POINTER(_CInputDeviceInfo), ctypes.c_void_p)


class MLRecorderError(RuntimeError):
    pass


class MLRecorder:
    def __init__(self, dll_path: Optional[str] = None) -> None:
        if os.name != "nt":
            raise OSError("mlrecorder is only supported on Windows")

        self._dll_path = self._resolve_dll_path(dll_path)
        self._dll_dir_handles = []
        self._prepare_dll_search_path(self._dll_path)
        self._dll = ctypes.WinDLL(str(self._dll_path))
        self._configure_signatures()

    @property
    def dll_path(self) -> Path:
        return self._dll_path

    def _resolve_dll_path(self, dll_path: Optional[str]) -> Path:
        if dll_path:
            p = Path(dll_path).expanduser().resolve()
            if not p.exists():
                raise FileNotFoundError(f"DLL not found: {p}")
            return p

        env_path = os.environ.get("MLRECORDER_DLL")
        if env_path:
            p = Path(env_path).expanduser().resolve()
            if p.exists():
                return p

        here = Path(__file__).resolve().parent
        candidates = [
            here / "bin" / "mlrecorder_core.dll",
            here.parent / "mlrecorder_core.dll",
            here.parent.parent / "build-release" / "bin" / "mlrecorder_core.dll",
            here.parent.parent / "build" / "bin" / "mlrecorder_core.dll",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate

        raise FileNotFoundError(
            "Could not locate mlrecorder_core.dll. "
            "Pass dll_path=... or set MLRECORDER_DLL."
        )

    def _prepare_dll_search_path(self, dll_path: Path) -> None:
        if not hasattr(os, "add_dll_directory"):
            return

        search_dirs = []
        if dll_path.parent.exists():
            search_dirs.append(dll_path.parent)

        package_bin = Path(__file__).resolve().parent / "bin"
        if package_bin.exists():
            search_dirs.append(package_bin)

        seen = set()
        for directory in search_dirs:
            resolved = str(directory.resolve())
            if resolved in seen:
                continue
            seen.add(resolved)
            self._dll_dir_handles.append(os.add_dll_directory(resolved))

    def _configure_signatures(self) -> None:
        self._dll.mlr_initialize.argtypes = []
        self._dll.mlr_initialize.restype = ctypes.c_int

        self._dll.mlr_shutdown.argtypes = []
        self._dll.mlr_shutdown.restype = None

        self._dll.mlr_is_initialized.argtypes = []
        self._dll.mlr_is_initialized.restype = ctypes.c_int

        self._dll.mlr_result_to_string.argtypes = [ctypes.c_int]
        self._dll.mlr_result_to_string.restype = ctypes.c_char_p

        self._dll.mlr_get_last_error.argtypes = []
        self._dll.mlr_get_last_error.restype = ctypes.c_char_p

        self._dll.mlr_list_processes.argtypes = [ctypes.c_int, _ProcessCallback, ctypes.c_void_p]
        self._dll.mlr_list_processes.restype = ctypes.c_int

        self._dll.mlr_list_input_devices.argtypes = [_InputDeviceCallback, ctypes.c_void_p]
        self._dll.mlr_list_input_devices.restype = ctypes.c_int

        self._dll.mlr_start_capture_to_file.argtypes = [
            ctypes.c_uint32,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
            ctypes.c_int,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_int,
        ]
        self._dll.mlr_start_capture_to_file.restype = ctypes.c_int

        self._dll.mlr_start_capture_to_directory.argtypes = [
            ctypes.c_uint32,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
            ctypes.c_int,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_int,
        ]
        self._dll.mlr_start_capture_to_directory.restype = ctypes.c_int

        self._dll.mlr_stop_capture.argtypes = [ctypes.c_uint32]
        self._dll.mlr_stop_capture.restype = ctypes.c_int

        self._dll.mlr_stop_all_captures.argtypes = []
        self._dll.mlr_stop_all_captures.restype = None

        self._dll.mlr_is_capturing.argtypes = [ctypes.c_uint32]
        self._dll.mlr_is_capturing.restype = ctypes.c_int

        self._dll.mlr_set_capture_volume.argtypes = [ctypes.c_uint32, ctypes.c_float]
        self._dll.mlr_set_capture_volume.restype = ctypes.c_int

        self._dll.mlr_get_active_session_count.argtypes = []
        self._dll.mlr_get_active_session_count.restype = ctypes.c_int

        self._dll.mlr_start_microphone_capture_to_file.argtypes = [
            ctypes.c_char_p,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
            ctypes.c_int,
            ctypes.c_int,
        ]
        self._dll.mlr_start_microphone_capture_to_file.restype = ctypes.c_int

        self._dll.mlr_start_microphone_capture_to_directory.argtypes = [
            ctypes.c_char_p,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
            ctypes.c_int,
            ctypes.c_int,
        ]
        self._dll.mlr_start_microphone_capture_to_directory.restype = ctypes.c_int

        self._dll.mlr_stop_microphone_capture.argtypes = [ctypes.c_char_p]
        self._dll.mlr_stop_microphone_capture.restype = ctypes.c_int

        self._dll.mlr_stop_all_microphone_captures.argtypes = []
        self._dll.mlr_stop_all_microphone_captures.restype = None

        self._dll.mlr_enable_mixed_recording_to_file.argtypes = [
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
        ]
        self._dll.mlr_enable_mixed_recording_to_file.restype = ctypes.c_int

        self._dll.mlr_enable_mixed_recording_to_directory.argtypes = [
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_uint32,
            ctypes.c_char_p,
        ]
        self._dll.mlr_enable_mixed_recording_to_directory.restype = ctypes.c_int

        self._dll.mlr_disable_mixed_recording.argtypes = []
        self._dll.mlr_disable_mixed_recording.restype = None

        self._dll.mlr_is_mixed_recording_active.argtypes = []
        self._dll.mlr_is_mixed_recording_active.restype = ctypes.c_int

    def initialize(self) -> None:
        code = self._dll.mlr_initialize()
        self._raise_if_error(code, "mlr_initialize failed")

    def shutdown(self) -> None:
        self._dll.mlr_shutdown()

    def is_initialized(self) -> bool:
        return bool(self._dll.mlr_is_initialized())

    def result_to_string(self, code: int) -> str:
        text = self._dll.mlr_result_to_string(code)
        return text.decode("utf-8", errors="replace") if text else ""

    def last_error(self) -> str:
        text = self._dll.mlr_get_last_error()
        return text.decode("utf-8", errors="replace") if text else ""

    def list_processes(self, only_active_audio: bool = False) -> List[ProcessInfo]:
        processes: List[ProcessInfo] = []

        @_ProcessCallback
        def _callback(info_ptr: ctypes.POINTER(_CProcessInfo), _user_data: ctypes.c_void_p) -> int:
            info = info_ptr.contents
            name = info.process_name_utf8.decode("utf-8", errors="replace") if info.process_name_utf8 else ""
            title = info.window_title_utf8.decode("utf-8", errors="replace") if info.window_title_utf8 else ""
            processes.append(
                ProcessInfo(
                    process_id=int(info.process_id),
                    process_name=name,
                    window_title=title,
                    has_active_audio=bool(info.has_active_audio),
                )
            )
            return 0

        code = self._dll.mlr_list_processes(int(only_active_audio), _callback, None)
        self._raise_if_error(code, "mlr_list_processes failed")
        return processes

    def list_input_devices(self) -> List[InputDeviceInfo]:
        devices: List[InputDeviceInfo] = []

        @_InputDeviceCallback
        def _callback(info_ptr: ctypes.POINTER(_CInputDeviceInfo), _user_data: ctypes.c_void_p) -> int:
            info = info_ptr.contents
            devices.append(
                InputDeviceInfo(
                    device_id=info.device_id_utf8.decode("utf-8", errors="replace")
                    if info.device_id_utf8
                    else "",
                    friendly_name=info.friendly_name_utf8.decode("utf-8", errors="replace")
                    if info.friendly_name_utf8
                    else "",
                    is_default=bool(info.is_default),
                )
            )
            return 0

        code = self._dll.mlr_list_input_devices(_callback, None)
        self._raise_if_error(code, "mlr_list_input_devices failed")
        return devices

    def start_capture_to_file(
        self,
        process_id: int,
        output_file: str,
        fmt: int = FORMAT_WAV,
        bitrate: int = 0,
        skip_silence: bool = False,
        passthrough_device_id: Optional[str] = None,
        monitor_only: bool = False,
        strict_process_isolation: bool = True,
    ) -> None:
        code = self._dll.mlr_start_capture_to_file(
            ctypes.c_uint32(process_id),
            self._encode(output_file),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
            ctypes.c_int(skip_silence),
            self._encode_optional(passthrough_device_id),
            ctypes.c_int(monitor_only),
            ctypes.c_int(strict_process_isolation),
        )
        self._raise_if_error(code, "mlr_start_capture_to_file failed")

    def start_capture_to_directory(
        self,
        process_id: int,
        output_dir: str,
        fmt: int = FORMAT_WAV,
        bitrate: int = 0,
        skip_silence: bool = False,
        passthrough_device_id: Optional[str] = None,
        monitor_only: bool = False,
        strict_process_isolation: bool = True,
    ) -> None:
        code = self._dll.mlr_start_capture_to_directory(
            ctypes.c_uint32(process_id),
            self._encode(output_dir),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
            ctypes.c_int(skip_silence),
            self._encode_optional(passthrough_device_id),
            ctypes.c_int(monitor_only),
            ctypes.c_int(strict_process_isolation),
        )
        self._raise_if_error(code, "mlr_start_capture_to_directory failed")

    def stop_capture(self, process_id: int) -> None:
        code = self._dll.mlr_stop_capture(ctypes.c_uint32(process_id))
        self._raise_if_error(code, "mlr_stop_capture failed")

    def stop_all_captures(self) -> None:
        self._dll.mlr_stop_all_captures()

    def is_capturing(self, process_id: int) -> bool:
        code = self._dll.mlr_is_capturing(ctypes.c_uint32(process_id))
        self._raise_if_error(code, "mlr_is_capturing failed")
        return bool(code)

    def set_capture_volume(self, process_id: int, volume_0_to_1: float) -> None:
        code = self._dll.mlr_set_capture_volume(ctypes.c_uint32(process_id), ctypes.c_float(volume_0_to_1))
        self._raise_if_error(code, "mlr_set_capture_volume failed")

    def active_session_count(self) -> int:
        code = self._dll.mlr_get_active_session_count()
        self._raise_if_error(code, "mlr_get_active_session_count failed")
        return int(code)

    def start_microphone_capture_to_file(
        self,
        output_file: str,
        input_device_id: Optional[str] = None,
        fmt: int = FORMAT_WAV,
        bitrate: int = 0,
        skip_silence: bool = False,
        monitor_only: bool = False,
    ) -> None:
        code = self._dll.mlr_start_microphone_capture_to_file(
            self._encode_optional(input_device_id),
            self._encode(output_file),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
            ctypes.c_int(skip_silence),
            ctypes.c_int(monitor_only),
        )
        self._raise_if_error(code, "mlr_start_microphone_capture_to_file failed")

    def start_microphone_capture_to_directory(
        self,
        output_dir: str,
        input_device_id: Optional[str] = None,
        fmt: int = FORMAT_WAV,
        bitrate: int = 0,
        skip_silence: bool = False,
        monitor_only: bool = False,
    ) -> None:
        code = self._dll.mlr_start_microphone_capture_to_directory(
            self._encode_optional(input_device_id),
            self._encode(output_dir),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
            ctypes.c_int(skip_silence),
            ctypes.c_int(monitor_only),
        )
        self._raise_if_error(code, "mlr_start_microphone_capture_to_directory failed")

    def stop_microphone_capture(self, input_device_id: Optional[str] = None) -> None:
        code = self._dll.mlr_stop_microphone_capture(self._encode_optional(input_device_id))
        self._raise_if_error(code, "mlr_stop_microphone_capture failed")

    def stop_all_microphone_captures(self) -> None:
        self._dll.mlr_stop_all_microphone_captures()

    def enable_mixed_recording_to_file(self, output_file: str, fmt: int = FORMAT_WAV, bitrate: int = 0) -> None:
        code = self._dll.mlr_enable_mixed_recording_to_file(
            self._encode(output_file),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
        )
        self._raise_if_error(code, "mlr_enable_mixed_recording_to_file failed")

    def enable_mixed_recording_to_directory(
        self,
        output_dir: str,
        fmt: int = FORMAT_WAV,
        bitrate: int = 0,
        base_name: str = "Mixed",
    ) -> None:
        code = self._dll.mlr_enable_mixed_recording_to_directory(
            self._encode(output_dir),
            ctypes.c_int(fmt),
            ctypes.c_uint32(bitrate),
            self._encode(base_name),
        )
        self._raise_if_error(code, "mlr_enable_mixed_recording_to_directory failed")

    def disable_mixed_recording(self) -> None:
        self._dll.mlr_disable_mixed_recording()

    def is_mixed_recording_active(self) -> bool:
        code = self._dll.mlr_is_mixed_recording_active()
        self._raise_if_error(code, "mlr_is_mixed_recording_active failed")
        return bool(code)

    def _raise_if_error(self, code: int, context: str) -> None:
        if code >= 0:
            return
        error_text = self.last_error()
        result_text = self.result_to_string(code)
        raise MLRecorderError(f"{context} [{code}: {result_text}] {error_text}".strip())

    @staticmethod
    def _encode(text: str) -> bytes:
        if text is None:
            raise ValueError("text cannot be None")
        return text.encode("utf-8")

    @staticmethod
    def _encode_optional(text: Optional[str]) -> Optional[bytes]:
        if text is None:
            return None
        return text.encode("utf-8")

    def __enter__(self) -> "MLRecorder":
        self.initialize()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.shutdown()
