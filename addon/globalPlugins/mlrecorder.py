# MLRecorder add-on for NVDA.
# Copyright (C) 2026

from __future__ import annotations

import ctypes
import datetime
import os
import struct
import sys
import traceback
from pathlib import Path
from typing import Optional, Tuple

import addonHandler
import api
import config
import globalPluginHandler
import globalVars
import ui
from scriptHandler import script

addonHandler.initTranslation()


def disableInSecureMode(decoratedCls):
	if globalVars.appArgs.secure:
		return globalPluginHandler.GlobalPlugin
	return decoratedCls


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("MLRecorder")

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self._mlr = None
		self._processSession = None
		self._microphoneSession = None
		self._mixedSession = None
		self._lastProcessPid: Optional[int] = None
		self._lastProcessName = ""
		self._lastDiagnosticPath: Optional[Path] = None

	def terminate(self):
		self._stopAllSessions()
		if self._mlr is not None:
			try:
				self._mlr.shutdown()
			except Exception:
				pass
			self._mlr = None
		super().terminate()

	def _speak(self, message: str) -> None:
		ui.message(message)

	def _addonDir(self) -> Path:
		return Path(__file__).resolve().parents[1]

	def _defaultOutputDir(self) -> str:
		userProfile = Path(os.environ.get("USERPROFILE", str(Path.home())))
		candidates = [
			userProfile / "Documents",
			userProfile / "Documentos",
			Path.home() / "Documents",
		]
		base = None
		for candidate in candidates:
			if candidate.exists():
				base = candidate
				break
		if base is None:
			base = userProfile

		output = base / "NVDA_MLRecorder"
		output.mkdir(parents=True, exist_ok=True)
		return str(output)

	def _diagnosticDir(self) -> Path:
		base = Path(config.getUserDefaultConfigPath())
		diag = base / "mlrecorder-diagnostics"
		diag.mkdir(parents=True, exist_ok=True)
		return diag

	def _runtimeArchFolder(self) -> str:
		return "x86" if (8 * struct.calcsize("P")) == 32 else "x64"

	def _addonBinRoot(self) -> Path:
		return self._addonDir() / "lib" / "mlrecorder" / "bin"

	def _selectedBinDir(self) -> Path:
		binRoot = self._addonBinRoot()
		archDir = binRoot / self._runtimeArchFolder()
		if archDir.exists():
			return archDir
		return binRoot

	def _addonDllPaths(self, binDir: Optional[Path] = None) -> Tuple[Path, Path, Path, Path]:
		if binDir is None:
			binDir = self._selectedBinDir()
		return (
			binDir / "mlrecorder_core.dll",
			binDir / "FLAC.dll",
			binDir / "ogg.dll",
			binDir / "opus.dll",
		)

	def _peArchitecture(self, path: Path) -> str:
		if not path.exists():
			return "missing"
		try:
			data = path.read_bytes()
			if data[:2] != b"MZ":
				return "not-pe"
			e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
			sig = data[e_lfanew:e_lfanew + 4]
			if sig != b"PE\x00\x00":
				return "bad-pe-signature"
			machine = struct.unpack_from("<H", data, e_lfanew + 4)[0]
			if machine == 0x14C:
				return "x86"
			if machine == 0x8664:
				return "x64"
			if machine == 0xAA64:
				return "arm64"
			return "machine-0x%X" % machine
		except Exception as exc:
			return "error-reading-pe: %s" % exc

	def _probeLoad(self, path: Path) -> str:
		if not path.exists():
			return "missing"
		try:
			ctypes.WinDLL(str(path))
			return "ok"
		except Exception as exc:
			return "error: %s" % exc

	def _writeDiagnosticReport(self, reason: str, exc: Optional[Exception] = None) -> Path:
		now = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
		reportPath = self._diagnosticDir() / ("mlrecorder-diag-%s.txt" % now)
		self._lastDiagnosticPath = reportPath

		binRoot = self._addonBinRoot()
		selectedBinDir = self._selectedBinDir()
		core, flac, ogg, opus = self._addonDllPaths(selectedBinDir)
		lines = []
		lines.append("MLRecorder NVDA diagnostic report")
		lines.append("timestamp=%s" % datetime.datetime.now().isoformat())
		lines.append("reason=%s" % reason)
		lines.append("")
		lines.append("python_version=%s" % sys.version.replace("\n", " "))
		lines.append("python_executable=%s" % sys.executable)
		lines.append("python_bitness=%s" % (8 * struct.calcsize("P")))
		lines.append("PROCESSOR_ARCHITECTURE=%s" % os.environ.get("PROCESSOR_ARCHITECTURE", ""))
		lines.append("PROCESSOR_ARCHITEW6432=%s" % os.environ.get("PROCESSOR_ARCHITEW6432", ""))
		lines.append("MLRECORDER_DLL=%s" % os.environ.get("MLRECORDER_DLL", ""))
		lines.append("addon_dir=%s" % self._addonDir())
		lines.append("selected_runtime_arch=%s" % self._runtimeArchFolder())
		lines.append("selected_bin_dir=%s" % selectedBinDir)
		lines.append("bin_root=%s" % binRoot)
		lines.append("bin_x86_exists=%s" % (binRoot / "x86").exists())
		lines.append("bin_x64_exists=%s" % (binRoot / "x64").exists())
		lines.append("")
		lines.append("dll_checks:")

		for dllPath in (core, flac, ogg, opus):
			size = dllPath.stat().st_size if dllPath.exists() else 0
			lines.append("  path=%s" % dllPath)
			lines.append("    exists=%s" % dllPath.exists())
			lines.append("    size=%s" % size)
			lines.append("    pe_arch=%s" % self._peArchitecture(dllPath))
			lines.append("    load_probe=%s" % self._probeLoad(dllPath))

		if exc is not None:
			lines.append("")
			lines.append("exception=%s" % repr(exc))
			lines.append("traceback:")
			lines.append(traceback.format_exc())

		reportPath.write_text("\n".join(lines), encoding="utf-8", errors="replace")
		return reportPath

	def _ensureRuntime(self) -> bool:
		if self._mlr is not None:
			return True

		libDir = self._addonDir() / "lib"
		if str(libDir) not in sys.path:
			sys.path.insert(0, str(libDir))
		dllPath = self._selectedBinDir() / "mlrecorder_core.dll"

		try:
			import mlrecorder as mlr  # type: ignore
		except Exception as exc:
			self._speak(_("MLRecorder no disponible: %s") % exc)
			return False

		if not dllPath.exists():
			self._speak(_("No se encontró la DLL de MLRecorder en el addon."))
			return False

		try:
			# Force addon-local DLL to avoid accidental MLRECORDER_DLL env overrides.
			mlr.initialize(dll_path=str(dllPath))
		except Exception as exc:
			reportPath = self._writeDiagnosticReport("initialize-failed", exc=exc)
			self._speak(_("No se pudo inicializar MLRecorder: %s") % exc)
			self._speak(_("Se generó reporte de depuración en: %s") % reportPath.name)
			return False

		self._mlr = mlr
		return True

	def _getFocusProcess(self) -> Tuple[int, str]:
		focus = api.getFocusObject()
		pid = int(getattr(focus, "processID", 0) or 0)
		name = str(getattr(focus, "appName", "") or "").strip()
		if not name:
			try:
				appModule = getattr(focus, "appModule", None)
				name = str(getattr(appModule, "appName", "") or "").strip()
			except Exception:
				pass
		return pid, name

	def _normalizeProcessLabel(self, rawName: str) -> str:
		name = (rawName or "").strip()
		if not name:
			return ""

		# Prefer display without file extension for spoken feedback.
		if name.lower().endswith(".exe"):
			name = name[:-4]
		return name.strip()

	def _resolveProcessLabel(self, pid: int, fallbackName: str) -> str:
		name = self._normalizeProcessLabel(fallbackName)
		if name:
			return name

		if self._mlr is not None:
			try:
				for proc in self._mlr.list_processes(only_active_audio=False):  # type: ignore[union-attr]
					if int(getattr(proc, "process_id", 0)) == pid:
						procName = self._normalizeProcessLabel(str(getattr(proc, "process_name", "") or ""))
						if procName:
							return procName
			except Exception:
				pass

		return _("aplicación actual")

	def _stopAllSessions(self):
		if self._mixedSession is not None:
			try:
				self._mixedSession.stop()
			except Exception:
				pass
			self._mixedSession = None

		if self._microphoneSession is not None:
			try:
				self._microphoneSession.stop()
			except Exception:
				pass
			self._microphoneSession = None

		if self._processSession is not None:
			try:
				self._processSession.stop()
			except Exception:
				pass
			self._processSession = None
			self._lastProcessPid = None
			self._lastProcessName = ""

	@script(
		description=_("Inicia grabación del proceso enfocado."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+r",
	)
	def script_startFocusedProcessRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._processSession is not None:
			try:
				self._processSession.stop()
				self._processSession = None
				self._lastProcessPid = None
				self._lastProcessName = ""
				self._speak(_("Grabación de proceso detenida."))
			except Exception as exc:
				self._speak(_("Error al detener grabación de proceso: %s") % exc)
			return

		if self._mixedSession is not None:
			self._speak(_("Detén la sesión mixta antes de iniciar grabación de proceso."))
			return

		pid, appName = self._getFocusProcess()
		if pid <= 0:
			self._speak(_("No se pudo resolver el proceso enfocado."))
			return

		try:
			processLabel = self._resolveProcessLabel(pid, appName)
			self._processSession = self._mlr.start_recorder(  # type: ignore[union-attr]
				pid=pid,
				output_dir=self._defaultOutputDir(),
				fmt="wav",
				strict_process_isolation=True,
			)
			self._lastProcessPid = pid
			self._lastProcessName = processLabel
			self._speak(_("Grabando audio de %s.") % processLabel)
		except Exception as exc:
			self._speak(_("Error al iniciar grabación de proceso: %s") % exc)

	@script(
		description=_("Alterna grabación de micrófono."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+m",
	)
	def script_toggleMicrophoneRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._mixedSession is not None:
			self._speak(_("Detén la sesión mixta antes de grabar micrófono por separado."))
			return

		if self._microphoneSession is not None:
			try:
				self._microphoneSession.stop()
				self._microphoneSession = None
				self._speak(_("Grabación de micrófono detenida."))
			except Exception as exc:
				self._speak(_("Error al detener micrófono: %s") % exc)
			return

		try:
			self._microphoneSession = self._mlr.start_microphone_recorder(  # type: ignore[union-attr]
				output_dir=self._defaultOutputDir(),
				fmt="wav",
			)
			self._speak(_("Grabación de micrófono iniciada."))
		except Exception as exc:
			self._speak(_("Error al iniciar micrófono: %s") % exc)

	@script(
		description=_("Alterna grabación mixta de proceso enfocado más micrófono."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+x",
	)
	def script_toggleMixedRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._processSession is not None or self._microphoneSession is not None:
			self._speak(_("Detén sesiones activas antes de iniciar mezcla."))
			return

		if self._mixedSession is not None:
			try:
				self._mixedSession.stop()
				self._mixedSession = None
				self._speak(_("Grabación mixta detenida."))
			except Exception as exc:
				self._speak(_("Error al detener mezcla: %s") % exc)
			return

		pid, appName = self._getFocusProcess()
		if pid <= 0:
			self._speak(_("No se pudo resolver el proceso enfocado."))
			return

		try:
			processLabel = self._resolveProcessLabel(pid, appName)
			self._mixedSession = self._mlr.start_mixed_recorder(  # type: ignore[union-attr]
				pid=pid,
				output_dir=self._defaultOutputDir(),
				fmt="wav",
				include_microphone=True,
				strict_process_isolation=True,
				base_name="NVDA-Mixed",
			)
			self._speak(_("Grabación mixta iniciada para %s.") % processLabel)
		except Exception as exc:
			self._speak(_("Error al iniciar mezcla: %s") % exc)

	@script(
		description=_("Detiene todas las grabaciones activas de MLRecorder."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+s",
	)
	def script_stopAllRecordings(self, gesture):
		del gesture
		self._stopAllSessions()
		self._speak(_("Todas las grabaciones detenidas."))

	@script(
		description=_("Informa estado de grabación de MLRecorder."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+i",
	)
	def script_reportStatus(self, gesture):
		del gesture
		parts = []
		if self._processSession is not None:
			parts.append(_("proceso activo"))
		if self._microphoneSession is not None:
			parts.append(_("micrófono activo"))
		if self._mixedSession is not None:
			parts.append(_("mezcla activa"))
		if not parts:
			parts.append(_("sin grabaciones activas"))
		self._speak(", ".join(parts))

	@script(
		description=_("Genera reporte de depuración de MLRecorder."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+d",
	)
	def script_dumpDiagnostics(self, gesture):
		del gesture
		try:
			path = self._writeDiagnosticReport("manual-dump")
			self._speak(_("Reporte de depuración generado: %s") % path.name)
			ui.browseableMessage(
				path.read_text(encoding="utf-8", errors="replace"),
				_("MLRecorder - Reporte de depuración"),
				isHtml=False,
			)
		except Exception as exc:
			self._speak(_("No se pudo generar el reporte: %s") % exc)
