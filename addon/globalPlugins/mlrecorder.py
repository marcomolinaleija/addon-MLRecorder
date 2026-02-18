# MLRecorder add-on for NVDA.
# Copyright (C) 2026

from __future__ import annotations

import ctypes
import struct
import os
import sys
from pathlib import Path
from typing import Optional, Tuple

import addonHandler
import api
import config
import globalPluginHandler
import globalVars
import gui
from gui import settingsDialogs
import wx
import ui
import tones
from scriptHandler import script

addonHandler.initTranslation()


def disableInSecureMode(decoratedCls):
	if globalVars.appArgs.secure:
		return globalPluginHandler.GlobalPlugin
	return decoratedCls


class MLRecorderSettingsPanel(settingsDialogs.SettingsPanel):
	title = _("MLRecorder")

	def makeSettings(self, settingsSizer):
		sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		# Format
		self.formatLabel = sHelper.addItem(wx.StaticText(self, label=_("&Formato de salida:")))
		self.formatChoice = sHelper.addItem(wx.Choice(self, choices=["wav", "mp3", "flac", "opus"]))
		self.formatChoice.SetStringSelection(config.conf["mlrecorder"]["outputFormat"])

		# Skip Silence
		self.skipSilenceCb = sHelper.addItem(wx.CheckBox(self, label=_("&Saltar silencios")))
		self.skipSilenceCb.SetValue(config.conf["mlrecorder"]["skipSilence"])

		# Process Volume
		self.volumeLabel = sHelper.addItem(wx.StaticText(self, label=_("&Volumen de proceso (%):")))
		self.volumeSlider = sHelper.addItem(wx.SpinCtrl(self, value=str(config.conf["mlrecorder"]["processVolume"]), min=0, max=200))

		# Microphone
		# We need to try to get devices. If runtime isn't loaded, we might not see them all,
		# but usually GlobalPlugin loads it.
		# For this panel, we'll try to use the global instance if available.
		currentMicId = config.conf["mlrecorder"]["microphoneId"]
		choices = [_("Predeterminado")]
		self.micIds = [""]
		
		# Try to list devices from the runtime if possible
		try:
			# We can't easily access the plugin instance here, so we might need a workaround 
			# or just rely on what's available. 
			# Ideally we would access GlobalPlugin instance but it's not global.
			# For now, we'll just show the Default option and if possible list others if we can get a handle.
			# Note: In a real implementation with valid dll, we'd call mlr.list_input_devices().
			# Since we are in a mocked env or standard NVDA env, we might not have the DLL loaded here.
			pass
		except Exception:
			pass

		self.micLabel = sHelper.addItem(wx.StaticText(self, label=_("&Micrófono:")))
		self.micChoice = sHelper.addItem(wx.Choice(self, choices=choices))
		self.micChoice.SetSelection(0) # Default to first

	def onSave(self):
		config.conf["mlrecorder"]["outputFormat"] = self.formatChoice.GetStringSelection()
		config.conf["mlrecorder"]["skipSilence"] = self.skipSilenceCb.GetValue()
		config.conf["mlrecorder"]["processVolume"] = self.volumeSlider.GetValue()
		# config.conf["mlrecorder"]["microphoneId"] = self.micIds[self.micChoice.GetSelection()]


@disableInSecureMode
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("MLRecorder")

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		config.conf.spec["mlrecorder"] = {
			"outputFormat": "string(default='wav')",
			"skipSilence": "boolean(default=False)",
			"processVolume": "integer(default=100, min=0, max=200)",
			"microphoneId": "string(default='')",
		}
		settingsDialogs.NVDASettingsDialog.categoryClasses.append(MLRecorderSettingsPanel)
		self._mlr = None
		self._processSession = None
		self._microphoneSession = None
		self._systemSession = None
		self._mixedSession = None
		self.toggling = False

	def getScript(self, gesture):
		if not self.toggling:
			return super().getScript(gesture)
		script = super().getScript(gesture)
		if not script:
			script = self.script_error
		return self.finish(script)

	def finish(self, script):
		def wrapper(*args, **kwargs):
			try:
				script(*args, **kwargs)
			finally:
				self.deactivateLayer()
		return wrapper

	def deactivateLayer(self):
		self.toggling = False
		self.clearGestureBindings()
		self.bindGestures(self.__gestures)

	def script_error(self, gesture):
		# Translators: Error message when no function is assigned to the pressed key in command layer
		ui.message(_("Tecla no válida en la capa de comandos"))
		# tones.beep(277.18, 110) # Optional beep
		self._mixedSession = None
		self._systemSession = None
		self._lastProcessPid: Optional[int] = None
		self._lastProcessName = ""


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
			self._speak(_("No se pudo inicializar MLRecorder: %s") % exc)
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
		if self._systemSession is not None:
			try:
				self._systemSession.stop()
			except Exception:
				pass
			self._systemSession = None

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
		description=_("Alterna grabación del proceso enfocado."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+r",
	)
	def script_toggleFocusedProcessRecording(self, gesture):
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
			self._speak(_("La grabación MIXTA está activa. Deténla primero."))
			return
		if self._systemSession is not None:
			self._speak(_("La grabación de SISTEMA está activa. Deténla primero."))
			return

		pid, appName = self._getFocusProcess()

		if pid <= 0:
			self._speak(_("No se pudo resolver el proceso enfocado."))
			return

		try:
			processLabel = self._resolveProcessLabel(pid, appName)
			fmt = config.conf["mlrecorder"]["outputFormat"]
			skipSilence = config.conf["mlrecorder"]["skipSilence"]

			self._processSession = self._mlr.start_recorder(  # type: ignore[union-attr]
				pid=pid,
				output_dir=self._defaultOutputDir(),
				fmt=fmt,
				skip_silence=skipSilence,
				strict_process_isolation=True,
			)
			self._lastProcessPid = pid
			self._lastProcessName = processLabel

			# Apply volume
			try:
				vol = config.conf["mlrecorder"]["processVolume"]
				# Convert 0-200 to 0.0-2.0
				volFloat = float(vol) / 100.0
				self._mlr.set_capture_volume(pid, volFloat)
			except Exception:
				pass

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
			fmt = config.conf["mlrecorder"]["outputFormat"]
			self._microphoneSession = self._mlr.start_microphone_recorder(  # type: ignore[union-attr]
				output_dir=self._defaultOutputDir(),
				fmt=fmt,
			)
			self._speak(_("Grabación de micrófono iniciada."))
		except Exception as exc:
			self._speak(_("Error al iniciar micrófono: %s") % exc)

	@script(
		description=_("Alterna grabación del audio del sistema (escritorio)."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+g",
	)
	def script_toggleSystemRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._systemSession is not None:
			try:
				self._systemSession.stop()
				self._systemSession = None
				self._speak(_("Grabación de sistema detenida."))
			except Exception as exc:
				self._speak(_("Error al detener grabación de sistema: %s") % exc)
			return

		if self._processSession is not None or self._mixedSession is not None:
			self._speak(_("Detén las sesiones de proceso o mixtas antes de iniciar grabación de sistema."))
			return

		try:
			fmt = config.conf["mlrecorder"]["outputFormat"]
			skipSilence = config.conf["mlrecorder"]["skipSilence"]

			# pid=0 indicates system audio (desktop)
			self._systemSession = self._mlr.start_recorder(  # type: ignore[union-attr]
				pid=0,
				output_dir=self._defaultOutputDir(),
				fmt=fmt,
				skip_silence=skipSilence,
				strict_process_isolation=False,
			)

			self._speak(_("Grabación de sistema iniciada."))
		except Exception as exc:
			self._speak(_("Error al iniciar grabación de sistema: %s") % exc)

	@script(
		description=_("Alterna grabación mixta de proceso enfocado más micrófono."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+x",
	)
	def script_toggleMixedRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._processSession is not None or self._microphoneSession is not None or self._systemSession is not None:
			self._speak(_("Detén sesiones activas (proceso, micrófono o sistema) antes de iniciar mezcla."))
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
			fmt = config.conf["mlrecorder"]["outputFormat"]
			self._mixedSession = self._mlr.start_mixed_recorder(  # type: ignore[union-attr]
				pid=pid,
				output_dir=self._defaultOutputDir(),
				fmt=fmt,
				include_microphone=True,
				strict_process_isolation=True,
				base_name="NVDA-Mixed",
			)

			# Apply volume for the process part if possible.
			try:
				vol = config.conf["mlrecorder"]["processVolume"]
				volFloat = float(vol) / 100.0
				self._mlr.set_capture_volume(pid, volFloat)
			except Exception:
				pass

			self._speak(_("Grabación mixta iniciada para %s.") % processLabel)
		except Exception as exc:
			self._speak(_("Error al iniciar mezcla: %s") % exc)

	@script(
		description=_("Alterna grabación mixta de Sistema (Escritorio) más micrófono."),
		category=_("MLRecorder"),
	)
	def script_toggleSystemMixedRecording(self, gesture):
		del gesture
		if not self._ensureRuntime():
			return

		if self._processSession is not None or self._microphoneSession is not None or self._systemSession is not None:
			self._speak(_("Detén sesiones activas (proceso, micrófono o sistema) antes de iniciar mezcla de sistema."))
			return

		if self._mixedSession is not None:
			# If it's a system mix, we stop it. If it's a process mix, we also stop it (toggle behavior).
			try:
				self._mixedSession.stop()
				self._mixedSession = None
				self._speak(_("Grabación mixta detenida."))
			except Exception as exc:
				self._speak(_("Error al detener mezcla: %s") % exc)
			return

		try:
			fmt = config.conf["mlrecorder"]["outputFormat"]
			# PID 0 is system audio
			self._mixedSession = self._mlr.start_mixed_recorder(  # type: ignore[union-attr]
				pid=0,
				output_dir=self._defaultOutputDir(),
				fmt=fmt,
				include_microphone=True,
				strict_process_isolation=False, # Not needed for system
				base_name="System-Mixed",
			)

			self._speak(_("Grabación mixta de Sistema iniciada."))
		except Exception as exc:
			self._speak(_("Error al iniciar mezcla de sistema: %s") % exc)


	@script(
		description=_("Reporta el estado actual de las grabaciones."),
		category=_("MLRecorder"),
		gesture="kb:NVDA+shift+i",
	)
	def script_reportStatus(self, gesture):
		del gesture
		msgs = []
		if self._processSession:
			msgs.append(_("Grabando proceso: %s") % self._lastProcessName)
		if self._systemSession:
			msgs.append(_("Grabando sistema"))
		if self._mixedSession:
			# Check internal process_id to distinguish system mix vs process mix if possible.
			# Using getattr with default in case types are loose.
			if getattr(self._mixedSession, "process_id", -1) == 0:
				msgs.append(_("Grabación mixta (Sistema + Mic)"))
			else:
				msgs.append(_("Grabación mixta (Proceso + Mic)"))
		if self._microphoneSession:
			msgs.append(_("Grabando micrófono"))

		if not msgs:
			self._speak(_("No hay grabaciones activas."))
		else:
			self._speak(", ".join(msgs))

	@script(
		description=_("Detiene la grabación activa, sea del tipo que sea."),
		category=_("MLRecorder"),
	)
	def script_stopActiveRecording(self, gesture):
		del gesture
		stopped_something = False
		
		# Stop System
		if self._systemSession:
			try:
				self._systemSession.stop()
				self._systemSession = None
				self._speak(_("Grabación de sistema detenida."))
				stopped_something = True
			except Exception as exc:
				self._speak(_("Error al detener sistema: %s") % exc)

		# Stop Process
		if self._processSession:
			try:
				self._processSession.stop()
				self._processSession = None
				self._speak(_("Grabación de proceso detenida."))
				stopped_something = True
			except Exception as exc:
				self._speak(_("Error al detener proceso: %s") % exc)

		# Stop Mixed
		if self._mixedSession:
			try:
				self._mixedSession.stop()
				self._mixedSession = None
				self._speak(_("Grabación mixta detenida."))
				stopped_something = True
			except Exception as exc:
				self._speak(_("Error al detener mezcla: %s") % exc)

		# Stop Mic
		if self._microphoneSession:
			try:
				self._microphoneSession.stop()
				self._microphoneSession = None
				self._speak(_("Grabación de micrófono detenida."))
				stopped_something = True
			except Exception as exc:
				self._speak(_("Error al detener micrófono: %s") % exc)
				
		if not stopped_something:
			self._speak(_("No hay grabaciones activas para detener."))

	@script(
		description=_("Detiene TODAS las grabaciones activas."),
		category=_("MLRecorder"),
	)
	def script_stopAllRecordings(self, gesture):
		del gesture
		self._stopAllSessions()
		self._speak(_("Todas las grabaciones detenidas."))

	@script(
		description=_("Abre la carpeta de grabaciones."),
		category=_("MLRecorder"),
	)
	def script_openRecordingsFolder(self, gesture):
		del gesture
		full_path = self._defaultOutputDir()
		if os.path.isdir(full_path):
			os.startfile(full_path)
		else:
			self._speak(_("La carpeta de grabaciones no existe o no se pudo encontrar."))

	@script(
		description=_("Activa la capa de comandos de MLRecorder"),
		category=_("MLRecorder"),
		gesture="kb:NVDA+alt+a",
	)
	def script_commandLayer(self, gesture):
		if self.toggling:
			self.script_error(gesture)
			return
		self.bindGestures(self.__commandGestures)
		self.toggling = True
		self._speak(_("Capa de audio activada."))
		tones.beep(349.23, 110) # Optional beep

	__commandGestures = {
		"kb:p": "toggleFocusedProcessRecording",
		"kb:alt+p": "toggleMixedRecording",
		"kb:s": "toggleSystemRecording",
		"kb:alt+s": "toggleSystemMixedRecording",
		"kb:m": "toggleMicrophoneRecording", # Kept 'm' for mic, as 's' is system
		"kb:d": "stopActiveRecording",
		"kb:alt+d": "stopAllRecordings",
		"kb:o": "openRecordingsFolder",
		"kb:i": "reportStatus",
	}
	__gestures = {
		"kb:NVDA+shift+r": "toggleFocusedProcessRecording",
		"kb:NVDA+shift+m": "toggleMicrophoneRecording",
		"kb:NVDA+shift+g": "toggleSystemRecording",
		"kb:NVDA+shift+x": "toggleMixedRecording",
		"kb:NVDA+shift+h": "toggleSystemMixedRecording",
		"kb:NVDA+alt+a": "commandLayer",
		"kb:NVDA+shift+i": "reportStatus",
		"kb:NVDA+shift+o": "openRecordingsFolder",
	}
