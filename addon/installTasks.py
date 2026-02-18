# -*- coding: utf-8 -*-

"""Install tasks for MLRecorder NVDA addon.

Intentionally minimal: no external UI prompts, no network calls.
"""

import addonHandler

addonHandler.initTranslation()


def onInstall():
	# Keep install task side-effect free.
	return
