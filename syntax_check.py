
import sys
from unittest.mock import MagicMock

# Mock NVDA modules
sys.modules['addonHandler'] = MagicMock()
sys.modules['api'] = MagicMock()
sys.modules['config'] = MagicMock()
sys.modules['globalPluginHandler'] = MagicMock()
sys.modules['globalVars'] = MagicMock()
sys.modules['ui'] = MagicMock()
sys.modules['scriptHandler'] = MagicMock()

# Append path and try import
sys.path.append('addon/globalPlugins')
try:
    import mlrecorder
    print("Syntax Check Passed")
except SyntaxError as e:
    print(f"Syntax Error: {e}")
except ImportError as e:
    print(f"Import Error (likely missing dependency, but syntax is ok): {e}")
except Exception as e:
    print(f"Runtime Error (syntax likely ok): {e}")
