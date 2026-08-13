import os
import sys
import types
import unittest
from unittest.mock import Mock, patch

import app


class AppDispatchTests(unittest.TestCase):
    def test_zhkh_daemon_skips_source_picker(self):
        daemon_main = Mock()
        daemon_module = types.ModuleType("flows.zhkh_daemon")
        daemon_module.main = daemon_main

        with patch.object(
                sys, "argv",
                ["app.py", "--mode=zhkh-daemon", "--headless"]), \
             patch.object(app.cfg, "get_base_dir", return_value=os.getcwd()), \
             patch.object(app.cfg, "load", return_value={}), \
             patch.object(
                 app, "pick_source",
                 side_effect=AssertionError("source picker must not be called")), \
             patch.dict(sys.modules, {"flows.zhkh_daemon": daemon_module}), \
             patch.dict(os.environ, {}, clear=False):
            app.main()

        daemon_main.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
