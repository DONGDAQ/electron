import sys
import os
from pathlib import Path

if getattr(sys, 'frozen', False):
    base = Path(sys._MEIPASS)
else:
    base = Path(__file__).parent

sys.path.insert(0, str(base))
os.chdir(str(base))

from quote_system.web_app import app
app.run(host="127.0.0.1", port=5000, debug=False)
