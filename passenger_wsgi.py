import os
import sys

INTERP = "/home/informatica/flota.guinezapp.cl/app/venv/bin/python"

if sys.executable != INTERP:
    os.execl(INTERP, INTERP, *sys.argv)

sys.path.insert(0, os.path.dirname(__file__))

os.environ.setdefault(
    "DJANGO_SETTINGS_MODULE",
    "config.settings"
)

from django.core.wsgi import get_wsgi_application
application = get_wsgi_application()