import os
from django.core.wsgi import get_wsgi_application
os.environ.setdefault('DJANGO_SETTINGS_MODULE','rail_casting_plan.settings')
application=get_wsgi_application()
