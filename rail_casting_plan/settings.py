from pathlib import Path
BASE_DIR=Path(__file__).resolve().parent.parent
SECRET_KEY='spm-local-development-key-change-for-production'
DEBUG=True
ALLOWED_HOSTS = ['172.17.18.13', '127.0.0.1', 'localhost', '*']
CSRF_TRUSTED_ORIGINS = [
    'http://172.17.18.13:5674',
    'http://172.17.18.13:8000',
    'http://127.0.0.1:5674',
    'http://127.0.0.1:8000',
    'http://localhost:5674',
    'http://localhost:8000',
]
INSTALLED_APPS=['django.contrib.contenttypes','django.contrib.auth','django.contrib.sessions','django.contrib.messages','django.contrib.staticfiles','casting_plan']
MIDDLEWARE=['django.middleware.security.SecurityMiddleware','django.contrib.sessions.middleware.SessionMiddleware','django.middleware.common.CommonMiddleware','django.middleware.csrf.CsrfViewMiddleware','django.contrib.auth.middleware.AuthenticationMiddleware','django.contrib.messages.middleware.MessageMiddleware']
ROOT_URLCONF='rail_casting_plan.urls'
TEMPLATES=[{'BACKEND':'django.template.backends.django.DjangoTemplates','DIRS':[BASE_DIR/'casting_plan'/'templates'],'APP_DIRS':True,'OPTIONS':{'context_processors':['django.template.context_processors.request','django.contrib.auth.context_processors.auth','django.contrib.messages.context_processors.messages','casting_plan.context_processors.notifications']}}]
WSGI_APPLICATION='rail_casting_plan.wsgi.application'
DATABASES={'default':{'ENGINE':'django.db.backends.sqlite3','NAME':BASE_DIR/'rail_mill.sqlite3'}}
LANGUAGE_CODE='en-us'; TIME_ZONE='Asia/Kolkata'; USE_I18N=True; USE_TZ=True
STATIC_URL='static/'
STATICFILES_DIRS=[BASE_DIR/'casting_plan'/'static']
DEFAULT_AUTO_FIELD='django.db.models.BigAutoField'

LOGIN_URL = 'login'
LOGIN_REDIRECT_URL = '/'

