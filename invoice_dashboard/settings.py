import os
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent.parent
SECRET_KEY = 'django-insecure-invoice-dashboard-2025-secret-key'
AUTH_USER_MODEL = 'invoices.CustomUser'
DEBUG = os.environ.get('DJANGO_DEBUG', 'False') == 'True'
ALLOWED_HOSTS = ['mahmoudramadan.pythonanywhere.com','localhost','127.0.0.1']
SECURE_BROWSER_XSS_FILTER = True
X_FRAME_OPTIONS = 'DENY'
SESSION_EXPIRE_AT_BROWSER_CLOSE = True  # Let your middleware handle this

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'invoices',
    'django_bootstrap5'
]
MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]
ROOT_URLCONF = 'invoice_dashboard.urls'
TEMPLATES = [{'BACKEND': 'django.template.backends.django.DjangoTemplates', 'DIRS': [BASE_DIR / 'templates'], 'APP_DIRS': True, 'OPTIONS': {'context_processors': ['django.template.context_processors.debug', 'django.template.context_processors.request', 'django.contrib.auth.context_processors.auth', 'django.contrib.messages.context_processors.messages']}}]
WSGI_APPLICATION = 'invoice_dashboard.wsgi.application'
DATABASES = {'default': {'ENGINE': 'django.db.backends.sqlite3', 'NAME': BASE_DIR / 'db.sqlite3'}}
LANGUAGE_CODE = 'ar'
TIME_ZONE = 'Asia/Riyadh'
USE_I18N = True
USE_TZ = True
STATIC_URL = 'static/'
STATICFILES_DIRS = [os.path.join(BASE_DIR, 'static')]
STATIC_ROOT = os.path.join(BASE_DIR, 'staticfiles')
MEDIA_URL = '/media/'
MEDIA_ROOT = os.path.join(BASE_DIR, 'media')
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
LOGIN_URL = 'login'
LOGIN_REDIRECT_URL =  'dashboard'
LOGOUT_REDIRECT_URL = 'login'
# ── ZATCA ────────────────────────────────────────────────
# أضف هذه الإعدادات في settings.py

# ZATCA Settings
ZATCA_ENVIRONMENT = os.getenv('ZATCA_ENVIRONMENT', 'sandbox')  # production for live
ZATCA_SANDBOX_BASE = "https://gw-fatoora.zatca.gov.sa/e-invoicing/developer-portal"
ZATCA_PROD_BASE = "https://gw-fatoora.zatca.gov.sa/e-invoicing/core"

# Security
ZATCA_PRIVATE_KEY_PATH = os.getenv('ZATCA_PRIVATE_KEY_PATH', '')
ZATCA_CERTIFICATE_PATH = os.getenv('ZATCA_CERTIFICATE_PATH', '')
ZATCA_CSID = os.getenv('ZATCA_CSID', '')
ZATCA_SECRET = os.getenv('ZATCA_SECRET', '')

# Retry Settings
ZATCA_MAX_RETRIES = 3
ZATCA_RETRY_DELAY = 5  # seconds

# Logging
# في ملف settings.py

LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose': {  # ✅ أضف هذا القسم
            'format': '{levelname} {asctime} {module} {process:d} {thread:d} {message}',
            'style': '{',
        },
        'simple': {
            'format': '{levelname} {message}',
            'style': '{',
        },
    },
    'handlers': {
        'zatca_file': {
            'level': 'INFO',
            'class': 'logging.FileHandler',
            'filename': os.path.join(BASE_DIR, 'logs', 'zatca.log'),  # استخدم os.path.join
            'formatter': 'verbose',  # الآن هذا سيعمل
        },
        'console': {
            'level': 'INFO',
            'class': 'logging.StreamHandler',
            'formatter': 'simple',
        },
    },
    'loggers': {
        'invoices.zatca_service': {
            'handlers': ['zatca_file', 'console'],
            'level': 'INFO',
            'propagate': True,
        },
        'django': {
            'handlers': ['console'],
            'level': 'INFO',
        },
    },
}