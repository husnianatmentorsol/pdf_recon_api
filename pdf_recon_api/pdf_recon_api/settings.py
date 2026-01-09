from pathlib import Path
import os

# Build paths inside the project like this: BASE_DIR / 'subdir'.
BASE_DIR = Path(__file__).resolve().parent.parent

# Quick-start development settings - unsuitable for production
SECRET_KEY = 'django-insecure-u0n+p-$vy%4r7v--ay3^_icf50e=l(zq_y*k4x(#q3e)$e*_b)'
DEBUG = True
# ALLOWED_HOSTS = [
#     host.strip()
#     for host in os.getenv(
#         "DJANGO_ALLOWED_HOSTS", "localhost,127.0.0.1,[::1]"
#     ).split(",")
#     if host.strip()
# ]
ALLOWED_HOSTS = ["*"]
# Application definition
INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',

    # Third-party apps
    "rest_framework",
    "corsheaders",  # for React frontend

    # Your apps
    "api",
]

MIDDLEWARE = [
    "corsheaders.middleware.CorsMiddleware",  # CORS middleware
    'django.middleware.security.SecurityMiddleware',
    "whitenoise.middleware.WhiteNoiseMiddleware",
    'django.contrib.sessions.middleware.SessionMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

ROOT_URLCONF = 'pdf_recon_api.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],  # You can add custom template directories here
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'pdf_recon_api.wsgi.application'

# Database
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

# Password validation
AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator',},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator',},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator',},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator',},
]

# Internationalization
LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True

# Static files (CSS, JavaScript, Images)
STATIC_URL = 'static/'
STATIC_ROOT = BASE_DIR / 'staticfiles'
STATICFILES_STORAGE = "whitenoise.storage.CompressedManifestStaticFilesStorage"

# Media files (uploaded PDFs)
# settings.py
MEDIA_ROOT = BASE_DIR / "media"
MEDIA_URL = '/media/'


# CORS configuration for React frontend
CORS_ALLOW_ALL_ORIGINS = True  # or specify origins like below
# CORS_ALLOWED_ORIGINS = [
#     "http://localhost:3000",
# ]

# Default primary key field type
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'

# Avoid automatic slash append in URLs (useful for APIs)
APPEND_SLASH = False
# Google Sheets credentials
GOOGLE_SERVICE_ACCOUNT_FILE = BASE_DIR / "credentials" / "ethereal-terra-441812-d5-22d30c1611b9.json"
MASTER_SHEET_ID= "1Z_ZKrKohFPQA_J4OKGviPBtLl7FexyKQuSbq-Hsa8JQ"
FOLDER_ID="1qAmDuqfK7oLzTBL04mdV2fA-ZbINBK-h"

X_FRAME_OPTIONS = 'ALLOWALL'
