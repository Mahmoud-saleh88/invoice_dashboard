from django.apps import AppConfig


class InvoicesConfig(AppConfig):
    name = 'invoices'
    
class InvoicesConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'invoices'

    def ready(self):
        # يُنفَّذ مرة واحدة عند بدء Django فقط — لا عند import
        from django.conf import settings
        import os
        logs_dir = settings.BASE_DIR / 'logs'
        os.makedirs(logs_dir, exist_ok=True)
