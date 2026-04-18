from django.contrib import admin
from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static
from django.views.generic import RedirectView

urlpatterns = [
    path('admin/', admin.site.urls),
    
    path('', include('invoices.urls')),
    path('redirect-admin/', RedirectView.as_view(url='/admin/', permanent=False), name='redirect-admin'),
    path('', RedirectView.as_view(url='/dashboard/', permanent=False), name='home'),
]

# 4. Static files - ALWAYS LAST
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)