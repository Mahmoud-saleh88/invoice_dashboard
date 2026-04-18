from django.contrib import admin
from .models import Client, Company, SalesInvoice, PurchaseInvoice, SalesInvoiceItem, Supplier
from django.contrib.auth.admin import UserAdmin
from .models import CustomUser


admin.site.register(SalesInvoice)
admin.site.register(PurchaseInvoice)
admin.site.register(Company)
admin.site.register(Client)
admin.site.register(Supplier)
admin.site.register(SalesInvoiceItem)

@admin.register(CustomUser)
class CustomUserAdmin(UserAdmin):
    fieldsets = UserAdmin.fieldsets + (
        ('معلومات إضافية', {'fields': ('mobile',)}),
    )
    add_fieldsets = UserAdmin.add_fieldsets + (
        ('معلومات إضافية', {'fields': ('mobile',)}),
    )
    list_display = ['username', 'email', 'get_full_name', 'mobile', 'is_staff']
    