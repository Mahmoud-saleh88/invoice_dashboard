# admin.py
from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from django.utils.html import format_html
from django.utils import timezone
from django.urls import reverse
from django.db.models import Sum
from decimal import Decimal
from .models import (AccountCategory, AccountGroup, Account,JournalEntry, JournalEntryLine,
                    CostCenter,Company, Client, Supplier,SalesInvoice, SalesInvoiceItem,
                    PurchaseInvoice,ZATCADevice, ZATCALog,FiscalYear, AccountingPeriod,
                    Payment, PaymentInvoiceAllocation,CustomUser)

# ============================================================
# CHART OF ACCOUNTS
# ============================================================

@admin.register(AccountCategory)
class AccountCategoryAdmin(admin.ModelAdmin):
    list_display = ['code', 'name_ar', 'name_en', 'normal_balance']
    list_filter = ['normal_balance']
    search_fields = ['code', 'name_ar', 'name_en']
    ordering = ['code']


@admin.register(AccountGroup)
class AccountGroupAdmin(admin.ModelAdmin):
    list_display = ['code', 'name_ar', 'name_en', 'category', 'normal_balance']
    list_filter = ['category', 'normal_balance']
    search_fields = ['code', 'name_ar', 'name_en']
    ordering = ['code']


@admin.register(Account)
class AccountAdmin(admin.ModelAdmin):
    list_display = ['code', 'name_ar', 'name_en', 'group', 'account_type', 'normal_balance', 'is_active', 'balance_display']
    list_filter = ['group__category', 'group', 'account_type', 'normal_balance', 'is_active']
    search_fields = ['code', 'name_ar', 'name_en']
    ordering = ['code']
    list_select_related = ['group', 'group__category', 'parent']
    readonly_fields = ['balance_display']
    fieldsets = (
        ('معلومات أساسية', {
            'fields': ('code', 'name_ar', 'name_en', 'group', 'parent')
        }),
        ('نوع الحساب', {
            'fields': ('account_type', 'normal_balance', 'is_active')
        }),
        ('الرصيد الافتتاحي', {
            'fields': ('opening_balance', 'opening_balance_date')
        }),
        ('ملاحظات', {
            'fields': ('notes',)
        }),
    )
    
    def balance_display(self, obj):
        if obj.account_type == 'detail':
            balance = obj.get_balance()
            color = 'green' if balance >= 0 else 'red'
            return format_html('<span style="color: {}; font-weight: bold;">{:,.2f}</span>', color, balance)
        return '-'
    balance_display.short_description = 'الرصيد الحالي'


# ============================================================
# JOURNAL ENTRIES
# ============================================================

class JournalEntryLineInline(admin.TabularInline):
    model = JournalEntryLine
    extra = 1
    fields = ['account', 'description', 'debit_amount', 'credit_amount', 'cost_center']
    autocomplete_fields = ['account', 'cost_center']
    
    def get_queryset(self, request):
        return super().get_queryset(request).select_related('account', 'cost_center')


@admin.register(JournalEntry)
class JournalEntryAdmin(admin.ModelAdmin):
    list_display = ['number', 'date', 'entry_type', 'description_preview', 'status', 'total_display', 'is_balanced_display']
    list_filter = ['status', 'entry_type', 'date']
    search_fields = ['number', 'description', 'reference']
    ordering = ['-date', '-number']
    date_hierarchy = 'date'
    readonly_fields = ['created_at', 'posted_at', 'total_debit', 'total_credit']
    autocomplete_fields = ['created_by']
    inlines = [JournalEntryLineInline]
    actions = ['post_selected', 'reverse_selected']
    fieldsets = (
        ('معلومات القيد', {
            'fields': ('number', 'date', 'entry_type', 'status')
        }),
        ('الوصف', {
            'fields': ('description', 'reference', 'notes')
        }),
        ('الارتباطات', {
            'fields': ('sales_invoice', 'purchase_invoice', 'reversed_entry')
        }),
        ('معلومات النظام', {
            'fields': ('created_by', 'created_at', 'posted_at')
        }),
    )
    
    def description_preview(self, obj):
        return obj.description[:50] + '...' if len(obj.description) > 50 else obj.description
    description_preview.short_description = 'البيان'
    
    def total_display(self, obj):
        return format_html(
            '<span style="color: #2E7D32;">{:,.2f}</span> / <span style="color: #C62828;">{:,.2f}</span>',
            obj.total_debit, obj.total_credit
        )
    total_display.short_description = 'مدين / دائن'
    
    def is_balanced_display(self, obj):
        if obj.is_balanced:
            return format_html('<span style="color: green;">✓ متوازن</span>')
        return format_html('<span style="color: red;">✗ غير متوازن</span>')
    is_balanced_display.short_description = 'التوازن'
    
    def post_selected(self, request, queryset):
        posted = 0
        for entry in queryset.filter(status='draft'):
            if entry.is_balanced:
                entry.status = 'posted'
                entry.posted_at = timezone.now()
                entry.save()
                posted += 1
        self.message_user(request, f'تم ترحيل {posted} قيد بنجاح.')
    post_selected.short_description = 'ترحيل القيود المحددة'
    
    def reverse_selected(self, request, queryset):
        reversed_count = 0
        for entry in queryset.filter(status='posted'):
            try:
                entry.reverse(request.user)
                reversed_count += 1
            except Exception as e:
                self.message_user(request, f'خطأ في عكس القيد {entry.number}: {e}', level='ERROR')
        self.message_user(request, f'تم عكس {reversed_count} قيد بنجاح.')
    reverse_selected.short_description = 'عكس القيود المحددة'


@admin.register(JournalEntryLine)
class JournalEntryLineAdmin(admin.ModelAdmin):
    list_display = ['journal_entry', 'account', 'description_preview', 'debit_amount', 'credit_amount', 'cost_center']
    list_filter = ['account__group__category', 'journal_entry__date']
    search_fields = ['account__name_ar', 'account__code', 'description']
    autocomplete_fields = ['journal_entry', 'account', 'cost_center']
    raw_id_fields = ['journal_entry']
    
    def description_preview(self, obj):
        return obj.description[:40] + '...' if len(obj.description) > 40 else obj.description
    description_preview.short_description = 'البيان'


# ============================================================
# COST CENTERS
# ============================================================

@admin.register(CostCenter)
class CostCenterAdmin(admin.ModelAdmin):
    list_display = ['code', 'name_ar', 'name_en', 'parent', 'is_active']
    list_filter = ['is_active']
    search_fields = ['code', 'name_ar', 'name_en']
    ordering = ['code']


# ============================================================
# COMPANY & PARTIES
# ============================================================

@admin.register(Company)
class CompanyAdmin(admin.ModelAdmin):
    list_display = ['name', 'vat_number', 'cr_number', 'city', 'email', 'phone', 'zatca_registered']
    search_fields = ['name', 'vat_number', 'cr_number']
    fieldsets = (
        ('معلومات أساسية', {
            'fields': ('name', 'name_en', 'logo')
        }),
        ('العنوان', {
            'fields': ('street_name', 'building_number', 'district', 'city', 'postal_code')
        }),
        ('معلومات قانونية', {
            'fields': ('vat_number', 'cr_number')
        }),
        ('الاتصال', {
            'fields': ('email', 'phone')
        }),
        ('زاتكا (قديم)', {
            'fields': ('zatca_csid', 'zatca_private_key', 'zatca_certificate', 'zatca_registered'),
            'classes': ('collapse',)
        }),
    )


@admin.register(Client)
class ClientAdmin(admin.ModelAdmin):
    list_display = ['name', 'vat_number', 'city', 'phone', 'credit_limit', 'balance_display', 'is_active']
    list_filter = ['is_active', 'city']
    search_fields = ['name', 'name_en', 'vat_number', 'phone', 'email']
    autocomplete_fields = ['ar_account']
    
    def balance_display(self, obj):
        if obj.ar_account:
            balance = obj.ar_account.get_balance()
            color = 'red' if balance > 0 else 'green'
            return format_html('<span style="color: {}; font-weight: bold;">{:,.2f}</span>', color, abs(balance))
        return '-'
    balance_display.short_description = 'الرصيد'


@admin.register(Supplier)
class SupplierAdmin(admin.ModelAdmin):
    list_display = ['name', 'vat_number', 'city', 'phone', 'balance_display', 'is_active']
    list_filter = ['is_active', 'city']
    search_fields = ['name', 'name_en', 'vat_number', 'phone', 'email']
    autocomplete_fields = ['ap_account']
    
    def balance_display(self, obj):
        if obj.ap_account:
            balance = obj.ap_account.get_balance()
            color = 'red' if balance > 0 else 'green'
            return format_html('<span style="color: {}; font-weight: bold;">{:,.2f}</span>', color, abs(balance))
        return '-'
    balance_display.short_description = 'الرصيد'


# ============================================================
# SALES INVOICES
# ============================================================

class SalesInvoiceItemInline(admin.TabularInline):
    model = SalesInvoiceItem
    extra = 1
    fields = ['description', 'quantity', 'unit_price', 'discount_percent', 'tax_rate', 'total', 'account']
    readonly_fields = ['total']
    autocomplete_fields = ['account']


@admin.register(SalesInvoice)
class SalesInvoiceAdmin(admin.ModelAdmin):
    list_display = ['invoice_number', 'date', 'client', 'transaction_type', 'total', 'amount_paid', 'balance_due', 'zatca_status_display']
    list_filter = ['transaction_type', 'zatca_status', 'payment_method', 'date']
    search_fields = ['invoice_number', 'client__name', 'notes']
    ordering = ['-date', '-invoice_number']
    date_hierarchy = 'date'
    readonly_fields = ['total_quantity', 'total_sales_excl_tax', 'taxable_amount', 'total_tax', 'total', 'created_at']
    autocomplete_fields = ['company', 'client', 'reference_invoice']
    inlines = [SalesInvoiceItemInline]
    actions = ['submit_to_zatca']
    fieldsets = (
        ('معلومات الفاتورة', {
            'fields': ('invoice_number', 'invoice_uuid', 'date', 'supply_date', 'due_date')
        }),
        ('الطرفان', {
            'fields': ('company', 'client')
        }),
        ('نوع المعاملة', {
            'fields': ('transaction_type', 'payment_method')
        }),
        ('المبالغ', {
            'fields': ('total_quantity', 'total_sales_excl_tax', 'discount_amount', 'taxable_amount', 'tax_rate', 'total_tax', 'total', 'amount_paid')
        }),
        ('زاتكا', {
            'fields': ('zatca_status', 'zatca_submission_id', 'zatca_invoice_hash', 'zatca_qr_code'),
            'classes': ('collapse',)
        }),
        ('معلومات إضافية', {
            'fields': ('created_by', 'created_at', 'notes')
        }),
    )
    
    def zatca_status_display(self, obj):
        colors = {
            'pending': 'orange',
            'submitted': 'blue',
            'accepted': 'green',
            'rejected': 'red',
            'cleared': 'darkgreen',
            'reported': 'purple',
        }
        color = colors.get(obj.zatca_status, 'gray')
        return format_html(
            '<span style="color: {}; font-weight: bold;">{}</span>',
            color, obj.get_zatca_status_display()
        )
    zatca_status_display.short_description = 'حالة زاتكا'
    
    def balance_due(self, obj):
        """Display the balance due for the invoice"""
        balance = obj.total - obj.amount_paid
        color = 'red' if balance > 0 else 'green'
        return format_html(
            '<span style="color: {}; font-weight: bold;">{:,.2f}</span>',
            color, balance
        )
    balance_due.short_description = 'المبلغ المتبقي'
    balance_due.admin_order_field = 'total'  # Allow sorting
    
    def submit_to_zatca(self, request, queryset):
        from .zatca_service import ZATCAInvoiceService
        submitted = 0
        for invoice in queryset.filter(zatca_status='pending'):
            device = ZATCADevice.objects.filter(company=invoice.company, is_active=True).first()
            if device:
                try:
                    service = ZATCAInvoiceService(device)
                    result = service.process(invoice)
                    if result['success']:
                        submitted += 1
                except Exception:
                    pass
        self.message_user(request, f'تم إرسال {submitted} فاتورة إلى زاتكا.')
    submit_to_zatca.short_description = 'إرسال إلى زاتكا'

@admin.register(SalesInvoiceItem)
class SalesInvoiceItemAdmin(admin.ModelAdmin):
    list_display = ['invoice', 'description_preview', 'quantity', 'unit_price', 'total', 'tax_amount']
    search_fields = ['description', 'invoice__invoice_number']
    autocomplete_fields = ['invoice', 'account']
    raw_id_fields = ['invoice']
    
    def description_preview(self, obj):
        return obj.description[:40] + '...' if len(obj.description) > 40 else obj.description
    description_preview.short_description = 'الوصف'


# ============================================================
# PURCHASE INVOICES
# ============================================================

@admin.register(PurchaseInvoice)
class PurchaseInvoiceAdmin(admin.ModelAdmin):
    list_display = ['invoice_number', 'date', 'supplier', 'transaction_type', 'total', 'amount_paid', 'balance_due']
    list_filter = ['transaction_type', 'date']
    search_fields = ['invoice_number', 'supplier__name', 'notes']
    ordering = ['-date', '-invoice_number']
    date_hierarchy = 'date'
    readonly_fields = ['created_at']
    autocomplete_fields = ['supplier']
    fieldsets = (
        ('معلومات الفاتورة', {
            'fields': ('invoice_number', 'date', 'entry_date', 'due_date')
        }),
        ('المورد', {
            'fields': ('supplier', 'tax_number')
        }),
        ('نوع المعاملة', {
            'fields': ('transaction_type',)
        }),
        ('المبالغ', {
            'fields': ('subtotal', 'discount_amount', 'tax_amount', 'total', 'amount_paid')
        }),
        ('مرفقات وملاحظات', {
            'fields': ('purchaseinvoice_attachments', 'notes')
        }),
        ('معلومات النظام', {
            'fields': ('created_by', 'created_at')
        }),
    )
    def balance_due(self, obj):
        """Display the balance due for the invoice"""
        balance = obj.total - obj.amount_paid
        color = 'red' if balance > 0 else 'green'
        return format_html(
            '<span style="color: {}; font-weight: bold;">{:,.2f}</span>',
            color, balance
        )
    balance_due.short_description = 'المبلغ المتبقي'
    balance_due.admin_order_field = 'total'

# ============================================================
# ZATCA INTEGRATION
# ============================================================

@admin.register(ZATCADevice)
class ZATCADeviceAdmin(admin.ModelAdmin):
    list_display = ['device_name', 'company', 'environment', 'is_active', 'registered_at']
    list_filter = ['environment', 'is_active', 'company']
    search_fields = ['device_name', 'serial_number']
    readonly_fields = ['csr', 'certificate', 'csid', 'secret', 'registered_at']


@admin.register(ZATCALog)
class ZATCALogAdmin(admin.ModelAdmin):
    list_display = ['action', 'device', 'sales_invoice', 'success_display', 'response_code', 'created_at']
    list_filter = ['action', 'success', 'device__environment']
    search_fields = ['sales_invoice__invoice_number', 'error_message']
    ordering = ['-created_at']
    readonly_fields = ['device', 'sales_invoice', 'action', 'request_payload', 'response_code', 'response_body', 'success', 'created_at', 'error_message']
    
    def success_display(self, obj):
        if obj.success:
            return format_html('<span style="color: green;">✓ ناجح</span>')
        return format_html('<span style="color: red;">✗ فشل</span>')
    success_display.short_description = 'النتيجة'


# ============================================================
# FISCAL YEAR
# ============================================================

class AccountingPeriodInline(admin.TabularInline):
    model = AccountingPeriod
    extra = 0
    fields = ['name', 'start_date', 'end_date', 'is_closed']


@admin.register(FiscalYear)
class FiscalYearAdmin(admin.ModelAdmin):
    list_display = ['name', 'company', 'start_date', 'end_date', 'is_closed_display']
    list_filter = ['is_closed', 'company']
    search_fields = ['name']
    ordering = ['-start_date']
    inlines = [AccountingPeriodInline]
    actions = ['close_fiscal_year']
    
    def is_closed_display(self, obj):
        if obj.is_closed:
            return format_html('<span style="color: red;">🔒 مغلقة</span>')
        return format_html('<span style="color: green;">🔓 مفتوحة</span>')
    is_closed_display.short_description = 'الحالة'
    
    def close_fiscal_year(self, request, queryset):
        for fy in queryset.filter(is_closed=False):
            fy.is_closed = True
            fy.closed_at = timezone.now()
            fy.save()
        self.message_user(request, f'تم إغلاق {queryset.count()} سنة مالية.')
    close_fiscal_year.short_description = 'إغلاق السنوات المالية المحددة'


@admin.register(AccountingPeriod)
class AccountingPeriodAdmin(admin.ModelAdmin):
    list_display = ['name', 'fiscal_year', 'start_date', 'end_date', 'is_closed']
    list_filter = ['is_closed', 'fiscal_year']
    search_fields = ['name']


# ============================================================
# PAYMENTS
# ============================================================

class PaymentInvoiceAllocationInline(admin.TabularInline):
    model = PaymentInvoiceAllocation
    extra = 1
    fields = ['sales_invoice', 'purchase_invoice', 'allocated_amount']
    autocomplete_fields = ['sales_invoice', 'purchase_invoice']


@admin.register(Payment)
class PaymentAdmin(admin.ModelAdmin):
    list_display = ['number', 'date', 'payment_type', 'payment_method', 'party_display', 'amount']
    list_filter = ['payment_type', 'payment_method', 'date']
    search_fields = ['number', 'client__name', 'supplier__name', 'description']
    ordering = ['-date']
    autocomplete_fields = ['client', 'supplier', 'bank_account', 'cash_account', 'journal_entry', 'created_by']
    inlines = [PaymentInvoiceAllocationInline]
    
    def party_display(self, obj):
        if obj.client:
            return f'عميل: {obj.client.name}'
        elif obj.supplier:
            return f'مورد: {obj.supplier.name}'
        return '-'
    party_display.short_description = 'الطرف'


# ============================================================
# CUSTOM USER
# ============================================================

@admin.register(CustomUser)
class CustomUserAdmin(UserAdmin):
    fieldsets = UserAdmin.fieldsets + (
        ('معلومات إضافية', {'fields': ('mobile',)}),
    )
    add_fieldsets = UserAdmin.add_fieldsets + (
        ('معلومات إضافية', {'fields': ('mobile',)}),
    )
    list_display = ['username', 'email', 'get_full_name', 'mobile', 'is_staff', 'is_active']
    search_fields = ['username', 'email', 'first_name', 'last_name', 'mobile']


# ============================================================
# ADMIN SITE CUSTOMIZATION
# ============================================================

admin.site.site_header = 'نظام المحاسبة المتكامل - Integrated Accounting System'
admin.site.site_title = 'النظام المحاسبي'
admin.site.index_title = 'لوحة التحكم الإدارية'