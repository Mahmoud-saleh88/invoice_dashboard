from django.db import models
from django.db.models import Sum
from django.contrib.auth.models import AbstractUser
import uuid
from decimal import Decimal, ROUND_HALF_UP
from django.db import models
from django.contrib.auth.models import AbstractUser
from django.core.validators import MinValueValidator
from django.utils import timezone


# ============================================================
# CHART OF ACCOUNTS - دليل الحسابات
# ============================================================

class AccountCategory(models.Model):
    NORMAL_BALANCE = (
        ('debit',  'مدين  - Debit'),
        ('credit', 'دائن - Credit'),
    )

    code = models.CharField(max_length=2, unique=True, verbose_name="كود الفئة")
    name_ar = models.CharField(max_length=100, verbose_name="اسم الفئة (عربي)")
    name_en = models.CharField(max_length=100, verbose_name="اسم الفئة (إنجليزي)")
    normal_balance = models.CharField(max_length=6, choices=NORMAL_BALANCE, verbose_name="الطبيعة")

    class Meta:
        verbose_name = "فئة رئيسية"
        verbose_name_plural = "الفئات الرئيسية"
        ordering = ['code']

    def __str__(self):
        return f"{self.code} - {self.name_ar}"


class AccountGroup(models.Model):
    category = models.ForeignKey(AccountCategory, on_delete=models.CASCADE,related_name='groups', verbose_name="الفئة الرئيسية")
    code = models.CharField(max_length=4, unique=True, verbose_name="كود المجموعة")
    name_ar = models.CharField(max_length=100, verbose_name="اسم المجموعة (عربي)")
    name_en = models.CharField(max_length=100, verbose_name="اسم المجموعة (إنجليزي)")
    normal_balance = models.CharField(max_length=6, choices=AccountCategory.NORMAL_BALANCE, verbose_name="الطبيعة")

    class Meta:
        verbose_name = "مجموعة حسابات"
        verbose_name_plural = "مجموعات الحسابات"
        ordering = ['code']

    def __str__(self):
        return f"{self.code} - {self.name_ar}"


class Account(models.Model):
    ACCOUNT_TYPES = (
        ('header',  'رأسي - Header'),
        ('detail',  'تفصيلي - Detail'),
    )

    group = models.ForeignKey(AccountGroup, on_delete=models.CASCADE,related_name='accounts', verbose_name="المجموعة")
    parent= models.ForeignKey('self', null=True, blank=True,on_delete=models.CASCADE, related_name='children',verbose_name="الحساب الأب")
    code = models.CharField(max_length=20, unique=True, verbose_name="كود الحساب")
    name_ar = models.CharField(max_length=200, verbose_name="اسم الحساب (عربي)")
    name_en = models.CharField(max_length=200, verbose_name="اسم الحساب (إنجليزي)")
    account_type = models.CharField(max_length=10, choices=ACCOUNT_TYPES, default='detail', verbose_name="نوع الحساب")
    normal_balance = models.CharField(max_length=6, choices=AccountCategory.NORMAL_BALANCE,verbose_name="الطبيعة")
    is_active = models.BooleanField(default=True, verbose_name="نشط")
    notes = models.TextField(blank=True, verbose_name="ملاحظات")
    opening_balance = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="الرصيد الافتتاحي")
    opening_balance_date = models.DateField(null=True, blank=True,verbose_name="تاريخ الرصيد الافتتاحي")

    class Meta:
        verbose_name = "حساب"
        verbose_name_plural = "الحسابات"
        ordering = ['code']
        permissions = [
            ('export_account_statement', 'Can export account statement'),
            ('view_account_balances',    'Can view account balances'),
        ]

    def __str__(self):
        return f"{self.code} - {self.name_ar}"

    def get_balance(self, date_from=None, date_to=None):
        from django.db.models import Sum, Q
        qs = JournalEntryLine.objects.filter(account=self,journal_entry__status='posted',)
        if date_from:
            qs = qs.filter(journal_entry__date__gte=date_from)
        if date_to:
            qs = qs.filter(journal_entry__date__lte=date_to)

        agg = qs.aggregate(
            total_debit=Sum('debit_amount'),
            total_credit=Sum('credit_amount'),
        )
        debit  = agg['total_debit']  or Decimal('0')
        credit = agg['total_credit'] or Decimal('0')
        opening = self.opening_balance if self.opening_balance else Decimal('0')

        if self.normal_balance == 'debit':
            return opening + debit - credit
        else:
            return opening + credit - debit

    def get_all_children(self):
        result = list(self.children.all())
        for child in list(result):
            result.extend(child.get_all_children())
        return result


# ============================================================
# JOURNAL ENTRIES - القيود المحاسبية
# ============================================================

class JournalEntry(models.Model):
    STATUS_CHOICES = (
        ('draft',    'مسودة - Draft'),
        ('posted',   'مرحّل - Posted'),
        ('reversed', 'معكوس - Reversed'),
        ('cancelled','ملغي  - Cancelled'),
    )
    ENTRY_TYPES = (
        ('manual',         'يدوي - Manual'),
        ('sales_invoice',  'فاتورة مبيعات'),
        ('purchase_invoice','فاتورة مشتريات'),
        ('payment',        'دفعة'),
        ('receipt',        'قبض'),
        ('depreciation',   'استهلاك'),
        ('adjustment',     'تسوية'),
        ('opening',        'قيد افتتاحي'),
    )

    number = models.CharField(max_length=30, unique=True, verbose_name="رقم القيد")
    date = models.DateField(verbose_name="تاريخ القيد")
    entry_type = models.CharField(max_length=30, choices=ENTRY_TYPES,default='manual', verbose_name="نوع القيد")
    status = models.CharField(max_length=10, choices=STATUS_CHOICES,default='draft', verbose_name="الحالة")
    description = models.TextField(verbose_name="البيان / الوصف")
    reference = models.CharField(max_length=100, blank=True, verbose_name="المرجع")
    sales_invoice = models.ForeignKey('SalesInvoice', null=True, blank=True,on_delete=models.SET_NULL,related_name='journal_entries', verbose_name="فاتورة المبيعات")
    purchase_invoice = models.ForeignKey('PurchaseInvoice', null=True, blank=True,on_delete=models.SET_NULL,related_name='journal_entries', verbose_name="فاتورة المشتريات")
    reversed_entry = models.OneToOneField('self', null=True, blank=True,on_delete=models.SET_NULL,related_name='reversal_of', verbose_name="عكس القيد")
    created_by = models.ForeignKey('CustomUser', on_delete=models.PROTECT,verbose_name="أنشأ بواسطة")
    created_at = models.DateTimeField(auto_now_add=True)
    posted_at = models.DateTimeField(null=True, blank=True)
    notes = models.TextField(blank=True, verbose_name="ملاحظات")

    class Meta:
        verbose_name = "قيد يومي"
        verbose_name_plural = "القيود اليومية"
        ordering = ['-date', '-number']
        permissions = [
            ('post_journalentry',    'Can post journal entries'),
            ('reverse_journalentry', 'Can reverse journal entries'),
            ('export_journal',       'Can export journal to Excel/PDF'),
        ]

    def __str__(self):
        return f"{self.number} - {self.date} - {self.description[:50]}"

    @property
    def total_debit(self):
        return self.lines.aggregate(t=models.Sum('debit_amount'))['t'] or Decimal('0')

    @property
    def total_credit(self):
        return self.lines.aggregate(t=models.Sum('credit_amount'))['t'] or Decimal('0')

    @property
    def is_balanced(self):
        return self.total_debit == self.total_credit

    def clean(self):
        from django.core.exceptions import ValidationError
        if not self.date:
            return
        closed_year = FiscalYear.objects.filter(
            is_closed=True,
            start_date__lte=self.date,
            end_date__gte=self.date
        ).first()
        if closed_year:
            raise ValidationError({
                'date': f'لا يمكن إنشاء أو تعديل قيود في تاريخ {self.date}. '
                        f'السنة المالية "{closed_year.name}" مغلقة.'
            })

    def save(self, *args, **kwargs):
        skip_validation = kwargs.pop('skip_validation', False)
        if not skip_validation:
            try:
                self.full_clean(exclude=['created_by', 'sales_invoice', 'purchase_invoice'])
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning(f'JournalEntry validation warning: {e}')
        super().save(*args, **kwargs)
        
    
    def post(self, user):
        """Post the journal entry with fiscal year validation"""
        from django.core.exceptions import ValidationError
        
        if not self.is_balanced:
            raise ValueError("القيد غير متوازن: مجموع المدين لا يساوي مجموع الدائن")
        
        # Check fiscal year is not closed
        closed_year = FiscalYear.objects.filter(
            is_closed=True,
            start_date__lte=self.date,
            end_date__gte=self.date
        ).first()
        
        if closed_year:
            raise ValidationError(f'لا يمكن ترحيل القيد. السنة المالية "{closed_year.name}" مغلقة.')
        
        self.status = 'posted'
        self.posted_at = timezone.now()
        self.save()

    def reverse(self, user, date=None, description=None):
        if self.status != 'posted':
            raise ValueError("يمكن عكس القيود المرحّلة فقط")
        rev = JournalEntry.objects.create(
            number=f"REV-{self.number}",
            date=date or timezone.now().date(),
            entry_type=self.entry_type,
            status='draft',
            description=description or f"عكس القيد: {self.description}",
            reference=self.reference,
            reversed_entry=self,
            created_by=user,
        )
        for line in self.lines.all():
            JournalEntryLine.objects.create(
                journal_entry=rev,
                account=line.account,
                description=line.description,
                debit_amount=line.credit_amount,
                credit_amount=line.debit_amount,
                cost_center=line.cost_center,
            )
        self.status = 'reversed'
        self.save()
        return rev
    
    
class JournalEntryLine(models.Model):
    journal_entry = models.ForeignKey(JournalEntry, on_delete=models.CASCADE,related_name='lines', verbose_name="القيد")
    account = models.ForeignKey(Account, on_delete=models.PROTECT,verbose_name="الحساب")
    description = models.CharField(max_length=300, blank=True, verbose_name="البيان")
    debit_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0,validators=[MinValueValidator(0)], verbose_name="مدين")
    credit_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0,validators=[MinValueValidator(0)], verbose_name="دائن")
    cost_center   = models.ForeignKey('CostCenter', null=True, blank=True,on_delete=models.SET_NULL, verbose_name="مركز التكلفة")

    class Meta:
        verbose_name = "سطر قيد"
        verbose_name_plural = "أسطر القيود"

    def __str__(self):
        return f"{self.account} | D:{self.debit_amount} C:{self.credit_amount}"

    def clean(self):
        from django.core.exceptions import ValidationError
        if self.debit_amount > 0 and self.credit_amount > 0:
            raise ValidationError("لا يمكن أن يكون السطر مدينًا ودائنًا في نفس الوقت")
        if self.debit_amount == 0 and self.credit_amount == 0:
            raise ValidationError("يجب أن يكون السطر مدينًا أو دائنًا")
        if self.account.account_type == 'header':
            raise ValidationError("لا يمكن القيد على حساب رأسي - استخدم الحسابات التفصيلية فقط")


# ============================================================
# COST CENTERS - مراكز التكلفة
# ============================================================

class CostCenter(models.Model):
    code = models.CharField(max_length=20, unique=True, verbose_name="الكود")
    name_ar = models.CharField(max_length=200, verbose_name="الاسم (عربي)")
    name_en = models.CharField(max_length=200, verbose_name="الاسم (إنجليزي)")
    parent = models.ForeignKey('self', null=True, blank=True,on_delete=models.CASCADE, related_name='children',verbose_name="المركز الأب")
    is_active = models.BooleanField(default=True, verbose_name="نشط")

    class Meta:
        verbose_name = "مركز تكلفة"
        verbose_name_plural = "مراكز التكلفة"

    def __str__(self):
        return f"{self.code} - {self.name_ar}"


# ============================================================
# COMPANY - بيانات الشركة
# ============================================================

class Company(models.Model):
    name = models.CharField(max_length=200, verbose_name="اسم الشركة")
    name_en = models.CharField(max_length=200, blank=True, verbose_name="الاسم الإنجليزي")
    street_name = models.CharField(max_length=200, verbose_name="اسم الشارع (BT-35)")
    building_number = models.CharField(max_length=4, verbose_name="رقم المبنى (KSA-17) — 4 أرقام")
    district = models.CharField(max_length=100, verbose_name="الحي (KSA-3)")
    city = models.CharField(max_length=100, verbose_name="المدينة (BT-37)")
    postal_code = models.CharField(max_length=5, verbose_name="الرمز البريدي (BT-38) — 5 أرقام")
    vat_number = models.CharField(max_length=15, verbose_name="الرقم الضريبي — 15 رقماً")
    cr_number = models.CharField(max_length=50, blank=True, verbose_name="السجل التجاري")
    email = models.EmailField(blank=True, verbose_name="البريد الإلكتروني")
    phone = models.CharField(max_length=20, blank=True, verbose_name="الهاتف")
    logo = models.ImageField(upload_to='company/', blank=True, verbose_name="الشعار")
    zatca_csid = models.TextField(blank=True, verbose_name="CSID زاتكا (legacy)")
    zatca_private_key = models.TextField(blank=True, verbose_name="المفتاح الخاص (legacy)")
    zatca_certificate = models.TextField(blank=True, verbose_name="الشهادة (legacy)")
    zatca_registered = models.BooleanField(default=False, verbose_name="مسجل في زاتكا")

    class Meta:
        verbose_name = "شركة"
        verbose_name_plural = "الشركات"

    def __str__(self):
        return self.name

    def clean(self):
        from django.core.exceptions import ValidationError
        errors = {}
        vat = (self.vat_number or '').strip()
        if not (len(vat) == 15 and vat.isdigit()):
            errors['vat_number'] = "الرقم الضريبي يجب أن يكون 15 رقماً"
        bld = ''.join(c for c in (self.building_number or '') if c.isdigit())
        if len(bld) != 4:
            errors['building_number'] = "رقم المبنى يجب أن يكون 4 أرقام بالضبط"
        pz = ''.join(c for c in (self.postal_code or '') if c.isdigit())
        if len(pz) != 5:
            errors['postal_code'] = "الرمز البريدي يجب أن يكون 5 أرقام بالضبط"
        if errors:
            raise ValidationError(errors)


# ============================================================
# CLIENT - العملاء
# ============================================================

class Client(models.Model):
    name = models.CharField(max_length=200, verbose_name="اسم العميل")
    name_en = models.CharField(max_length=200, blank=True, verbose_name="الاسم الإنجليزي")
    street_name = models.CharField(max_length=200, blank=True, verbose_name="اسم الشارع (BT-50)")
    building_number = models.CharField(max_length=20, blank=True, verbose_name="رقم المبنى (KSA-18)")
    district = models.CharField(max_length=100, blank=True, verbose_name="الحي (KSA-4)")
    city = models.CharField(max_length=100, blank=True, verbose_name="المدينة (BT-52)")
    country = models.CharField(max_length=100, default="Saudi Arabia", verbose_name="الدولة")
    postal_code = models.CharField(max_length=20, blank=True, verbose_name="الرمز البريدي (BT-53)")
    vat_number = models.CharField(max_length=50, blank=True, verbose_name="الرقم الضريبي")
    cr_number = models.CharField(max_length=50, blank=True, verbose_name="السجل التجاري")
    email = models.EmailField(blank=True, verbose_name="البريد الإلكتروني")
    phone = models.CharField(max_length=20, blank=True, verbose_name="الهاتف")
    ar_account = models.ForeignKey(Account, null=True, blank=True,on_delete=models.SET_NULL,verbose_name="حساب المدين (العميل)",limit_choices_to={'account_type': 'detail'})
    credit_limit = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="حد الائتمان")
    is_active   = models.BooleanField(default=True, verbose_name="نشط")

    class Meta:
        verbose_name = "عميل"
        verbose_name_plural = "العملاء"

    def __str__(self):
        return self.name

    @property
    def effective_street_name(self):
        """Returns street_name if set."""
        return (self.street_name or "").strip()


# ============================================================
# SUPPLIER - الموردون
# ============================================================

class Supplier(models.Model):
    name = models.CharField(max_length=200, verbose_name="اسم المورد")
    name_en = models.CharField(max_length=200, blank=True, verbose_name="الاسم الإنجليزي")
    street_name = models.CharField(max_length=200, blank=True, verbose_name="اسم الشارع")
    building_number = models.CharField(max_length=20, blank=True, verbose_name="رقم المبنى")
    district = models.CharField(max_length=100, blank=True, verbose_name="الحي")
    city = models.CharField(max_length=100, blank=True, verbose_name="المدينة")
    postal_code = models.CharField(max_length=20, blank=True, verbose_name="الرمز البريدي")
    vat_number = models.CharField(max_length=50, blank=True, verbose_name="الرقم الضريبي")
    cr_number = models.CharField(max_length=50, blank=True, verbose_name="السجل التجاري")
    email = models.EmailField(blank=True, verbose_name="البريد الإلكتروني")
    phone = models.CharField(max_length=20, blank=True, verbose_name="الهاتف")
    ap_account = models.ForeignKey(Account, null=True, blank=True,on_delete=models.SET_NULL,verbose_name="حساب الدائن (المورد)",limit_choices_to={'account_type': 'detail'})
    is_active   = models.BooleanField(default=True, verbose_name="نشط")

    class Meta:
        verbose_name = "مورد"
        verbose_name_plural = "الموردون"

    def __str__(self):
        return self.name


# ============================================================
# SALES INVOICE - فاتورة المبيعات 
# ============================================================

class SalesInvoice(models.Model):
    TRANSACTION_TYPES = (
        ("invoice",       "فاتورة ضريبية"),
        ("simplified",    "فاتورة مبسطة"),
        ("credit",        "إشعار دائن"),
        ("debit",         "إشعار مدين"),
    )
    ZATCA_STATUS = (
        ("pending",   "في الانتظار"),
        ("submitted", "مُرسلة"),
        ("accepted",  "مقبولة"),
        ("rejected",  "مرفوضة"),
        ("cleared",   "مُخلّصة"),
        ("reported",  "مُبلّغ عنها"),
    )
    PAYMENT_METHODS = (
        ("cash",     "نقداً"),
        ("credit",   "آجل"),
        ("bank",     "تحويل بنكي"),
        ("cheque",   "شيك"),
        ("card",     "بطاقة"),
    )

    invoice_number = models.CharField(max_length=50, unique=True, verbose_name="رقم الفاتورة")
    invoice_uuid = models.UUIDField(default=uuid.uuid4, unique=True,editable=False, verbose_name="UUID الفاتورة")
    date = models.DateField(verbose_name="تاريخ الفاتورة")
    supply_date = models.DateField(null=True, blank=True, verbose_name="تاريخ التوريد")
    due_date = models.DateField(null=True, blank=True, verbose_name="تاريخ الاستحقاق")
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES,default="invoice", verbose_name="نوع المعاملة")
    payment_method = models.CharField(max_length=10, choices=PAYMENT_METHODS,default="credit", verbose_name="طريقة الدفع")
    company = models.ForeignKey(Company, on_delete=models.CASCADE, verbose_name="الشركة")
    client = models.ForeignKey(Client,  on_delete=models.CASCADE, verbose_name="العميل")
    reference_invoice = models.ForeignKey('self', null=True, blank=True,on_delete=models.SET_NULL,verbose_name="الفاتورة المرجعية")
    total_quantity = models.PositiveIntegerField(default=0, verbose_name="إجمالي الكمية")
    total_sales_excl_tax = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="إجمالي بدون ضريبة")
    discount_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="قيمة الخصم")
    taxable_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="الوعاء الضريبي")
    tax_rate = models.DecimalField(max_digits=5, decimal_places=2, default=15, verbose_name="نسبة الضريبة %")
    total_tax = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="إجمالي الضريبة")
    total = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="المجموع الكلي")
    amount_paid = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="المبلغ المدفوع")
    zatca_status = models.CharField(max_length=20, choices=ZATCA_STATUS,default="pending", verbose_name="حالة زاتكا")
    zatca_submission_id = models.CharField(max_length=200, blank=True,verbose_name="معرف الإرسال (ZATCA)")
    zatca_invoice_hash = models.CharField(max_length=500, blank=True,verbose_name="Hash الفاتورة")
    zatca_qr_code = models.TextField(blank=True, verbose_name="QR Code")
    zatca_xml = models.TextField(blank=True, verbose_name="XML الفاتورة")
    zatca_signed_xml = models.TextField(blank=True, verbose_name="XML الموقّع")
    zatca_clearance_status = models.TextField(blank=True, verbose_name="حالة التخليص")
    zatca_submitted_at = models.DateTimeField(null=True, blank=True,verbose_name="وقت الإرسال لزاتكا")
    zatca_response = models.JSONField(null=True, blank=True, verbose_name="استجابة زاتكا")
    invoice_counter_value = models.PositiveIntegerField(default=0,verbose_name="رقم الفاتورة التسلسلي (KSA-16)",help_text="يُعيَّن تلقائياً عند إرسال الفاتورة لزاتكا. لا تعدّل يدوياً.",)
    previous_invoice_hash   = models.CharField(max_length=500,blank=True,verbose_name="Hash الفاتورة السابقة (KSA-13)",help_text="يُعيَّن تلقائياً عند إرسال الفاتورة لزاتكا. لا تعدّل يدوياً.",)
    created_by  = models.CharField(max_length=100, verbose_name="أنشأ بواسطة")
    created_at  = models.DateTimeField(auto_now_add=True)
    notes       = models.TextField(blank=True, verbose_name="ملاحظات")

    class Meta:
        verbose_name = "فاتورة مبيعات"
        verbose_name_plural = "فواتير المبيعات"
        ordering = ['-date', '-invoice_number']
        permissions = [
            ('exportexcel_salesinvoice',  'Can export sales invoices to Excel'),
            ('view_salesinvoice_pdf',     'Can view sales invoice PDF'),
            ('exportpdf_salesinvoice',    'Can export sales invoice to PDF'),
            ('submit_zatca',              'Can submit invoice to ZATCA'),
        ]

    def calculate_totals_from_items(self):
        items = self.items.all()
        total_qty = sum(item.quantity for item in items)
        total_gross = sum(item.total for item in items)
        total_disc = sum(item.discount_amount for item in items)
        return total_qty, total_gross, total_disc

    def save(self, *args, **kwargs):
        skip_validation = kwargs.pop('skip_validation', False)
        
        if self.pk:
            try:
                qty, gross, disc = self.calculate_totals_from_items()
                self.total_quantity = int(qty)
                self.total_sales_excl_tax = Decimal(str(gross)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                self.discount_amount = Decimal(str(disc)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            except Exception:
                pass

        total_sales = Decimal(str(self.total_sales_excl_tax or 0))
        discount = Decimal(str(self.discount_amount or 0))
        tax_rate = Decimal(str(self.tax_rate or 15))

        self.taxable_amount = (total_sales - discount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        self.total_tax = (self.taxable_amount * tax_rate / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        self.total = (self.taxable_amount + self.total_tax).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        if not skip_validation:
            try:
                self.full_clean(exclude=['invoice_uuid', 'company', 'client'])
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning(f'SalesInvoice validation warning: {e}')

        super().save(*args, **kwargs)

    def clean(self):
        from django.core.exceptions import ValidationError
        if not self.date:
            return
        closed_year = FiscalYear.objects.filter(
            is_closed=True,
            start_date__lte=self.date,
            end_date__gte=self.date
        ).first()
        if closed_year:
            raise ValidationError({
                'date': f'لا يمكن إنشاء فاتورة في تاريخ {self.date}. '
                        f'السنة المالية "{closed_year.name}" مغلقة.'
            })

    
    def __str__(self):
        return f"{self.invoice_number} - {self.client}"


class SalesInvoiceItem(models.Model):
    invoice = models.ForeignKey(SalesInvoice, related_name="items", on_delete=models.CASCADE, verbose_name="الفاتورة")
    description = models.TextField(verbose_name="تفاصيل السلعة أو الخدمة")
    unit = models.CharField(max_length=50, blank=True, verbose_name="الوحدة")
    quantity = models.DecimalField(max_digits=10, decimal_places=3, default=1, verbose_name="الكمية")
    unit_price = models.DecimalField(max_digits=18, decimal_places=2, verbose_name="سعر الوحدة")
    discount_percent = models.DecimalField(max_digits=5, decimal_places=2, default=0, verbose_name="نسبة الخصم %")
    discount_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="مبلغ الخصم")
    total = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="الإجمالي بدون ضريبة")
    tax_rate = models.DecimalField(max_digits=5, decimal_places=2, default=15, verbose_name="نسبة الضريبة %")
    tax_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="قيمة الضريبة")
    total_with_tax = models.DecimalField(max_digits=18, decimal_places=2, default=0, verbose_name="الإجمالي شامل الضريبة")
    account = models.ForeignKey(Account, null=True, blank=True, on_delete=models.SET_NULL, verbose_name="حساب الإيراد", limit_choices_to={'account_type': 'detail'})

    class Meta:
        verbose_name = "بند الفاتورة"
        verbose_name_plural = "بنود الفاتورة"

    def save(self, *args, **kwargs):
        quantity = Decimal(str(self.quantity))
        unit_price = Decimal(str(self.unit_price))
        discount_pct = Decimal(str(self.discount_percent or 0))
        discount_amt = Decimal(str(self.discount_amount or 0))
        tax_rate = Decimal(str(self.tax_rate or 15))

        gross = (quantity * unit_price).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        disc  = ((gross * discount_pct / Decimal('100')) + discount_amt).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        self.total = (gross - disc).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        self.tax_amount = (self.total * tax_rate / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        self.total_with_tax = (self.total + self.tax_amount).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        super().save(*args, **kwargs)
        self.invoice.save(skip_validation=True)

    def delete(self, *args, **kwargs):
        invoice = self.invoice
        super().delete(*args, **kwargs)
        invoice.save(skip_validation=True)


    def __str__(self):
        return f"{self.description[:50]} - {self.total}"


# ============================================================
# ZATCA INTEGRATION - تكامل زاتكا
# ============================================================

class ZATCADevice(models.Model):
    environment_choices = (
        ('sandbox', 'Sandbox'),
        ('simulation', 'Simulation'),
        ('production', 'Production'),   
    )
    company = models.ForeignKey(Company, on_delete=models.CASCADE,related_name='zatca_devices', verbose_name="الشركة")
    device_name = models.CharField(max_length=100, verbose_name="اسم الجهاز")
    serial_number = models.CharField(max_length=200, unique=True, verbose_name="الرقم التسلسلي")
    private_key = models.TextField(verbose_name="المفتاح الخاص")
    csr = models.TextField(blank=True, verbose_name="CSR")
    certificate = models.TextField(blank=True, verbose_name="الشهادة")
    csid = models.TextField(blank=True, verbose_name="CSID")
    secret = models.TextField(blank=True, verbose_name="Secret")
    environment = models.CharField(max_length=20, choices=environment_choices, default='sandbox', verbose_name="البيئة")
    is_active = models.BooleanField(default=True)
    registered_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "جهاز زاتكا"
        verbose_name_plural = "أجهزة زاتكا"

    def __str__(self):
        return f"{self.device_name} ({self.environment})"


class ZATCALog(models.Model):
    
    ACTION_TYPES = (
        ('onboarding',  'Onboarding'),
        ('renewal',     'Renewal'),
        ('clearance',   'Clearance'),
        ('reporting',   'Reporting'),
        ('compliance',  'Compliance Check'),
    )

    device = models.ForeignKey(ZATCADevice, on_delete=models.CASCADE,related_name='logs', verbose_name="الجهاز")
    sales_invoice = models.ForeignKey(SalesInvoice, null=True, blank=True,on_delete=models.SET_NULL, verbose_name="الفاتورة")
    action = models.CharField(max_length=20, choices=ACTION_TYPES)
    request_payload = models.TextField(blank=True, verbose_name="الـ XML المرسل (base64)")
    response_code = models.IntegerField(null=True, blank=True, verbose_name="كود الاستجابة")
    response_body = models.TextField(blank=True, verbose_name="استجابة زاتكا")
    success = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    error_message = models.TextField(blank=True)

    class Meta:
        verbose_name = "سجل زاتكا"
        verbose_name_plural = "سجلات زاتكا"
        ordering = ['-created_at']

    def __str__(self):
        return f"{self.action} - {self.created_at.strftime('%Y-%m-%d %H:%M')} - {'✓' if self.success else '✗'}"


# ============================================================
# PURCHASE INVOICE - فاتورة المشتريات
# ============================================================

class PurchaseInvoice(models.Model):
    TRANSACTION_TYPES = (
        ("invoice", "فاتورة"),
        ("credit",  "إشعار دائن"),
        ("debit",   "إشعار مدين"),
    )

    invoice_number = models.CharField(max_length=50, verbose_name="رقم الفاتورة")
    date = models.DateField(verbose_name="تاريخ الفاتورة")
    entry_date = models.DateField(default=timezone.now, verbose_name="تاريخ الإدخال")
    due_date = models.DateField(null=True, blank=True, verbose_name="تاريخ الاستحقاق")
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES,default="invoice", verbose_name="نوع المعاملة")
    supplier = models.ForeignKey(Supplier, on_delete=models.CASCADE,verbose_name="المورد")
    tax_number = models.CharField(max_length=50, blank=True,verbose_name="الرقم الضريبي للمورد")
    subtotal = models.DecimalField(max_digits=18, decimal_places=2,verbose_name="إجمالي بدون ضريبة")
    discount_amount = models.DecimalField(max_digits=18, decimal_places=2, default=0,null=True, blank=True, verbose_name="الخصم")
    tax_amount = models.DecimalField(max_digits=18, decimal_places=2,verbose_name="إجمالي الضريبة")
    total = models.DecimalField(max_digits=18, decimal_places=2,verbose_name="الإجمالي")
    amount_paid = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="المبلغ المدفوع")
    purchaseinvoice_attachments = models.FileField(upload_to='purchase_invoices/',blank=True, null=True,verbose_name="مرفقات الفاتورة")
    notes = models.TextField(blank=True, verbose_name="ملاحظات")
    created_by = models.CharField(max_length=100, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "فاتورة مشتريات"
        verbose_name_plural = "فواتير المشتريات"
        ordering = ['-date']
        permissions = [
            ('exportexcel_purchaseinvoice', 'Can export purchase invoices to Excel'),
        ]
        unique_together = [['supplier', 'invoice_number']]

    @property
    def balance_due(self):
        return self.total - self.amount_paid
    
    def clean(self):
        from django.core.exceptions import ValidationError
        if not self.date:
            return
        closed_year = FiscalYear.objects.filter(
            is_closed=True,
            start_date__lte=self.date,
            end_date__gte=self.date
        ).first()
        if closed_year:
            raise ValidationError({
                'date': f'لا يمكن إنشاء فاتورة مشتريات في تاريخ {self.date}. '
                        f'السنة المالية "{closed_year.name}" مغلقة.'
            })

    def save(self, *args, **kwargs):
        skip_validation = kwargs.pop('skip_validation', False)
        if not skip_validation:
            try:
                self.full_clean(exclude=['supplier'])
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning(f'PurchaseInvoice validation warning: {e}')
        super().save(*args, **kwargs)
        
    def __str__(self):
        return f"{self.invoice_number} - {self.supplier}"


# ============================================================
# FISCAL YEAR & ACCOUNTING PERIOD
# ============================================================

class FiscalYear(models.Model):
    company = models.ForeignKey(Company, on_delete=models.CASCADE)
    name = models.CharField(max_length=50, verbose_name="اسم السنة المالية")
    start_date = models.DateField(verbose_name="تاريخ البداية")
    end_date = models.DateField(verbose_name="تاريخ النهاية")
    is_closed = models.BooleanField(default=False, verbose_name="مغلقة")
    closed_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "سنة مالية"
        verbose_name_plural = "السنوات المالية"
        permissions = [
            ('close_fiscalyear', 'Can close fiscal year'),
        ]

    def __str__(self):
        return f"{self.name} ({self.start_date} → {self.end_date})"
    def contains_date(self, date):
        """Check if a date is within this fiscal year"""
        return self.start_date <= date <= self.end_date
    
    @classmethod
    def validate_date_not_in_closed_year(cls, date, company=None):
        """Raise ValidationError if date is in a closed fiscal year"""
        from django.core.exceptions import ValidationError
        
        if company is None:
            company = Company.objects.first()
        
        closed_years = cls.objects.filter(
            company=company,
            is_closed=True,
            start_date__lte=date,
            end_date__gte=date
        )
        
        if closed_years.exists():
            year = closed_years.first()
            raise ValidationError(
                f'لا يمكن إجراء عمليات في تاريخ {date}. '
                f'السنة المالية "{year.name}" مغلقة.'
            )
    
    @classmethod
    def get_active_for_date(cls, date, company=None):
        """Get the active (open) fiscal year containing the date"""
        if company is None:
            company = Company.objects.first()
        
        return cls.objects.filter(
            company=company,
            is_closed=False,
            start_date__lte=date,
            end_date__gte=date
        ).first()

class AccountingPeriod(models.Model):
    fiscal_year = models.ForeignKey(FiscalYear, on_delete=models.CASCADE,related_name='periods', verbose_name="السنة المالية")
    name = models.CharField(max_length=50, verbose_name="اسم الفترة")
    start_date = models.DateField(verbose_name="تاريخ البداية")
    end_date = models.DateField(verbose_name="تاريخ النهاية")
    is_closed = models.BooleanField(default=False, verbose_name="مغلقة")

    class Meta:
        verbose_name = "فترة محاسبية"
        verbose_name_plural = "الفترات المحاسبية"

    def __str__(self):
        return f"{self.name}"

# ============================================================
# PAYMENT / RECEIPT
# ============================================================

class Payment(models.Model):
    PAYMENT_TYPES = (
        ('receipt', 'قبض - Receipt'),
        ('payment', 'صرف - Payment'),
    )
    PAYMENT_METHODS = (
        ('cash',   'نقداً'),
        ('bank',   'تحويل بنكي'),
        ('cheque', 'شيك'),
        ('card',   'بطاقة'),
    )

    number = models.CharField(max_length=30, unique=True)
    date = models.DateField(verbose_name="التاريخ")
    payment_type = models.CharField(max_length=10, choices=PAYMENT_TYPES)
    payment_method = models.CharField(max_length=10, choices=PAYMENT_METHODS, default='cash')
    amount = models.DecimalField(max_digits=18, decimal_places=2,validators=[MinValueValidator(Decimal('0.01'))])
    client = models.ForeignKey(Client, null=True, blank=True,on_delete=models.PROTECT, verbose_name="العميل")
    supplier = models.ForeignKey(Supplier, null=True, blank=True,on_delete=models.PROTECT, verbose_name="المورد")
    bank_account = models.ForeignKey(Account, null=True, blank=True,on_delete=models.PROTECT,related_name='payments_bank',verbose_name="الحساب البنكي")
    cash_account = models.ForeignKey(Account, null=True, blank=True,on_delete=models.PROTECT,related_name='payments_cash',verbose_name="حساب الصندوق")
    description = models.TextField(blank=True, verbose_name="البيان")
    reference = models.CharField(max_length=100, blank=True, verbose_name="رقم الشيك/المرجع")
    journal_entry = models.ForeignKey(JournalEntry, null=True, blank=True,on_delete=models.SET_NULL,verbose_name="القيد المحاسبي")
    created_by = models.ForeignKey('CustomUser', on_delete=models.PROTECT,verbose_name="أنشأ بواسطة")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "دفعة / قبض"
        verbose_name_plural = "المدفوعات والمقبوضات"
        ordering = ['-date']

    def __str__(self):
        return f"{self.number} - {self.payment_type} - {self.amount}"


class PaymentInvoiceAllocation(models.Model):
    payment = models.ForeignKey(Payment, on_delete=models.CASCADE,related_name='allocations')
    sales_invoice = models.ForeignKey(SalesInvoice, null=True, blank=True,on_delete=models.CASCADE)
    purchase_invoice= models.ForeignKey(PurchaseInvoice, null=True, blank=True,on_delete=models.CASCADE)
    allocated_amount= models.DecimalField(max_digits=18, decimal_places=2)

    class Meta:
        verbose_name = "تسوية دفعة"
        verbose_name_plural = "تسويات الدفعات"


# ============================================================
# FIXED ASSETS - الأصول الثابتة
# ============================================================

class AssetCategory(models.Model):
    name = models.CharField(max_length=100, verbose_name="الفئة")
    useful_life_years = models.PositiveIntegerField(default=5, verbose_name="العمر الإنتاجي (سنوات)")
    depreciation_rate = models.DecimalField(max_digits=5, decimal_places=2, default=20,verbose_name="نسبة الاستهلاك %")
    asset_account = models.ForeignKey(Account, on_delete=models.PROTECT,related_name='asset_categories',verbose_name="حساب الأصل")
    accum_dep_account = models.ForeignKey(Account, on_delete=models.PROTECT,related_name='accum_dep_categories',verbose_name="حساب مجمع الاستهلاك")
    dep_expense_account = models.ForeignKey(Account, on_delete=models.PROTECT,related_name='dep_expense_categories',verbose_name="حساب مصروف الاستهلاك")

    class Meta:
        verbose_name = "فئة أصل"
        verbose_name_plural = "فئات الأصول"

    def __str__(self):
        return self.name


class FixedAsset(models.Model):
    DEPRECIATION_METHODS = (
        ('straight_line',    'القسط الثابت'),
        ('declining_balance','الرصيد المتناقص'),
    )
    STATUS = (
        ('active',   'نشط'),
        ('disposed', 'مُتخلص منه'),
        ('scrapped', 'خُردة'),
    )

    code = models.CharField(max_length=30, unique=True, verbose_name="كود الأصل")
    name = models.CharField(max_length=200, verbose_name="اسم الأصل")
    category = models.ForeignKey(AssetCategory, on_delete=models.PROTECT,verbose_name="الفئة")
    purchase_date = models.DateField(verbose_name="تاريخ الشراء")
    purchase_cost = models.DecimalField(max_digits=18, decimal_places=2,verbose_name="تكلفة الشراء")
    salvage_value = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="القيمة التخريدية")
    useful_life_years   = models.PositiveIntegerField(verbose_name="العمر الإنتاجي (سنوات)")
    depreciation_method = models.CharField(max_length=20, choices=DEPRECIATION_METHODS,default='straight_line')
    accumulated_depreciation = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="مجمع الاستهلاك")
    net_book_value = models.DecimalField(max_digits=18, decimal_places=2, default=0,verbose_name="صافي القيمة الدفترية")
    status = models.CharField(max_length=10, choices=STATUS, default='active')
    location = models.CharField(max_length=200, blank=True, verbose_name="الموقع")
    notes = models.TextField(blank=True)

    class Meta:
        verbose_name = "أصل ثابت"
        verbose_name_plural = "الأصول الثابتة"

    def __str__(self):
        return f"{self.code} - {self.name}"

    def annual_depreciation(self):
        depreciable = self.purchase_cost - self.salvage_value
        if self.depreciation_method == 'straight_line':
            return depreciable / self.useful_life_years
        return Decimal('0')


# ============================================================
# CUSTOM USER MODEL
# ============================================================

class CustomUser(AbstractUser):
    mobile = models.CharField(max_length=20, blank=True, null=True,verbose_name="رقم الهاتف المحمول")

    class Meta:
        verbose_name = "مستخدم"
        verbose_name_plural = "المستخدمون"

    def __str__(self):
        return self.get_full_name() or self.username