"""
forms.py — شركة عقارات سعودية
"""
from decimal import Decimal

from django import forms
from django.forms import inlineformset_factory, BaseInlineFormSet
from django.utils import timezone
from .models import (
    SalesInvoice, SalesInvoiceItem,
    PurchaseInvoice,
    Client, Supplier,
    Account, AccountCategory, AccountGroup,
    FiscalYear, JournalEntry, JournalEntryLine,
    CostCenter,
)

_cls = 'form-control'

# ─────────────────────────────────────────────────────────
# Sales Invoice
# ─────────────────────────────────────────────────────────



class SalesInvoiceItemForm(forms.ModelForm):
    """نموذج بند الفاتورة - مبسط بدون discount_percent و tax_rate"""
    
    class Meta:
        model = SalesInvoiceItem
        # ✅ فقط الحقول الأساسية - إزالة discount_percent و tax_rate
        fields = ['description', 'quantity', 'unit_price']
        widgets = {
            'description': forms.Textarea(attrs={
                'rows': 2, 
                'class': _cls,
                'placeholder': 'تفاصيل السلعة أو الخدمة'
            }),
            'quantity': forms.NumberInput(attrs={
                'class': f'{_cls} item-quantity',
                'step': '1',
                'min': '1',
                'value': '1'
            }),
            'unit_price': forms.NumberInput(attrs={
                'class': f'{_cls} item-unit-price',
                'step': '0.01',
                'min': '0',
                'value': '0.00'
            }),
        }
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # جعل الحقول مطلوبة
        self.fields['description'].required = True
        self.fields['quantity'].required = True
        self.fields['unit_price'].required = True


class SalesInvoiceForm(forms.ModelForm):
    class Meta:
        from .models import SalesInvoice
        model  = SalesInvoice
        fields = [
            'invoice_number', 'date', 'supply_date', 'due_date',
            'transaction_type', 'payment_method', 'client',
            'created_by', 'tax_rate', 'notes',
            # Hidden calculated fields
            'total_quantity', 'total_sales_excl_tax', 'discount_amount',
            'taxable_amount', 'total_tax', 'total', 'amount_paid',
        ]
        widgets = {
            'invoice_number': forms.TextInput(attrs={
                'class': _cls,
                'readonly': 'readonly',
                'style': 'background-color: #e9ecef;'
            }),
            'date': forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            # ✅ supply_date — REQUIRED by ZATCA BR-KSA-15
            'supply_date': forms.DateInput(attrs={
                'type': 'date',
                'class': _cls,
                'placeholder': 'تاريخ التوريد (مطلوب للفواتير الضريبية)',
            }),
            # ✅ due_date — useful for credit management
            'due_date': forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'transaction_type': forms.Select(attrs={'class': _cls}),
            # ✅ payment_method — required for ZATCA PaymentMeans
            'payment_method': forms.Select(attrs={'class': _cls}),
            'client': forms.Select(attrs={'class': _cls}),
            'created_by': forms.TextInput(attrs={
                'class': _cls,
                'readonly': 'readonly',
                'style': 'background-color: #e9ecef;'
            }),
            'tax_rate': forms.NumberInput(attrs={
                'class': _cls,
                'step': '0.01',
                'min': '0',
                'id': 'id_tax_rate'
            }),
            'notes': forms.Textarea(attrs={'class': _cls, 'rows': 2}),
            # Hidden calculated fields
            'total_quantity':       forms.NumberInput(attrs={'type': 'hidden'}),
            'total_sales_excl_tax': forms.NumberInput(attrs={'type': 'hidden'}),
            'discount_amount':      forms.NumberInput(attrs={'type': 'hidden'}),
            'taxable_amount':       forms.NumberInput(attrs={'type': 'hidden'}),
            'total_tax':            forms.NumberInput(attrs={'type': 'hidden'}),
            'total':                forms.NumberInput(attrs={'type': 'hidden'}),
            'amount_paid':          forms.NumberInput(attrs={'type': 'hidden'}),
        }
 
    def __init__(self, *args, **kwargs):
        self.user = kwargs.pop('user', None)
        super().__init__(*args, **kwargs)
 
        # Make calculated fields optional
        for field_name in [
            'total_quantity', 'total_sales_excl_tax', 'discount_amount',
            'taxable_amount', 'total_tax', 'total', 'amount_paid',
        ]:
            self.fields[field_name].required = False
 
        # supply_date and due_date are optional in model but recommended
        self.fields['supply_date'].required = False
        self.fields['due_date'].required = False
 
        # New invoice defaults
        if not self.instance.pk:
            today = timezone.now().date()
            self.fields['invoice_number'].initial  = self._next_number()
            self.fields['date'].initial            = today
            self.fields['supply_date'].initial     = today   # ← default to today
            self.fields['tax_rate'].initial        = Decimal('15')
            self.fields['payment_method'].initial  = 'credit'
            self.fields['transaction_type'].initial = 'invoice'
            self.fields['discount_amount'].initial  = Decimal('0')
            self.fields['taxable_amount'].initial   = Decimal('0')
            self.fields['amount_paid'].initial      = Decimal('0')
            self.fields['total_quantity'].initial   = 0
            self.fields['total_sales_excl_tax'].initial = Decimal('0')
            self.fields['total_tax'].initial        = Decimal('0')
            self.fields['total'].initial            = Decimal('0')
 
            if self.user:
                self.fields['created_by'].initial = (
                    self.user.get_full_name() or self.user.username
                )
 
    @staticmethod
    def _next_number():
        from invoices.models import SalesInvoice
        last = SalesInvoice.objects.order_by('-id').first()
        if last and last.invoice_number:
            try:
                inv_str = str(last.invoice_number)
                if inv_str.startswith('INV-'):
                    return f"INV-{int(inv_str.split('-')[1]) + 1:04d}"
                elif inv_str.isdigit():
                    return str(int(inv_str) + 1)
                else:
                    return f"INV-{last.id + 1:04d}"
            except (ValueError, TypeError):
                return "INV-0001"
        return "INV-0001"
 
    def clean(self):
        cleaned_data = super().clean()
        date       = cleaned_data.get('date')
        supply_date = cleaned_data.get('supply_date')
        due_date   = cleaned_data.get('due_date')
        transaction_type = cleaned_data.get('transaction_type')

        if date and due_date and due_date < date:
            raise forms.ValidationError('تاريخ الاستحقاق لا يمكن أن يكون قبل تاريخ الفاتورة')
 
        tax_rate = cleaned_data.get('tax_rate')
        if tax_rate and tax_rate < 0:
            raise forms.ValidationError('نسبة الضريبة لا يمكن أن تكون سالبة')

        # supply_date مطلوب للفواتير الضريبية
        if transaction_type == 'invoice' and not supply_date:
            self.add_error('supply_date', 'تاريخ التوريد مطلوب للفواتير الضريبية')
        
        # إذا لم يتم تعيين supply_date، استخدم تاريخ الفاتورة
        if transaction_type == 'invoice' and not supply_date and date:
            cleaned_data['supply_date'] = date
 
        # Default calculated fields
        for field, default in {
            'discount_amount':      Decimal('0'),
            'taxable_amount':       Decimal('0'),
            'amount_paid':          Decimal('0'),
            'total_quantity':       0,
            'total_sales_excl_tax': Decimal('0'),
            'total_tax':            Decimal('0'),
            'total':                Decimal('0'),
        }.items():
            if not cleaned_data.get(field):
                cleaned_data[field] = default
 
        return cleaned_data
 

def get_sales_invoice_item_formset(instance=None, extra=None):
    _extra = extra if extra is not None else (0 if (instance and instance.pk) else 1)
    ItemFormSet = inlineformset_factory(
        SalesInvoice,
        SalesInvoiceItem,
        form=SalesInvoiceItemForm,
        fields=['description', 'quantity', 'unit_price'],
        extra=_extra,
        can_delete=True,
        min_num=0,
        validate_min=False
    )
    if instance and instance.pk:
        return ItemFormSet(instance=instance)
    return ItemFormSet()
# ─────────────────────────────────────────────────────────
# Purchase Invoice
# ─────────────────────────────────────────────────────────

class PurchaseInvoiceForm(forms.ModelForm):
    class Meta:
        model  = PurchaseInvoice
        fields = [
            'invoice_number', 'date', 'entry_date', 'due_date',
            'transaction_type', 'supplier', 'tax_number',
            'subtotal', 'discount_amount', 'tax_amount', 'total', 'amount_paid',
            'notes', 'purchaseinvoice_attachments',
        ]
        widgets = {
            'date':           forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'entry_date':     forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'due_date':       forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'supplier':       forms.Select(attrs={'class': _cls}),
            'transaction_type': forms.Select(attrs={'class': _cls}),
            'subtotal':       forms.NumberInput(attrs={'step': '0.01', 'min': '0', 'class': _cls}),
            # 'discount_amount': forms.NumberInput(attrs={'step': '0.01', 'min': '0', 'class': _cls}),
            'tax_amount':     forms.NumberInput(attrs={'step': '0.01', 'min': '0', 'class': _cls}),
            'total':          forms.NumberInput(attrs={'step': '0.01', 'min': '0', 'class': _cls, 'readonly': 'readonly'}),
            'amount_paid':    forms.NumberInput(attrs={'step': '0.01', 'min': '0', 'class': _cls}),
            'notes':          forms.Textarea(attrs={'rows': 2, 'class': _cls}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field in self.fields.values():
            field.widget.attrs.setdefault('class', _cls)
        self.fields['total'].widget.attrs.update({
            'readonly': True,
            'style': 'background:#e9ecef;cursor:not-allowed;font-weight:bold',
        })
        for fname in ('subtotal', 'tax_amount', 'total', 'amount_paid'):
            if self.instance and self.instance.pk:
                val = getattr(self.instance, fname, None)
                if val is not None:
                    self.initial[fname] = f'{float(val):.2f}'


# ─────────────────────────────────────────────────────────
# Client
# ─────────────────────────────────────────────────────────

class ClientForm(forms.ModelForm):
    class Meta:
        from .models import Client
        model  = Client
        fields = [
            'name', 'name_en',
            # ZATCA address fields (added by migration 0002)
            'street_name', 'building_number', 'district',
            'city', 'postal_code',
            # Other fields
            'vat_number', 'cr_number',
            'email', 'phone', 'credit_limit',
        ]
        widgets = {
            'name':            forms.TextInput(attrs={'class': _cls, 'placeholder': 'اسم العميل'}),
            'name_en':         forms.TextInput(attrs={'class': _cls, 'placeholder': 'Client Name (English)'}),
            # ✅ ZATCA address fields
            'street_name':     forms.TextInput(attrs={
                'class': _cls,
                'placeholder': 'اسم الشارع — مطلوب لزاتكا (BR-KSA-63)',
            }),
            'building_number': forms.TextInput(attrs={
                'class': _cls,
                'placeholder': '4 أرقام — مثل: 1234',
                'maxlength': '4',
            }),
            'district':        forms.TextInput(attrs={
                'class': _cls,
                'placeholder': 'الحي — مطلوب لزاتكا',
            }),
            'city':            forms.TextInput(attrs={'class': _cls, 'placeholder': 'المدينة'}),
            'postal_code':     forms.TextInput(attrs={
                'class': _cls,
                'placeholder': '5 أرقام — مثل: 12345',
                'maxlength': '5',
            }),
            'vat_number':      forms.TextInput(attrs={
                'class': _cls,
                'placeholder': '15 رقماً يبدأ وينتهي بـ 3',
            }),
            'cr_number':       forms.TextInput(attrs={'class': _cls}),
            'email':           forms.EmailInput(attrs={'class': _cls}),
            'phone':           forms.TextInput(attrs={'class': _cls}),
            'credit_limit':    forms.NumberInput(attrs={'class': _cls, 'step': '0.01'}),
        }
        error_messages = {
            'name': {'required': 'يرجى إدخال اسم العميل'},
        }
 
    def clean_name(self):
        from .models import Client
        name = ' '.join(self.cleaned_data.get('name', '').split())
        qs = Client.objects.filter(name__iexact=name)
        if self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise forms.ValidationError('هذا الاسم موجود بالفعل')
        return name
 
    def clean_vat_number(self):
        from .models import Client
        vat = self.cleaned_data.get('vat_number', '').strip()
        if vat:
            if len(vat) != 15 or not vat.isdigit() or not (vat.startswith('3') and vat.endswith('3')):
                raise forms.ValidationError('الرقم الضريبي يجب أن يكون 15 رقماً يبدأ وينتهي بـ 3')
            qs = Client.objects.filter(vat_number=vat)
            if self.instance.pk:
                qs = qs.exclude(pk=self.instance.pk)
            if qs.exists():
                raise forms.ValidationError('هذا الرقم الضريبي مسجل بالفعل')
        return vat
 
    def clean_building_number(self):
        bld = self.cleaned_data.get('building_number', '').strip()
        if bld and (not bld.isdigit() or len(bld) != 4):
            raise forms.ValidationError('رقم المبنى يجب أن يكون 4 أرقام بالضبط')
        return bld
 
    def clean_postal_code(self):
        pz = self.cleaned_data.get('postal_code', '').strip()
        if pz and (not pz.isdigit() or len(pz) != 5):
            raise forms.ValidationError('الرمز البريدي يجب أن يكون 5 أرقام')
        return pz
 

# ─────────────────────────────────────────────────────────
# Supplier
# ─────────────────────────────────────────────────────────

class SupplierForm(forms.ModelForm):
    class Meta:
        model  = Supplier
        fields = ['name', 'name_en','street_name','building_number', 'district','city','vat_number', 'cr_number',
                  'email', 'phone',]
        widgets = {
            'name':       forms.TextInput(attrs={'class': _cls, 'placeholder': 'اسم المورد'}),
            'name_en':    forms.TextInput(attrs={'class': _cls}),
            'street_name': forms.TextInput(attrs={'class': _cls}),
            'building_number': forms.TextInput(attrs={'class': _cls}),
            'district':    forms.TextInput(attrs={'class': _cls}),
            'city':       forms.TextInput(attrs={'class': _cls}),
            'vat_number': forms.TextInput(attrs={'class': _cls}),
            'cr_number':  forms.TextInput(attrs={'class': _cls}),
            'email':      forms.EmailInput(attrs={'class': _cls}),
            'phone':      forms.TextInput(attrs={'class': _cls}),
        }

    def clean_name(self):
        name = ' '.join(self.cleaned_data.get('name', '').split())
        qs = Supplier.objects.filter(name__iexact=name)
        if self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise forms.ValidationError('هذا الاسم موجود بالفعل')
        return name


# ─────────────────────────────────────────────────────────
# Account
# ─────────────────────────────────────────────────────────

class AccountForm(forms.ModelForm):
    class Meta:
        model  = Account
        fields = ['group', 'parent', 'code', 'name_ar', 'name_en',
                  'account_type', 'normal_balance', 'is_active',
                  'opening_balance', 'opening_balance_date', 'notes']
        widgets = {
            'group':                forms.Select(attrs={'class': _cls}),
            'parent':               forms.Select(attrs={'class': _cls}),
            'code':                 forms.TextInput(attrs={'class': _cls, 'placeholder': 'مثال: 01010101'}),
            'name_ar':              forms.TextInput(attrs={'class': _cls}),
            'name_en':              forms.TextInput(attrs={'class': _cls}),
            'account_type':         forms.Select(attrs={'class': _cls}),
            'normal_balance':       forms.Select(attrs={'class': _cls}),
            'opening_balance':      forms.NumberInput(attrs={'class': _cls, 'step': '0.01'}),
            'opening_balance_date': forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'notes':                forms.Textarea(attrs={'class': _cls, 'rows': 2}),
        }

    def clean_code(self):
        code = self.cleaned_data.get('code', '').strip()
        if not code.isdigit():
            raise forms.ValidationError('كود الحساب يجب أن يحتوي على أرقام فقط')
        qs = Account.objects.filter(code=code)
        if self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            raise forms.ValidationError('هذا الكود موجود بالفعل')
        return code


# ─────────────────────────────────────────────────────────
# Fiscal Year
# ─────────────────────────────────────────────────────────

class FiscalYearForm(forms.ModelForm):
    class Meta:
        model  = FiscalYear
        fields = ['name', 'start_date', 'end_date']
        widgets = {
            'name':       forms.TextInput(attrs={'class': _cls}),
            'start_date': forms.DateInput(attrs={'type': 'date', 'class': _cls}),
            'end_date':   forms.DateInput(attrs={'type': 'date', 'class': _cls}),
        }

    def clean(self):
        cleaned = super().clean()
        sd = cleaned.get('start_date')
        ed = cleaned.get('end_date')
        if sd and ed and sd >= ed:
            raise forms.ValidationError('تاريخ البداية يجب أن يكون قبل تاريخ النهاية')
        return cleaned


# ─────────────────────────────────────────────────────────
# Journal Entry
# ─────────────────────────────────────────────────────────

from django import forms
from django.forms import inlineformset_factory, BaseInlineFormSet
from .models import JournalEntry, JournalEntryLine, Account, CostCenter
 
_cls = 'form-control'
 
# ─────────────────────────────────────────────────────────
# Journal Entry Form - النموذج الرئيسي للقيد
# ─────────────────────────────────────────────────────────

class JournalEntryForm(forms.ModelForm):
    class Meta:
        model = JournalEntry
        fields = ['number', 'date', 'entry_type', 'description', 'reference', 'notes']
        widgets = {
            'number': forms.TextInput(attrs={
                'class': _cls,
                'readonly': 'readonly',
                'style': 'background-color: #e9ecef;',
                'placeholder': 'سيتم توليده تلقائياً'
            }),
            'date': forms.DateInput(attrs={
                'type': 'date',
                'class': _cls,
            }),
            'entry_type': forms.Select(attrs={
                'class': _cls,
            }),
            'description': forms.Textarea(attrs={
                'class': _cls,
                'rows': 3,
                'placeholder': 'أدخل وصف القيد المحاسبي',
            }),
            'reference': forms.TextInput(attrs={
                'class': _cls,
                'placeholder': 'رقم المستند / المرجع (اختياري)',
            }),
            'notes': forms.Textarea(attrs={
                'class': _cls,
                'rows': 2,
                'placeholder': 'ملاحظات إضافية (اختياري)',
            }),
        }
        labels = {
            'number': 'رقم القيد',
            'date': 'التاريخ',
            'entry_type': 'نوع القيد',
            'description': 'البيان / الوصف',
            'reference': 'المرجع',
            'notes': 'ملاحظات',
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fields['date'].required = True
        self.fields['entry_type'].required = True
        self.fields['description'].required = True
        self.fields['number'].required = False  # مهم: اجعله غير مطلوب
        self.fields['reference'].required = False
        self.fields['notes'].required = False
        
        if not self.instance.pk:
            # للقيد الجديد فقط
            self.fields['number'].initial = self._generate_entry_number()
            self.fields['date'].initial = timezone.now().date()
            self.fields['entry_type'].initial = 'manual'
    
    def _generate_entry_number(self):
        """توليد رقم قيد جديد فريد"""
        from .models import JournalEntry
        import uuid
        from datetime import datetime
        
        year = timezone.now().year
        month = timezone.now().month
        
        # البحث عن آخر قيد في نفس الشهر والسنة
        last_entry = JournalEntry.objects.filter(
            date__year=year,
            date__month=month
        ).order_by('-number').first()
        
        if last_entry and last_entry.number:
            try:
                # محاولة استخراج الرقم التسلسلي من الصيغة JV-YYYYMM-XXXXX
                parts = last_entry.number.split('-')
                if len(parts) >= 3:
                    seq = int(parts[-1]) + 1
                    return f"JV-{year}{month:02d}-{seq:05d}"
            except (ValueError, IndexError, AttributeError):
                pass
        
        # إذا لم يتم العثور على قيد سابق، ابدأ من 00001
        return f"JV-{year}{month:02d}-00001"
    
    def clean_number(self):
        """التحقق من أن رقم القيد فريد"""
        number = self.cleaned_data.get('number')
        if not number:
            # توليد رقم تلقائياً إذا كان فارغاً
            number = self._generate_entry_number()
        
        # التحقق من عدم وجود تكرار
        qs = JournalEntry.objects.filter(number=number)
        if self.instance.pk:
            qs = qs.exclude(pk=self.instance.pk)
        if qs.exists():
            # إذا كان الرقم مكرراً، أضف لاحقة عشوائية
            import uuid
            number = f"{number}-{str(uuid.uuid4())[:4].upper()}"
        
        return number
    
    def clean(self):
        cleaned_data = super().clean()
        date = cleaned_data.get('date')
        
        if date and date > timezone.now().date():
            self.add_error('date', 'لا يمكن أن يكون تاريخ القيد في المستقبل.')
        
        # تأكد من وجود رقم قيد
        if not cleaned_data.get('number'):
            cleaned_data['number'] = self._generate_entry_number()
        
        return cleaned_data
    
class JournalEntryLineForm(forms.ModelForm):
    class Meta:
        model  = JournalEntryLine
        fields = ['account', 'description', 'debit_amount', 'credit_amount', 'cost_center']
        widgets = {
            'account': forms.Select(attrs={
                'class': f'{_cls} line-account',
            }),
            'description': forms.TextInput(attrs={
                'class': _cls,
                'placeholder': 'البيان',
            }),
            'debit_amount': forms.NumberInput(attrs={
                'class': f'{_cls} line-debit debit-input',
                'step': '0.01',
                'min': '0',
            }),
            'credit_amount': forms.NumberInput(attrs={
                'class': f'{_cls} line-credit credit-input',
                'step': '0.01',
                'min': '0',
            }),
            'cost_center': forms.Select(attrs={'class': _cls}),
        }
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['account'].queryset = Account.objects.filter(
            account_type='detail', is_active=True
        ).order_by('code')
        self.fields['account'].required  = True
        self.fields['cost_center'].queryset = CostCenter.objects.filter(is_active=True)
        self.fields['cost_center'].required = False
        # Allow zero debit/credit — validation handled at formset level
        self.fields['debit_amount'].required  = False
        self.fields['credit_amount'].required = False
 
    def clean(self):
        cleaned = super().clean()
        # إذا كان الصف فارغاً تماماً (بدون account، بدون مبالغ) → تجاهله
        account      = cleaned.get('account')
        debit_amount  = cleaned.get('debit_amount')  or 0
        credit_amount = cleaned.get('credit_amount') or 0
        if not account and debit_amount == 0 and credit_amount == 0:
            # علّم هذا النموذج كـ "فارغ" — سيتم تجاهله في formset.clean()
            cleaned['_is_empty'] = True
            return cleaned
        # التحقق من صحة البيانات إذا كان الصف غير فارغ
        if account and debit_amount == 0 and credit_amount == 0:
            raise forms.ValidationError('يجب إدخال مبلغ مدين أو دائن.')
        if debit_amount > 0 and credit_amount > 0:
            raise forms.ValidationError('لا يمكن أن يكون السطر مدينًا ودائنًا في نفس الوقت.')
        cleaned['_is_empty'] = False
        return cleaned
 
 
class BaseJournalEntryLineFormSet(BaseInlineFormSet):
    """
    Formset مُحسَّن للقيود اليومية:
    - يتجاهل الصفوف الفارغة تلقائياً (لا يُطلق أخطاء عليها)
    - يتحقق من توازن المدين والدائن
    - يشترط صفّين ممتلئَين على الأقل
    """
 
    def _should_delete_form(self, form):
        """تجاهل الصفوف الفارغة تلقائياً دون اشتراط تعبئتها."""
        if not form.has_changed():
            return True  # لا تغييرات → تجاهل
        cd = getattr(form, 'cleaned_data', {})
        if cd.get('_is_empty'):
            return True  # علّمه JournalEntryLineForm.clean() كفارغ
        return super()._should_delete_form(form)
 
    def clean(self):
        if any(self.errors):
            # لا تتحقق من التوازن إذا كان هناك أخطاء في الحقول
            return
 
        filled_forms = []
        total_debit  = 0
        total_credit = 0
 
        for form in self.forms:
            if not form.cleaned_data:
                continue
            if form.cleaned_data.get('DELETE'):
                continue
            if form.cleaned_data.get('_is_empty'):
                continue
            if not form.has_changed():
                continue
 
            account = form.cleaned_data.get('account')
            dr      = form.cleaned_data.get('debit_amount')  or 0
            cr      = form.cleaned_data.get('credit_amount') or 0
 
            if account or dr > 0 or cr > 0:
                filled_forms.append(form)
                total_debit  += dr
                total_credit += cr
 
        # اشتراط صفّين ممتلئَين على الأقل
        if len(filled_forms) < 2:
            raise forms.ValidationError(
                'يجب إضافة بندَين على الأقل في القيد اليومي (مدين ودائن).'
            )
 
        # التحقق من التوازن
        from decimal import Decimal
        if abs(Decimal(str(total_debit)) - Decimal(str(total_credit))) > Decimal('0.005'):
            raise forms.ValidationError(
                f'القيد غير متوازن: المدين ({total_debit:.2f}) ≠ الدائن ({total_credit:.2f})'
            )
 
 
# ════════════════════════════════════════════════════════
#  الـ FormSet النهائي — استبدل به الموجود في forms.py
# ════════════════════════════════════════════════════════
JournalEntryLineFormSet = inlineformset_factory(
    JournalEntry,
    JournalEntryLine,
    form=JournalEntryLineForm,
    formset=BaseJournalEntryLineFormSet,
    fields=['account', 'description', 'debit_amount', 'credit_amount', 'cost_center'],
    extra=0,           # ← كان 2 — الصفر يمنع الصفوف الفارغة الابتدائية
    can_delete=True,   # ← مطلوب لعمل زر الحذف
    min_num=1,         # ← اشتراط بندَين على الأقل
    validate_min=True, # ← تفعيل الاشتراط
    max_num=100,       # ← حد أقصى معقول
)