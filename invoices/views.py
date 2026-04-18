"""
=============================================================
 views.py — شركة عقارات سعودية
 Full accounting + ZATCA Phase 2 Integration
=============================================================
"""
from django.forms import inlineformset_factory
from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, JsonResponse
from django.db.models import Sum, Q, F
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.forms import AuthenticationForm
from django.contrib import messages
from django.views.decorators.http import require_http_methods, require_POST
from django.views.decorators.csrf import csrf_protect
from django.views.decorators.cache import never_cache
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.utils import timezone
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.core.serializers.json import DjangoJSONEncoder
from django.template.loader import render_to_string
from django.db.models.functions import TruncMonth
from django.core.exceptions import ValidationError
from decimal import Decimal
from datetime import date, timedelta, datetime
import io, json, uuid, base64, logging, urllib.parse

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .models import (
    SalesInvoice, SalesInvoiceItem, PurchaseInvoice,
    Company, Client, Supplier,
    Account, AccountCategory, AccountGroup,
    FiscalYear, AccountingPeriod,
    JournalEntry, JournalEntryLine,
    CostCenter, Payment,
    ZATCADevice, ZATCALog,
    CustomUser,
)
from .forms import *

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────

def check_perm(request, perm):
    if not request.user.has_perm(perm):
        return render(request, 'invoices/403.html', status=403)
    return None

_perm = check_perm


def _paginate(qs, request, per_page=20):
    p  = request.GET.get('page', 1)
    sz = int(request.GET.get('page_size', per_page) or per_page)
    paginator = Paginator(qs, sz)
    try:
        page_obj = paginator.page(p)
    except PageNotAnInteger:
        page_obj = paginator.page(1)
    except EmptyPage:
        page_obj = paginator.page(paginator.num_pages)
    return paginator, page_obj, sz


def _page_range(paginator, page_obj):
    current = page_obj.number
    total   = paginator.num_pages
    if total <= 7:
        return list(range(1, total + 1))
    if current <= 3:
        return list(range(1, 6)) + ['...', total]
    if current >= total - 2:
        return [1, '...'] + list(range(total - 4, total + 1))
    return [1, '...'] + list(range(current - 1, current + 2)) + ['...', total]


def permission_denied_view(request, exception=None):
    return render(request, 'invoices/403.html', status=403)


# ─────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────

@csrf_protect
@never_cache
def login_view(request):
    # إذا كان المستخدم مسجل الدخول بالفعل، أعد توجيهه إلى dashboard
    if request.user.is_authenticated:
        return redirect('dashboard')
    
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            messages.success(request, f"مرحباً {user.get_full_name() or user.username}!")
            
            # التحقق من وجود next في الطلب
            next_url = request.POST.get('next')
            if next_url:
                return redirect(next_url)
            return redirect('dashboard')
        else:
            messages.error(request, "اسم المستخدم أو كلمة المرور غير صحيحة.")
    else:
        form = AuthenticationForm()
    
    return render(request, 'login.html', {'form': form})


def logout_view(request):
    logout(request)
    messages.info(request, "تم تسجيل الخروج بنجاح.")
    return redirect('login')


# ─────────────────────────────────────────────────────────
# Dashboard
# ─────────────────────────────────────────────────────────

@login_required
def dashboard(request):
    yr = timezone.now().year
    sales_qs = SalesInvoice.objects.filter(date__year=yr)
    purch_qs = PurchaseInvoice.objects.filter(date__year=yr)

    sales_total  = sales_qs.aggregate(t=Sum('total'))['t'] or 0
    purch_total  = purch_qs.aggregate(t=Sum('total'))['t'] or 0
    sales_tax    = sales_qs.aggregate(t=Sum('total_tax'))['t'] or 0
    purch_tax    = purch_qs.aggregate(t=Sum('tax_amount'))['t'] or 0
    sales_count  = sales_qs.count()
    purch_count  = purch_qs.count()

    six_ago = timezone.now() - timedelta(days=180)
    sbm = (SalesInvoice.objects.filter(date__gte=six_ago)
           .annotate(month=TruncMonth('date')).values('month')
           .annotate(total=Sum('total')).order_by('month'))
    pbm = (PurchaseInvoice.objects.filter(date__gte=six_ago)
           .annotate(month=TruncMonth('date')).values('month')
           .annotate(total=Sum('total')).order_by('month'))

    months = sorted(set(
        [e['month'].strftime('%Y-%m') for e in sbm] +
        [e['month'].strftime('%Y-%m') for e in pbm]
    ))
    sm = {e['month'].strftime('%Y-%m'): float(e['total']) for e in sbm}
    pm = {e['month'].strftime('%Y-%m'): float(e['total']) for e in pbm}

    zatca_pending  = SalesInvoice.objects.filter(zatca_status='pending').count()
    zatca_cleared  = SalesInvoice.objects.filter(zatca_status__in=['cleared','reported']).count()
    zatca_rejected = SalesInvoice.objects.filter(zatca_status='rejected').count()

    return render(request, 'invoices/dashboard.html', {
        'current_year':         yr,
        'sales_total':          sales_total,
        'purch_total':          purch_total,
        'purchases_total':      purch_total,
        'sales_tax':            sales_tax,
        'purch_tax':            purch_tax,
        'sales_tax_total':      sales_tax,
        'purchases_tax_total':  purch_tax,
        'sales_count':          sales_count,
        'purch_count':          purch_count,
        'purchases_count':      purch_count,
        'avg_sales':            round(float(sales_total) / sales_count, 2) if sales_count else 0,
        'avg_purch':            round(float(purch_total) / purch_count, 2) if purch_count else 0,
        'avg_sales_invoice':    round(float(sales_total) / sales_count, 2) if sales_count else 0,
        'avg_purchase_invoice': round(float(purch_total) / purch_count, 2) if purch_count else 0,
        'recent_sales':         sales_qs.select_related('client').order_by('-date')[:5],
        'recent_purchases':     purch_qs.select_related('supplier').order_by('-date')[:5],
        'recent_purch':         purch_qs.select_related('supplier').order_by('-date')[:5],
        'top_clients':          sales_qs.values('client__name').annotate(total=Sum('total')).order_by('-total')[:5],
        'top_suppliers':        purch_qs.values('supplier__name').annotate(total=Sum('total')).order_by('-total')[:5],
        'zatca_pending':        zatca_pending,
        'zatca_cleared':        zatca_cleared,
        'zatca_rejected':       zatca_rejected,
        'chart_data_json': json.dumps({
            'labels':              months,
            'sales':               [float(sm.get(m, 0)) for m in months],
            'purchases':           [float(pm.get(m, 0)) for m in months],
            'sales_total':         float(sales_total),
            'purchases_total':     float(purch_total),
            'sales_tax_total':     float(sales_tax),
            'purchases_tax_total': float(purch_tax),
        }, cls=DjangoJSONEncoder),
    })

# ─────────────────────────────────────────────────────────
# Sales Invoices
# ─────────────────────────────────────────────────────────

@login_required
def sales_list(request):
    deny = check_perm(request, 'invoices.view_salesinvoice')
    if deny: return deny

    q     = request.GET.get('q', '')
    df    = request.GET.get('date_from', '')
    dt    = request.GET.get('date_to', '')
    ttype = request.GET.get('transaction_type', '')
    cid   = request.GET.get('client', '')
    zst   = request.GET.get('zatca_status', '')

    ps = int(request.GET.get('page_size') or request.session.get('sales_ps', 10))
    request.session['sales_ps'] = ps

    qs = SalesInvoice.objects.select_related('client', 'company').order_by('-date', '-id')
    if q:     qs = qs.filter(Q(client__name__icontains=q)|Q(invoice_number__icontains=q))
    if df:    qs = qs.filter(date__gte=df)
    if dt:    qs = qs.filter(date__lte=dt)
    if ttype: qs = qs.filter(transaction_type=ttype)
    if cid:   qs = qs.filter(client_id=cid)
    if zst:   qs = qs.filter(zatca_status=zst)

    total = qs.aggregate(t=Sum('total'))['t'] or 0
    paginator, page_obj, ps = _paginate(qs, request, ps)

    return render(request, 'invoices/sales_list.html', {
        'invoices':         page_obj,
        'q': q, 'date_from': df, 'date_to': dt,
        'transaction_type': ttype, 'client_id': cid, 'zatca_status': zst,
        'clients':          Client.objects.all().order_by('name'),
        'zatca_statuses':   SalesInvoice.ZATCA_STATUS,
        'total':            total,
        'paginator':        paginator,
        'page_range':       _page_range(paginator, page_obj),
        'current_page':     page_obj.number,
        'total_pages':      paginator.num_pages,
        'page_size':        ps,
        'page_sizes':       [10, 25, 50, 100],
        'start_index':      page_obj.start_index(),
        'end_index':        page_obj.end_index(),
        'total_count':      paginator.count,
    })


def generate_invoice_number():
    """توليد رقم فاتورة جديد بتنسيق متوافق مع ZATCA"""
    
    # الحصول على آخر فاتورة
    last_invoice = SalesInvoice.objects.order_by('-id').first()
    
    if last_invoice and last_invoice.invoice_number:
        try:
            # محاولة استخراج الرقم من التنسيقات المختلفة
            invoice_str = str(last_invoice.invoice_number)
            
            # إذا كان الرقم بصيغة INV-XXXX
            if invoice_str.startswith('INV-'):
                last_number = int(invoice_str.split('-')[1])
                new_number = last_number + 1
                return f"INV-{new_number:04d}"  # INV-0001, INV-0002, ...
            
            # إذا كان الرقم بصيغة INV-20240001
            elif invoice_str.startswith('INV-') and len(invoice_str) > 8:
                last_number = int(invoice_str.split('-')[1])
                new_number = last_number + 1
                return f"INV-{new_number:08d}"
            
            # إذا كان الرقم رقم فقط
            elif invoice_str.isdigit():
                last_number = int(invoice_str)
                new_number = last_number + 1
                return str(new_number)
            
            # أي تنسيق آخر
            else:
                return f"INV-{int(last_invoice.id) + 1:04d}"
                
        except (ValueError, TypeError, IndexError):
            return "INV-0001"
    
    # أول فاتورة في النظام
    return "INV-0001"


def _next_invoice_number():
    """نفس وظيفة generate_invoice_number - للتوافق"""
    return generate_invoice_number()


@login_required
def sales_add(request):
    """إضافة فاتورة مبيعات جديدة"""
    perm_check = check_perm(request, 'invoices.add_salesinvoice')
    if perm_check:
        return perm_check
    
    if request.method == 'POST':
        form = SalesInvoiceForm(request.POST, user=request.user)
        formset = get_sales_invoice_item_formset()
        
        # ✅ ربط البيانات بالفورمست
        formset = formset.__class__(request.POST, request.FILES)
        
        if form.is_valid() and formset.is_valid():
            try:
                with transaction.atomic():
                    invoice = form.save(commit=False)
                    
                    # تعيين الشركة الافتراضية
                    default_company = Company.objects.first()
                    if not default_company:
                        messages.error(request, 'لا توجد شركة مسجلة. يرجى إضافة بيانات الشركة أولاً.')
                        return redirect('sales_list')
                    invoice.company = default_company
                    
                    # تعيين منشئ الفاتورة
                    if not invoice.created_by:
                        invoice.created_by = request.user.get_full_name() or request.user.username
                    
                    # تعيين رقم الفاتورة
                    if not invoice.invoice_number:
                        invoice.invoice_number = generate_invoice_number()
                    
                    # حفظ الفاتورة
                    invoice.save()
                    
                    # حفظ بنود الفاتورة
                    total_quantity = 0
                    total_sales_excl_tax = Decimal('0')
                    
                    for item_form in formset:
                        if item_form.cleaned_data and not item_form.cleaned_data.get('DELETE', False):
                            item = item_form.save(commit=False)
                            item.invoice = invoice
                            
                            # ✅ تعيين القيم الافتراضية للحقول المحسوبة
                            item.discount_percent = Decimal('0')
                            item.discount_amount = Decimal('0')
                            item.tax_rate = invoice.tax_rate or Decimal('15')
                            
                            # حساب الإجمالي
                            gross = item.quantity * item.unit_price
                            item.total = gross
                            
                            # حساب الضريبة
                            item.tax_amount = item.total * item.tax_rate / Decimal('100')
                            item.total_with_tax = item.total + item.tax_amount
                            
                            item.save()
                            
                            total_quantity += int(item.quantity)
                            total_sales_excl_tax += item.total
                    
                    # تحديث إجماليات الفاتورة
                    invoice.total_quantity = total_quantity
                    invoice.total_sales_excl_tax = total_sales_excl_tax
                    invoice.discount_amount = Decimal('0')
                    invoice.taxable_amount = total_sales_excl_tax
                    invoice.total_tax = total_sales_excl_tax * invoice.tax_rate / Decimal('100')
                    invoice.total = total_sales_excl_tax + invoice.total_tax
                    invoice.amount_paid = Decimal('0')
                    
                    invoice.save()
                    
                    messages.success(request, f'تم إنشاء الفاتورة {invoice.invoice_number} بنجاح')
                    return redirect('sales_list')
                    
            except Exception as e:
                messages.error(request, f'خطأ في حفظ الفاتورة: {str(e)}')
                import traceback
                traceback.print_exc()
        else:
            # عرض أخطاء النموذج
            if form.errors:
                for field, errors in form.errors.items():
                    for error in errors:
                        messages.error(request, f'{field}: {error}')
            if formset.errors:
                for idx, errors in enumerate(formset.errors):
                    if errors:
                        messages.error(request, f'البند {idx+1}: {errors}')
    else:
        form = SalesInvoiceForm(user=request.user)
        formset = get_sales_invoice_item_formset()
    
    return render(request, 'invoices/sales_form.html', {
        'form': form,
        'formset': formset,
        'title': 'إضافة فاتورة مبيعات',
        'back_url': 'sales_list'
    })

@login_required
def sales_edit(request, pk):
    perm_check = check_perm(request, 'invoices.change_salesinvoice')
    if perm_check:
        return perm_check

    invoice = get_object_or_404(SalesInvoice, pk=pk)
    ItemFormSet = get_sales_invoice_item_formset(instance=invoice).__class__

    if request.method == 'POST':
        form    = SalesInvoiceForm(request.POST, instance=invoice, user=request.user)
        formset = ItemFormSet(request.POST, request.FILES, instance=invoice)

        if form.is_valid() and formset.is_valid():
            try:
                with transaction.atomic():
                    updated_invoice = form.save(commit=False)
                    updated_invoice.save()

                    total_quantity       = Decimal('0')
                    total_sales_excl_tax = Decimal('0')
                    total_discount       = Decimal('0')

                    for item_form in formset:
                        if not item_form.cleaned_data:
                            continue
                        if item_form.cleaned_data.get('DELETE', False):
                            if item_form.instance and item_form.instance.pk:
                                item_form.instance.delete()
                            continue

                        item = item_form.save(commit=False)
                        item.invoice = updated_invoice
                        gross = item.quantity * item.unit_price
                        item.total = gross
                        tax_rate = updated_invoice.tax_rate or Decimal('15')
                        item.tax_rate = tax_rate
                        item.tax_amount = item.total * tax_rate / Decimal('100')
                        item.total_with_tax = item.total + item.tax_amount
                        item.discount_percent = Decimal('0')
                        item.discount_amount  = Decimal('0')
                        item.save()

                        total_quantity += item.quantity
                        total_sales_excl_tax += item.total

                    updated_invoice.total_quantity = int(total_quantity)
                    updated_invoice.total_sales_excl_tax = total_sales_excl_tax
                    updated_invoice.discount_amount = Decimal('0')
                    updated_invoice.taxable_amount = total_sales_excl_tax
                    updated_invoice.total_tax = total_sales_excl_tax * (updated_invoice.tax_rate or Decimal('15')) / Decimal('100')
                    updated_invoice.total = updated_invoice.taxable_amount + updated_invoice.total_tax

                    # Reset ZATCA status if invoice data changed
                    if updated_invoice.zatca_status in ('cleared', 'reported'):
                        pass  # Don't reset already accepted invoices
                    else:
                        updated_invoice.zatca_status = 'pending'
                        updated_invoice.zatca_invoice_hash = ''
                        updated_invoice.zatca_signed_xml = ''
                        updated_invoice.invoice_counter_value = 0
                        updated_invoice.previous_invoice_hash = ''

                    updated_invoice.save()

                    messages.success(request, f'تم تعديل الفاتورة {updated_invoice.invoice_number} بنجاح')
                    return redirect('sales_list')

            except Exception as e:
                messages.error(request, f'خطأ في تعديل الفاتورة: {str(e)}')
                log.error(f'Error editing invoice {pk}: {e}')
        else:
            if form.errors:
                for field, errors in form.errors.items():
                    for error in errors:
                        messages.error(request, f'{field}: {error}')
            if formset.errors:
                for form_errors in formset.errors:
                    for field, errors in form_errors.items():
                        for error in errors:
                            messages.error(request, f'بند الفاتورة - {field}: {error}')
    else:
        form    = SalesInvoiceForm(instance=invoice, user=request.user)
        formset = ItemFormSet(instance=invoice)

    return render(request, 'invoices/sales_form.html', {
        'form':     form,
        'formset':  formset,
        'title':    'تعديل فاتورة مبيعات',
        'back_url': 'sales_list',
        'invoice':  invoice,
    })

@login_required
def sales_invoice_item_delete(request, pk):
    """حذف بند من الفاتورة (AJAX)"""
    if request.method != 'POST':
        return JsonResponse({'success': False, 'error': 'Method not allowed'}, status=405)
    
    try:
        item = get_object_or_404(SalesInvoiceItem, pk=pk)
        invoice = item.invoice
        
        # حذف البند
        item.delete()
        
        # إعادة حساب إجماليات الفاتورة
        items = invoice.items.all()
        
        if items.exists():
            total_quantity = sum(item.quantity for item in items)
            total_sales_excl_tax = sum(item.total for item in items)
            total_discount = sum(item.discount_amount for item in items)
            
            invoice.total_quantity = total_quantity
            invoice.total_sales_excl_tax = total_sales_excl_tax
            invoice.discount_amount = total_discount
            invoice.taxable_amount = total_sales_excl_tax - total_discount
            invoice.total_tax = invoice.taxable_amount * (invoice.tax_rate or Decimal('15')) / Decimal('100')
            invoice.total = invoice.taxable_amount + invoice.total_tax
        else:
            # إذا لم يبق أي بنود، قم بتصفير الإجماليات
            invoice.total_quantity = 0
            invoice.total_sales_excl_tax = Decimal('0')
            invoice.discount_amount = Decimal('0')
            invoice.taxable_amount = Decimal('0')
            invoice.total_tax = Decimal('0')
            invoice.total = Decimal('0')
        
        invoice.save()
        
        return JsonResponse({
            'success': True,
            'total': float(invoice.total),
            'total_sales_excl_tax': float(invoice.total_sales_excl_tax),
            'total_tax': float(invoice.total_tax)
        })
        
    except SalesInvoiceItem.DoesNotExist:
        return JsonResponse({'success': False, 'error': 'البند غير موجود'}, status=404)
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@login_required
def sales_delete(request, pk):
    perm_check = check_perm(request, 'invoices.delete_salesinvoice')
    if perm_check:
        return perm_check
    inv = get_object_or_404(SalesInvoice, pk=pk)
    if request.method == 'POST':
        inv.delete()
        return redirect('sales_list')
    return render(request, 'invoices/confirm_delete.html', {'obj': inv, 'back_url': 'sales_list'})


@login_required
def sales_detail(request, pk):
    inv = get_object_or_404(SalesInvoice.objects.select_related('company','client').prefetch_related('items'), pk=pk)
    tax_rate = float(inv.tax_rate)
    item_rows = []
    for idx, item in enumerate(inv.items.all(), 1):
        sub = float(item.total)
        tax = round(sub * tax_rate / 100, 2)
        item_rows.append({
            'idx': idx, 'description': item.description,
            'unit_price': f'{float(item.unit_price):,.2f}',
            'quantity': item.quantity,
            'subtotal': f'{sub:,.2f}',
            'tax': f'{tax:,.2f} @{tax_rate:.0f}%',
            'total': f'{sub + tax:,.2f}',
        })

    def _words(amount, lang='ar'):
        try:
            from num2words import num2words as n2w
            r = int(amount); h = round((amount - r) * 100)
            if lang == 'ar':
                t = n2w(r, lang='ar') + ' ريال سعودي'
                if h: t += ' و' + n2w(h, lang='ar') + ' هللة'
            else:
                t = n2w(r, lang='en').title() + ' Saudi Riyal'
                if h: t += ' And ' + n2w(h, lang='en').title() + ' Halala'
            return t
        except Exception:
            return f'{amount:,.2f}'

    # Generate QR for display
    qr_b64 = ''
    try:
        import qrcode, io, base64
        from .zatca_service import ZATCAQRGenerator
        qr_data = inv.zatca_qr_code or ZATCAQRGenerator().generate(inv)
        qr = qrcode.QRCode(version=2, box_size=4, border=2)
        qr.add_data(qr_data)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        qr_b64 = 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        pass

    return render(request, 'invoices/sales_detail.html', {
        'inv': inv, 'item_rows': item_rows,
        'words_ar': _words(float(inv.total), 'ar'),
        'words_en': _words(float(inv.total), 'en'),
        'qr_b64':   qr_b64,
    })

@login_required
def sales_pdf(request, pk):
    import pathlib
    import qrcode
    import base64
    import io
    import logging
    from weasyprint import HTML
    from weasyprint.text.fonts import FontConfiguration
    from django.template.loader import render_to_string
    from django.conf import settings

    log = logging.getLogger(__name__)

    deny = check_perm(request, 'invoices.view_salesinvoice_pdf')
    if deny: 
        return deny

    inv = get_object_or_404(SalesInvoice.objects.select_related('company', 'client').prefetch_related('items'), pk=pk)

    static_dir = pathlib.Path(settings.BASE_DIR) / 'static'
    if not static_dir.exists():
        static_dir = pathlib.Path(settings.BASE_DIR) / 'Alwessam_contracting' / 'static'
    
    fonts_dir = static_dir / 'fonts'
    log.warning(f"Looking for fonts in: {fonts_dir}")
    
    def _font_css():
        faces = []
        font_files = {
            400: ('Regular', 'Cairo-Regular.ttf'),
            600: ('SemiBold', 'Cairo-SemiBold.ttf'),
            700: ('Bold', 'Cairo-Bold.ttf'),
        }
        
        for weight, (suffix, filename) in font_files.items():
            font_path = fonts_dir / filename
            if font_path.exists():
                log.warning(f"Found font: {font_path}")
                faces.append(
                    f"@font-face {{"
                    f"font-family:'Cairo';"
                    f"src:url('{font_path.as_uri()}') format('truetype');"
                    f"font-weight:{weight};"
                    f"font-style:normal;"
                    f"}}"
                )
            else:
                log.warning(f"Font not found: {font_path}")
        
        if not faces:
            # Fallback to system fonts
            log.warning("No Cairo fonts found, using system fallback")
            faces.append(
                "@font-face {"
                "font-family:'Cairo';"
                "src:local('Segoe UI'), local('Arial');"
                "}"
            )
        
        return '\n'.join(faces)

    def _generate_zatca_qr():
        try:
            from .zatca_service import ZATCAQRGenerator
            qr_data = inv.zatca_qr_code or ZATCAQRGenerator().generate(inv)
            qr = qrcode.QRCode(version=2, box_size=4, border=2,
                            error_correction=qrcode.constants.ERROR_CORRECT_M)
            qr.add_data(qr_data)
            qr.make(fit=True)
            img = qr.make_image(fill_color='black', back_color='white')
            buffer = io.BytesIO()
            img.save(buffer, format='PNG')
            return 'data:image/png;base64,' + base64.b64encode(buffer.getvalue()).decode()
        except Exception as e:
            log.warning(f'QR generation failed: {e}')
            return ''
    def _generate_simple_qr():
        """Generate simple QR code as fallback"""
        try:
            lines = [
                f"Seller: {inv.company.name}",
                f"VAT: {inv.company.vat_number}",
                f"Date: {inv.date}",
                f"Total: {float(inv.total):.2f} SAR",
                f"VAT: {float(inv.total_tax):.2f} SAR",
            ]
            qr = qrcode.QRCode(version=2, box_size=4, border=1)
            qr.add_data('\n'.join(lines))
            qr.make(fit=True)
            img = qr.make_image(fill_color='black', back_color='white')
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return 'data:image/png;base64,' + base64.b64encode(buf.getvalue()).decode()
        except Exception as e:
            log.warning(f'Simple QR failed: {e}')
            return ''

    def _create_qr_png(qr_data):
        """Create QR code PNG from data string"""
        qr = qrcode.QRCode(version=2, box_size=4, border=1)
        qr.add_data(qr_data)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        return base64.b64encode(buf.getvalue()).decode()

    def _get_logo():
        """Get company logo"""
        logo_path = static_dir / 'images' / 'logo.png'
        if logo_path.exists():
            return 'data:image/png;base64,' + base64.b64encode(logo_path.read_bytes()).decode()
        return ''

    def _words_ar(amount):
        """Convert amount to Arabic words"""
        try:
            from num2words import num2words as n2w
            riyal = int(amount)
            halalas = round((amount - riyal) * 100)
            txt = n2w(riyal, lang='ar') + ' ريال سعودي'
            if halalas:
                txt += ' و' + n2w(halalas, lang='ar') + ' هللة'
            return txt
        except Exception:
            return f'{amount:,.2f} ريال سعودي'

    def _words_en(amount):
        """Convert amount to English words"""
        try:
            from num2words import num2words as n2w
            riyal = int(amount)
            halalas = round((amount - riyal) * 100)
            txt = n2w(riyal, lang='en').title() + ' Saudi Riyal'
            if halalas:
                txt += ' And ' + n2w(halalas, lang='en').title() + ' Halala'
            return txt
        except Exception:
            return f'{amount:,.2f} Saudi Riyal'

    # Prepare invoice items
    tax_rate = float(inv.tax_rate)
    item_rows = []
    for idx, item in enumerate(inv.items.all(), 1):
        sub = float(item.total)
        tax = round(sub * tax_rate / 100, 2)
        item_rows.append({
            'idx': idx,
            'description': item.description,
            'unit_price': f'{float(item.unit_price):,.2f}',
            'quantity': item.quantity,
            'subtotal': f'{sub:,.2f}',
            'tax': f'{tax:,.2f} @{tax_rate:.0f}%',
            'tax_rate': f'{tax_rate:.0f}%',
            'total': f'{sub + tax:,.2f}',
        })

    context = {
        'inv': inv,
        'item_rows': item_rows,
        'subtotal': f'{float(inv.total_sales_excl_tax):,.2f}',
        'discount': f'{float(inv.discount_amount):,.2f}',
        'taxable_amount': f'{float(inv.taxable_amount):,.2f}',
        'tax_amount': f'{float(inv.total_tax):,.2f}',
        'total': f'{float(inv.total):,.2f}',
        'words_ar': _words_ar(float(inv.total)),
        'words_en': _words_en(float(inv.total)),
        'qr_b64': _generate_zatca_qr(),
        'logo_b64': _get_logo(),
        'font_css': _font_css(),
        'zatca_cleared': inv.zatca_status in ('cleared', 'reported'),
        'current_date': timezone.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    # Render HTML to PDF
    html_str = render_to_string('invoices/sales_pdf.html', context)
    
    # Configure fonts for WeasyPrint
    font_config = FontConfiguration()
    
    # Create PDF with proper font configuration
    pdf_bytes = HTML(string=html_str, base_url=settings.BASE_DIR).write_pdf(
        font_config=font_config,
        presentational_hints=True  # This helps with table styling
    )

    response = HttpResponse(pdf_bytes, content_type='application/pdf')
    response['Content-Disposition'] = f'inline; filename="invoice_{inv.invoice_number}.pdf"'
    return response

@login_required
def export_sales_excel(request):
    deny = check_perm(request, 'invoices.exportexcel_salesinvoice')
    if deny: return deny

    q = request.GET.get('q', '')
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    transaction_type = request.GET.get('transaction_type', '')
    client_id = request.GET.get('client', '')

    qs = SalesInvoice.objects.select_related('client', 'company').order_by('-date', '-id')
    if q: qs = qs.filter(Q(client__name__icontains=q) | Q(invoice_number__icontains=q))
    if date_from: qs = qs.filter(date__gte=date_from)
    if date_to: qs = qs.filter(date__lte=date_to)
    if transaction_type: qs = qs.filter(transaction_type=transaction_type)
    if client_id: qs = qs.filter(client_id=client_id)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "فواتير المبيعات"
    ws.sheet_view.rightToLeft = True

    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_fill  = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bdr = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'),  bottom=Side(style='thin')
    )

    HEADERS = [
        'رقم الفاتورة', 'التاريخ', 'نوع المعاملة', 'العميل',
        'إجمالي بدون ضريبة', 'نسبة الضريبة', 'قيمة الضريبة',
        'الإجمالي', 'حالة زاتكا', 'أنشأ بواسطة',
    ]
    NUM_COLS = len(HEADERS)

    # ── Row 1: Title ────────────────────────────────────────
    ws.merge_cells(f'A1:{get_column_letter(NUM_COLS)}1')
    t = ws['A1']
    t.value = "تقرير فواتير المبيعات"
    t.font = Font(size=16, bold=True, color="000000")
    t.alignment = Alignment(horizontal='center', vertical='center')
    t.border = Border()
    ws.row_dimensions[1].height = 32

    # ── Row 2: Filter info ──────────────────────────────────
    filter_parts = []
    if date_from: filter_parts.append(f"من: {date_from}")
    if date_to: filter_parts.append(f"إلى: {date_to}")
    if transaction_type: filter_parts.append(f"النوع: {dict(SalesInvoice.TRANSACTION_TYPES).get(transaction_type, transaction_type)}")
    if q: filter_parts.append(f"بحث: {q}")
    filter_parts.append(f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    ws.merge_cells(f'A2:{get_column_letter(NUM_COLS)}2')
    fi = ws['A2']
    fi.value = "   |   ".join(filter_parts)
    fi.font = Font(size=9, bold=True, color="555555")
    fi.alignment = Alignment(horizontal='center', vertical='center')
    t.border = Border()
    ws.row_dimensions[2].height = 18

    # ── Row 3: Headers ──────────────────────────────────────
    for ci, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=ci)
        cell.value = h
        cell.font = Font(bold=True, size=10, color="000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bdr
    ws.row_dimensions[3].height = 22

    # ── Data rows ───────────────────────────────────────────
    total_sum  = 0
    data_start = 4

    for idx, inv in enumerate(qs, start=data_start):
        vals = [
            inv.invoice_number,
            inv.date.strftime('%Y-%m-%d') if inv.date else '',
            inv.get_transaction_type_display(),
            inv.client.name if inv.client else '',
            float(inv.total_sales_excl_tax),
            f'{float(inv.tax_rate):.0f}%',
            float(inv.total_tax),
            float(inv.total),
            inv.get_zatca_status_display(),
            inv.created_by or '',
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=idx, column=ci, value=v)
            cell.border = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
                cell.alignment     = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[idx].height = 16
        total_sum += float(inv.total)

    # ── Total row ───────────────────────────────────────────
    total_row = qs.count() + data_start
    ws.merge_cells(
        start_row=total_row, start_column=1,
        end_row=total_row,   end_column=7
    )
    for ci in range(1, NUM_COLS + 1):
        cell = ws.cell(row=total_row, column=ci)
        cell.fill = total_fill
        cell.border = bdr
        cell.alignment = Alignment(horizontal='center', vertical='center')

    lc = ws.cell(row=total_row, column=1)
    lc.value = "الإجمالي الكلي"
    lc.font  = Font(bold=True, size=11, color="000000")

    vc = ws.cell(row=total_row, column=8)
    vc.value = total_sum
    vc.font = Font(bold=True, size=11, color="000000")
    vc.number_format = '#,##0.00'
    vc.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[total_row].height = 20

    # ── Column widths ───────────────────────────────────────
    min_widths = [16, 13, 16, 28, 18, 10, 14, 14, 14, 18]
    for ci, min_w in enumerate(min_widths, 1):
        col_letter = get_column_letter(ci)
        max_len    = min_w
        for row in range(data_start, total_row):
            val = ws.cell(row=row, column=ci).value
            if val:
                max_len = max(max_len, len(str(val)) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, 40)

    ws.freeze_panes = 'A4'

    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    filename = f"sales_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    resp = HttpResponse(
        excel_file.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp
# ─────────────────────────────────────────────────────────
# ZATCA Integration Views
# ─────────────────────────────────────────────────────────

@login_required
def zatca_submit_invoice(request, pk):
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny: return deny

    inv = get_object_or_404(SalesInvoice, pk=pk)

    if inv.zatca_status in ('cleared', 'reported'):
        messages.info(request, f'الفاتورة {inv.invoice_number} مُرسلة مسبقاً لزاتكا.')
        return redirect('sales_detail', pk=pk)

    device = ZATCADevice.objects.filter(company=inv.company, is_active=True).first()
    if not device:
        messages.error(request, 'لا يوجد جهاز زاتكا مُفعَّل.')
        return redirect('sales_detail', pk=pk)

    if not device.csid or not device.secret:
        messages.error(request, 'جهاز زاتكا غير مكتمل الإعداد.')
        return redirect('sales_detail', pk=pk)

    # Reset so process() starts fresh
    inv.zatca_status          = 'pending'
    inv.zatca_invoice_hash    = ''
    inv.zatca_signed_xml      = ''
    inv.invoice_counter_value = 0
    inv.previous_invoice_hash = ''
    inv.save(update_fields=[
        'zatca_status', 'zatca_invoice_hash', 'zatca_signed_xml',
        'invoice_counter_value', 'previous_invoice_hash'
    ])

    try:
        from .zatca_service import ZATCAInvoiceService
        service = ZATCAInvoiceService(device)
        result  = service.process(inv)

        if result['success']:
            messages.success(request,
                f'✓ تم إرسال الفاتورة {inv.invoice_number} إلى زاتكا بنجاح. '
                f'الحالة: {inv.get_zatca_status_display()}')
        else:
            resp_data = result.get('result', {}).get('data', {})
            status_code = result.get('result', {}).get('status_code', '')
            vr = resp_data.get('validationResults', {})
            errors   = vr.get('errorMessages', [])
            warnings = vr.get('warningMessages', [])
            
            err_str  = ' | '.join(f"{e.get('code')}: {e.get('message')}" for e in errors[:3]) if errors else ''
            warn_str = ' | '.join(f"{w.get('code')}: {w.get('message')}" for w in warnings[:3]) if warnings else ''
            
            disposition = resp_data.get('reportingStatus') or resp_data.get('clearanceStatus') or resp_data.get('dispositionMessage', '')
            
            if status_code == 401:
                messages.warning(request,
                    f'HTTP 401 — بيانات اعتماد زاتكا منتهية الصلاحية. '
                    f'يرجى تسجيل جهاز جديد بـ OTP جديد من بوابة فاتورة.')
            elif status_code == 202:
                messages.success(request,
                    f'✓ مقبولة مع تحذيرات (202). {warn_str}')
            else:
                messages.warning(request,
                    f'HTTP {status_code} — رُفضت. {err_str or disposition}')
    except Exception as e:
        messages.error(request, f'خطأ أثناء الإرسال لزاتكا: {str(e)}')
        log.exception('ZATCA submission failed for invoice %s', inv.invoice_number)

    return redirect('sales_detail', pk=pk)

@login_required
def zatca_bulk_submit(request):
    """Submit all pending invoices to ZATCA."""
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny: return deny

    if request.method != 'POST':
        pending = SalesInvoice.objects.filter(zatca_status='pending').count()
        return render(request, 'invoices/zatca_bulk.html', {'pending_count': pending})

    company = Company.objects.first()
    device  = ZATCADevice.objects.filter(company=company, is_active=True).first() if company else None

    if not device:
        messages.error(request, 'لا يوجد جهاز زاتكا مُفعَّل.')
        return redirect('sales_list')

    from .zatca_service import ZATCAInvoiceService
    service   = ZATCAInvoiceService(device)
    invoices  = SalesInvoice.objects.filter(zatca_status='pending')
    success_n = fail_n = 0

    for inv in invoices:
        try:
            result = service.process(inv)
            if result['success']:
                success_n += 1
            else:
                fail_n += 1
        except Exception as e:
            log.warning('ZATCA bulk: invoice %s failed: %s', inv.invoice_number, e)
            fail_n += 1

    messages.success(request, f'تم الإرسال: {success_n} ناجح، {fail_n} فاشل.')
    return redirect('sales_list')



import base64
import logging
from django.contrib import messages
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required

log = logging.getLogger(__name__)


@login_required
def zatca_onboarding(request):
    """ZATCA device onboarding page (Production-ready)."""

    from .models import Company, ZATCADevice, ZATCALog
    from .views import check_perm

    deny = check_perm(request, 'invoices.submit_zatca')
    if deny:
        return deny

    company = Company.objects.first()
    devices = ZATCADevice.objects.filter(company=company).order_by('-registered_at') if company else []
    logs = ZATCALog.objects.order_by('-created_at')[:20]

    # 🔴 VALIDATION FUNCTION
    def validate_company(company):
        errors = []

        if not company.vat_number or len(company.vat_number) != 15:
            errors.append("الرقم الضريبي يجب أن يكون 15 رقم")

        for field, label in [
            ("street_name", "اسم الشارع"),
            ("building_number", "رقم المبنى"),
            ("city", "المدينة"),
            ("postal_code", "الرمز البريدي"),
        ]:
            if not getattr(company, field):
                errors.append(f"{label} مطلوب")

        return errors

    if request.method == 'POST':
        otp = request.POST.get('otp', '').strip()
        device_name = request.POST.get('device_name', 'EGS-001').strip()
        environment = request.POST.get('environment', 'simulation')

        csr_pem_input = request.POST.get('csr_pem', '').strip()
        priv_key_input = request.POST.get('private_key_pem', '').strip()

        csr_file = request.FILES.get('csr_file')
        key_file = request.FILES.get('key_file')

        # 🔹 READ FILES
        if csr_file:
            try:
                csr_pem_input = csr_file.read().decode('utf-8')
            except Exception:
                csr_pem_input = base64.b64decode(csr_file.read()).decode()

        if key_file:
            try:
                priv_key_input = key_file.read().decode('utf-8')
            except Exception:
                priv_key_input = base64.b64decode(key_file.read()).decode()

        # 🔴 COMPANY VALIDATION
        if not company:
            messages.error(request, 'يرجى إضافة بيانات الشركة أولاً.')
            return redirect('zatca_onboarding')

        company_errors = validate_company(company)
        if company_errors:
            for err in company_errors:
                messages.error(request, err)
            return redirect('zatca_onboarding')

        # 🔴 OTP VALIDATION
        if environment != 'sandbox':
            if not otp or len(otp) != 6 or not otp.isdigit():
                messages.error(request, 'OTP يجب أن يكون 6 أرقام صحيحة.')
                return redirect('zatca_onboarding')

        try:
            from .zatca_service import onboard_device, generate_csr

            # 🔵 CASE 1: CSR uploaded
            if csr_pem_input and priv_key_input:
                device = onboard_device(
                    company=company,
                    device_name=device_name,
                    otp=otp,
                    environment=environment,
                    csr_pem=csr_pem_input,
                    private_key_pem=priv_key_input,
                )

            # 🟡 CASE 2: AUTO CSR
            else:
                csr_pem, private_key = generate_csr(
                    company_name=company.name_en or company.name,
                    vat_number=company.vat_number,
                    serial_number=device_name,
                )

                device = onboard_device(
                    company=company,
                    device_name=device_name,
                    otp=otp,
                    environment=environment,
                    csr_pem=csr_pem,
                    private_key_pem=private_key,
                )

            messages.success(request, f'✓ تم تسجيل الجهاز {device.device_name} بنجاح')

        except Exception as e:
            error_msg = str(e)

            # 🔥 IMPORTANT DEBUG
            if hasattr(e, 'response'):
                try:
                    error_msg = e.response.text
                except:
                    pass

            messages.error(request, f'❌ فشل التسجيل:\n{error_msg}')
            log.exception('ZATCA onboarding failed')

        return redirect('zatca_onboarding')

    return render(request, 'invoices/zatca_onboarding.html', {
        'company': company,
        'devices': devices,
        'logs': logs,
        'environments': [
            ('sandbox', 'Sandbox (اختبار فقط)'),
            ('simulation', 'Simulation (اختبار حقيقي)'),
            ('production', 'Production (إنتاج)'),
        ],
    })
    
    
@login_required
def zatca_logs(request):
    logs = ZATCALog.objects.select_related('sales_invoice', 'device').order_by('-created_at')
    paginator, page_obj, ps = _paginate(logs, request, 20)
    return render(request, 'invoices/zatca_logs.html', {
        'logs': page_obj, 'paginator': paginator,
        'page_range': _page_range(paginator, page_obj),
    })


# ─────────────────────────────────────────────────────────
# Purchase Invoices
# ─────────────────────────────────────────────────────────

@login_required
def purchases_list(request):
    deny = check_perm(request, 'invoices.view_purchaseinvoice')
    if deny: return deny

    q     = request.GET.get('q', '').strip()
    df    = request.GET.get('date_from', '')
    dt    = request.GET.get('date_to', '')
    ttype = request.GET.get('transaction_type', '')
    sid   = request.GET.get('supplier', '')

    ps = int(request.GET.get('page_size') or request.session.get('purch_ps', 10))
    request.session['purch_ps'] = ps

    qs = PurchaseInvoice.objects.select_related('supplier').order_by('-date', '-id')
    if q:     qs = qs.filter(Q(supplier__name__icontains=q)|Q(invoice_number__icontains=q))
    if df:    qs = qs.filter(date__gte=df)
    if dt:    qs = qs.filter(date__lte=dt)
    if ttype: qs = qs.filter(transaction_type=ttype)
    if sid:   qs = qs.filter(supplier_id=sid)

    agg = qs.aggregate(total_sum=Sum('total'), sub_sum=Sum('subtotal'), tax_sum=Sum('tax_amount'))
    paginator, page_obj, ps = _paginate(qs, request, ps)

    return render(request, 'invoices/purchases_list.html', {
        'invoices':          page_obj,
        'q': q, 'date_from': df, 'date_to': dt,
        'transaction_type':  ttype, 'supplier_id': sid,
        'suppliers':         Supplier.objects.all().order_by('name'),
        'transaction_types': PurchaseInvoice.TRANSACTION_TYPES,  # ← added
        'total_sum':         agg['total_sum'] or 0,
        'subtotal_sum':      agg['sub_sum']   or 0,
        'tax_sum':           agg['tax_sum']   or 0,
        'paginator':         paginator,
        'page_range':        _page_range(paginator, page_obj),
        'current_page':      page_obj.number,
        'total_pages':       paginator.num_pages,
        'page_size':         ps,
        'page_sizes':        [10, 25, 50, 100],
        'start_index':       page_obj.start_index(),
        'end_index':         page_obj.end_index(),
        'total_count':       paginator.count,
        'base_query_string': urllib.parse.urlencode({
            k: v for k, v in request.GET.items() if k != 'page'
        }),
    })

@login_required
def purchases_add(request):
    deny = check_perm(request, 'invoices.add_purchaseinvoice')
    if deny: return deny

    if request.method == 'POST':
        form = PurchaseInvoiceForm(request.POST, request.FILES)
        if form.is_valid():
            with transaction.atomic():
                inv = form.save(commit=False)
                inv.created_by = request.user.get_full_name() or request.user.username
                inv.save()
            messages.success(request, f'تمت إضافة الفاتورة {inv.invoice_number}')
            return redirect('purchases_list')
    else:
        form = PurchaseInvoiceForm(initial={'entry_date': timezone.now().date()})

    suppliers_vat = {s.pk: s.vat_number or '' for s in Supplier.objects.all()}
    return render(request, 'invoices/purchases_form.html', {
        'form': form, 'title': 'إضافة فاتورة مشتريات',
        'back_url': 'purchases_list',
        'suppliers_vat': json.dumps(suppliers_vat),
    })


@login_required
def purchases_edit(request, pk):
    deny = check_perm(request, 'invoices.change_purchaseinvoice')
    if deny: return deny

    inv = get_object_or_404(PurchaseInvoice, pk=pk)

    if request.method == 'POST':
        form = PurchaseInvoiceForm(request.POST, request.FILES, instance=inv)
        if form.is_valid():
            form.save()
            messages.success(request, f'تم تعديل الفاتورة {inv.invoice_number}')
            return redirect('purchases_list')
        else:
            for field, errors in form.errors.items():
                for error in errors:
                    messages.error(request, f'{field}: {error}')
    else:
        form = PurchaseInvoiceForm(instance=inv)

    suppliers_vat = {s.pk: s.vat_number or '' for s in Supplier.objects.all()}
    return render(request, 'invoices/purchases_form.html', {
        'form':         form,
        'title':        f'تعديل فاتورة: {inv.invoice_number}',
        'back_url':     'purchases_list',
        'inv':          inv,
        'suppliers_vat': json.dumps(suppliers_vat),
    })

@login_required
def purchases_delete(request, pk):
    deny = check_perm(request, 'invoices.delete_purchaseinvoice')
    if deny: return deny
    inv = get_object_or_404(PurchaseInvoice, pk=pk)
    if request.method == 'POST':
        num = inv.invoice_number
        inv.delete()
        messages.success(request, f'تم حذف الفاتورة {num}')
        return redirect('purchases_list')
    return render(request, 'invoices/confirm_delete.html', {'obj': inv, 'back_url': 'purchases_list'})


@login_required
def export_purchases_excel(request):
    deny = check_perm(request, 'invoices.exportexcel_purchaseinvoice')
    if deny: return deny

    q     = request.GET.get('q', '')
    df    = request.GET.get('date_from', '')
    dt    = request.GET.get('date_to', '')
    ttype = request.GET.get('transaction_type', '')
    sid   = request.GET.get('supplier', '')

    qs = PurchaseInvoice.objects.select_related('supplier').order_by('-date', '-id')
    if q:     qs = qs.filter(Q(supplier__name__icontains=q)|Q(invoice_number__icontains=q))
    if df:    qs = qs.filter(date__gte=df)
    if dt:    qs = qs.filter(date__lte=dt)
    if ttype: qs = qs.filter(transaction_type=ttype)
    if sid:   qs = qs.filter(supplier_id=sid)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "فواتير المشتريات"
    ws.sheet_view.rightToLeft = True

    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_fill  = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bdr = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'),  bottom=Side(style='thin')
    )

    HEADERS = [
        'رقم الفاتورة', 'التاريخ', 'تاريخ الإدخال', 'نوع المعاملة',
        'المورد', 'الرقم الضريبي', 'بدون ضريبة',
        'قيمة الضريبة', 'الإجمالي', 'المدفوع',
    ]
    NUM_COLS = len(HEADERS)

    # Row 1: Title
    ws.merge_cells(f'A1:{get_column_letter(NUM_COLS)}1')
    t           = ws['A1']
    t.value     = "تقرير فواتير المشتريات"
    t.font      = Font(size=16, bold=True, color="000000")
    t.alignment = Alignment(horizontal='center', vertical='center')
    t.border    = Border()
    ws.row_dimensions[1].height = 32

    # Row 2: Filter info
    filter_parts = []
    if df:    filter_parts.append(f"من: {df}")
    if dt:    filter_parts.append(f"إلى: {dt}")
    if ttype: filter_parts.append(f"النوع: {dict(PurchaseInvoice.TRANSACTION_TYPES).get(ttype, ttype)}")
    if q:     filter_parts.append(f"بحث: {q}")
    filter_parts.append(f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    ws.merge_cells(f'A2:{get_column_letter(NUM_COLS)}2')
    fi           = ws['A2']
    fi.value     = "   |   ".join(filter_parts)
    fi.font      = Font(size=9, bold=True, color="555555")
    fi.alignment = Alignment(horizontal='center', vertical='center')
    fi.border    = Border()
    ws.row_dimensions[2].height = 18

    # Row 3: Headers
    for ci, h in enumerate(HEADERS, 1):
        cell           = ws.cell(row=3, column=ci)
        cell.value     = h
        cell.font      = Font(bold=True, size=10, color="000000")
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = bdr
    ws.row_dimensions[3].height = 22

    # Data rows
    total_sum  = 0
    data_start = 4

    for idx, inv in enumerate(qs, start=data_start):
        vals = [
            inv.invoice_number,
            inv.date.strftime('%Y-%m-%d') if inv.date else '',
            inv.entry_date.strftime('%Y-%m-%d') if inv.entry_date else '',
            inv.get_transaction_type_display(),
            inv.supplier.name if inv.supplier else '',
            inv.tax_number or '',
            float(inv.subtotal),
            float(inv.tax_amount),
            float(inv.total),
            float(inv.amount_paid),
        ]
        for ci, v in enumerate(vals, 1):
            cell           = ws.cell(row=idx, column=ci, value=v)
            cell.border    = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
        ws.row_dimensions[idx].height = 16
        total_sum += float(inv.total)

    # Total row
    total_row = qs.count() + data_start
    ws.merge_cells(
        start_row=total_row, start_column=1,
        end_row=total_row,   end_column=8
    )
    for ci in range(1, NUM_COLS + 1):
        cell           = ws.cell(row=total_row, column=ci)
        cell.fill      = total_fill
        cell.border    = bdr
        cell.alignment = Alignment(horizontal='center', vertical='center')

    lc       = ws.cell(row=total_row, column=1)
    lc.value = "الإجمالي الكلي"
    lc.font  = Font(bold=True, size=11, color="000000")

    vc               = ws.cell(row=total_row, column=9)
    vc.value         = total_sum
    vc.font          = Font(bold=True, size=11, color="000000")
    vc.number_format = '#,##0.00'
    vc.alignment     = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[total_row].height = 20

    # Column widths
    min_widths = [16, 13, 13, 16, 26, 16, 14, 14, 14, 14]
    for ci, min_w in enumerate(min_widths, 1):
        col_letter = get_column_letter(ci)
        max_len    = min_w
        for row in range(data_start, total_row):
            val = ws.cell(row=row, column=ci).value
            if val:
                max_len = max(max_len, len(str(val)) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, 40)

    ws.freeze_panes = 'A4'

    ef = io.BytesIO()
    wb.save(ef)
    ef.seek(0)

    filename = f"purchases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    resp = HttpResponse(
        ef.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp

# ─────────────────────────────────────────────────────────
# Clients
# ─────────────────────────────────────────────────────────

@login_required
def client_list(request):
    q = request.GET.get('search', '')
    qs = Client.objects.filter(
        Q(name__icontains=q)|Q(vat_number__icontains=q)|Q(email__icontains=q)
    ) if q else Client.objects.all()
    return render(request, 'invoices/client_list.html', {'clients': qs, 'search_query': q})


@login_required
def client_create(request):
    if request.method == 'POST':
        form = ClientForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'تم إضافة العميل بنجاح')
            return redirect('client_list')
        for field, errors in form.errors.items():
            for error in errors:
                messages.error(request, f'{form.fields[field].label}: {error}')
    else:
        form = ClientForm()
    return render(request, 'invoices/client_form.html', {
        'form': form, 'title': 'إضافة عميل جديد', 'button_text': 'إضافة'
    })


@login_required
def client_edit(request, pk):
    client = get_object_or_404(Client, pk=pk)
    if request.method == 'POST':
        form = ClientForm(request.POST, instance=client)
        if form.is_valid():
            form.save()
            messages.success(request, 'تم تحديث بيانات العميل')
            return redirect('client_list')
    else:
        form = ClientForm(instance=client)
    return render(request, 'invoices/client_form.html', {
        'form': form, 'client': client,
        'title': f'تعديل: {client.name}', 'button_text': 'تحديث'
    })


@login_required
def client_delete(request, pk):
    client = get_object_or_404(Client, pk=pk)
    if request.method == 'POST':
        name = client.name; client.delete()
        messages.success(request, f'تم حذف العميل "{name}"')
        return redirect('client_list')
    return render(request, 'invoices/client_confirm_delete.html', {'client': client})


# ─────────────────────────────────────────────────────────
# Suppliers
# ─────────────────────────────────────────────────────────

@login_required
def supplier_list(request):
    q = request.GET.get('search', '')
    qs = Supplier.objects.filter(
        Q(name__icontains=q)|Q(vat_number__icontains=q)
    ) if q else Supplier.objects.all()
    return render(request, 'invoices/supplier_list.html', {'suppliers': qs, 'search_query': q})


@login_required
def supplier_create(request):
    form = SupplierForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, 'تم إضافة المورد')
        return redirect('supplier_list')
    return render(request, 'invoices/supplier_form.html', {
        'form': form, 'title': 'إضافة مورد', 'button_text': 'إضافة'
    })


@login_required
def supplier_edit(request, pk):
    supplier = get_object_or_404(Supplier, pk=pk)
    form = SupplierForm(request.POST or None, instance=supplier)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, 'تم تعديل بيانات المورد')
        return redirect('supplier_list')
    return render(request, 'invoices/supplier_form.html', {
        'form': form, 'supplier': supplier,
        'title': f'تعديل: {supplier.name}', 'button_text': 'تحديث'
    })


@login_required
def supplier_delete(request, pk):
    supplier = get_object_or_404(Supplier, pk=pk)
    if request.method == 'POST':
        name = supplier.name; supplier.delete()
        messages.success(request, f'تم حذف المورد "{name}"')
        return redirect('supplier_list')
    return render(request, 'invoices/supplier_confirm_delete.html', {'supplier': supplier})


# ─────────────────────────────────────────────────────────
# Chart of Accounts - دليل الحسابات
# ─────────────────────────────────────────────────────────

@login_required
def chart_of_accounts(request):
    categories = AccountCategory.objects.prefetch_related(
        'groups__accounts__children'
    ).order_by('code')
    return render(request, 'invoices/chart_of_accounts.html', {
        'categories': categories
    })


@login_required
def account_add(request):
    deny = _perm(request, 'invoices.add_account')
    if deny: return deny
    form = AccountForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, 'تم إضافة الحساب بنجاح')
        return redirect('chart_of_accounts')
    return render(request, 'invoices/account_form.html', {
        'form': form, 'title': 'إضافة حساب جديد'
    })


@login_required
def account_edit(request, pk):
    deny = _perm(request, 'invoices.change_account')
    if deny: return deny
    account = get_object_or_404(Account, pk=pk)
    form = AccountForm(request.POST or None, instance=account)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, f'تم تعديل الحساب {account.name_ar}')
        return redirect('chart_of_accounts')
    return render(request, 'invoices/account_form.html', {
        'form': form, 'account': account, 'title': f'تعديل: {account.name_ar}'
    })


@login_required
def account_statement(request, pk):
    account   = get_object_or_404(Account, pk=pk)
    date_from = request.GET.get('date_from') or date.today().replace(month=1, day=1).isoformat()
    date_to   = request.GET.get('date_to')   or date.today().isoformat()
    export    = request.GET.get('export')

    lines = (JournalEntryLine.objects
             .filter(
                 account=account,
                 journal_entry__status='posted',
                 journal_entry__date__gte=date_from,
                 journal_entry__date__lte=date_to,
             )
             .select_related('journal_entry')
             .order_by('journal_entry__date', 'journal_entry__id'))

    running = account.opening_balance or Decimal('0')
    entries = []
    for line in lines:
        if account.normal_balance == 'debit':
            running += line.debit_amount - line.credit_amount
        else:
            running += line.credit_amount - line.debit_amount
        entries.append({
            'date':            line.journal_entry.date,
            'entry_id':        line.journal_entry.pk,
            'entry_number':    line.journal_entry.number,
            'description':     line.description or line.journal_entry.description,
            'debit':           line.debit_amount,
            'credit':          line.credit_amount,
            'running_balance': running,
        })

    totals        = lines.aggregate(
        total_debit=Sum('debit_amount'),
        total_credit=Sum('credit_amount'),
    )
    closing_balance = running
    balance_type    = 'مدين' if closing_balance >= 0 else 'دائن'

    if export == 'excel':
        rows_for_export = [{
            'date': e['date'], 'number': e['entry_number'],
            'description': e['description'],
            'debit': e['debit'], 'credit': e['credit'],
            'balance': e['running_balance'],
        } for e in entries]
        return _export_account_statement_excel(
            account, rows_for_export, date_from, date_to, closing_balance
        )

    return render(request, 'invoices/account_ledger.html', {
        'account':         account,
        'entries':         entries,
        'date_from':       date_from,
        'date_to':         date_to,
        'total_debit':     totals['total_debit']  or 0,
        'total_credit':    totals['total_credit'] or 0,
        'closing_balance': closing_balance,
        'opening_balance': account.opening_balance or 0,
        'balance_type':    balance_type,
    })

def _export_account_statement_excel(account, rows, date_from, date_to, closing):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = f"كشف حساب {account.code}"
    ws.sheet_view.rightToLeft = True
    hf  = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    bdr = Border(*[Side(style='thin')] * 4)

    ws.merge_cells('A1:F1')
    c = ws['A1']
    c.value = f"كشف حساب: {account.name_ar} ({account.code})"
    c.font  = Font(size=14, bold=True, color="C9A84C")
    c.fill  = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:F2')
    c2 = ws['A2']
    c2.value = f"من {date_from} إلى {date_to}"
    c2.font  = Font(size=10, color="C9A84C")
    c2.fill  = hf
    c2.alignment = Alignment(horizontal='center')

    headers = ['التاريخ','رقم القيد','البيان','مدين','دائن','الرصيد']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.font  = Font(bold=True, size=10, color="C9A84C")
        cell.fill  = hf
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bdr

    for ri, row in enumerate(rows, start=4):
        vals = [str(row['date']), row['number'], row['description'],
                float(row['debit']), float(row['credit']), float(row['balance'])]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')

    last = len(rows) + 4
    ws.cell(row=last, column=1, value="الرصيد الختامي").font = Font(bold=True)
    ws.cell(row=last, column=6, value=float(closing)).font = Font(bold=True)

    for ci in range(1, 7):
        ws.column_dimensions[get_column_letter(ci)].width = 20

    ef = io.BytesIO()
    wb.save(ef); ef.seek(0)
    resp = HttpResponse(ef.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = (
        f'attachment; filename="statement_{account.code}_{date_from}_{date_to}.xlsx"'
    )
    return resp


# ─────────────────────────────────────────────────────────
# Journal Entries - اليومية الأمريكية
# ─────────────────────────────────────────────────────────

@login_required
def journal_entries_list(request):
    q     = request.GET.get('q', '')
    df    = request.GET.get('date_from', '')
    dt    = request.GET.get('date_to', '')
    etype = request.GET.get('entry_type', '')
    status= request.GET.get('status', '')

    qs = JournalEntry.objects.select_related('created_by').prefetch_related('lines').order_by('-date', '-number')
    if q:     qs = qs.filter(Q(number__icontains=q)|Q(description__icontains=q))
    if df:    qs = qs.filter(date__gte=df)
    if dt:    qs = qs.filter(date__lte=dt)
    if etype: qs = qs.filter(entry_type=etype)
    if status: qs = qs.filter(status=status)

    paginator, page_obj, ps = _paginate(qs, request, 20)
    return render(request, 'invoices/journal_list.html', {
        'entries':     page_obj,
        'q': q, 'date_from': df, 'date_to': dt,
        'entry_type': etype, 'status': status,
        'entry_types': JournalEntry.ENTRY_TYPES,
        'statuses':    JournalEntry.STATUS_CHOICES,
        'paginator':   paginator,
        'page_range':  _page_range(paginator, page_obj),
        'page_size':   ps, 'page_sizes': [10, 25, 50],
        'total_count': paginator.count,
    })


# Update journal_entry_add view to pass cost_centers and account levels
@login_required
def journal_entry_add(request):
    deny = _perm(request, 'invoices.add_journalentry')
    if deny: 
        return deny

    if request.method == 'POST':
        form = JournalEntryForm(request.POST)
        formset = JournalEntryLineFormSet(request.POST)
        
        if form.is_valid() and formset.is_valid():
            try:
                with transaction.atomic():
                    entry = form.save(commit=False)
                    entry.created_by = request.user
                    
                    # ✅ Check fiscal year before saving
                    active_year = FiscalYear.get_active_for_date(entry.date)
                    if not active_year:
                        # Create fiscal year automatically or show error
                        messages.error(
                            request, 
                            f'لا توجد سنة مالية نشطة تحتوي على تاريخ {entry.date}. '
                            'يرجى إنشاء سنة مالية أولاً.'
                        )
                        return redirect('fiscal_years_list')
                    
                    entry.save()
                    
                    lines = formset.save(commit=False)
                    for line in lines:
                        line.journal_entry = entry
                        line.save()
                    
                    if not entry.is_balanced:
                        raise ValueError(
                            f'القيد غير متوازن: المدين={entry.total_debit} ≠ الدائن={entry.total_credit}'
                        )
                
                messages.success(request, f'تم إنشاء القيد {entry.number}')
                return redirect('journal_entry_detail', pk=entry.pk)
                
            except ValidationError as e:
                messages.error(request, str(e))
            except ValueError as e:
                messages.error(request, str(e))
            except Exception as e:
                messages.error(request, f'خطأ: {e}')
        else:
            for field, errors in form.errors.items():
                for error in errors:
                    messages.error(request, f'{field}: {error}')
    else:
        form = JournalEntryForm(initial={'date': date.today()})
        formset = JournalEntryLineFormSet()

    # Get accounts with level information for tree display
    accounts = Account.objects.filter(
        account_type='detail', 
        is_active=True
    ).select_related('parent', 'group__category').order_by('code')
    
    account_list = []
    for acc in accounts:
        level = 0
        parent = acc.parent
        while parent:
            level += 1
            parent = parent.parent
        acc.level = level
        account_list.append(acc)
    
    cost_centers = CostCenter.objects.filter(is_active=True).order_by('code')

    return render(request, 'invoices/journal_entry_form.html', {
        'form': form,
        'formset': formset,
        'title': 'إضافة قيد يومي',
        'button_text': 'حفظ القيد',
        'accounts': account_list,
        'cost_centers': cost_centers,
    })
        
# Update journal_entry_detail view
@login_required
def journal_entry_detail(request, pk):
    entry = get_object_or_404(
        JournalEntry.objects.select_related('created_by').prefetch_related('lines__account', 'lines__cost_center'),
        pk=pk
    )
    total_debit  = sum(line.debit_amount for line in entry.lines.all())
    total_credit = sum(line.credit_amount for line in entry.lines.all())
    is_balanced  = abs(total_debit - total_credit) < Decimal('0.01')
    difference = abs(total_debit - total_credit)
    return render(request, 'invoices/journal_entry_detail.html', {
        'entry':        entry,
        'total_debit':  total_debit,
        'total_credit': total_credit,
        'is_balanced':  is_balanced,
        'difference': difference,  # ← أضف هذا
    })
    
@login_required
def journal_entry_post(request, pk):
    deny = _perm(request, 'invoices.post_journalentry')
    if deny: 
        return deny
    
    entry = get_object_or_404(JournalEntry, pk=pk)
    
    if entry.status != 'draft':
        messages.error(request, 'يمكن ترحيل القيود في حالة مسودة فقط.')
        return redirect('journal_entry_detail', pk=pk)
    
    try:
        # ✅ Check fiscal year
        closed_year = FiscalYear.objects.filter(
            is_closed=True,
            start_date__lte=entry.date,
            end_date__gte=entry.date
        ).first()
        
        if closed_year:
            messages.error(
                request, 
                f'لا يمكن ترحيل القيد. السنة المالية "{closed_year.name}" مغلقة.'
            )
            return redirect('journal_entry_detail', pk=pk)
        
        entry.post(request.user)
        messages.success(request, f'تم ترحيل القيد {entry.number}')
        
    except ValidationError as e:
        messages.error(request, str(e))
    except ValueError as e:
        messages.error(request, str(e))
    
    return redirect('journal_entry_detail', pk=pk)

@login_required
def journal_entry_reverse(request, pk):
    deny = _perm(request, 'invoices.reverse_journalentry')
    if deny: return deny
    entry = get_object_or_404(JournalEntry, pk=pk)
    if request.method == 'POST':
        try:
            rev = entry.reverse(request.user, description=request.POST.get('description'))
            messages.success(request, f'تم إنشاء قيد العكس {rev.number}')
            return redirect('journal_entry_detail', pk=rev.pk)
        except ValueError as e:
            messages.error(request, str(e))
    return render(request, 'invoices/journal_entry_reverse_confirm.html', {'entry': entry})

@login_required
def journal_entry_delete(request, pk):
    """حذف قيد يومي"""
    deny = _perm(request, 'invoices.delete_journalentry')
    if deny: 
        return deny
    
    entry = get_object_or_404(JournalEntry, pk=pk)
    
    # لا يمكن حذف القيود المرحلة
    if entry.status == 'posted':
        messages.error(request, 'لا يمكن حذف قيد مرحّل. يجب عكسه أولاً.')
        return redirect('journal_entry_detail', pk=pk)
    
    if request.method == 'POST':
        entry_number = entry.number
        entry.delete()
        messages.success(request, f'تم حذف القيد {entry_number} بنجاح')
        return redirect('journal_entries_list')
    
    return render(request, 'invoices/journal_confirm_delete.html', {
        'entry': entry,
        'back_url': 'journal_entries_list'
    })

@login_required
def journal_export_excel(request):
    deny = _perm(request, 'invoices.export_journal')
    if deny: 
        return deny
    
    # ── Get ALL filter parameters from request ─────────────────────
    q = request.GET.get('q', '').strip()
    date_from = request.GET.get('date_from', '')
    date_to = request.GET.get('date_to', '')
    entry_type = request.GET.get('entry_type', '')
    status = request.GET.get('status', '')
    
    # ── Base queryset ──────────────────────────────────────────────
    qs = JournalEntry.objects.select_related('created_by').prefetch_related(
        'lines__account'
    ).order_by('-date', '-number')
    
    # ── Apply all filters (same as journal_entries_list) ──────────
    if q:
        qs = qs.filter(Q(number__icontains=q) | Q(description__icontains=q))
    if date_from:
        qs = qs.filter(date__gte=date_from)
    if date_to:
        qs = qs.filter(date__lte=date_to)
    if entry_type:
        qs = qs.filter(entry_type=entry_type)
    if status:
        qs = qs.filter(status=status)
    
    # ── Create workbook ────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "اليومية الأمريكية"
    ws.sheet_view.rightToLeft = True
    
    # ── Styles ────────────────────────────────────────────────────
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bdr = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # ── Headers ───────────────────────────────────────────────────
    HEADERS = ['التاريخ', 'رقم القيد', 'البيان', 'الحساب', 'مدين (ر.س)', 'دائن (ر.س)', 'الحالة']
    NUM_COLS = len(HEADERS)
    
    # ── Row 1: Title ──────────────────────────────────────────────
    ws.merge_cells(f'A1:{get_column_letter(NUM_COLS)}1')
    title_cell = ws['A1']
    title_cell.value = "تقرير اليومية الأمريكية"
    title_cell.font = Font(size=16, bold=True, color="000000")
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.border = Border()
    ws.row_dimensions[1].height = 32
    
    # ── Row 2: Filter Information ─────────────────────────────────
    filter_parts = []
    if date_from:
        filter_parts.append(f"من تاريخ: {date_from}")
    if date_to:
        filter_parts.append(f"إلى تاريخ: {date_to}")
    if entry_type:
        type_display = dict(JournalEntry.ENTRY_TYPES).get(entry_type, entry_type)
        filter_parts.append(f"نوع القيد: {type_display}")
    if status:
        status_display = dict(JournalEntry.STATUS_CHOICES).get(status, status)
        filter_parts.append(f"الحالة: {status_display}")
    if q:
        filter_parts.append(f"بحث: {q}")
    filter_parts.append(f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    
    ws.merge_cells(f'A2:{get_column_letter(NUM_COLS)}2')
    filter_cell = ws['A2']
    filter_cell.value = "   |   ".join(filter_parts) if filter_parts else f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    filter_cell.font = Font(size=9, bold=True, color="555555")
    filter_cell.alignment = Alignment(horizontal='center', vertical='center')
    filter_cell.border = Border()
    ws.row_dimensions[2].height = 20
    
    # ── Row 3: Column Headers ─────────────────────────────────────
    for ci, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=ci)
        cell.value = h
        cell.font = Font(bold=True, size=10, color="000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bdr
    ws.row_dimensions[3].height = 22
    
    # ── Data Rows ─────────────────────────────────────────────────
    total_debit_sum = Decimal('0')
    total_credit_sum = Decimal('0')
    data_start = 4
    ri = data_start
    
    for entry in qs:
        lines = entry.lines.all()
        if not lines:
            continue
            
        for line in lines:
            vals = [
                entry.date.strftime('%Y-%m-%d') if entry.date else '',
                entry.number,
                entry.description[:50] + ('...' if len(entry.description) > 50 else ''),
                f"{line.account.code} - {line.account.name_ar}",
                float(line.debit_amount),
                float(line.credit_amount),
                entry.get_status_display(),
            ]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=v)
                cell.border = bdr
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # تنسيق الأرقام
                if isinstance(v, float):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # تلوين المدين باللون الأخضر
                    if ci == 5 and v > 0:  # عمود مدين
                        cell.font = Font(color="2E7D32", bold=True)
                    # تلوين الدائن باللون الأحمر
                    elif ci == 6 and v > 0:  # عمود دائن
                        cell.font = Font(color="C62828", bold=True)
            
            total_debit_sum += line.debit_amount
            total_credit_sum += line.credit_amount
            ri += 1
    
    # إضافة صف فارغ بين البيانات والإجمالي
    if ri > data_start:
        ri += 1
    
    # ── Total Row ─────────────────────────────────────────────────
    total_row = ri
    ws.merge_cells(
        start_row=total_row, start_column=1,
        end_row=total_row, end_column=4
    )
    
    for ci in range(1, NUM_COLS + 1):
        cell = ws.cell(row=total_row, column=ci)
        cell.fill = total_fill
        cell.border = bdr
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # خلية "الإجمالي الكلي"
    label_cell = ws.cell(row=total_row, column=1)
    label_cell.value = "الإجمالي الكلي"
    label_cell.font = Font(bold=True, size=11, color="000000")
    
    # إجمالي المدين
    debit_cell = ws.cell(row=total_row, column=5)
    debit_cell.value = float(total_debit_sum)
    debit_cell.font = Font(bold=True, size=11, color="2E7D32")
    debit_cell.number_format = '#,##0.00'
    debit_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # إجمالي الدائن
    credit_cell = ws.cell(row=total_row, column=6)
    credit_cell.value = float(total_credit_sum)
    credit_cell.font = Font(bold=True, size=11, color="C62828")
    credit_cell.number_format = '#,##0.00'
    credit_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # حالة التوازن
    balance_cell = ws.cell(row=total_row, column=7)
    is_balanced = abs(total_debit_sum - total_credit_sum) < Decimal('0.01')
    balance_cell.value = "✓ متوازن" if is_balanced else "✗ غير متوازن"
    balance_cell.font = Font(bold=True, size=10, color="2E7D32" if is_balanced else "C62828")
    balance_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.row_dimensions[total_row].height = 22
    
    # ── Additional Summary Row (Count and Balance) ────────────────
    summary_row = total_row + 1
    ws.merge_cells(f'A{summary_row}:{get_column_letter(NUM_COLS)}{summary_row}')
    summary_cell = ws.cell(row=summary_row, column=1)
    
    entry_count = qs.count()
    difference = abs(total_debit_sum - total_credit_sum)
    
    summary_parts = [
        f"عدد القيود: {entry_count}",
        f"إجمالي المدين: {float(total_debit_sum):,.2f} ر.س",
        f"إجمالي الدائن: {float(total_credit_sum):,.2f} ر.س",
    ]
    if not is_balanced:
        summary_parts.append(f"الفرق: {float(difference):,.2f} ر.س")
    
    summary_cell.value = "   |   ".join(summary_parts)
    summary_cell.font = Font(size=9, color="555555")
    summary_cell.alignment = Alignment(horizontal='center', vertical='center')
    summary_cell.border = Border()
    ws.row_dimensions[summary_row].height = 18
    
    # ── Column Widths ─────────────────────────────────────────────
    min_widths = [13, 16, 35, 30, 16, 16, 14]
    for ci, min_w in enumerate(min_widths, 1):
        col_letter = get_column_letter(ci)
        max_len = min_w
        for row in range(data_start, total_row):
            val = ws.cell(row=row, column=ci).value
            if val:
                max_len = max(max_len, len(str(val)) + 2)
        ws.column_dimensions[col_letter].width = min(max_len, 45)
    
    # ── Freeze Panes ──────────────────────────────────────────────
    ws.freeze_panes = 'A4'
    
    # ── Save and Return ───────────────────────────────────────────
    excel_file = io.BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    
    # إنشاء اسم ملف وصفي
    filename_parts = ["journal"]
    if date_from:
        filename_parts.append(f"from_{date_from}")
    if date_to:
        filename_parts.append(f"to_{date_to}")
    filename_parts.append(datetime.now().strftime('%Y%m%d_%H%M%S'))
    filename = "_".join(filename_parts) + ".xlsx"
    
    resp = HttpResponse(
        excel_file.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="{filename}"'
    return resp

# ─────────────────────────────────────────────────────────
# Financial Reports - التقارير المالية
# ─────────────────────────────────────────────────────────

@login_required
def accounting_dashboard(request):
    from decimal import Decimal

    def _cat_balance(cat_code):
        total = Decimal('0')
        for acc in Account.objects.filter(
            group__category__code=cat_code, account_type='detail', is_active=True
        ):
            total += acc.get_balance()
        return total

    total_assets      = _cat_balance('01')
    total_liabilities = _cat_balance('02')
    total_equity      = _cat_balance('03')
    total_rev         = _cat_balance('04')
    total_exp         = _cat_balance('05')
    net_income        = total_rev - total_exp

    unposted_sales     = SalesInvoice.objects.filter(zatca_status='pending').count()
    unposted_purchases = PurchaseInvoice.objects.filter(
        amount_paid=0, total__gt=0
    ).count()

    recent_entries = JournalEntry.objects.select_related(
        'created_by'
    ).order_by('-date', '-id')[:10]

    return render(request, 'invoices/accounting_dashboard.html', {
        'total_assets':       total_assets,
        'total_liab':         total_liabilities,
        'total_equity':       total_equity,
        'total_rev':          total_rev,
        'total_exp':          total_exp,
        'net_income':         net_income,
        'unposted_sales':     unposted_sales,
        'unposted_purchases': unposted_purchases,
        'recent_entries':     recent_entries,
        'total_accounts':     Account.objects.count(),
        'posted_entries':     JournalEntry.objects.filter(status='posted').count(),
        'draft_entries':      JournalEntry.objects.filter(status='draft').count(),
    })

@login_required
def trial_balance(request):
    """ميزان المراجعة - مع حركات الفترة والأرصدة الختامية"""
    from datetime import datetime, date
    from django.db.models import Sum, Q
    from decimal import Decimal
    
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to = request.GET.get('date_to', date.today().isoformat())
    export = request.GET.get('export')
    
    # تحويل التواريخ إلى كائنات date للمقارنة
    from datetime import datetime
    df_obj = datetime.strptime(date_from, '%Y-%m-%d').date() if isinstance(date_from, str) else date_from
    dt_obj = datetime.strptime(date_to, '%Y-%m-%d').date() if isinstance(date_to, str) else date_to
    
    # جلب جميع الحسابات التفصيلية النشطة
    accounts = Account.objects.filter(
        account_type='detail', 
        is_active=True
    ).select_related('group__category').order_by('code')
    
    rows = []
    total_debit_move = Decimal('0')
    total_credit_move = Decimal('0')
    total_debit_bal = Decimal('0')
    total_credit_bal = Decimal('0')
    
    for acc in accounts:
        # ── حساب حركات الفترة ─────────────────────────────────────
        period_moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=df_obj,
            journal_entry__date__lte=dt_obj
        ).aggregate(
            total_debit=Sum('debit_amount'),
            total_credit=Sum('credit_amount')
        )
        
        dr_move = period_moves['total_debit'] or Decimal('0')
        cr_move = period_moves['total_credit'] or Decimal('0')
        
        # ── حساب الرصيد الافتتاحي (قبل تاريخ البداية) ──────────────
        opening_moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__lt=df_obj
        ).aggregate(
            total_debit=Sum('debit_amount'),
            total_credit=Sum('credit_amount')
        )
        
        opening_dr = opening_moves['total_debit'] or Decimal('0')
        opening_cr = opening_moves['total_credit'] or Decimal('0')
        
        # الرصيد الافتتاحي (مع مراعاة الرصيد الافتتاحي المدخل يدوياً)
        opening_balance = acc.opening_balance or Decimal('0')
        
        if acc.normal_balance == 'debit':
            opening_bal = opening_balance + opening_dr - opening_cr
        else:
            opening_bal = opening_balance + opening_cr - opening_dr
        
        # ── حساب الرصيد الختامي ───────────────────────────────────
        if acc.normal_balance == 'debit':
            closing_bal = opening_bal + dr_move - cr_move
        else:
            closing_bal = opening_bal + cr_move - dr_move
        
        # ── تحديد إذا كان هناك حركات في الفترة ────────────────────
        has_moves = dr_move > 0 or cr_move > 0
        
        # إظهار الحساب إذا كان له رصيد أو حركات
        if closing_bal != 0 or has_moves:
            # تحديد المدين والدائن للرصيد الختامي
            if closing_bal > 0:
                bal_dr = closing_bal if acc.normal_balance == 'debit' else Decimal('0')
                bal_cr = closing_bal if acc.normal_balance == 'credit' else Decimal('0')
            else:
                bal_dr = abs(closing_bal) if acc.normal_balance == 'credit' else Decimal('0')
                bal_cr = abs(closing_bal) if acc.normal_balance == 'debit' else Decimal('0')
            
            total_debit_move += dr_move
            total_credit_move += cr_move
            total_debit_bal += bal_dr
            total_credit_bal += bal_cr
            
            rows.append({
                'account_pk': acc.pk,
                'code': acc.code,
                'name': acc.name_ar,
                'category': acc.group.category.name_ar,
                'dr_move': dr_move,
                'cr_move': cr_move,
                'bal_dr': bal_dr,
                'bal_cr': bal_cr,
                'has_moves': has_moves,
            })
    
    # التحقق من التوازن
    is_balanced = abs(total_debit_bal - total_credit_bal) < Decimal('0.01')
    
    if export == 'excel':
        return _export_trial_balance_excel_v2(rows, total_debit_move, total_credit_move, 
                                               total_debit_bal, total_credit_bal, 
                                               date_from, date_to)
    
    return render(request, 'invoices/trial_balance.html', {
        'rows': rows,
        'date_from': date_from,
        'date_to': date_to,
        'total_debit_move': total_debit_move,
        'total_credit_move': total_credit_move,
        'total_debit_bal': total_debit_bal,
        'total_credit_bal': total_credit_bal,
        'is_balanced': is_balanced,
    })


def _export_trial_balance_excel_v2(rows, total_dr_move, total_cr_move, 
                                    total_dr_bal, total_cr_bal, df, dt):
    """نسخة محسنة من تصدير ميزان المراجعة"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ميزان المراجعة"
    ws.sheet_view.rightToLeft = True
    
    hf = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    bdr = Border(*[Side(style='thin')] * 4)
    
    # العنوان الرئيسي
    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f"ميزان المراجعة من {df} إلى {dt}"
    c.font = Font(size=16, bold=True, color="C9A84C")
    c.fill = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30
    
    # رؤوس الأعمدة
    headers = ['كود الحساب', 'اسم الحساب', 'الفئة', 'حركات مدين', 'حركات دائن', 'رصيد مدين', 'رصيد دائن']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=ci, value=h)
        cell.font = Font(bold=True, size=11, color="000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bdr
    ws.row_dimensions[2].height = 22
    
    # البيانات
    for ri, row in enumerate(rows, start=3):
        vals = [
            row['code'],
            row['name'],
            row['category'],
            float(row['dr_move']) if row['dr_move'] > 0 else '',
            float(row['cr_move']) if row['cr_move'] > 0 else '',
            float(row['bal_dr']) if row['bal_dr'] > 0 else '',
            float(row['bal_cr']) if row['bal_cr'] > 0 else '',
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
    
    # صف الإجماليات
    last = len(rows) + 3
    ws.merge_cells(f'A{last}:C{last}')
    label_cell = ws.cell(row=last, column=1, value="الإجمالي الكلي")
    label_cell.font = Font(bold=True, size=12)
    label_cell.fill = header_fill
    label_cell.border = bdr
    label_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    totals = [
        float(total_dr_move), float(total_cr_move),
        float(total_dr_bal), float(total_cr_bal)
    ]
    for i, val in enumerate(totals, start=4):
        cell = ws.cell(row=last, column=i, value=val)
        cell.font = Font(bold=True, size=12)
        cell.fill = header_fill
        cell.border = bdr
        cell.number_format = '#,##0.00'
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    ws.row_dimensions[last].height = 22
    
    # عرض الأعمدة
    widths = [14, 30, 18, 16, 16, 16, 16]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    
    # تجميد الألواح
    ws.freeze_panes = 'A3'
    
    ef = io.BytesIO()
    wb.save(ef)
    ef.seek(0)
    
    resp = HttpResponse(
        ef.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="trial_balance_{df}_{dt}.xlsx"'
    return resp

@login_required
def income_statement(request):
    """قائمة الدخل (الأرباح والخسائر)"""
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to   = request.GET.get('date_to', date.today().isoformat())

    # Revenue accounts (category code 04)
    rev_accounts  = Account.objects.filter(
        group__category__code='04', account_type='detail', is_active=True)
    # Expense accounts (category code 05)
    exp_accounts  = Account.objects.filter(
        group__category__code='05', account_type='detail', is_active=True)

    revenues  = []
    expenses  = []
    total_rev = total_exp = Decimal('0')

    for acc in rev_accounts:
        bal = acc.get_balance(date_from, date_to)
        if bal != 0:
            revenues.append({'code': acc.code, 'name': acc.name_ar, 'amount': bal})
            total_rev += bal

    for acc in exp_accounts:
        bal = acc.get_balance(date_from, date_to)
        if bal != 0:
            expenses.append({'code': acc.code, 'name': acc.name_ar, 'amount': bal})
            total_exp += bal

    net_income = total_rev - total_exp

    return render(request, 'invoices/income_statement.html', {
        'revenues': revenues, 'expenses': expenses,
        'total_rev': total_rev, 'total_exp': total_exp,
        'net_income': net_income,
        'date_from': date_from, 'date_to': date_to,
    })


@login_required
def balance_sheet(request):
    as_of = request.GET.get('as_of', date.today().isoformat())

    def _get_accounts(cat_code):
        result = []
        total  = Decimal('0')
        for acc in Account.objects.filter(
            group__category__code=cat_code,
            account_type='detail', is_active=True
        ):
            bal = acc.get_balance(date_to=as_of)
            if bal != 0:
                result.append({'account': acc, 'balance': bal})
                total += bal
        return result, total

    # Assets — split into current (01-01) and fixed (01-02) by group code
    all_assets, total_assets = _get_accounts('01')
    current_assets = [r for r in all_assets if r['account'].group.code.startswith('01-01') or '01-01' in r['account'].group.code]
    fixed_assets   = [r for r in all_assets if r not in current_assets]
    intangible_assets = []

    if not current_assets and not fixed_assets:
        current_assets = all_assets

    cur_assets_total   = sum(r['balance'] for r in current_assets)
    fixed_assets_total = sum(r['balance'] for r in fixed_assets)

    all_liab, total_liabilities = _get_accounts('02')
    current_liab   = all_liab
    cur_liab_total = total_liabilities
    long_term_liab = []
    lt_liab_total  = Decimal('0')

    equity_rows, total_equity = _get_accounts('03')

    total_liab_equity = total_liabilities + total_equity
    is_balanced       = abs(total_assets - total_liab_equity) < Decimal('0.01')
    difference        = total_assets - total_liab_equity

    return render(request, 'invoices/balance_sheet.html', {
        'current_assets':    current_assets,
        'fixed_assets':      fixed_assets,
        'intangible_assets': intangible_assets,
        'cur_assets_total':  cur_assets_total,
        'fixed_assets_total': fixed_assets_total,
        'total_assets':      total_assets,
        'current_liab':      current_liab,
        'long_term_liab':    long_term_liab,
        'cur_liab_total':    cur_liab_total,
        'lt_liab_total':     lt_liab_total,
        'total_liabilities': total_liabilities,
        'equity_rows':       equity_rows,
        'total_equity':      total_equity,
        'total_liab_equity': total_liab_equity,
        'is_balanced':       is_balanced,
        'difference':        difference,
        'as_of':             as_of,
    })

# ─────────────────────────────────────────────────────────
# Fiscal Years - السنوات المالية
# ─────────────────────────────────────────────────────────

@login_required
def fiscal_years_list(request):
    company = Company.objects.first()
    fy_list = FiscalYear.objects.filter(company=company).order_by('-start_date') if company else []
    return render(request, 'invoices/fiscal_years_list.html', {
        'fiscal_years': fy_list, 'company': company
    })


@login_required
def fiscal_year_add(request):
    deny = _perm(request, 'invoices.add_fiscalyear')
    if deny: return deny
    form = FiscalYearForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        fy = form.save(commit=False)
        fy.company = Company.objects.first()
        fy.save()
        messages.success(request, f'تم إضافة السنة المالية {fy.name}')
        return redirect('fiscal_years_list')
    return render(request, 'invoices/fiscal_year_form.html', {
        'form': form, 'title': 'إضافة سنة مالية'
    })

# ─────────────────────────────────────────────────────────
# Fiscal Years - Additional Views
# ─────────────────────────────────────────────────────────

@login_required
def fiscal_year_edit(request, pk):
    """تعديل سنة مالية"""
    deny = _perm(request, 'invoices.change_fiscalyear')
    if deny: 
        return deny
    
    fy = get_object_or_404(FiscalYear, pk=pk)
    
    if fy.is_closed:
        messages.error(request, 'لا يمكن تعديل سنة مالية مغلقة.')
        return redirect('fiscal_years_list')
    
    if request.method == 'POST':
        form = FiscalYearForm(request.POST, instance=fy)
        if form.is_valid():
            form.save()
            messages.success(request, f'تم تعديل السنة المالية {fy.name}')
            return redirect('fiscal_years_list')
    else:
        form = FiscalYearForm(instance=fy)
    
    return render(request, 'invoices/fiscal_year_form.html', {
        'form': form,
        'title': f'تعديل سنة مالية: {fy.name}',
        'button_text': 'حفظ التعديلات',
        'fiscal_year': fy,
    })


@login_required
def fiscal_year_delete(request, pk):
    """حذف سنة مالية"""
    deny = _perm(request, 'invoices.delete_fiscalyear')
    if deny: 
        return deny
    
    fy = get_object_or_404(FiscalYear, pk=pk)
    
    # Check if there are any journal entries in this fiscal year
    entries_count = JournalEntry.objects.filter(
        date__gte=fy.start_date,
        date__lte=fy.end_date,
        status='posted'
    ).count()
    
    if entries_count > 0:
        messages.error(request, f'لا يمكن حذف السنة المالية لوجود {entries_count} قيد مرحل فيها.')
        return redirect('fiscal_years_list')
    
    if request.method == 'POST':
        fy_name = fy.name
        fy.delete()
        messages.success(request, f'تم حذف السنة المالية {fy_name}')
        return redirect('fiscal_years_list')
    
    return render(request, 'invoices/fiscal_year_confirm_delete.html', {
        'fiscal_year': fy,
        'entries_count': entries_count,
    })


@login_required
def fiscal_year_close(request, pk):
    """إغلاق سنة مالية"""
    deny = _perm(request, 'invoices.close_fiscalyear')
    if deny: 
        return deny
    
    fy = get_object_or_404(FiscalYear, pk=pk)
    
    if fy.is_closed:
        messages.warning(request, f'السنة المالية {fy.name} مغلقة بالفعل.')
        return redirect('fiscal_years_list')
    
    # Check if there are any unposted entries
    draft_entries = JournalEntry.objects.filter(
        date__gte=fy.start_date,
        date__lte=fy.end_date,
        status='draft'
    ).count()
    
    if request.method == 'POST':
        # Perform year-end closing
        try:
            with transaction.atomic():
                # Post all remaining draft entries if requested
                if request.POST.get('post_drafts') == 'yes':
                    JournalEntry.objects.filter(
                        date__gte=fy.start_date,
                        date__lte=fy.end_date,
                        status='draft'
                    ).update(status='posted', posted_at=timezone.now())
                
                # Calculate net income for the year
                # Revenue - Expenses
                from django.db.models import Sum
                
                revenues = JournalEntryLine.objects.filter(
                    account__group__category__code='04',
                    journal_entry__status='posted',
                    journal_entry__date__gte=fy.start_date,
                    journal_entry__date__lte=fy.end_date,
                ).aggregate(
                    d=Sum('debit_amount'),
                    c=Sum('credit_amount')
                )
                total_rev = (revenues['c'] or 0) - (revenues['d'] or 0)
                
                expenses = JournalEntryLine.objects.filter(
                    account__group__category__code='05',
                    journal_entry__status='posted',
                    journal_entry__date__gte=fy.start_date,
                    journal_entry__date__lte=fy.end_date,
                ).aggregate(
                    d=Sum('debit_amount'),
                    c=Sum('credit_amount')
                )
                total_exp = (expenses['d'] or 0) - (expenses['c'] or 0)
                
                net_income = total_rev - total_exp
                
                # Close the fiscal year
                fy.is_closed = True
                fy.closed_at = timezone.now()
                fy.save()
                
                # Optionally: Create closing entry to transfer net income to retained earnings
                if request.POST.get('create_closing_entry') == 'yes' and net_income != 0:
                    retained_earnings = Account.objects.filter(
                        code='030201',  # أرباح محتجزة
                        account_type='detail'
                    ).first()
                    
                    income_summary = Account.objects.filter(
                        code__startswith='03',
                        name_ar__icontains='ملخص الدخل'
                    ).first()
                    
                    if retained_earnings:
                        user = request.user
                        closing_entry = JournalEntry.objects.create(
                            number=f"CLOSE-{fy.name}",
                            date=fy.end_date,
                            entry_type='adjustment',
                            status='posted',
                            description=f'قيد إقفال السنة المالية {fy.name} - صافي {"الربح" if net_income >= 0 else "الخسارة"}',
                            created_by=user,
                            posted_at=timezone.now()
                        )
                        
                        if net_income > 0:
                            # Profit: Debit Income Summary, Credit Retained Earnings
                            JournalEntryLine.objects.create(
                                journal_entry=closing_entry,
                                account=income_summary or retained_earnings,
                                description='إقفال صافي الربح',
                                debit_amount=net_income,
                                credit_amount=0
                            )
                            JournalEntryLine.objects.create(
                                journal_entry=closing_entry,
                                account=retained_earnings,
                                description='زيادة الأرباح المحتجزة',
                                debit_amount=0,
                                credit_amount=net_income
                            )
                        else:
                            # Loss: Debit Retained Earnings, Credit Income Summary
                            JournalEntryLine.objects.create(
                                journal_entry=closing_entry,
                                account=retained_earnings,
                                description='تخفيض الأرباح المحتجزة',
                                debit_amount=abs(net_income),
                                credit_amount=0
                            )
                            JournalEntryLine.objects.create(
                                journal_entry=closing_entry,
                                account=income_summary or retained_earnings,
                                description='إقفال صافي الخسارة',
                                debit_amount=0,
                                credit_amount=abs(net_income)
                            )
                
                messages.success(request, f'تم إغلاق السنة المالية {fy.name} بنجاح. صافي {"الربح" if net_income >= 0 else "الخسارة"}: {abs(net_income):,.2f} ر.س')
                
        except Exception as e:
            messages.error(request, f'خطأ في إغلاق السنة المالية: {str(e)}')
        
        return redirect('fiscal_years_list')
    
    return render(request, 'invoices/fiscal_year_close_confirm.html', {
        'fiscal_year': fy,
        'draft_entries': draft_entries,
    })


@login_required
def fiscal_year_detail(request, pk):
    """تفاصيل السنة المالية"""
    fy = get_object_or_404(FiscalYear, pk=pk)
    
    # Get statistics for this fiscal year
    from django.db.models import Sum, Count
    from django.db.models.functions import TruncMonth
    
    entries = JournalEntry.objects.filter(
        date__gte=fy.start_date,
        date__lte=fy.end_date
    )
    
    stats = {
        'total_entries': entries.count(),
        'posted_entries': entries.filter(status='posted').count(),
        'draft_entries': entries.filter(status='draft').count(),
        'reversed_entries': entries.filter(status='reversed').count(),
    }
    
    # Get monthly totals
    monthly_totals = entries.filter(status='posted').annotate(
        month=TruncMonth('date')
    ).values('month').annotate(
        count=Count('id'),
        debit=Sum('lines__debit_amount'),
        credit=Sum('lines__credit_amount')
    ).order_by('month')
    
    # Get recent entries
    recent_entries = entries.select_related('created_by').order_by('-date', '-id')[:10]
    
    return render(request, 'invoices/fiscal_year_detail.html', {
        'fiscal_year': fy,
        'stats': stats,
        'monthly_totals': monthly_totals,
        'recent_entries': recent_entries,
    })
# ─────────────────────────────────────────────────────────
# AJAX helpers
# ─────────────────────────────────────────────────────────

@login_required
def ajax_account_balance(request, pk):
    """Return current balance for an account as JSON."""
    account   = get_object_or_404(Account, pk=pk)
    date_from = request.GET.get('date_from')
    date_to   = request.GET.get('date_to')
    balance   = account.get_balance(date_from, date_to)
    return JsonResponse({
        'account_code': account.code,
        'account_name': account.name_ar,
        'balance':      float(balance),
        'normal_balance': account.normal_balance,
    })
    
    
# for testing only :
# قبل إرسال فواتير حقيقية، يجب اختبار التوافق
# أضف هذا الكود مؤقتاً للاختبار:

# أضف هذه الدالة في views.py بعد دوال ZATCA

@login_required
def zatca_compliance_test(request, pk):
    """اختبار توافق الفاتورة مع متطلبات ZATCA قبل الإرسال"""
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny: return deny
    
    inv = get_object_or_404(SalesInvoice, pk=pk)
    device = ZATCADevice.objects.filter(company=inv.company, is_active=True).first()
    
    if not device:
        messages.error(request, 'لا يوجد جهاز زاتكا مُفعَّل.')
        return redirect('sales_detail', pk=pk)
    
    try:
        from .zatca_service import ZATCAInvoiceService
        service = ZATCAInvoiceService(device)
        
        # توليد XML بدون إرسال
        xml_str = service.generator.generate(inv)
        
        # التوقيع
        signed_xml, invoice_hash = service.signer.sign_xml(xml_str)
        
        # فحص التوافق
        result = service.client.compliance_check(inv, signed_xml, invoice_hash)
        
        if result['success']:
            messages.success(request, '✅ الفاتورة متوافقة مع متطلبات زاتكا')
            # عرض تفاصيل التحقق
            return render(request, 'invoices/zatca_compliance_result.html', {
                'invoice': inv,
                'result': result['data'],
                'xml_preview': xml_str[:500],
                'hash': invoice_hash,
            })
        else:
            messages.error(request, f'❌ الفاتورة غير متوافقة: {result["data"]}')
            return render(request, 'invoices/zatca_compliance_result.html', {
                'invoice': inv,
                'error': result['data'],
                'status_code': result['status_code'],
            })
            
    except Exception as e:
        messages.error(request, f'خطأ في فحص التوافق: {str(e)}')
        log.exception('Compliance check failed for invoice %s', inv.invoice_number)
    
    return redirect('sales_detail', pk=pk)

@login_required
def zatca_view_xml(request, pk):
    """عرض XML للفاتورة للتحقق منها"""
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny: return deny
    
    inv = get_object_or_404(SalesInvoice, pk=pk)
    device = ZATCADevice.objects.filter(company=inv.company, is_active=True).first()
    
    if not device:
        messages.error(request, 'لا يوجد جهاز زاتكا مُفعَّل.')
        return redirect('sales_detail', pk=pk)
    
    try:
        from .zatca_service import ZATCAInvoiceService
        service = ZATCAInvoiceService(device)
        
        # توليد XML
        xml_str = service.generator.generate(inv)
        
        # تنسيق XML للعرض
        from lxml import etree
        root = etree.fromstring(xml_str.encode())
        pretty_xml = etree.tostring(root, pretty_print=True, encoding='unicode')
        
        return render(request, 'invoices/zatca_xml_view.html', {
            'invoice': inv,
            'xml_content': pretty_xml,
        })
        
    except Exception as e:
        messages.error(request, f'خطأ في توليد XML: {str(e)}')
        return redirect('sales_detail', pk=pk)
    
# أضف هذه الدالة في views.py لاختبار QR

@login_required
def test_qr_generation(request, pk):
    """اختبار توليد QR Code مع التحقق"""
    inv = get_object_or_404(SalesInvoice, pk=pk)
    
    from .zatca_service import ImprovedZATCAQRGenerator, ZATCAValidationError
    
    qr_generator = ImprovedZATCAQRGenerator()
    
    try:
        # اختبار التحقق
        qr_generator.validate_tlv_data(inv)
        
        # توليد QR
        qr_text = qr_generator.generate_with_validation(inv)
        
        # إنشاء صورة QR
        import qrcode
        import io
        import base64
        
        qr = qrcode.QRCode(version=2, box_size=4, border=1)
        qr.add_data(qr_text)
        qr.make(fit=True)
        img = qr.make_image(fill_color='black', back_color='white')
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        qr_image = base64.b64encode(buf.getvalue()).decode()
        
        return render(request, 'invoices/test_qr.html', {
            'invoice': inv,
            'qr_image': qr_image,
            'qr_data': qr_text,
            'validation_passed': True,
        })
        
    except ZATCAValidationError as e:
        return render(request, 'invoices/test_qr.html', {
            'invoice': inv,
            'validation_passed': False,
            'validation_errors': str(e),
        })
        
"""
Patch for views.py — replace these functions with the corrected versions.
The main issues fixed:
1. zatca_compliance_test used signer.sign_xml() — method is now signer.sign()
2. export_invoice_xml must set invoice_counter_value before generating XML
3. zatca_view_xml uses correct generator
"""

# ── zatca_compliance_test (replace the existing one) ──────────────────────

@login_required
def zatca_compliance_test(request, pk):
    """اختبار توافق الفاتورة مع متطلبات ZATCA قبل الإرسال الرسمي."""
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny:
        return deny

    inv    = get_object_or_404(SalesInvoice, pk=pk)
    device = ZATCADevice.objects.filter(company=inv.company, is_active=True).first()

    if not device:
        messages.error(request, 'لا يوجد جهاز زاتكا مُفعَّل.')
        return redirect('sales_detail', pk=pk)

    if not device.csid or not device.certificate:
        messages.error(request, 'بيانات جهاز زاتكا غير مكتملة (CSID أو الشهادة مفقودة).')
        return redirect('sales_detail', pk=pk)

    try:
        from .zatca_service import (
            ZATCAXMLGenerator, ZATCASigner, ZATCAQRGenerator,
            ZATCAAPIClient, validate_invoice,
        )

        # Pre-flight
        validate_invoice(inv)

        # Assign counter and chain hash (mirrors production flow)
        from .zatca_service import (
            ZATCAXMLGenerator, ZATCASigner, ZATCAQRGenerator,
            ZATCAAPIClient, ZATCAInvoiceService, validate_invoice,
        )
        ZATCAInvoiceService(device)._assign_counter_and_hash(inv)

        # Generate
        qr_gen  = ZATCAQRGenerator()
        gen     = ZATCAXMLGenerator()
        signer  = ZATCASigner(device.private_key, device.certificate)
        client  = ZATCAAPIClient(device.csid, device.secret, device.environment)

        qr_code             = qr_gen.generate(inv)
        xml_str             = gen.generate(inv)
        signed_xml, inv_hash = signer.sign(xml_str, qr_code)

        result = client.compliance_check(inv, signed_xml, inv_hash)

        return render(request, 'invoices/zatca_compliance_result.html', {
            'invoice':     inv,
            'result':      result.get('data'),
            'success':     result.get('success'),
            'status_code': result.get('status_code'),
            'xml_preview': xml_str[:1000],
            'hash':        inv_hash,
        })

    except Exception as exc:
        messages.error(request, f'خطأ في فحص التوافق: {exc}')
        log.exception('Compliance check failed for invoice %s', inv.invoice_number)
        return redirect('sales_detail', pk=pk)


# ── export_invoice_xml (replace the existing one) ─────────────────────────

@login_required
def export_invoice_xml(request, pk):
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny:
        return deny
 
    invoice = get_object_or_404(SalesInvoice, pk=pk)
 
    from django.db import transaction as db_transaction
    from .zatca_service import ZATCAXMLGenerator, ZATCA_GENESIS_HASH
 
    changed = False
 
    # Ensure UUID
    if not invoice.invoice_uuid:
        import uuid as _uuid
        invoice.invoice_uuid = _uuid.uuid4()
        changed = True
 
    # Atomically assign counter + chain hash (same logic as the service)
    with db_transaction.atomic():
        locked = SalesInvoice.objects.select_for_update().get(pk=invoice.pk)
 
        if not locked.invoice_counter_value or locked.invoice_counter_value < 1:
            last = (SalesInvoice.objects
                    .filter(transaction_type=invoice.transaction_type)
                    .exclude(pk=invoice.pk)
                    .filter(invoice_counter_value__gte=1)
                    .select_for_update()
                    .order_by('-invoice_counter_value')
                    .first())
            locked.invoice_counter_value = (last.invoice_counter_value + 1) if last else 1
            changed = True
 
        if not locked.previous_invoice_hash:
            prev = (SalesInvoice.objects
                    .filter(transaction_type=invoice.transaction_type,
                            zatca_invoice_hash__isnull=False)
                    .exclude(zatca_invoice_hash='')
                    .exclude(pk=invoice.pk)
                    .order_by('-invoice_counter_value')
                    .first())
            locked.previous_invoice_hash = prev.zatca_invoice_hash if prev else ZATCA_GENESIS_HASH
            changed = True
 
        if changed:
            update_fields = ['invoice_counter_value', 'previous_invoice_hash']
            if not invoice.invoice_uuid or str(invoice.invoice_uuid) != str(locked.invoice_uuid):
                locked.invoice_uuid = invoice.invoice_uuid
                update_fields.append('invoice_uuid')
            locked.save(update_fields=update_fields)
 
        invoice.invoice_counter_value = locked.invoice_counter_value
        invoice.previous_invoice_hash = locked.previous_invoice_hash
 
    try:
        xml_str  = ZATCAXMLGenerator().generate(invoice)
        filename = f"invoice_{invoice.invoice_number}.xml"
        response = HttpResponse(xml_str, content_type='application/xml; charset=utf-8')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
 
    except Exception as exc:
        messages.error(request, f'خطأ في تصدير XML: {exc}')
        log.exception('XML export failed for invoice %s', invoice.invoice_number)
        return redirect('sales_detail', pk=pk)
 
 
# ─────────────────────────────────────────────────────────
# [FIX-5] zatca_view_xml — confirmed correct, no changes.
# Included here for completeness.
# ─────────────────────────────────────────────────────────
 
@login_required
def zatca_view_xml(request, pk):
    """عرض XML للفاتورة في المتصفح (للمراجعة)."""
    deny = check_perm(request, 'invoices.submit_zatca')
    if deny:
        return deny
 
    invoice = get_object_or_404(SalesInvoice, pk=pk)
 
    try:
        from .zatca_service import ZATCAXMLGenerator
        from lxml import etree
 
        xml_str    = ZATCAXMLGenerator().generate(invoice)
        root       = etree.fromstring(xml_str.encode())
        pretty_xml = etree.tostring(root, pretty_print=True, encoding='unicode')
 
        return render(request, 'invoices/zatca_xml_view.html', {
            'invoice':     invoice,
            'xml_content': pretty_xml,
        })
 
    except Exception as exc:
        messages.error(request, f'خطأ في توليد XML: {exc}')
        return redirect('sales_detail', pk=pk)
    
    
    
"""
=============================================================
 views_additions.py — إضافات لملف views.py الرئيسي
 1. حساب الأستاذ العام (General Ledger)
 2. ميزان المراجعة المحسّن (4 أعمدة)
 3. تصدير Excel لجميع التقارير
 4. كشف حساب العميل/المورد
=============================================================
ADD THESE IMPORTS TO views.py:
    from .models import FixedAsset, AssetCategory, PaymentInvoiceAllocation
"""

from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, JsonResponse
from django.db.models import Sum, Q, F
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.core.paginator import Paginator
from decimal import Decimal
from datetime import date, timedelta
import io, openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .models import (
    Account, AccountCategory, AccountGroup,
    JournalEntry, JournalEntryLine,
    SalesInvoice, PurchaseInvoice,
    Client, Supplier, Company,
    Payment, PaymentInvoiceAllocation,
    CostCenter,
)

# ─────────────────────────────────────────────────────────
# حساب الأستاذ العام - General Ledger
# ─────────────────────────────────────────────────────────

"""
=============================================================
 views_additions.py — إضافات لملف views.py الرئيسي
 1. حساب الأستاذ العام (General Ledger)
 2. ميزان المراجعة المحسّن (4 أعمدة)
 3. تصدير Excel لجميع التقارير
 4. كشف حساب العميل/المورد
=============================================================
ADD THESE IMPORTS TO views.py:
    from .models import FixedAsset, AssetCategory, PaymentInvoiceAllocation
"""

from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, JsonResponse
from django.db.models import Sum, Q, F
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.core.paginator import Paginator
from decimal import Decimal
from datetime import date, timedelta
import io, openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .models import (
    Account, AccountCategory, AccountGroup,
    JournalEntry, JournalEntryLine,
    SalesInvoice, PurchaseInvoice,
    Client, Supplier, Company,
    Payment, PaymentInvoiceAllocation,
    CostCenter,
)

# ─────────────────────────────────────────────────────────
# حساب الأستاذ العام - General Ledger
# ─────────────────────────────────────────────────────────

@login_required
def general_ledger(request):
    """
    حساب الأستاذ العام — يعرض جميع الحسابات التفصيلية مع حركاتها
    خلال فترة زمنية محددة مع الأرصدة الجارية
    """
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to   = request.GET.get('date_to',   date.today().isoformat())
    account_id = request.GET.get('account_id', '')
    cat_code   = request.GET.get('category', '')
    export     = request.GET.get('export', '')

    # بناء قائمة الحسابات
    accounts_qs = Account.objects.filter(
        account_type='detail', is_active=True
    ).select_related('group__category').order_by('code')

    if cat_code:
        accounts_qs = accounts_qs.filter(group__category__code=cat_code)
    if account_id:
        accounts_qs = accounts_qs.filter(pk=account_id)

    ledger_data = []

    for account in accounts_qs:
        # رصيد أول المدة: كل الحركات قبل date_from
        opening_lines = JournalEntryLine.objects.filter(
            account=account,
            journal_entry__status='posted',
            journal_entry__date__lt=date_from,
        ).aggregate(
            d=Sum('debit_amount'),
            c=Sum('credit_amount'),
        )
        ob_debit  = opening_lines['d'] or Decimal('0')
        ob_credit = opening_lines['c'] or Decimal('0')
        ob_base   = account.opening_balance or Decimal('0')

        if account.normal_balance == 'debit':
            opening_balance = ob_base + ob_debit - ob_credit
        else:
            opening_balance = ob_base + ob_credit - ob_debit

        # حركات الفترة
        period_lines = (
            JournalEntryLine.objects
            .filter(
                account=account,
                journal_entry__status='posted',
                journal_entry__date__gte=date_from,
                journal_entry__date__lte=date_to,
            )
            .select_related('journal_entry')
            .order_by('journal_entry__date', 'journal_entry__id')
        )

        if not period_lines.exists() and opening_balance == 0:
            continue  # تخطي الحسابات الصفرية

        # بناء الحركات مع الرصيد الجاري
        entries       = []
        running       = opening_balance
        period_debit  = Decimal('0')
        period_credit = Decimal('0')

        for line in period_lines:
            if account.normal_balance == 'debit':
                running += line.debit_amount - line.credit_amount
            else:
                running += line.credit_amount - line.debit_amount

            period_debit  += line.debit_amount
            period_credit += line.credit_amount

            entries.append({
                'date':           line.journal_entry.date,
                'entry_id':       line.journal_entry.pk,
                'entry_number':   line.journal_entry.number,
                'description':    line.description or line.journal_entry.description,
                'debit':          line.debit_amount,
                'credit':         line.credit_amount,
                'running_balance': running,
            })

        closing_balance = running

        ledger_data.append({
            'account':          account,
            'opening_balance':  opening_balance,
            'entries':          entries,
            'period_debit':     period_debit,
            'period_credit':    period_credit,
            'closing_balance':  closing_balance,
        })

    # إجماليات كلية
    grand_opening = sum(x['opening_balance'] for x in ledger_data)
    grand_debit   = sum(x['period_debit']    for x in ledger_data)
    grand_credit  = sum(x['period_credit']   for x in ledger_data)
    grand_closing = sum(x['closing_balance'] for x in ledger_data)

    if export == 'excel':
        return _export_general_ledger_excel(
            ledger_data, date_from, date_to,
            grand_opening, grand_debit, grand_credit, grand_closing
        )

    return render(request, 'invoices/general_ledger.html', {
        'ledger_data':    ledger_data,
        'date_from':      date_from,
        'date_to':        date_to,
        'account_id':     account_id,
        'cat_code':       cat_code,
        'categories':     AccountCategory.objects.order_by('code'),
        'accounts_list':  Account.objects.filter(account_type='detail', is_active=True).order_by('code'),
        'grand_opening':  grand_opening,
        'grand_debit':    grand_debit,
        'grand_credit':   grand_credit,
        'grand_closing':  grand_closing,
    })


def _export_general_ledger_excel(ledger_data, date_from, date_to,
                                  grand_opening, grand_debit, grand_credit, grand_closing):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "حساب الأستاذ"
    ws.sheet_view.rightToLeft = True

    # أنماط
    hdr_fill  = PatternFill("solid", fgColor="1a1a2e")
    acc_fill  = PatternFill("solid", fgColor="16213e")
    open_fill = PatternFill("solid", fgColor="f0f4ff")
    close_fill= PatternFill("solid", fgColor="fff3e0")
    total_fill= PatternFill("solid", fgColor="0f3460")
    bdr = Border(
        left=Side(style='thin'),  right=Side(style='thin'),
        top=Side(style='thin'),   bottom=Side(style='thin'),
    )
    hdr_font   = Font(bold=True, color="C9A84C", size=10)
    acc_font   = Font(bold=True, color="C9A84C", size=11)
    white_font = Font(bold=True, color="FFFFFF",  size=10)

    # ──────────────── صف العنوان ────────────────
    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f"حساب الأستاذ العام  |  من {date_from}  إلى {date_to}"
    c.font  = Font(bold=True, color="C9A84C", size=14)
    c.fill  = hdr_fill
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ri = 2  # الصف الحالي

    for block in ledger_data:
        acc = block['account']

        # ──────── رأس الحساب ────────
        ws.merge_cells(f'A{ri}:G{ri}')
        hc = ws[f'A{ri}']
        hc.value = f"{acc.code}  —  {acc.name_ar}  |  {acc.name_en}"
        hc.font  = acc_font
        hc.fill  = acc_fill
        hc.alignment = Alignment(horizontal='right', vertical='center')
        ws.row_dimensions[ri].height = 22
        ri += 1

        # ──────── رؤوس الأعمدة ────────
        headers = ['التاريخ', 'رقم القيد', 'البيان', 'مدين', 'دائن', 'الرصيد', '']
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=ri, column=ci, value=h)
            cell.font      = hdr_font
            cell.fill      = hdr_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border    = bdr
        ws.row_dimensions[ri].height = 18
        ri += 1

        # ──────── رصيد أول المدة ────────
        ws.merge_cells(f'A{ri}:C{ri}')
        ob_cell = ws[f'A{ri}']
        ob_cell.value  = "رصيد أول المدة"
        ob_cell.font   = Font(bold=True, size=9)
        ob_cell.fill   = open_fill
        ob_cell.border = bdr
        ob_cell.alignment = Alignment(horizontal='center')

        for ci in range(4, 8):
            ws.cell(row=ri, column=ci).fill   = open_fill
            ws.cell(row=ri, column=ci).border = bdr

        bal_cell = ws.cell(row=ri, column=6, value=float(block['opening_balance']))
        bal_cell.font         = Font(bold=True, size=9)
        bal_cell.fill         = open_fill
        bal_cell.number_format = '#,##0.00'
        bal_cell.alignment    = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 16
        ri += 1

        # ──────── بنود الحركات ────────
        for entry in block['entries']:
            vals = [
                str(entry['date']),
                entry['entry_number'],
                entry['description'],
                float(entry['debit'])   if entry['debit']  else None,
                float(entry['credit'])  if entry['credit'] else None,
                float(entry['running_balance']),
                '',
            ]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=v)
                cell.border    = bdr
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font      = Font(size=9)
                if isinstance(v, float):
                    cell.number_format = '#,##0.00'
                    if ci == 4:  # مدين
                        cell.font = Font(color="1B5E20", size=9)
                    elif ci == 5:  # دائن
                        cell.font = Font(color="B71C1C", size=9)
                    elif ci == 6:  # رصيد
                        cell.font = Font(bold=True, size=9,
                                         color="0D47A1" if v >= 0 else "B71C1C")
            ws.row_dimensions[ri].height = 15
            ri += 1

        # ──────── إجمالي الحساب ────────
        ws.merge_cells(f'A{ri}:C{ri}')
        tot_label = ws[f'A{ri}']
        tot_label.value = "إجمالي الفترة"
        tot_label.font  = Font(bold=True, color="FFFFFF", size=9)
        tot_label.fill  = total_fill
        tot_label.border= bdr
        tot_label.alignment = Alignment(horizontal='center')

        for ci, v in enumerate([float(block['period_debit']), float(block['period_credit']),
                                  float(block['closing_balance']), ''], 1):
            cell = ws.cell(row=ri, column=ci+3, value=v if v != '' else '')
            cell.font   = Font(bold=True, color="C9A84C", size=9)
            cell.fill   = total_fill
            cell.border = bdr
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 18
        ri += 1

        # فاصل
        ws.row_dimensions[ri].height = 6
        ri += 1

    # ──────── إجمالي عام ────────
    ws.merge_cells(f'A{ri}:C{ri}')
    grand = ws[f'A{ri}']
    grand.value     = "الإجمالي العام"
    grand.font      = Font(bold=True, color="C9A84C", size=11)
    grand.fill      = hdr_fill
    grand.border    = bdr
    grand.alignment = Alignment(horizontal='center')

    for ci, v in enumerate([float(grand_debit), float(grand_credit), float(grand_closing), ''], 1):
        cell = ws.cell(row=ri, column=ci+3, value=v if v != '' else '')
        cell.font      = Font(bold=True, color="C9A84C", size=11)
        cell.fill      = hdr_fill
        cell.border    = bdr
        cell.number_format = '#,##0.00'
        cell.alignment = Alignment(horizontal='center')
    ws.row_dimensions[ri].height = 24

    # ──────── عرض الأعمدة ────────
    col_widths = [13, 16, 38, 14, 14, 14, 4]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="general_ledger_{date_from}_{date_to}.xlsx"'
    return resp


# ─────────────────────────────────────────────────────────
# ميزان المراجعة المحسّن (4 أعمدة: حركات + أرصدة)
# ─────────────────────────────────────────────────────────

@login_required
def trial_balance_v2(request):
    """
    ميزان المراجعة بـ 4 أعمدة:
    حركات الفترة (مدين/دائن) + أرصدة الختام (مدين/دائن)
    """
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to   = request.GET.get('date_to',   date.today().isoformat())
    export    = request.GET.get('export', '')

    accounts = (
        Account.objects
        .filter(account_type='detail', is_active=True)
        .select_related('group__category')
        .order_by('code')
    )

    rows = []
    total_dr_move = total_cr_move = Decimal('0')
    total_dr_bal  = total_cr_bal  = Decimal('0')

    for acc in accounts:
        agg = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=date_from,
            journal_entry__date__lte=date_to,
        ).aggregate(d=Sum('debit_amount'), c=Sum('credit_amount'))

        dr_move = agg['d'] or Decimal('0')
        cr_move = agg['c'] or Decimal('0')

        if dr_move == 0 and cr_move == 0:
            continue

        ob = acc.opening_balance or Decimal('0')
        net = ob + (dr_move - cr_move if acc.normal_balance == 'debit' else cr_move - dr_move)

        bal_dr = max(net,  Decimal('0')) if acc.normal_balance == 'debit'  else max(-net, Decimal('0'))
        bal_cr = max(-net, Decimal('0')) if acc.normal_balance == 'debit'  else max(net,  Decimal('0'))

        total_dr_move += dr_move
        total_cr_move += cr_move
        total_dr_bal  += bal_dr
        total_cr_bal  += bal_cr

        rows.append({
            'code':    acc.code,
            'name':    acc.name_ar,
            'category': acc.group.category.name_ar,
            'account': acc,
            'account_pk': acc.pk,          # explicit pk — avoids NoReverseMatch
            'dr_move': dr_move,
            'cr_move': cr_move,
            'bal_dr':  bal_dr,
            'bal_cr':  bal_cr,
        })

    if export == 'excel':
        return _export_trial_balance_v2_excel(
            rows, date_from, date_to,
            total_dr_move, total_cr_move, total_dr_bal, total_cr_bal
        )

    return render(request, 'invoices/trial_balance.html', {
        'rows':           rows,
        'date_from':      date_from,
        'date_to':        date_to,
        'total_debit_move':  total_dr_move,
        'total_credit_move': total_cr_move,
        'total_debit_bal':   total_dr_bal,
        'total_credit_bal':  total_cr_bal,
        'is_balanced':    abs(total_dr_bal - total_cr_bal) < Decimal('0.01'),
    })


def _export_trial_balance_v2_excel(rows, date_from, date_to,
                                    total_dr_move, total_cr_move,
                                    total_dr_bal,  total_cr_bal):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "ميزان المراجعة"
    ws.sheet_view.rightToLeft = True

    hf  = PatternFill("solid", fgColor="1a1a2e")
    tf  = PatternFill("solid", fgColor="16213e")
    bdr = Border(
        left=Side(style='thin'),  right=Side(style='thin'),
        top=Side(style='thin'),   bottom=Side(style='thin'),
    )

    # عنوان
    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f"ميزان المراجعة  |  من {date_from}  إلى {date_to}"
    c.font  = Font(bold=True, color="C9A84C", size=13)
    c.fill  = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # رؤوس مدمجة
    merge_headers = [
        ('A2', 'A3', 'كود الحساب'),
        ('B2', 'B3', 'اسم الحساب'),
        ('C2', 'C3', 'الفئة'),
        ('D2', 'E2', 'حركات الفترة'),
        ('F2', 'G2', 'الأرصدة'),
    ]
    for start, end, title in merge_headers:
        ws.merge_cells(f'{start}:{end}')
        cell = ws[start]
        cell.value     = title
        cell.font      = Font(bold=True, color="C9A84C", size=10)
        cell.fill      = hf
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border    = bdr

    sub_headers = [('D3','مدين'), ('E3','دائن'), ('F3','مدين'), ('G3','دائن')]
    for addr, label in sub_headers:
        cell = ws[addr]
        cell.value     = label
        cell.font      = Font(bold=True, color="C9A84C", size=9)
        cell.fill      = hf
        cell.alignment = Alignment(horizontal='center')
        cell.border    = bdr

    for ci in ['A','B','C']:
        ws[f'{ci}3'].border = bdr
        ws[f'{ci}3'].fill   = hf

    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 18

    # بيانات
    for ri, row in enumerate(rows, start=4):
        vals = [
            row['code'], row['name'], row['category'],
            float(row['dr_move']) if row['dr_move'] else None,
            float(row['cr_move']) if row['cr_move'] else None,
            float(row['bal_dr'])  if row['bal_dr']  else None,
            float(row['bal_cr'])  if row['bal_cr']  else None,
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border    = bdr
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font      = Font(size=9)
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
        ws.row_dimensions[ri].height = 15

    # إجماليات
    last = len(rows) + 4
    ws.merge_cells(f'A{last}:C{last}')
    lc = ws[f'A{last}']
    lc.value     = "الإجمالي"
    lc.font      = Font(bold=True, color="C9A84C", size=10)
    lc.fill      = tf
    lc.border    = bdr
    lc.alignment = Alignment(horizontal='center')

    for ci, v in enumerate([float(total_dr_move), float(total_cr_move),
                              float(total_dr_bal),  float(total_cr_bal)], start=4):
        cell = ws.cell(row=last, column=ci, value=v)
        cell.font          = Font(bold=True, color="C9A84C", size=10)
        cell.fill          = tf
        cell.border        = bdr
        cell.number_format = '#,##0.00'
        cell.alignment     = Alignment(horizontal='center')
    ws.row_dimensions[last].height = 22

    col_widths = [14, 32, 18, 14, 14, 14, 14]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = 'A4'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="trial_balance_{date_from}_{date_to}.xlsx"'
    return resp


# ─────────────────────────────────────────────────────────
# قائمة الدخل — تصدير Excel
# ─────────────────────────────────────────────────────────

@login_required
def income_statement_v2(request):
    """قائمة الدخل (الأرباح والخسائر)"""
    from datetime import datetime, date
    from django.db.models import Sum, Q
    from decimal import Decimal
    
    # Get date range
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to = request.GET.get('date_to', date.today().isoformat())
    export = request.GET.get('export', '')
    
    # Convert to date objects
    if isinstance(date_from, str):
        df_obj = datetime.strptime(date_from, '%Y-%m-%d').date()
    else:
        df_obj = date_from
        
    if isinstance(date_to, str):
        dt_obj = datetime.strptime(date_to, '%Y-%m-%d').date()
    else:
        dt_obj = date_to
    
    # Get revenue accounts (category code 04)
    rev_accounts = Account.objects.filter(
        group__category__code='04', 
        account_type='detail', 
        is_active=True
    ).select_related('group__category').order_by('code')
    
    # Get expense accounts (category code 05)
    exp_accounts = Account.objects.filter(
        group__category__code='05', 
        account_type='detail', 
        is_active=True
    ).select_related('group__category').order_by('code')
    
    rev_rows = []
    exp_rows = []
    total_rev = Decimal('0')
    total_exp = Decimal('0')
    
    # Calculate revenues
    for acc in rev_accounts:
        # Get movements for the period
        moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=df_obj,
            journal_entry__date__lte=dt_obj
        ).aggregate(
            total_debit=Sum('debit_amount'),
            total_credit=Sum('credit_amount')
        )
        
        debit = moves['total_debit'] or Decimal('0')
        credit = moves['total_credit'] or Decimal('0')
        
        # Revenue accounts have credit balance (credit increases revenue)
        balance = credit - debit
        
        if balance != 0:
            rev_rows.append({
                'account': acc,
                'code': acc.code,
                'name': acc.name_ar,
                'amount': balance,
                'debit': debit,
                'credit': credit,
            })
            total_rev += balance
    
    # Calculate expenses
    for acc in exp_accounts:
        # Get movements for the period
        moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=df_obj,
            journal_entry__date__lte=dt_obj
        ).aggregate(
            total_debit=Sum('debit_amount'),
            total_credit=Sum('credit_amount')
        )
        
        debit = moves['total_debit'] or Decimal('0')
        credit = moves['total_credit'] or Decimal('0')
        
        # Expense accounts have debit balance (debit increases expense)
        balance = debit - credit
        
        if balance != 0:
            exp_rows.append({
                'account': acc,
                'code': acc.code,
                'name': acc.name_ar,
                'amount': balance,
                'debit': debit,
                'credit': credit,
            })
            total_exp += balance
    
    # Calculate net income
    net_income = total_rev - total_exp
    
    # Sort by account code
    rev_rows.sort(key=lambda x: x['code'])
    exp_rows.sort(key=lambda x: x['code'])
    
    if export == 'excel':
        return _export_income_statement_excel(
            rev_rows, exp_rows, total_rev, total_exp, net_income, date_from, date_to
        )
    
    return render(request, 'invoices/income_statement.html', {
        'rev_rows': rev_rows,
        'exp_rows': exp_rows,
        'total_revenue': total_rev,
        'total_expenses': total_exp,
        'net_income': net_income,
        'from_date': date_from,
        'to_date': date_to,
    })

def _export_income_statement_excel(rev_rows, exp_rows, total_rev, total_exp,
                                    net_income, date_from, date_to):
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    import io
    from django.http import HttpResponse
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "قائمة الدخل"
    ws.sheet_view.rightToLeft = True

    # Correct PatternFill syntax
    hf = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    rev_fill = PatternFill(start_color="1B5E20", end_color="1B5E20", fill_type="solid")
    exp_fill = PatternFill(start_color="B71C1C", end_color="B71C1C", fill_type="solid")
    net_fill_good = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    net_fill_bad = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
    
    bdr = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )

    # Title
    ws.merge_cells('A1:C1')
    c = ws['A1']
    c.value = f"قائمة الدخل | من {date_from} إلى {date_to}"
    c.font = Font(bold=True, color="C9A84C", size=13)
    c.fill = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ri = 2

    def write_section_header(title, fill):
        nonlocal ri
        ws.merge_cells(f'A{ri}:C{ri}')
        cell = ws[f'A{ri}']
        cell.value = title
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = fill
        cell.border = bdr
        cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 22
        ri += 1

    def write_row(code, name, amount, color="000000"):
        nonlocal ri
        for ci, v in enumerate([code, name, float(amount)], 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border = bdr
            cell.alignment = Alignment(horizontal='center' if ci != 2 else 'right')
            cell.font = Font(color=color, size=9)
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
        ws.row_dimensions[ri].height = 15
        ri += 1

    def write_total(label, amount, fill, color="FFFFFF"):
        nonlocal ri
        ws.merge_cells(f'A{ri}:B{ri}')
        cell = ws[f'A{ri}']
        cell.value = label
        cell.font = Font(bold=True, color=color, size=10)
        cell.fill = fill
        cell.border = bdr
        cell.alignment = Alignment(horizontal='center')
        vc = ws.cell(row=ri, column=3, value=float(amount))
        vc.font = Font(bold=True, color=color, size=10)
        vc.fill = fill
        vc.border = bdr
        vc.number_format = '#,##0.00'
        vc.alignment = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 20
        ri += 1

    # إيرادات
    write_section_header("الإيرادات", rev_fill)
    for row in rev_rows:
        write_row(row['account'].code, row['account'].name_ar, row['amount'], "1B5E20")
    write_total("إجمالي الإيرادات", total_rev, rev_fill)

    ri += 1  # فاصل

    # مصروفات
    write_section_header("المصروفات", exp_fill)
    for row in exp_rows:
        write_row(row['account'].code, row['account'].name_ar, row['amount'], "B71C1C")
    write_total("إجمالي المصروفات", total_exp, exp_fill)

    ri += 1  # فاصل

    # صافي الدخل
    net_fill = net_fill_good if net_income >= 0 else net_fill_bad
    label = "صافي الربح" if net_income >= 0 else "صافي الخسارة"
    write_total(label, net_income, net_fill)

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 36
    ws.column_dimensions['C'].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="income_statement_{date_from}_{date_to}.xlsx"'
    return resp
# ─────────────────────────────────────────────────────────
# الميزانية العمومية — تصدير Excel
# ─────────────────────────────────────────────────────────

@login_required
def balance_sheet_v2(request):
    as_of  = request.GET.get('as_of', date.today().isoformat())
    export = request.GET.get('export', '')

    def _get_accounts(cat_code):
        result = []
        total  = Decimal('0')
        for acc in Account.objects.filter(
            group__category__code=cat_code,
            account_type='detail', is_active=True
        ).order_by('code'):
            bal = acc.get_balance(date_to=as_of)
            if bal != 0:
                result.append({'account': acc, 'balance': bal})
                total += bal
        return result, total

    all_assets,      total_assets      = _get_accounts('01')
    all_liabilities, total_liabilities = _get_accounts('02')
    equity_rows,     total_equity      = _get_accounts('03')

    # تصنيف الأصول
    current_assets = [r for r in all_assets if '0101' in r['account'].code[:4]]
    fixed_assets   = [r for r in all_assets if r not in current_assets]
    if not current_assets:
        current_assets = all_assets
        fixed_assets   = []

    # تصنيف الالتزامات
    current_liab  = [r for r in all_liabilities if '0201' in r['account'].code[:4]]
    long_term_liab= [r for r in all_liabilities if r not in current_liab]
    if not current_liab:
        current_liab   = all_liabilities
        long_term_liab = []

    cur_assets_total   = sum(r['balance'] for r in current_assets)
    fixed_assets_total = sum(r['balance'] for r in fixed_assets)
    cur_liab_total     = sum(r['balance'] for r in current_liab)
    lt_liab_total      = sum(r['balance'] for r in long_term_liab)
    total_liab_equity  = total_liabilities + total_equity
    is_balanced        = abs(total_assets - total_liab_equity) < Decimal('0.01')
    difference         = total_assets - total_liab_equity

    ctx = {
        'current_assets':     current_assets,
        'fixed_assets':       fixed_assets,
        'intangible_assets':  [],
        'cur_assets_total':   cur_assets_total,
        'fixed_assets_total': fixed_assets_total,
        'total_assets':       total_assets,
        'current_liab':       current_liab,
        'long_term_liab':     long_term_liab,
        'cur_liab_total':     cur_liab_total,
        'lt_liab_total':      lt_liab_total,
        'total_liabilities':  total_liabilities,
        'equity_rows':        equity_rows,
        'total_equity':       total_equity,
        'total_liab_equity':  total_liab_equity,
        'is_balanced':        is_balanced,
        'difference':         difference,
        'as_of':              as_of,
    }

    if export == 'excel':
        return _export_balance_sheet_excel(ctx)

    return render(request, 'invoices/balance_sheet.html', ctx)


def _export_balance_sheet_excel(ctx):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "الميزانية العمومية"
    ws.sheet_view.rightToLeft = True

    hf       = PatternFill("solid", fgColor="1a1a2e")
    asset_f  = PatternFill("solid", fgColor="0D47A1")
    liab_f   = PatternFill("solid", fgColor="B71C1C")
    eq_f     = PatternFill("solid", fgColor="1B5E20")
    sub_f    = PatternFill("solid", fgColor="E3F2FD")
    bdr = Border(
        left=Side(style='thin'),  right=Side(style='thin'),
        top=Side(style='thin'),   bottom=Side(style='thin'),
    )

    as_of = ctx['as_of']
    ws.merge_cells('A1:D1')
    c = ws['A1']
    c.value = f"الميزانية العمومية بتاريخ {as_of}"
    c.font  = Font(bold=True, color="C9A84C", size=13)
    c.fill  = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ri = 2

    def section(title, fill, rows, subtotal_label, subtotal):
        nonlocal ri
        ws.merge_cells(f'A{ri}:D{ri}')
        cell = ws[f'A{ri}']
        cell.value     = title
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = fill
        cell.border    = bdr
        cell.alignment = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 20
        ri += 1

        for row in rows:
            vals = [row['account'].code, row['account'].name_ar, '', float(row['balance'])]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=v)
                cell.border    = bdr
                cell.alignment = Alignment(horizontal='center' if ci != 2 else 'right')
                cell.font      = Font(size=9)
                if isinstance(v, float):
                    cell.number_format = '#,##0.00'
            ws.row_dimensions[ri].height = 15
            ri += 1

        ws.merge_cells(f'A{ri}:C{ri}')
        lc = ws[f'A{ri}']
        lc.value     = subtotal_label
        lc.font      = Font(bold=True, size=10)
        lc.fill      = sub_f
        lc.border    = bdr
        lc.alignment = Alignment(horizontal='center')
        vc = ws.cell(row=ri, column=4, value=float(subtotal))
        vc.font          = Font(bold=True, size=10)
        vc.fill          = sub_f
        vc.border        = bdr
        vc.number_format = '#,##0.00'
        vc.alignment     = Alignment(horizontal='center')
        ws.row_dimensions[ri].height = 18
        ri += 1

    # ASSETS
    section("الأصول المتداولة", asset_f, ctx['current_assets'], "إجمالي الأصول المتداولة", ctx['cur_assets_total'])
    if ctx['fixed_assets']:
        section("الأصول الثابتة", asset_f, ctx['fixed_assets'], "إجمالي الأصول الثابتة", ctx['fixed_assets_total'])

    ws.merge_cells(f'A{ri}:C{ri}')
    tc = ws[f'A{ri}']
    tc.value = "إجمالي الأصول"
    tc.font  = Font(bold=True, color="C9A84C", size=11)
    tc.fill  = hf
    tc.border= bdr
    tc.alignment = Alignment(horizontal='center')
    tv = ws.cell(row=ri, column=4, value=float(ctx['total_assets']))
    tv.font          = Font(bold=True, color="C9A84C", size=11)
    tv.fill          = hf
    tv.border        = bdr
    tv.number_format = '#,##0.00'
    tv.alignment     = Alignment(horizontal='center')
    ws.row_dimensions[ri].height = 22
    ri += 2

    # LIABILITIES
    section("الالتزامات المتداولة", liab_f, ctx['current_liab'], "إجمالي الالتزامات المتداولة", ctx['cur_liab_total'])
    if ctx['long_term_liab']:
        section("الالتزامات طويلة الأجل", liab_f, ctx['long_term_liab'], "إجمالي الالتزامات طويلة الأجل", ctx['lt_liab_total'])

    # EQUITY
    section("حقوق الملكية", eq_f, ctx['equity_rows'], "إجمالي حقوق الملكية", ctx['total_equity'])

    # TOTAL L+E
    ws.merge_cells(f'A{ri}:C{ri}')
    tc2 = ws[f'A{ri}']
    tc2.value = "إجمالي الالتزامات وحقوق الملكية"
    tc2.font  = Font(bold=True, color="C9A84C", size=11)
    tc2.fill  = hf
    tc2.border= bdr
    tc2.alignment = Alignment(horizontal='center')
    tv2 = ws.cell(row=ri, column=4, value=float(ctx['total_liab_equity']))
    tv2.font          = Font(bold=True, color="C9A84C", size=11)
    tv2.fill          = hf
    tv2.border        = bdr
    tv2.number_format = '#,##0.00'
    tv2.alignment     = Alignment(horizontal='center')
    ws.row_dimensions[ri].height = 22
    ri += 2

    status_val  = "✓ الميزانية متوازنة" if ctx['is_balanced'] else f"✗ غير متوازنة — الفرق: {float(ctx['difference']):.2f}"
    status_fill = PatternFill("solid", fgColor="1B5E20" if ctx['is_balanced'] else "B71C1C")
    ws.merge_cells(f'A{ri}:D{ri}')
    sc = ws[f'A{ri}']
    sc.value     = status_val
    sc.font      = Font(bold=True, color="FFFFFF", size=10)
    sc.fill      = status_fill
    sc.alignment = Alignment(horizontal='center')
    sc.border    = bdr

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="balance_sheet_{ctx["as_of"]}.xlsx"'
    return resp


# ─────────────────────────────────────────────────────────
# كشف حساب العميل
# ─────────────────────────────────────────────────────────

@login_required
def client_statement(request, pk):
    client    = get_object_or_404(Client, pk=pk)
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to   = request.GET.get('date_to',   date.today().isoformat())
    export    = request.GET.get('export', '')

    invoices = SalesInvoice.objects.filter(
        client=client, date__gte=date_from, date__lte=date_to,
        transaction_type__in=['invoice', 'simplified']
    ).order_by('date')
    payments = Payment.objects.filter(
        client=client, date__gte=date_from, date__lte=date_to,
        payment_type='receipt'
    ).order_by('date')

    # بناء كشف الحساب
    rows = []
    for inv in invoices:
        rows.append({
            'date':        inv.date,
            'type':        'invoice',
            'reference':   inv.invoice_number,
            'description': f'فاتورة مبيعات — {inv.get_transaction_type_display()}',
            'debit':       inv.total,
            'credit':      None,
            'url':         f'/sales/{inv.pk}/',
        })

    for pay in payments:
        rows.append({
            'date':        pay.date,
            'type':        'payment',
            'reference':   pay.number,
            'description': pay.description or 'قبض دفعة',
            'debit':       None,
            'credit':      pay.amount,
            'url':         '#',
        })

    rows.sort(key=lambda x: x['date'])

    # رصيد أول المدة (الفواتير غير المدفوعة قبل الفترة)
    opening_inv = SalesInvoice.objects.filter(
        client=client, date__lt=date_from,
        transaction_type__in=['invoice', 'simplified']
    ).aggregate(t=Sum('total'))['t'] or Decimal('0')

    opening_pay = Payment.objects.filter(
        client=client, date__lt=date_from, payment_type='receipt'
    ).aggregate(t=Sum('amount'))['t'] or Decimal('0')

    opening_balance = opening_inv - opening_pay

    running = opening_balance
    for row in rows:
        if row['debit']:
            running += row['debit']
        if row['credit']:
            running -= row['credit']
        row['running_balance'] = running

    total_invoices = sum(r['debit'] or Decimal('0') for r in rows)
    total_payments = sum(r['credit'] or Decimal('0') for r in rows)
    balance        = running

    if export == 'excel':
        return _export_client_statement_excel(
            client, rows, opening_balance, total_invoices, total_payments, balance, date_from, date_to
        )

    return render(request, 'invoices/client_statement.html', {
        'client':          client,
        'statement_rows':  rows,
        'opening_balance': opening_balance,
        'total_invoices':  total_invoices,
        'total_payments':  total_payments,
        'balance':         balance,
        'date_from':       date_from,
        'date_to':         date_to,
    })


def _export_client_statement_excel(client, rows, opening_balance,
                                    total_invoices, total_payments, balance,
                                    date_from, date_to):
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = f"كشف {client.name[:20]}"
    ws.sheet_view.rightToLeft = True

    hf  = PatternFill("solid", fgColor="1a1a2e")
    bdr = Border(
        left=Side(style='thin'),  right=Side(style='thin'),
        top=Side(style='thin'),   bottom=Side(style='thin'),
    )

    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = f"كشف حساب العميل: {client.name}  |  من {date_from}  إلى {date_to}"
    c.font  = Font(bold=True, color="C9A84C", size=12)
    c.fill  = hf
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    headers = ['التاريخ','النوع','المرجع','البيان','مدين (ر.س)','دائن (ر.س)','الرصيد (ر.س)']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=ci, value=h)
        cell.font      = Font(bold=True, color="C9A84C", size=9)
        cell.fill      = hf
        cell.alignment = Alignment(horizontal='center')
        cell.border    = bdr
    ws.row_dimensions[2].height = 18

    # رصيد أول المدة
    ws.merge_cells('A3:D3')
    ob = ws['A3']
    ob.value     = "رصيد أول المدة"
    ob.font      = Font(bold=True, size=9)
    ob.border    = bdr
    ob.alignment = Alignment(horizontal='center')
    for ci in [5, 6]:
        ws.cell(row=3, column=ci).border = bdr
    obv = ws.cell(row=3, column=7, value=float(opening_balance))
    obv.font          = Font(bold=True, size=9)
    obv.border        = bdr
    obv.number_format = '#,##0.00'
    obv.alignment     = Alignment(horizontal='center')
    ws.row_dimensions[3].height = 15

    for ri, row in enumerate(rows, start=4):
        vals = [
            str(row['date']),
            'فاتورة' if row['type'] == 'invoice' else 'دفعة',
            row['reference'],
            row['description'],
            float(row['debit'])   if row['debit']  else None,
            float(row['credit'])  if row['credit'] else None,
            float(row['running_balance']),
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=v)
            cell.border    = bdr
            cell.alignment = Alignment(horizontal='center')
            cell.font      = Font(size=9)
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
        ws.row_dimensions[ri].height = 15

    last = len(rows) + 4
    ws.merge_cells(f'A{last}:D{last}')
    lc = ws[f'A{last}']
    lc.value     = "الإجمالي"
    lc.font      = Font(bold=True, color="C9A84C", size=10)
    lc.fill      = hf
    lc.border    = bdr
    lc.alignment = Alignment(horizontal='center')

    for ci, v in enumerate([float(total_invoices), float(total_payments), float(balance)], start=5):
        cell = ws.cell(row=last, column=ci, value=v)
        cell.font          = Font(bold=True, color="C9A84C", size=10)
        cell.fill          = hf
        cell.border        = bdr
        cell.number_format = '#,##0.00'
        cell.alignment     = Alignment(horizontal='center')
    ws.row_dimensions[last].height = 20

    col_widths = [12, 10, 14, 32, 14, 14, 14]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    resp = HttpResponse(
        buf.read(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    resp['Content-Disposition'] = f'attachment; filename="client_statement_{client.pk}_{date_from}_{date_to}.xlsx"'
    return resp



# Add this temporary view to check account categories
@login_required
def debug_accounts(request):
    from django.http import JsonResponse
    
    # Check all accounts
    accounts = Account.objects.filter(account_type='detail', is_active=True).select_related('group__category')
    
    data = []
    for acc in accounts:
        data.append({
            'code': acc.code,
            'name': acc.name_ar,
            'category_code': acc.group.category.code if acc.group and acc.group.category else 'None',
            'category_name': acc.group.category.name_ar if acc.group and acc.group.category else 'None',
            'group_code': acc.group.code if acc.group else 'None',
            'normal_balance': acc.normal_balance,
        })
    
    # Check journal entry lines
    lines = JournalEntryLine.objects.filter(
        journal_entry__status='posted'
    ).select_related('account', 'journal_entry')[:20]
    
    line_data = []
    for line in lines:
        line_data.append({
            'date': str(line.journal_entry.date),
            'account': f"{line.account.code} - {line.account.name_ar}",
            'debit': float(line.debit_amount),
            'credit': float(line.credit_amount),
            'entry_number': line.journal_entry.number,
        })
    
    return JsonResponse({
        'accounts': data,
        'recent_lines': line_data,
        'total_accounts': accounts.count(),
        'total_lines': JournalEntryLine.objects.filter(journal_entry__status='posted').count(),
    })
    

@login_required
def debug_income_statement(request):
    """تشخيص قائمة الدخل"""
    from django.http import JsonResponse
    from datetime import date, datetime
    from decimal import Decimal
    
    date_from = request.GET.get('date_from', date.today().replace(month=1, day=1).isoformat())
    date_to = request.GET.get('date_to', date.today().isoformat())
    
    df_obj = datetime.strptime(date_from, '%Y-%m-%d').date()
    dt_obj = datetime.strptime(date_to, '%Y-%m-%d').date()
    
    # Get revenue accounts
    rev_accounts = Account.objects.filter(
        group__category__code='04', 
        account_type='detail', 
        is_active=True
    )
    
    # Get expense accounts
    exp_accounts = Account.objects.filter(
        group__category__code='05', 
        account_type='detail', 
        is_active=True
    )
    
    result = {
        'date_range': f"{date_from} to {date_to}",
        'revenue_accounts': [],
        'expense_accounts': [],
        'all_posted_entries': []
    }
    
    # Check revenue accounts
    for acc in rev_accounts:
        # Get movements directly from JournalEntryLine
        moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=df_obj,
            journal_entry__date__lte=dt_obj
        ).aggregate(
            d=Sum('debit_amount'),
            c=Sum('credit_amount')
        )
        
        debit = moves['d'] or Decimal('0')
        credit = moves['c'] or Decimal('0')
        
        # Revenue: credit increases, debit decreases
        balance = credit - debit
        
        # Also check get_balance() method
        get_bal = acc.get_balance(date_from, date_to)
        
        result['revenue_accounts'].append({
            'code': acc.code,
            'name': acc.name_ar,
            'normal_balance': acc.normal_balance,
            'debit': float(debit),
            'credit': float(credit),
            'calculated_balance': float(balance),
            'get_balance_method': float(get_bal),
        })
    
    # Check expense accounts
    for acc in exp_accounts:
        moves = JournalEntryLine.objects.filter(
            account=acc,
            journal_entry__status='posted',
            journal_entry__date__gte=df_obj,
            journal_entry__date__lte=dt_obj
        ).aggregate(
            d=Sum('debit_amount'),
            c=Sum('credit_amount')
        )
        
        debit = moves['d'] or Decimal('0')
        credit = moves['c'] or Decimal('0')
        
        # Expense: debit increases, credit decreases
        balance = debit - credit
        
        get_bal = acc.get_balance(date_from, date_to)
        
        result['expense_accounts'].append({
            'code': acc.code,
            'name': acc.name_ar,
            'normal_balance': acc.normal_balance,
            'debit': float(debit),
            'credit': float(credit),
            'calculated_balance': float(balance),
            'get_balance_method': float(get_bal),
        })
    
    # Check all posted entries in the period (last 20)
    entries = JournalEntry.objects.filter(
        status='posted',
        date__gte=df_obj,
        date__lte=dt_obj
    ).order_by('-date')[:20]
    
    for entry in entries:
        lines = []
        for line in entry.lines.all():
            lines.append({
                'account': f"{line.account.code} - {line.account.name_ar}",
                'category': line.account.group.category.code if line.account.group and line.account.group.category else 'None',
                'debit': float(line.debit_amount),
                'credit': float(line.credit_amount),
            })
        result['all_posted_entries'].append({
            'number': entry.number,
            'date': str(entry.date),
            'description': entry.description,
            'status': entry.status,
            'lines': lines,
        })
    
    return JsonResponse(result, json_dumps_params={'ensure_ascii': False, 'indent': 2})
