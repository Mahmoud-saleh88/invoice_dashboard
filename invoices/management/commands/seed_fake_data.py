# management/commands/seed_fake_data.py
"""
Management command to load fake data into the database.
Usage: python manage.py seed_fake_data --keep-company --keep-users
"""

from django.core.management.base import BaseCommand
from django.core.management import call_command
from django.db import transaction
from django.utils import timezone
from decimal import Decimal
from datetime import datetime, timedelta
import logging

from invoices.models import (
    AccountCategory, AccountGroup, Account, CostCenter,
    Client, Supplier, FiscalYear, AccountingPeriod,
    SalesInvoice, SalesInvoiceItem, PurchaseInvoice,
    AssetCategory, FixedAsset, JournalEntry, JournalEntryLine,
    Payment, PaymentInvoiceAllocation, Company, CustomUser
)

logger = logging.getLogger(__name__)


class Command(BaseCommand):
    help = 'Load fake data into the database (clears existing data first)'

    def add_arguments(self, parser):
        parser.add_argument(
            '--keep-company',
            action='store_true',
            help='Keep existing Company data (default: False)',
        )
        parser.add_argument(
            '--keep-users',
            action='store_true',
            help='Keep existing CustomUser data (default: False)',
        )

    def handle(self, *args, **options):
        self.stdout.write(self.style.WARNING('Starting data load...'))
        
        # Store existing company and user if needed
        self.existing_company = None
        self.existing_user = None
        
        if options.get('keep_company'):
            self.existing_company = Company.objects.first()
            if self.existing_company:
                self.stdout.write(f'  - Keeping existing company: {self.existing_company.name}')
        
        if options.get('keep_users'):
            self.existing_user = CustomUser.objects.first()
            if self.existing_user:
                self.stdout.write(f'  - Keeping existing user: {self.existing_user.username}')
        
        # Clear existing data (excluding Company and CustomUser if specified)
        self.clear_existing_data(
            keep_company=options.get('keep_company', False),
            keep_users=options.get('keep_users', False)
        )
        
        # Load new data
        self.load_account_categories()
        self.load_account_groups()
        self.load_accounts()
        self.load_cost_centers()
        self.load_clients()
        self.load_suppliers()
        self.load_fiscal_years()
        self.load_asset_categories()
        self.load_fixed_assets()
        self.load_sales_invoices()
        self.load_purchase_invoices()
        self.load_journal_entries()
        self.load_payments()
        
        self.stdout.write(self.style.SUCCESS('Successfully loaded all fake data!'))

    def clear_existing_data(self, keep_company=False, keep_users=False):
        """Clear existing data in reverse order of dependencies"""
        self.stdout.write('Clearing existing data...')
        
        # Order matters due to foreign key constraints
        models_to_clear = [
            PaymentInvoiceAllocation,
            Payment,
            JournalEntryLine,
            JournalEntry,
            SalesInvoiceItem,
            SalesInvoice,
            PurchaseInvoice,
            FixedAsset,
            AssetCategory,
            AccountingPeriod,
            FiscalYear,
            Client,
            Supplier,
            CostCenter,
            Account,
            AccountGroup,
            AccountCategory,
        ]
        
        deleted_counts = {}
        for model in models_to_clear:
            try:
                count = model.objects.count()
                if count > 0:
                    model.objects.all().delete()
                    deleted_counts[model.__name__] = count
                    self.stdout.write(f'  - Cleared {model.__name__}: {count} records')
            except Exception as e:
                self.stdout.write(self.style.WARNING(f'  - Could not clear {model.__name__}: {e}'))
        
        self.stdout.write(self.style.SUCCESS(f'Data cleared successfully! Cleared {len(deleted_counts)} model types'))

    def load_account_categories(self):
        self.stdout.write('Loading Account Categories...')
        categories = [
            ('01', 'الأصول', 'Assets', 'debit'),
            ('02', 'الالتزامات', 'Liabilities', 'credit'),
            ('03', 'حقوق الملكية', 'Equity', 'credit'),
            ('04', 'الإيرادات', 'Revenue', 'credit'),
            ('05', 'المصروفات', 'Expenses', 'debit'),
        ]
        
        created_categories = []
        for code, name_ar, name_en, balance in categories:
            category = AccountCategory.objects.create(
                code=code,
                name_ar=name_ar,
                name_en=name_en,
                normal_balance=balance
            )
            created_categories.append(category)
            self.stdout.write(f'    - Created: {code} - {name_ar}')
        
        self.stdout.write(f'  - Created {len(created_categories)} account categories')
        return created_categories

    def load_account_groups(self):
        self.stdout.write('Loading Account Groups...')
        
        # First, get the created categories
        try:
            assets_cat = AccountCategory.objects.get(code='01')
            liabilities_cat = AccountCategory.objects.get(code='02')
            equity_cat = AccountCategory.objects.get(code='03')
            revenue_cat = AccountCategory.objects.get(code='04')
            expenses_cat = AccountCategory.objects.get(code='05')
        except AccountCategory.DoesNotExist as e:
            self.stdout.write(self.style.ERROR(f'Category not found: {e}'))
            raise
        
        groups = [
            # Assets Groups
            ('0101', 'الأصول المتداولة', 'Current Assets', 'debit', assets_cat),
            ('0102', 'الأصول الثابتة', 'Fixed Assets', 'debit', assets_cat),
            ('0103', 'الأصول غير الملموسة', 'Intangible Assets', 'debit', assets_cat),
            # Liabilities Groups
            ('0201', 'الالتزامات المتداولة', 'Current Liabilities', 'credit', liabilities_cat),
            ('0202', 'الالتزامات طويلة الأجل', 'Long-term Liabilities', 'credit', liabilities_cat),
            # Equity Groups
            ('0301', 'رأس المال', 'Capital', 'credit', equity_cat),
            ('0302', 'الأرباح المحتجزة', 'Retained Earnings', 'credit', equity_cat),
            # Revenue Groups
            ('0401', 'إيرادات المبيعات', 'Sales Revenue', 'credit', revenue_cat),
            ('0402', 'إيرادات أخرى', 'Other Revenue', 'credit', revenue_cat),
            # Expenses Groups
            ('0501', 'مصروفات تشغيلية', 'Operating Expenses', 'debit', expenses_cat),
            ('0502', 'مصروفات إدارية', 'Administrative Expenses', 'debit', expenses_cat),
        ]
        
        created_groups = []
        for code, name_ar, name_en, balance, category in groups:
            group = AccountGroup.objects.create(
                code=code,
                name_ar=name_ar,
                name_en=name_en,
                normal_balance=balance,
                category=category
            )
            created_groups.append(group)
            self.stdout.write(f'    - Created: {code} - {name_ar}')
        
        self.stdout.write(f'  - Created {len(created_groups)} account groups')
        return created_groups

    def load_accounts(self):
        self.stdout.write('Loading Accounts...')
        
        # Get groups
        current_assets = AccountGroup.objects.get(code='0101')
        fixed_assets = AccountGroup.objects.get(code='0102')
        current_liabilities = AccountGroup.objects.get(code='0201')
        capital_group = AccountGroup.objects.get(code='0301')
        retained_earnings = AccountGroup.objects.get(code='0302')
        revenue_group = AccountGroup.objects.get(code='0401')
        other_revenue = AccountGroup.objects.get(code='0402')
        operating_expenses = AccountGroup.objects.get(code='0501')
        admin_expenses = AccountGroup.objects.get(code='0502')
        
        accounts = [
            # Current Assets
            ('010101', 'الصندوق', 'Cash on Hand', 'detail', 'debit', current_assets, 50000),
            ('010102', 'البنك - الراجحي', 'Bank - Al Rajhi', 'detail', 'debit', current_assets, 250000),
            ('010103', 'البنك - الأهلي', 'Bank - Al Ahli', 'detail', 'debit', current_assets, 150000),
            ('010104', 'مدينون - عملاء', 'Accounts Receivable', 'detail', 'debit', current_assets, 75000),
            ('010105', 'مخزون', 'Inventory', 'detail', 'debit', current_assets, 120000),
            # Fixed Assets
            ('010201', 'مباني', 'Buildings', 'detail', 'debit', fixed_assets, 1000000),
            ('010202', 'معدات', 'Equipment', 'detail', 'debit', fixed_assets, 300000),
            ('010203', 'مركبات', 'Vehicles', 'detail', 'debit', fixed_assets, 150000),
            ('010204', 'مجمع إهلاك المباني', 'Accumulated Depreciation - Buildings', 'detail', 'credit', fixed_assets, 50000),
            # Current Liabilities
            ('020101', 'دائنون - موردون', 'Accounts Payable', 'detail', 'credit', current_liabilities, 45000),
            ('020102', 'ضريبة القيمة المضافة', 'VAT Payable', 'detail', 'credit', current_liabilities, 15000),
            ('020103', 'رواتب مستحقة', 'Salaries Payable', 'detail', 'credit', current_liabilities, 0),
            # Capital and Equity
            ('030101', 'رأس المال', 'Capital', 'detail', 'credit', capital_group, 2000000),
            ('030201', 'أرباح محتجزة', 'Retained Earnings', 'detail', 'credit', retained_earnings, 250000),
            # Revenue
            ('040101', 'مبيعات', 'Sales', 'detail', 'credit', revenue_group, 0),
            ('040102', 'خدمات', 'Service Revenue', 'detail', 'credit', revenue_group, 0),
            ('040201', 'إيرادات إيجار', 'Rental Income', 'detail', 'credit', other_revenue, 0),
            # Expenses
            ('050101', 'مشتريات', 'Purchases', 'detail', 'debit', operating_expenses, 0),
            ('050102', 'إيجار', 'Rent Expense', 'detail', 'debit', operating_expenses, 0),
            ('050201', 'رواتب وأجور', 'Salaries Expense', 'detail', 'debit', admin_expenses, 0),
            ('050202', 'كهرباء وماء', 'Utilities Expense', 'detail', 'debit', admin_expenses, 0),
            ('050203', 'اتصالات', 'Telecommunications', 'detail', 'debit', admin_expenses, 0),
            ('050204', 'مصروفات تسويق', 'Marketing Expenses', 'detail', 'debit', admin_expenses, 0),
        ]
        
        created_accounts = []
        for code, name_ar, name_en, acc_type, balance, group, opening in accounts:
            account = Account.objects.create(
                code=code,
                name_ar=name_ar,
                name_en=name_en,
                account_type=acc_type,
                normal_balance=balance,
                group=group,
                opening_balance=Decimal(str(opening)),
                opening_balance_date='2024-01-01',
                is_active=True
            )
            created_accounts.append(account)
            self.stdout.write(f'    - Created: {code} - {name_ar}')
        
        self.stdout.write(f'  - Created {len(created_accounts)} accounts')
        return created_accounts

    def load_cost_centers(self):
        self.stdout.write('Loading Cost Centers...')
        centers = [
            ('CC-001', 'الإدارة العامة', 'General Administration'),
            ('CC-002', 'المبيعات', 'Sales'),
            ('CC-003', 'التسويق', 'Marketing'),
            ('CC-004', 'تكنولوجيا المعلومات', 'IT Department'),
            ('CC-005', 'الموارد البشرية', 'Human Resources'),
            ('CC-006', 'المشتريات', 'Procurement'),
            ('CC-007', 'الخدمات اللوجستية', 'Logistics'),
        ]
        
        created_centers = []
        for code, name_ar, name_en in centers:
            center = CostCenter.objects.create(
                code=code,
                name_ar=name_ar,
                name_en=name_en,
                is_active=True
            )
            created_centers.append(center)
            self.stdout.write(f'    - Created: {code} - {name_ar}')
        
        self.stdout.write(f'  - Created {len(created_centers)} cost centers')
        return created_centers

    def load_clients(self):
        self.stdout.write('Loading Clients...')
        clients = [
            ('شركة الاتصالات السعودية', 'Saudi Telecom Company', 'Prince Turki Street', '1234', 'Al Olaya', 'Riyadh', '12211', '310234567891234', 'accounts@stc.sa', '0112345678', 1000000),
            ('مجموعة الراجحي القابضة', 'Al Rajhi Holding Group', 'King Fahd Road', '5678', 'Al Malaz', 'Riyadh', '11321', '311234567891234', 'finance@alrajhi.com', '0118765432', 2500000),
            ('شركة الزامل للصناعة', 'Zamil Industrial', 'Dammam Street', '9012', 'Al Khobar', 'Dammam', '31952', '312345678912345', 'accounts@zamil.com', '0138765432', 800000),
            ('بنك البلاد', 'Bank Al Bilad', 'Tahlia Street', '3456', 'Al Andalus', 'Jeddah', '23324', '313456789123456', 'finance@bankalbilad.com', '0126654321', 3000000),
            ('شركة سابك', 'SABIC', 'Eastern Ring Road', '7890', 'Al Waha', 'Riyadh', '11426', '314567891234567', 'accounts@sabic.com', '0112349876', 5000000),
            ('مستشفى السعودي الألماني', 'Saudi German Hospital', 'Madinah Road', '2345', 'Al Faisaliyah', 'Jeddah', '23444', '315678912345678', 'finance@saudigerman.com', '0126688000', 1500000),
            ('شركة موبايلي', 'Mobily', 'King Abdulaziz Road', '6789', 'Al Shatea', 'Dammam', '32414', '316789123456789', 'accounting@mobily.com', '0124567890', 1200000),
        ]
        
        created_clients = []
        for name, name_en, street, building, district, city, postal, vat, email, phone, credit in clients:
            client = Client.objects.create(
                name=name,
                name_en=name_en,
                street_name=street,
                building_number=building,
                district=district,
                city=city,
                postal_code=postal,
                vat_number=vat,
                email=email,
                phone=phone,
                credit_limit=Decimal(str(credit)),
                is_active=True
            )
            created_clients.append(client)
            self.stdout.write(f'    - Created: {name}')
        
        self.stdout.write(f'  - Created {len(created_clients)} clients')
        return created_clients

    def load_suppliers(self):
        self.stdout.write('Loading Suppliers...')
        suppliers = [
            ('شركة التجهيزات المكتبية', 'Office Supplies Co.', 'الرياض - حي المروج', '401234567891234', 'sales@officesupplies.com', '0112345678'),
            ('مؤسسة مواد البناء', 'Building Materials Est.', 'جدة - حي البلد', '402345678912345', 'info@buildingmaterials.com', '0123456789'),
            ('شركة التقنية المتقدمة', 'Advanced Technology Co.', 'الخبر - حي العقربية', '403456789123456', 'sales@advancedtech.com', '0133456789'),
            ('مطابع الجزيرة', 'Al Jazeera Press', 'الرياض - حي الملز', '404567891234567', 'info@aljazeerapress.com', '0114567890'),
            ('شركة الأثاث الحديث', 'Modern Furniture Co.', 'جدة - حي الروضة', '405678912345678', 'sales@modernfurniture.com', '0125678901'),
        ]
        
        created_suppliers = []
        for name, name_en, address, vat, email, phone in suppliers:
            supplier = Supplier.objects.create(
                name=name,
                name_en=name_en,
                address=address,
                vat_number=vat,
                email=email,
                phone=phone,
                is_active=True
            )
            created_suppliers.append(supplier)
            self.stdout.write(f'    - Created: {name}')
        
        self.stdout.write(f'  - Created {len(created_suppliers)} suppliers')
        return created_suppliers

    def load_fiscal_years(self):
        self.stdout.write('Loading Fiscal Years and Accounting Periods...')
        
        # Get the first company (or use existing one)
        company = Company.objects.first()
        if not company:
            self.stdout.write(self.style.ERROR('No company found. Please create a company first.'))
            return
        
        fiscal_2024 = FiscalYear.objects.create(
            company=company,
            name='2024',
            start_date='2024-01-01',
            end_date='2024-12-31',
            is_closed=False
        )
        
        months = [
            ('يناير 2024', '2024-01-01', '2024-01-31'),
            ('فبراير 2024', '2024-02-01', '2024-02-29'),
            ('مارس 2024', '2024-03-01', '2024-03-31'),
            ('أبريل 2024', '2024-04-01', '2024-04-30'),
            ('مايو 2024', '2024-05-01', '2024-05-31'),
            ('يونيو 2024', '2024-06-01', '2024-06-30'),
            ('يوليو 2024', '2024-07-01', '2024-07-31'),
            ('أغسطس 2024', '2024-08-01', '2024-08-31'),
            ('سبتمبر 2024', '2024-09-01', '2024-09-30'),
            ('أكتوبر 2024', '2024-10-01', '2024-10-31'),
            ('نوفمبر 2024', '2024-11-01', '2024-11-30'),
            ('ديسمبر 2024', '2024-12-01', '2024-12-31'),
        ]
        
        created_periods = []
        for name, start, end in months:
            period = AccountingPeriod.objects.create(
                fiscal_year=fiscal_2024,
                name=name,
                start_date=start,
                end_date=end,
                is_closed=False
            )
            created_periods.append(period)
        
        self.stdout.write(f'  - Created 1 fiscal year and {len(created_periods)} accounting periods')
        return fiscal_2024, created_periods

    def load_asset_categories(self):
        self.stdout.write('Loading Asset Categories...')
        
        # Get accounts
        try:
            equipment_account = Account.objects.get(code='010202')
            vehicles_account = Account.objects.get(code='010203')
            buildings_account = Account.objects.get(code='010201')
            accum_dep_account = Account.objects.get(code='010204')
            expense_account = Account.objects.get(code='050201')
        except Account.DoesNotExist as e:
            self.stdout.write(self.style.ERROR(f'Account not found: {e}'))
            raise
        
        categories = [
            ('معدات حاسوبية', 3, 33.33, equipment_account, accum_dep_account, expense_account),
            ('مركبات', 5, 20, vehicles_account, accum_dep_account, expense_account),
            ('مباني', 25, 4, buildings_account, accum_dep_account, expense_account),
            ('معدات مكتبية', 5, 20, equipment_account, accum_dep_account, expense_account),
        ]
        
        created_categories = []
        for name, years, rate, asset_acc, accum_acc, expense_acc in categories:
            category = AssetCategory.objects.create(
                name=name,
                useful_life_years=years,
                depreciation_rate=Decimal(str(rate)),
                asset_account=asset_acc,
                accum_dep_account=accum_acc,
                dep_expense_account=expense_acc
            )
            created_categories.append(category)
            self.stdout.write(f'    - Created: {name}')
        
        self.stdout.write(f'  - Created {len(created_categories)} asset categories')
        return created_categories

    def load_fixed_assets(self):
        self.stdout.write('Loading Fixed Assets...')
        
        # Get categories
        computer_cat = AssetCategory.objects.get(name='معدات حاسوبية')
        vehicles_cat = AssetCategory.objects.get(name='مركبات')
        buildings_cat = AssetCategory.objects.get(name='مباني')
        office_cat = AssetCategory.objects.get(name='معدات مكتبية')
        
        assets = [
            ('FA-2024-001', 'خادم HP ProLiant', computer_cat, '2024-01-15', 25000, 2000, 3, 'غرفة الخوادم - الطابق الثاني'),
            ('FA-2024-002', 'تويوتا كامري 2024', vehicles_cat, '2024-02-01', 125000, 25000, 5, 'مواقف الشركة'),
            ('FA-2024-003', 'مبنى المكاتب الرئيسي', buildings_cat, '2020-01-01', 2000000, 200000, 25, 'الرياض - حي المروج'),
            ('FA-2024-004', 'أثاث مكتبي تنفيذي', office_cat, '2024-03-10', 35000, 3000, 5, 'مكاتب الإدارة - الطابق الثالث'),
            ('FA-2024-005', 'أجهزة حاسب آلي محمول - 10 وحدات', computer_cat, '2024-04-20', 42000, 4000, 3, 'مستودع الأجهزة'),
        ]
        
        created_assets = []
        for code, name, category, purchase_date, cost, salvage, years, location in assets:
            asset = FixedAsset.objects.create(
                code=code,
                name=name,
                category=category,
                purchase_date=purchase_date,
                purchase_cost=Decimal(str(cost)),
                salvage_value=Decimal(str(salvage)),
                useful_life_years=years,
                depreciation_method='straight_line',
                location=location,
                status='active'
            )
            created_assets.append(asset)
            self.stdout.write(f'    - Created: {code} - {name}')
        
        self.stdout.write(f'  - Created {len(created_assets)} fixed assets')
        return created_assets

    def generate_invoice_number(self):
        """توليد رقم فاتورة جديد بتنسيق متوافق مع ZATCA"""
        last_invoice = SalesInvoice.objects.order_by('-id').first()
        
        if last_invoice and last_invoice.invoice_number:
            try:
                invoice_str = str(last_invoice.invoice_number)
                
                if invoice_str.startswith('INV-'):
                    last_number = int(invoice_str.split('-')[1])
                    new_number = last_number + 1
                    return f"INV-{new_number:04d}"
                elif invoice_str.isdigit():
                    return str(int(invoice_str) + 1)
                else:
                    return f"INV-{last_invoice.id + 1:04d}"
            except (ValueError, TypeError, IndexError):
                return "INV-0001"
        
        return "INV-0001"

    def load_sales_invoices(self):
        self.stdout.write('Loading Sales Invoices...')
        
        # Get company and clients
        company = Company.objects.first()
        if not company:
            self.stdout.write(self.style.ERROR('No company found. Please create a company first.'))
            return
        
        clients = Client.objects.all()
        if not clients.exists():
            self.stdout.write(self.style.ERROR('No clients found.'))
            return
        
        invoices_data = [
            ('2024-06-15', '2024-06-15', '2024-07-15', 'invoice', 'credit', clients[0], 'Admin User', 15, 'فاتورة خدمات اتصالات - شهر يونيو',
             [('خدمات اتصالات وإنترنت - حزمة الأعمال', 1, 15000),
              ('خدمات صيانة تقنية', 2, 2500)]),
            
            ('2024-06-20', '2024-06-20', '2024-07-20', 'invoice', 'bank', clients[1], 'Admin User', 15, 'معدات صناعية',
             [('معدات تصنيع - خط إنتاج', 1, 75000)]),
            
            ('2024-07-01', '2024-07-01', '2024-08-01', 'invoice', 'credit', clients[2], 'Admin User', 15, 'خدمات استشارية',
             [('استشارات إدارية ومالية', 1, 25000)]),
            
            ('2024-07-05', '2024-07-05', '2024-08-05', 'invoice', 'bank', clients[3], 'Admin User', 15, 'خدمات مصرفية',
             [('خدمات تقنية مصرفية', 1, 45000)]),
            
            ('2024-07-10', '2024-07-10', '2024-08-10', 'invoice', 'credit', clients[4], 'Admin User', 15, 'منتجات بتروكيماوية',
             [('منتجات بلاستيكية', 100, 500),
              ('مواد كيميائية', 50, 1000)]),
            
            ('2024-07-15', '2024-07-15', '2024-08-15', 'invoice', 'credit', clients[5], 'Admin User', 15, 'خدمات طبية',
             [('خدمات استشارية طبية', 1, 35000)]),
        ]
        
        created_invoices = []
        for inv_data in invoices_data:
            date, supply_date, due_date, trans_type, pay_method, client, created_by, tax_rate, notes, items = inv_data
            
            invoice = SalesInvoice.objects.create(
                invoice_number=self.generate_invoice_number(),
                date=date,
                supply_date=supply_date,
                due_date=due_date,
                transaction_type=trans_type,
                payment_method=pay_method,
                company=company,
                client=client,
                created_by=created_by,
                tax_rate=Decimal(str(tax_rate)),
                notes=notes,
                amount_paid=Decimal('0')
            )
            
            total_quantity = 0
            total_sales_excl_tax = Decimal('0')
            
            for desc, qty, price in items:
                item = SalesInvoiceItem.objects.create(
                    invoice=invoice,
                    description=desc,
                    quantity=Decimal(str(qty)),
                    unit_price=Decimal(str(price)),
                    discount_percent=Decimal('0'),
                    discount_amount=Decimal('0'),
                    tax_rate=Decimal(str(tax_rate))
                )
                total_quantity += int(qty)
                total_sales_excl_tax += item.total
            
            # Update invoice totals
            invoice.total_quantity = total_quantity
            invoice.total_sales_excl_tax = total_sales_excl_tax
            invoice.discount_amount = Decimal('0')
            invoice.taxable_amount = total_sales_excl_tax
            invoice.total_tax = total_sales_excl_tax * invoice.tax_rate / Decimal('100')
            invoice.total = total_sales_excl_tax + invoice.total_tax
            invoice.save()
            
            created_invoices.append(invoice)
            self.stdout.write(f'    - Created: {invoice.invoice_number} - {client.name}')
        
        self.stdout.write(f'  - Created {len(created_invoices)} sales invoices')
        return created_invoices

    def load_purchase_invoices(self):
        self.stdout.write('Loading Purchase Invoices...')
        
        suppliers = Supplier.objects.all()
        if not suppliers.exists():
            self.stdout.write(self.style.ERROR('No suppliers found.'))
            return
        
        purchases = [
            ('PO-2024-001', '2024-06-10', '2024-06-10', '2024-07-10', 'invoice', suppliers[0], '401234567891234', 15000, 500, 2175, 16675, 16675, 'أجهزة كمبيوتر ومعدات مكتبية'),
            ('PO-2024-002', '2024-06-15', '2024-06-15', '2024-07-15', 'invoice', suppliers[1], '402345678912345', 45000, 2000, 6450, 49450, 20000, 'مواد بناء'),
            ('PO-2024-003', '2024-06-20', '2024-06-20', '2024-07-20', 'invoice', suppliers[3], '404567891234567', 8500, 0, 1275, 9775, 9775, 'مواد دعائية وإعلانية'),
            ('PO-2024-004', '2024-07-01', '2024-07-01', '2024-08-01', 'invoice', suppliers[2], '403456789123456', 35000, 1500, 5025, 38525, 38525, 'برامج وتقنيات معلوماتية'),
            ('PO-2024-005', '2024-07-05', '2024-07-05', '2024-08-05', 'invoice', suppliers[4], '405678912345678', 28000, 1000, 4050, 31050, 15000, 'أثاث مكاتب جديد'),
        ]
        
        created_purchases = []
        for purchase in purchases:
            inv = PurchaseInvoice.objects.create(
                invoice_number=purchase[0],
                date=purchase[1],
                entry_date=purchase[2],
                due_date=purchase[3],
                transaction_type=purchase[4],
                supplier=purchase[5],
                tax_number=purchase[6],
                subtotal=Decimal(str(purchase[7])),
                discount_amount=Decimal(str(purchase[8])),
                tax_amount=Decimal(str(purchase[9])),
                total=Decimal(str(purchase[10])),
                amount_paid=Decimal(str(purchase[11])),
                notes=purchase[12],
                created_by='Admin User'
            )
            created_purchases.append(inv)
            self.stdout.write(f'    - Created: {purchase[0]} - {purchase[5].name}')
        
        self.stdout.write(f'  - Created {len(created_purchases)} purchase invoices')
        return created_purchases

    def load_journal_entries(self):
        self.stdout.write('Loading Journal Entries...')
        
        # Get the user (admin)
        user = CustomUser.objects.first()
        if not user:
            self.stdout.write(self.style.ERROR('No user found. Please create a user first.'))
            return
        
        # Opening Balance Journal Entry
        opening_journal = JournalEntry.objects.create(
            number='JV-00001',
            date='2024-01-01',
            entry_type='opening',
            status='posted',
            description='قيد افتتاحي للسنة المالية 2024',
            created_by=user,
            posted_at=timezone.now()
        )
        
        # Get accounts
        try:
            cash_account = Account.objects.get(code='010101')
            bank_account = Account.objects.get(code='010102')
            receivables = Account.objects.get(code='010104')
            payables = Account.objects.get(code='020101')
            vat_payable = Account.objects.get(code='020102')
            capital = Account.objects.get(code='030101')
        except Account.DoesNotExist as e:
            self.stdout.write(self.style.ERROR(f'Account not found for opening entry: {e}'))
            return
        
        opening_lines = [
            (cash_account, 50000, 0, 'رصيد افتتاحي - صندوق'),
            (bank_account, 250000, 0, 'رصيد افتتاحي - بنك الراجحي'),
            (receivables, 75000, 0, 'رصيد افتتاحي - مدينون'),
            (payables, 0, 45000, 'رصيد افتتاحي - دائنون'),
            (vat_payable, 0, 15000, 'رصيد افتتاحي - ضريبة القيمة المضافة'),
            (capital, 0, 2000000, 'رصيد افتتاحي - رأس المال'),
        ]
        
        for acc, debit, credit, desc in opening_lines:
            JournalEntryLine.objects.create(
                journal_entry=opening_journal,
                account=acc,
                debit_amount=Decimal(str(debit)),
                credit_amount=Decimal(str(credit)),
                description=desc
            )
        
        self.stdout.write(f'    - Created opening journal entry: {opening_journal.number}')
        
        # Sales Journal Entry (for first invoice)
        sales_journal = JournalEntry.objects.create(
            number='JV-00002',
            date='2024-06-15',
            entry_type='sales_invoice',
            status='posted',
            description='قيد فاتورة مبيعات',
            created_by=user,
            posted_at=timezone.now()
        )
        
        invoice = SalesInvoice.objects.first()
        if invoice:
            sales_journal.sales_invoice = invoice
            sales_journal.save()
            
            # Get accounts for sales entry
            sales_account = Account.objects.get(code='040101')
            
            # Calculate amounts
            total_amount = invoice.total
            tax_amount = invoice.total_tax
            
            JournalEntryLine.objects.create(
                journal_entry=sales_journal,
                account=receivables,
                debit_amount=total_amount,
                credit_amount=0,
                description=f'فاتورة مبيعات رقم {invoice.invoice_number}'
            )
            
            JournalEntryLine.objects.create(
                journal_entry=sales_journal,
                account=sales_account,
                debit_amount=0,
                credit_amount=invoice.taxable_amount,
                description=f'قيمة المبيعات - {invoice.invoice_number}'
            )
            
            JournalEntryLine.objects.create(
                journal_entry=sales_journal,
                account=vat_payable,
                debit_amount=0,
                credit_amount=tax_amount,
                description=f'ضريبة القيمة المضافة - {invoice.invoice_number}'
            )
            
            self.stdout.write(f'    - Created sales journal entry: {sales_journal.number}')
        
        self.stdout.write(f'  - Created {JournalEntry.objects.count()} journal entries')

    def load_payments(self):
        self.stdout.write('Loading Payments...')
        
        # Get user and clients/suppliers
        user = CustomUser.objects.first()
        if not user:
            self.stdout.write(self.style.ERROR('No user found.'))
            return
        
        clients = Client.objects.all()
        suppliers = Supplier.objects.all()
        
        if not clients.exists() or not suppliers.exists():
            self.stdout.write(self.style.WARNING('Skipping payments - no clients or suppliers found'))
            return
        
        # Receipt from client
        receipt = Payment.objects.create(
            number='RCT-2024-001',
            date='2024-06-25',
            payment_type='receipt',
            payment_method='bank',
            amount=Decimal('25000'),
            client=clients[0],
            description='دفعة مقدمة من شركة الاتصالات',
            reference='TRF-123456',
            created_by=user
        )
        
        # Payment to supplier
        payment = Payment.objects.create(
            number='PAY-2024-001',
            date='2024-06-18',
            payment_type='payment',
            payment_method='bank',
            amount=Decimal('16675'),
            supplier=suppliers[0],
            description='دفع فاتورة المشتريات PO-2024-001',
            reference='TRF-234567',
            created_by=user
        )
        
        # Allocations
        invoice = SalesInvoice.objects.first()
        if invoice:
            PaymentInvoiceAllocation.objects.create(
                payment=receipt,
                sales_invoice=invoice,
                allocated_amount=Decimal('25000')
            )
            self.stdout.write(f'    - Allocated {receipt.number} to {invoice.invoice_number}')
        
        purchase = PurchaseInvoice.objects.filter(invoice_number='PO-2024-001').first()
        if purchase:
            PaymentInvoiceAllocation.objects.create(
                payment=payment,
                purchase_invoice=purchase,
                allocated_amount=Decimal('16675')
            )
            self.stdout.write(f'    - Allocated {payment.number} to {purchase.invoice_number}')
        
        self.stdout.write(f'  - Created {Payment.objects.count()} payments and {PaymentInvoiceAllocation.objects.count()} allocations')