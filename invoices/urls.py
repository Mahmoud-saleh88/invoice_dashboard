"""urls.py — شركة عقارات سعودية"""
from django.urls import path
from . import views

urlpatterns = [
    # ── Auth ──────────────────────────────────────────────────────────────
    path('login/',  views.login_view,  name='login'),
    path('logout/', views.logout_view, name='logout'),

    # ── Dashboard ─────────────────────────────────────────────────────────
    path('dashboard/', views.dashboard, name='dashboard'),

    # ── Sales Invoices ─────────────────────────────────────────────────────
    path('sales/', views.sales_list, name='sales_list'),
    path('sales/add/', views.sales_add, name='sales_add'),
    path('sales/<int:pk>/edit/', views.sales_edit, name='sales_edit'),
    path('sales/<int:pk>/delete/', views.sales_delete, name='sales_delete'),
    path('sales/item/<int:pk>/delete/', views.sales_invoice_item_delete, name='sales_invoice_item_delete'),
    path('sales/<int:pk>/', views.sales_detail, name='sales_detail'),
    path('sales/<int:pk>/pdf/', views.sales_pdf, name='sales_pdf'),
    path('sales/export/', views.export_sales_excel, name='sales_export'),
    # ZATCA per invoice
    path('sales/<int:pk>/zatca/submit/',    views.zatca_submit_invoice,       name='zatca_submit_invoice'),

    # ── ZATCA ──────────────────────────────────────────────────────────────
    path('zatca/onboarding/',               views.zatca_onboarding,           name='zatca_onboarding'),
    path('zatca/bulk-submit/',              views.zatca_bulk_submit,          name='zatca_bulk_submit'),
    path('zatca/logs/',                     views.zatca_logs,                 name='zatca_logs'),

    # ── Purchase Invoices ──────────────────────────────────────────────────
    path('purchases/',                      views.purchases_list,             name='purchases_list'),
    path('purchases/add/',                  views.purchases_add,              name='purchases_add'),
    path('purchases/<int:pk>/edit/',        views.purchases_edit,             name='purchases_edit'),
    path('purchases/<int:pk>/delete/',      views.purchases_delete,           name='purchases_delete'),
    path('purchases/export/', views.export_purchases_excel, name='purchases_export'),

    # ── Clients ────────────────────────────────────────────────────────────
    path('clients/',                        views.client_list,                name='client_list'),
    path('clients/add/',                    views.client_create,              name='client_create'),
    path('clients/<int:pk>/edit/',          views.client_edit,                name='client_edit'),
    path('clients/<int:pk>/delete/',        views.client_delete,              name='client_delete'),

    # ── Suppliers ──────────────────────────────────────────────────────────
    path('suppliers/',                      views.supplier_list,              name='supplier_list'),
    path('suppliers/add/',                  views.supplier_create,            name='supplier_create'),
    path('suppliers/<int:pk>/edit/',        views.supplier_edit,              name='supplier_edit'),
    path('suppliers/<int:pk>/delete/',      views.supplier_delete,            name='supplier_delete'),

    # ── Chart of Accounts ──────────────────────────────────────────────────
    path('accounts/',                       views.chart_of_accounts,          name='chart_of_accounts'),
    path('accounts/add/',                   views.account_add,                name='account_add'),
    path('accounts/<int:pk>/edit/',         views.account_edit,               name='account_edit'),
    path('accounts/<int:pk>/statement/',    views.account_statement,          name='account_statement'),
    # AJAX
    path('api/accounts/<int:pk>/balance/',  views.ajax_account_balance,       name='ajax_account_balance'),

    # ── Journal Entries ────────────────────────────────────────────────────
    path('journal/',                        views.journal_entries_list,       name='journal_entries_list'),
    path('journal/add/',                    views.journal_entry_add,          name='journal_entry_add'),
    path('journal/<int:pk>/',               views.journal_entry_detail,       name='journal_entry_detail'),
    path('journal/<int:pk>/post/',          views.journal_entry_post,         name='journal_entry_post'),
    path('journal/<int:pk>/reverse/',       views.journal_entry_reverse,      name='journal_entry_reverse'),
    path('journal/export/excel/',           views.journal_export_excel,       name='journal_export_excel'),
    path('journal/<int:pk>/delete/', views.journal_entry_delete, name='journal_entry_delete'),
    # ── Accounting Dashboard ───────────────────────────────────────────────
    path('accounting/',                     views.accounting_dashboard,       name='accounting_dashboard'),

    # ── Financial Reports ──────────────────────────────────────────────────
    path('reports/trial-balance/',          views.trial_balance,              name='trial_balance'),
     path('accounts/<int:pk>/statement/', views.account_statement, name='account_statement'),
    path('reports/income-statement/', views.income_statement_v2, name='income_statement'),
    path('reports/export/excel/', views._export_income_statement_excel, name='_export_income_statement_excel'),
    path('reports/balance-sheet/',          views.balance_sheet,              name='balance_sheet'),
    

    # ── Fiscal Years ───────────────────────────────────────────────────────
    path('fiscal-years/', views.fiscal_years_list, name='fiscal_years_list'),
    path('fiscal-years/add/', views.fiscal_year_add, name='fiscal_year_add'),
    path('fiscal-years/<int:pk>/', views.fiscal_year_detail, name='fiscal_year_detail'),
    path('fiscal-years/<int:pk>/edit/', views.fiscal_year_edit, name='fiscal_year_edit'),
    path('fiscal-years/<int:pk>/delete/', views.fiscal_year_delete, name='fiscal_year_delete'),
    path('fiscal-years/<int:pk>/close/', views.fiscal_year_close, name='fiscal_year_close_confirm'),
    
    # ── ZATCA per invoice ─────────────────────────────────────────
    path('sales/<int:pk>/test-qr/',    views.test_qr_generation,    name='test_qr'),
    path('sales/<int:pk>/view-xml/',   views.zatca_view_xml,         name='view_xml'),
    path('sales/<int:pk>/test-xml/',   views.zatca_compliance_test,  name='test_xml'),
    path('sales/<int:pk>/export-xml/', views.export_invoice_xml,     name='export_invoice_xml'),
    path('sales/<int:pk>/submit-zatca/', views.zatca_submit_invoice, name='zatca_submit_invoice'),
    
    # ══════════════════════════════════════════════
    # حساب الأستاذ العام
    # ══════════════════════════════════════════════
    path('accounting/general-ledger/',views.general_ledger,name='general_ledger'),
    
    # ══════════════════════════════════════════════
    # التقارير المحسّنة (مع تصدير Excel)
    # ══════════════════════════════════════════════
    
    
    path('reports/balance-sheet/',views.balance_sheet_v2,name='balance_sheet'),
    
    # ══════════════════════════════════════════════
    # كشف حساب العميل
    # ══════════════════════════════════════════════
    path('clients/<int:pk>/statement/',views.client_statement,name='client_statement'),
    path('debug/accounts/', views.debug_accounts, name='debug_accounts'),
    path('debug/income-statement/', views.debug_income_statement, name='debug_income_statement'),
]