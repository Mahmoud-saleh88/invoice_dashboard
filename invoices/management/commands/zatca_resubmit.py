from django.core.management.base import BaseCommand
from invoices.models import SalesInvoice, ZATCADevice
from invoices.zatca_service import ZATCAInvoiceService, ZATCAValidationError
import json


class Command(BaseCommand):
    help = 'Resubmit a specific invoice to ZATCA and show full response'

    def add_arguments(self, parser):
        parser.add_argument('invoice_number', type=str)
        parser.add_argument(
            '--dry-run', action='store_true',
            help='Generate and validate XML only, do not submit'
        )

    def handle(self, *args, **options):
        inv_num = options['invoice_number']
        dry_run = options['dry_run']

        try:
            inv = SalesInvoice.objects.get(invoice_number=inv_num)
        except SalesInvoice.DoesNotExist:
            self.stdout.write(f'Invoice not found: {inv_num}')
            return

        self.stdout.write(f'\nInvoice   : {inv.invoice_number}')
        self.stdout.write(f'Client    : {inv.client.name}')
        self.stdout.write(f'Date      : {inv.date}')
        self.stdout.write(f'Total     : {inv.total} SAR')
        self.stdout.write(f'ZATCA now : {inv.zatca_status}')

        device = ZATCADevice.objects.filter(
            company=inv.company, is_active=True
        ).first()

        if not device:
            self.stdout.write('ERROR: No active ZATCA device for this company.')
            return

        self.stdout.write(f'Device    : {device.device_name} ({device.environment})')

        if dry_run:
            self.stdout.write('\n--- DRY RUN: generating XML only ---')
            from invoices.zatca_service import ZATCAXMLGenerator, validate_invoice
            try:
                validate_invoice(inv)
                xml = ZATCAXMLGenerator().generate(inv)
                self.stdout.write('Validation : PASS')
                self.stdout.write(f'XML length : {len(xml)} chars')
                self.stdout.write('\n--- XML preview (first 800 chars) ---')
                self.stdout.write(xml[:800])
            except ZATCAValidationError as e:
                self.stdout.write(f'Validation FAIL: {e}')
            return

        # Reset ZATCA fields so process() treats it as fresh
        inv.zatca_status         = 'pending'
        inv.zatca_invoice_hash   = ''
        inv.zatca_signed_xml     = ''
        inv.zatca_submission_id  = ''
        inv.invoice_counter_value = 0
        inv.previous_invoice_hash = ''
        inv.save(update_fields=[
            'zatca_status', 'zatca_invoice_hash', 'zatca_signed_xml',
            'zatca_submission_id', 'invoice_counter_value', 'previous_invoice_hash'
        ])

        self.stdout.write('\nSubmitting...')
        try:
            svc    = ZATCAInvoiceService(device)
            result = svc.process(inv)

            self.stdout.write(f'\nSuccess       : {result["success"]}')
            self.stdout.write(f'Invoice hash  : {result["invoice_hash"]}')
            self.stdout.write(f'ZATCA status  : {inv.zatca_status}')

            resp_data = result.get('result', {}).get('data', {})
            status_code = result.get('result', {}).get('status_code')
            self.stdout.write(f'HTTP status   : {status_code}')

            vr = resp_data.get('validationResults', {})
            warnings = vr.get('warningMessages', [])
            errors   = vr.get('errorMessages',   [])

            if errors:
                self.stdout.write(f'\nErrors ({len(errors)}):')
                for e in errors:
                    self.stdout.write(f'  [{e.get("code")}] {e.get("message")}')
            else:
                self.stdout.write('\nErrors: none')

            if warnings:
                self.stdout.write(f'\nWarnings ({len(warnings)}):')
                for w in warnings:
                    self.stdout.write(f'  [{w.get("code")}] {w.get("message")}')
            else:
                self.stdout.write('\nWarnings: none — all clean!')

            self.stdout.write(f'\nDisposition: {resp_data.get("reportingStatus") or resp_data.get("clearanceStatus") or resp_data.get("dispositionMessage", "")}')

        except ZATCAValidationError as e:
            self.stdout.write(f'\nValidation error: {e}')
        except Exception as e:
            self.stdout.write(f'\nError: {e}')