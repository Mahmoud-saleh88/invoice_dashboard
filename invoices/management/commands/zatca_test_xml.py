from django.core.management.base import BaseCommand
from invoices.models import SalesInvoice
from invoices.zatca_service import ZATCAXMLGenerator, validate_invoice, ZATCAValidationError
from lxml import etree

REQUIRED_FIELDS = {
    'AccountingSupplierParty': ['StreetName','BuildingNumber','CitySubdivisionName','CityName','PostalZone'],
    'AccountingCustomerParty': ['StreetName','BuildingNumber','CitySubdivisionName','CityName','PostalZone'],
}

NS = {
    'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
    'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
}

PLACEHOLDER_VALUES = {'شارع غير محدد', '0000', '00000', 'غير محدد'}

class Command(BaseCommand):
    help = 'Generate XML for latest invoice and check for warning-causing values'

    def add_arguments(self, parser):
        parser.add_argument('--invoice', type=str, help='Invoice number (default: latest)')

    def handle(self, *args, **options):
        inv_num = options.get('invoice')
        if inv_num:
            inv = SalesInvoice.objects.get(invoice_number=inv_num)
        else:
            inv = SalesInvoice.objects.order_by('-date').first()

        if not inv:
            self.stdout.write('No invoices found.')
            return

        self.stdout.write(f'\nTesting invoice: {inv.invoice_number}')

        # Step 1: validate_invoice check
        self.stdout.write('\n--- validate_invoice() ---')
        try:
            validate_invoice(inv)
            self.stdout.write('  PASS: no validation errors')
        except ZATCAValidationError as e:
            self.stdout.write(f'  FAIL: {e}')
            return

        # Step 2: generate XML and check for placeholders
        self.stdout.write('\n--- XML field check ---')
        try:
            gen = ZATCAXMLGenerator()
            xml = gen.generate(inv)
        except ZATCAValidationError as e:
            self.stdout.write(f'  FAIL generating XML: {e}')
            return

        tree = etree.fromstring(xml.encode())

        all_ok = True
        for party_tag, fields in REQUIRED_FIELDS.items():
            party_el = tree.find(f'.//cac:{party_tag}', NS)
            if party_el is None:
                self.stdout.write(f'  MISSING element: {party_tag}')
                continue
            addr = party_el.find('.//cac:PostalAddress', NS)
            if addr is None:
                self.stdout.write(f'  MISSING PostalAddress in {party_tag}')
                continue
            for field in fields:
                el = addr.find(f'cbc:{field}', NS)
                val = (el.text or '').strip() if el is not None else ''
                if not val:
                    self.stdout.write(f'  EMPTY  [{party_tag}] {field}')
                    all_ok = False
                elif val in PLACEHOLDER_VALUES:
                    self.stdout.write(f'  PLACEHOLDER [{party_tag}] {field} = "{val}"')
                    all_ok = False
                else:
                    self.stdout.write(f'  OK     [{party_tag}] {field} = "{val}"')

        # Step 3: supply_date check
        self.stdout.write('\n--- supply_date check ---')
        if inv.supply_date:
            self.stdout.write(f'  OK: supply_date = {inv.supply_date}')
        else:
            self.stdout.write('  WARN: supply_date is null — will use invoice.date as fallback')

        if all_ok:
            self.stdout.write('\nRESULT: all fields OK — warnings should be gone')
        else:
            self.stdout.write('\nRESULT: fix the fields above before submitting')