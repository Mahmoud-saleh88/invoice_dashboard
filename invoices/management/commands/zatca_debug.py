from django.core.management.base import BaseCommand
from invoices.models import Company, Client, SalesInvoice

class Command(BaseCommand):
    help = 'Debug ZATCA address fields and simulate XML generation'

    def handle(self, *args, **options):
        self.stdout.write('\n=== COMPANY DATA ===')
        for co in Company.objects.all():
            fields = ['street_name','building_number','district','city','postal_code','vat_number']
            for f in fields:
                val = getattr(co, f, '')
                status = 'OK' if (val or '').strip() else 'EMPTY!'
                self.stdout.write(f'  {f}: [{val}] → {status}')

        self.stdout.write('\n=== CLIENT DATA (Saudi clients) ===')
        for cl in Client.objects.filter(is_active=True):
            country = (cl.country or '').strip().lower()
            is_saudi = 'saudi' in country or country == 'sa'
            if not is_saudi:
                continue
            self.stdout.write(f'\n  Client: {cl.name}')
            for f in ['street_name','building_number','district','city','postal_code']:
                val = getattr(cl, f, '')
                status = 'OK' if (val or '').strip() else 'EMPTY!'
                self.stdout.write(f'    {f}: [{val}] → {status}')

        self.stdout.write('\n=== RECENT INVOICES supply_date ===')
        for inv in SalesInvoice.objects.order_by('-date')[:10]:
            sd = inv.supply_date
            status = 'OK' if sd else 'NULL!'
            self.stdout.write(f'  {inv.invoice_number} | supply_date: {sd} → {status}')