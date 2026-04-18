# In Django shell ----> In production
from invoices.models import Company
from invoices.zatca_service import onboard_device

company = Company.objects.first()
device  = onboard_device(
    company=company,
    device_name='EGS-001',
    otp='XXXXXX',              
    environment='production',
    csr_pem=open('zatca_sdk/zatca-einvoicing-sdk-Java-238-R3.4.8/Apps/new_csr2.pem').read(),
    private_key_pem=open('zatca_sdk/zatca-einvoicing-sdk-Java-238-R3.4.8/Apps/new_key2.pem').read(),
)
print("Device registered:", device.id, device.device_name)