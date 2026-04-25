from twilio.rest import Client
from django.conf import settings
import re

def enviar_whatsapp(telefone, mensagem):
    """Envia mensagem via WhatsApp usando Twilio"""
    if not telefone:
        return None
    
    # Limpar telefone
    telefone_limpo = re.sub(r'[^0-9]', '', str(telefone))
    if not telefone_limpo.startswith('55'):
        telefone_limpo = '55' + telefone_limpo
    
    client = Client(
        settings.TWILIO_API_KEY_SID,
        settings.TWILIO_API_KEY_SECRET,
        settings.TWILIO_ACCOUNT_SID
    )
    
    message = client.messages.create(
        body=mensagem,
        from_=settings.TWILIO_WHATSAPP_NUMBER,
        to=f'whatsapp:{telefone_limpo}'
    )
    return message.sid