"""
Script para generar token.json localmente para autenticación con Google Drive.
Ejecuta este script en tu máquina local (no en el servidor) para generar el token.
"""

import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/drive.file']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'

def generar_token():
    """Genera el token.json usando el flujo OAuth interactivo."""
    creds = None
    
    # Si ya existe un token, intentamos usarlo
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        if creds and creds.valid:
            print("Token ya existe y es válido. No es necesario generar uno nuevo.")
            return True
        
        # Si el token está expirado, intentamos renovarlo
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                with open(TOKEN_FILE, 'w') as token:
                    token.write(creds.to_json())
                print("Token renovado exitosamente.")
                return True
            except Exception as e:
                print(f"Error al renovar token: {e}")
                print("Generando nuevo token...")
    
    # Si no hay token válido, generamos uno nuevo
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"Error: No se encontró el archivo {CREDENTIALS_FILE}")
        print("Asegúrate de tener el archivo credentials.json en el mismo directorio.")
        return False
    
    try:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        
        # Guardar el token
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
        
        print(f"Token generado exitosamente y guardado en {TOKEN_FILE}")
        print("Ahora puedes subir este archivo al servidor.")
        return True
        
    except Exception as e:
        print(f"Error al generar token: {e}")
        return False

if __name__ == '__main__':
    print("=== Generador de Token para Google Drive ===")
    print("Este script abrirá una ventana del navegador para autenticarte con Google.")
    print("Asegúrate de tener el archivo credentials.json en el mismo directorio.")
    print()
    
    if generar_token():
        print("\n¡Éxito! Ahora sube el archivo token.json al servidor.")
    else:
        print("\nError: No se pudo generar el token.")
