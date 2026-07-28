#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================================
    Générateur de Certificats SSL Auto-signés
==========================================================================
    Ce script génère un certificat SSL auto-signé pour le Dashboard RH
    À utiliser uniquement si vous n'avez pas de certificat officiel
==========================================================================
"""

import subprocess
import os
from pathlib import Path

# Configuration
APP_DIR = Path(__file__).parent
CERT_FILE = APP_DIR / 'cert.pem'
KEY_FILE = APP_DIR / 'key.pem'

# Informations du certificat
CERT_INFO = {
    'country': 'MA',
    'state': 'Casablanca-Settat',
    'city': 'Casablanca',
    'organization': 'SONASID',
    'organizational_unit': 'Ressources Humaines',
    'common_name': 'portailrh.sonasid.ma',
    'email': 'rh@sonasid.ma'
}

def generate_certificate():
    """Générer un certificat SSL auto-signé"""
    print("\n" + "=" * 70)
    print("🔐 Génération du Certificat SSL")
    print("=" * 70)
    
    # Vérifier si les fichiers existent déjà
    if CERT_FILE.exists() and KEY_FILE.exists():
        print(f"\n⚠️  Les certificats existent déjà:")
        print(f"   📜 {CERT_FILE}")
        print(f"   🔑 {KEY_FILE}")
        
        response = input("\nVoulez-vous les remplacer? (o/n): ").strip().lower()
        if response != 'o':
            print("❌ Opération annulée")
            return False
    
    # Construire le sujet du certificat
    subject = f"/C={CERT_INFO['country']}/ST={CERT_INFO['state']}/L={CERT_INFO['city']}/O={CERT_INFO['organization']}/OU={CERT_INFO['organizational_unit']}/CN={CERT_INFO['common_name']}/emailAddress={CERT_INFO['email']}"
    
    # Commande OpenSSL
    cmd = [
        'openssl', 'req', '-x509', '-newkey', 'rsa:4096',
        '-keyout', str(KEY_FILE),
        '-out', str(CERT_FILE),
        '-days', '365',
        '-nodes',
        '-subj', subject,
        '-addext', f'subjectAltName=DNS:{CERT_INFO["common_name"]},DNS:localhost,IP:10.222.36.9,IP:127.0.0.1'
    ]
    
    print(f"\n📝 Informations du certificat:")
    print(f"   Pays: {CERT_INFO['country']}")
    print(f"   État/Province: {CERT_INFO['state']}")
    print(f"   Ville: {CERT_INFO['city']}")
    print(f"   Organisation: {CERT_INFO['organization']}")
    print(f"   Unité: {CERT_INFO['organizational_unit']}")
    print(f"   Nom commun (CN): {CERT_INFO['common_name']}")
    print(f"   Validité: 365 jours")
    
    print("\n⏳ Génération en cours...")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("\n✅ Certificat généré avec succès!")
            print(f"\n📁 Fichiers créés:")
            print(f"   📜 Certificat: {CERT_FILE}")
            print(f"   🔑 Clé privée: {KEY_FILE}")
            
            # Définir les permissions
            os.chmod(KEY_FILE, 0o600)
            os.chmod(CERT_FILE, 0o644)
            print("\n🔒 Permissions définies (clé: 600, cert: 644)")
            
            print("\n" + "=" * 70)
            print("✅ Vous pouvez maintenant démarrer le serveur HTTPS:")
            print("   python3 app.py")
            print("=" * 70 + "\n")
            return True
        else:
            print(f"\n❌ Erreur lors de la génération:")
            print(result.stderr)
            return False
            
    except FileNotFoundError:
        print("\n❌ OpenSSL n'est pas installé!")
        print("\n📋 Installation:")
        print("   Ubuntu/Debian: sudo apt install openssl")
        print("   CentOS/RHEL: sudo yum install openssl")
        print("   Windows: Téléchargez depuis https://slproweb.com/products/Win32OpenSSL.html")
        return False
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        return False

def convert_pfx_to_pem(pfx_file, password=None):
    """Convertir un certificat PFX/P12 en PEM"""
    print("\n" + "=" * 70)
    print("🔄 Conversion PFX → PEM")
    print("=" * 70)
    
    pfx_path = Path(pfx_file)
    if not pfx_path.exists():
        print(f"\n❌ Fichier non trouvé: {pfx_file}")
        return False
    
    # Demander le mot de passe si non fourni
    if password is None:
        import getpass
        password = getpass.getpass("🔑 Mot de passe du fichier PFX: ")
    
    try:
        # Extraire le certificat
        cert_cmd = [
            'openssl', 'pkcs12', '-in', str(pfx_path),
            '-out', str(CERT_FILE), '-clcerts', '-nokeys',
            '-passin', f'pass:{password}'
        ]
        
        # Extraire la clé privée
        key_cmd = [
            'openssl', 'pkcs12', '-in', str(pfx_path),
            '-out', str(KEY_FILE), '-nocerts', '-nodes',
            '-passin', f'pass:{password}'
        ]
        
        print("\n⏳ Extraction du certificat...")
        result1 = subprocess.run(cert_cmd, capture_output=True, text=True)
        
        print("⏳ Extraction de la clé privée...")
        result2 = subprocess.run(key_cmd, capture_output=True, text=True)
        
        if result1.returncode == 0 and result2.returncode == 0:
            os.chmod(KEY_FILE, 0o600)
            os.chmod(CERT_FILE, 0o644)
            print("\n✅ Conversion réussie!")
            print(f"   📜 Certificat: {CERT_FILE}")
            print(f"   🔑 Clé privée: {KEY_FILE}")
            return True
        else:
            print("\n❌ Erreur lors de la conversion")
            if result1.stderr:
                print(result1.stderr)
            if result2.stderr:
                print(result2.stderr)
            return False
            
    except Exception as e:
        print(f"\n❌ Erreur: {e}")
        return False

def main():
    """Menu principal"""
    print("\n" + "=" * 70)
    print("🔐 Utilitaire de Gestion des Certificats SSL")
    print("=" * 70)
    print("\n1. Générer un certificat auto-signé")
    print("2. Convertir un fichier PFX/P12 en PEM")
    print("3. Vérifier les certificats existants")
    print("4. Quitter")
    
    choice = input("\nVotre choix (1-4): ").strip()
    
    if choice == '1':
        generate_certificate()
    elif choice == '2':
        pfx_file = input("Chemin du fichier PFX: ").strip()
        convert_pfx_to_pem(pfx_file)
    elif choice == '3':
        print(f"\n📁 Vérification des certificats:")
        print(f"   📜 cert.pem: {'✅ Existe' if CERT_FILE.exists() else '❌ Non trouvé'}")
        print(f"   🔑 key.pem: {'✅ Existe' if KEY_FILE.exists() else '❌ Non trouvé'}")
        
        if CERT_FILE.exists():
            print("\n📋 Informations du certificat:")
            subprocess.run(['openssl', 'x509', '-in', str(CERT_FILE), '-text', '-noout', '-subject', '-dates'])
    elif choice == '4':
        print("\n👋 Au revoir!")
    else:
        print("\n❌ Choix invalide")

if __name__ == '__main__':
    main()
