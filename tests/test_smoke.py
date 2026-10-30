import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server

class ProjectSmokeTests(unittest.TestCase):
    def test_auth_disabled_by_default(self):
        for key in ('DUARTE_ENTRA_TENANT_ID','DUARTE_ENTRA_CLIENT_ID','DUARTE_ENTRA_CLIENT_SECRET','DUARTE_SESSION_SECRET','DUARTE_AUTH_REQUIRED','RAILWAY_ENVIRONMENT','DUARTE_ALLOWED_USERS'):
            os.environ.pop(key, None)
        cfg = server._auth_settings()
        self.assertFalse(cfg['enabled'])
        self.assertFalse(cfg['configured'])

    def test_email_normalization_and_deduplication(self):
        rows, duplicates = server.email_normalize([
            {'date':'2026-10-01','from':'a@example.com','subject':'Status','message id':'m-1'},
            {'date':'2026-10-01','from':'a@example.com','subject':'Status','message id':'m-1'},
        ], 'teste.csv')
        self.assertEqual(len(rows), 1)
        self.assertEqual(duplicates, 1)

    def test_exports_are_valid_containers(self):
        payload={'emails':[{'data':'2026-10-01','remetente':'a@exemplo.com','destinatarios':'b@exemplo.com','assunto':'Teste','direcao':'Recebido','cliente':'','conversa':'t1','id':'m1','fonte':'teste.eml'}], 'meta':{'period':'2026-10-01','source':'teste.eml'}}
        xlsx=server.make_xlsx(payload)
        pdf=server.make_pdf(payload)
        self.assertTrue(xlsx.startswith(b'PK'))
        self.assertTrue(pdf.startswith(b'%PDF-'))

    def test_landing_page_routes_to_explicit_login(self):
        source=(ROOT/'server.py').read_text()
        self.assertIn("if path in ('/','/index.html')", source)
        self.assertIn('Entrar com Microsoft', source)
        self.assertIn("AUTH_SCOPES=['User.Read']", source)
        self.assertIn('DUARTE_ALLOWED_USERS', source)

if __name__ == '__main__':
    unittest.main()
