#!/usr/bin/env python3
"""
TEST TIZIMI — Ishga tushirish skripti
=====================================
Ishlatish:
  python start.py

Keyin brauzerda oching:
  http://localhost:5050

Administrator panel:
  http://localhost:5050/admin
  Parol: admin123 (birinchi kirishda o'zgartiring!)
"""

import subprocess
import sys
import os

def check_requirements():
    required = ['flask', 'openpyxl', 'pandas']
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            print(f"[O'rnatilmoqda] {pkg}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg])

if __name__ == '__main__':
    print("=" * 50)
    print("📝 TEST TIZIMI")
    print("=" * 50)
    
    check_requirements()
    
    os.makedirs('tests', exist_ok=True)
    os.makedirs('results', exist_ok=True)
    
    print("\n✅ Dastur ishga tushirilmoqda...")
    print("🌐 Brauzerda oching: http://localhost:5050")
    print("⚙️  Admin panel:     http://localhost:5050/admin")
    print("🔑 Admin parol:     admin123")
    print("\n[To'xtatish uchun: Ctrl+C]\n")
    
    from app import app
    app.run(debug=False, host='0.0.0.0', port=5050)
