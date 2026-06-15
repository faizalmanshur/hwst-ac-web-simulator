CATATAN DEPLOY VERCEL - SIMULASI AC + HWST

File tambahan yang sudah dimasukkan:
1. backend/app.py
   - Entry sederhana: from main import app

2. backend/api/index.py
   - Entry Vercel Python Function.
   - Mengarahkan request Vercel ke FastAPI app dari backend/main.py.

3. backend/vercel.json
   - Routing semua request backend ke api/index.py.

4. frontend/.env.local
   - Untuk testing lokal:
     VITE_API_BASE_URL=http://127.0.0.1:8000

5. frontend/.env.example
   - Contoh environment variable.

6. backend/main.py
   - CORS sudah diubah agar mendukung localhost, FRONTEND_ORIGIN, dan domain *.vercel.app.

CARA DEPLOY RINGKAS:
A. Deploy backend di Vercel
   Root Directory: backend
   Framework: Other
   Environment Variable opsional:
   FRONTEND_ORIGIN=https://link-frontend-kamu.vercel.app

B. Deploy frontend di Vercel
   Root Directory: frontend
   Framework: Vite
   Build Command: npm run build
   Output Directory: dist
   Environment Variable wajib:
   VITE_API_BASE_URL=https://link-backend-kamu.vercel.app

C. Setelah deploy frontend, update FRONTEND_ORIGIN di backend Vercel
   FRONTEND_ORIGIN=https://link-frontend-kamu.vercel.app

D. Redeploy backend dan frontend.
