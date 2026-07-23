# Avtogaz Service — joylashtirish yo'riqnomasi

## Vercel'ga joylashtirish (10 daqiqa)

### 1-qadam: GitHub'ga yuklash

1. https://github.com da yangi repository yarating (nomi: `avtogaz-service`)
2. Bu papkadagi barcha fayllarni o'sha repositoryga yuklang (GitHub saytida "uploading an existing file" orqali, yoki `git` orqali)

### 2-qadam: Vercel'ga ulash

1. https://vercel.com ga kiring, GitHub orqali ro'yxatdan o'ting
2. "Add New Project" tugmasini bosing
3. GitHub repositoryingizni tanlang (`avtogaz-service`)
4. **Environment Variables** bo'limida quyidagilarni qo'shing:
   - `NEXT_PUBLIC_SUPABASE_URL` = `https://qjnwgnwlaqkjuzlwmkbd.supabase.co`
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY` = `sb_publishable_PyfEUPxkNG-5jgX7ho7hWQ_eF8pj5D8`
5. "Deploy" tugmasini bosing

2-3 daqiqadan keyin sizga `https://avtogaz-service-xxxx.vercel.app` kabi havola beriladi — shu orqali istalgan qurilmadan kirish mumkin.

## Muhim eslatmalar

- `.env.local` fayli — bu sizning maxfiy kalitlaringiz, uni GitHub'ga yuklamang (`.gitignore` avtomatik oldini oladi)
- Vercel'da esa Environment Variables orqali xavfsiz saqlanadi
- PIN kodlar: Admin `1111`, Kassir `2211`, Usta `3311` — birinchi kirgach o'zgartiring
