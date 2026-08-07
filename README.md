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
   - `NEXT_PUBLIC_SUPABASE_URL` = `https://aowobgqxpqzhpoqygyrt.supabase.co`
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY` = `sb_publishable_h0Ldy7XLssQUHUYG2McIWw_Y29ImVF6`
5. "Deploy" tugmasini bosing

2-3 daqiqadan keyin sizga `https://avtogaz-service-xxxx.vercel.app` kabi havola beriladi — shu orqali istalgan qurilmadan kirish mumkin.

### 3-qadam: Google orqali kirishni sozlash

Ilova endi PIN kod o'rniga **Google orqali kirish**dan foydalanadi. Buni ishga tushirish uchun quyidagi qadamlarni qo'lda bajarish kerak (kod bilan qilib bo'lmaydi):

1. **Google Cloud Console** (https://console.cloud.google.com) da yangi OAuth 2.0 Client ID yarating (turi: Web application). Authorized redirect URI sifatida shuni kiriting:
   `https://aowobgqxpqzhpoqygyrt.supabase.co/auth/v1/callback`
2. **Supabase Dashboard → Authentication → Providers → Google** bo'limida olingan Client ID va Client Secret'ni kiritib, provayderni yoqing.
3. **Supabase Dashboard → Authentication → URL Configuration** bo'limida Site URL'ni Vercel'dagi haqiqiy domeningizga o'rnating (masalan `https://avtogaz-service-xxxx.vercel.app`). Lokal test uchun `http://localhost:3000`ni Additional Redirect URLs'ga qo'shing.
4. **Xodimlarni qo'shish**: faqat `staff` jadvalida ro'yxatdan o'tgan Gmail manzillari kira oladi. Supabase Dashboard → Table Editor → `staff` jadvaliga xodimning email manzilini va rolini (`admin`, `kassir` yoki `usta`) qo'lda qo'shing.

## Muhim eslatmalar

- `.env.local` fayli — bu sizning maxfiy kalitlaringiz, uni GitHub'ga yuklamang (`.gitignore` avtomatik oldini oladi)
- Vercel'da esa Environment Variables orqali xavfsiz saqlanadi
- Kirish endi Google hisob orqali amalga oshiriladi — rol (`admin`/`kassir`/`usta`) `staff` jadvalidagi email yozuviga qarab avtomatik aniqlanadi
