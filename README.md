# Ombor Hisobi — yangi firma uchun joylashtirish yo'riqnomasi

Bu — to'liq ERP tizimi (Ombor, Ishlab chiqarish, Faktura, Kassa, Hisob-kitob, Ta'minotchilar, Analiz).
Har bir yangi firma uchun bu paketni ishlatib, **alohida, mustaqil nusxa** yarating (o'z bazasi, o'z sayti).

## 1-QADAM: Supabase'da ma'lumotlar bazasi

1. https://supabase.com → yangi loyiha yarating (bepul tarif yetarli)
2. **SQL Editor** bo'limini oching
3. `supabase-full-schema.sql` faylining **butun matnini** joylashtirib, **Run** bosing (bu bitta faylda barcha 17 ta jadvalni yaratadi)
4. **Project Settings → API** bo'limidan **Project URL** va **anon public key**ni nusxalab oling

## 2-QADAM: GitHub

1. Yangi (yoki fork qilingan) repository yarating
2. Shu papkadagi barcha fayllarni yuklang (`.env` dan tashqari)

## 3-QADAM: Vercel

1. https://vercel.com → repositoryni import qiling
2. **Environment Variables**ga qo'shing:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
3. **Deploy** bosing — 1-2 daqiqada tayyor

## Birinchi marta kirish

Sayt ochilgach, yuqoridagi logotipga bosib **"Korxona ma'lumotlari"** oynasida:
- Firma nomi va telefonini kiriting
- **Ilova paroli** (ixtiyoriy) — agar kiritsangiz, saytga kirishda shu parol so'raladi (oddiy himoya, kuchli xavfsizlik emas — havolani bilgan va parolni bilgan har kim kira oladi)
- **Telegram bot token va chat ID** (ixtiyoriy) — kam qolgan xom ashyo haqida avtomatik xabar yuborish uchun

## Har doim eslab qoling

- Bu — **bitta firma = bitta Supabase loyihasi = bitta Vercel loyihasi**. Yangi mijoz kelsa, shu 3 qadamni boshidan takrorlang (~15-20 daqiqa)
- Ilova parolsiz (yoki oddiy parol bilan) ishlaydi — kuchli maxfiylik kerak bo'lsa, Vercel'ning "Deployment Protection" (pullik) yoki to'liq login tizimi (Supabase Auth) alohida qo'shilishi kerak
