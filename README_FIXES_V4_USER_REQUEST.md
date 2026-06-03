# Tg-bot-main fixed v3

Ushbu versiyada foydalanuvchi aytgan 3 ta asosiy muammo tuzatildi:

1. Shikoyat yuborilgandan keyin bot qotib qolishi
- Complaint FSM `state.clear()` bilan to'liq yopiladi.
- Reply keyboard tozalanadi va foydalanuvchiga yangi asosiy inline menu qayta yuboriladi.
- Adminlarga yangi shikoyat bo'yicha xabar yuborish qo'shildi, lekin xabar yuborishda xato bo'lsa bot to'xtamaydi.

2. Admin panel tugmalari ishlamasligi
- `admin_required` dekoratori `functools.wraps` bilan tuzatildi. Aiogram endi `state` kabi dependency argumentlarini to'g'ri uzatadi.
- Eskirgan yoki noma'lum admin callbacklari uchun fallback qo'shildi.
- `edit_or_send` yanada xavfsiz qilindi: eski xabarni edit qilishda xato bo'lsa yangi xabar yuboradi.

3. Ovoz berish/baholash funksiyasi yo'qdek ko'rinishi
- `🗳 Ovoz berish`, `Ovoz berish`, `Baholash` kabi eski tugma nomlari ham rating flowga ulandi.
- Rating boshlanishida eski FSM holatlari tozalanadi.
- Obuna va voting_open tekshiruvi saqlandi.

Deploy:
- Railway/VPSga ZIP ichidagi fayllarni yuklang.
- BOT_TOKEN, ADMIN_IDS, CHANNEL_USERNAME env variables tekshiring.
- Eski votes.db buzilmaydi; migratsiya avtomatik ishlaydi.
